from __future__ import annotations

import functools
import json
import multiprocessing
import re
import shutil
import subprocess
import sys
import tempfile
import time

from concurrent.futures import (
    ProcessPoolExecutor,
    as_completed,
)
from pathlib import Path

import fitz
import pandas as pd
from tqdm import tqdm

from config import (
    BASE_DIR,
    DOWNLOADS_DIR,
    OUTPUT_DIR,
    PYTHON_YOLO,
    PYTHON_OCR,
    DETECTOR_PY,
    OCR_PY,
    BATCH_SIZE,
    QTD_IMAGENS_POR_PASSAGEM,
    MAX_DIFERENCAS_APROVACAO,
    ensure_parent,
)


# ============================================================
# CONFIG
# ============================================================

REGEX_PLACA = re.compile(
    r"([A-Z]{3}\d[A-Z\d]\d{2}|[A-Z]{3}\d{4})"
)

REGEX_DATA = re.compile(
    r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})"
)

TAMANHO_PLACA = 7


# ============================================================
# MAPAS DE CORREÇÃO POR POSIÇÃO
# ============================================================

_MAPA_LETRAS = {
    "0": "O",
    "1": "I",
    "2": "Z",
    "4": "A",
    "5": "S",
    "6": "G",
    "7": "Z",
    "8": "B",
}

_MAPA_DIGITOS = {
    "O": "0",
    "Q": "0",
    "D": "0",
    "I": "1",
    "L": "1",
    "T": "7",
    "J": "1",
    "Z": "2",
    "A": "4",
    "S": "5",
    "G": "6",
    "B": "8",
    "E": "8",
    "U": "0",
}


# ============================================================
# IMAGEMAGICK
# ============================================================

def _encontrar_imagemagick():
    caminho = shutil.which("magick")

    if caminho:
        return caminho

    bases = [
        Path(r"C:\Program Files"),
        Path(r"C:\Program Files (x86)"),
    ]

    for base in bases:
        if not base.exists():
            continue

        try:
            for pasta in base.iterdir():
                if not pasta.is_dir():
                    continue

                if not pasta.name.lower().startswith(
                    "imagemagick"
                ):
                    continue

                candidato = pasta / "magick.exe"

                if candidato.is_file():
                    return str(candidato)

        except Exception:
            continue

    return None


# ============================================================
# LEVENSHTEIN
# ============================================================

_CONFUNDIVEIS = {
    ("0", "O"),
    ("O", "0"),
    ("1", "I"),
    ("I", "1"),
    ("1", "L"),
    ("L", "1"),
    ("T", "I"),
    ("I", "T"),
    ("2", "Z"),
    ("Z", "2"),
    ("Z", "7"),
    ("7", "Z"),
    ("4", "A"),
    ("A", "4"),
    ("5", "S"),
    ("S", "5"),
    ("6", "G"),
    ("G", "6"),
    ("8", "B"),
    ("B", "8"),
    ("B", "E"),
    ("E", "B"),
    ("D", "0"),
    ("0", "D"),
    ("D", "O"),
    ("O", "D"),
}


@functools.lru_cache(maxsize=8192)
def _levenshtein(
    a: str,
    b: str,
) -> float:
    anterior = list(
        range(len(b) + 1)
    )

    for i, ca in enumerate(
        a,
        start=1,
    ):
        atual = [i]

        for j, cb in enumerate(
            b,
            start=1,
        ):
            if ca == cb:
                custo = 0.0

            elif (
                ca,
                cb,
            ) in _CONFUNDIVEIS:
                custo = 0.35

            else:
                custo = 1.0

            atual.append(
                min(
                    anterior[j] + 1,
                    atual[j - 1] + 1,
                    anterior[j - 1] + custo,
                )
            )

        anterior = atual

    return anterior[-1]


def distancia_placas(
    placa_a: str,
    placa_b: str,
) -> float:
    a = re.sub(
        r"[^A-Z0-9]",
        "",
        str(placa_a or "").upper(),
    )

    b = re.sub(
        r"[^A-Z0-9]",
        "",
        str(placa_b or "").upper(),
    )

    if not a or not b:
        return 999.0

    return _levenshtein(
        a,
        b,
    )


# ============================================================
# DESCOBRE MÁSCARA DA PLACA
# ============================================================

def descobrir_mascara(
    placa: str,
) -> str:
    placa = re.sub(
        r"[^A-Z0-9]",
        "",
        str(placa or "").upper(),
    )

    if len(placa) != 7:
        return "LLLDLDD"

    # Mercosul
    # LLLDLDD
    if placa[4].isalpha():
        return "LLLDLDD"

    # Antiga
    # LLLDDDD
    return "LLLDDDD"


# ============================================================
# APLICA MÁSCARA
# ============================================================

def aplicar_mascara_placa(
    candidato: str,
    mascara: str,
) -> str:
    candidato = re.sub(
        r"[^A-Z0-9]",
        "",
        str(candidato or "").upper(),
    )

    if len(candidato) != 7:
        return candidato

    resultado = []

    for caractere, tipo in zip(
        candidato,
        mascara,
    ):
        if tipo == "L":
            resultado.append(
                _MAPA_LETRAS.get(
                    caractere,
                    caractere,
                )
            )

        else:
            resultado.append(
                _MAPA_DIGITOS.get(
                    caractere,
                    caractere,
                )
            )

    return "".join(resultado)


# ============================================================
# NORMALIZA OCR
# ============================================================

def normalizar_ocr(
    texto: str,
) -> str:
    return re.sub(
        r"[^A-Z0-9]",
        "",
        str(texto or "").upper(),
    )


# ============================================================
# ENCONTRA A MELHOR JANELA DE 7 CARACTERES
# ============================================================

def encontrar_melhor_trecho(
    placa_esperada: str,
    texto_ocr: str,
    mascara: str | None = None,
):
    placa_esperada = normalizar_ocr(
        placa_esperada
    )

    texto_ocr = normalizar_ocr(
        texto_ocr
    )

    if not placa_esperada or not texto_ocr:
        return {
            "ocr_bruto_trecho": "",
            "ocr_limpo": "",
            "distancia": 999.0,
            "posicao": -1,
        }

    if mascara is None:
        mascara = descobrir_mascara(
            placa_esperada
        )

    # ========================================================
    # OCR MENOR QUE 7
    # ========================================================

    if len(texto_ocr) < TAMANHO_PLACA:
        candidato_corrigido = aplicar_mascara_placa(
            texto_ocr,
            mascara,
        )

        return {
            "ocr_bruto_trecho": texto_ocr,
            "ocr_limpo": candidato_corrigido,
            "distancia": distancia_placas(
                placa_esperada,
                candidato_corrigido,
            ),
            "posicao": 0,
        }

    # ========================================================
    # TESTA TODAS AS JANELAS DE 7
    # ========================================================

    candidatos = []

    for posicao in range(
        len(texto_ocr) - TAMANHO_PLACA + 1
    ):
        trecho = texto_ocr[
            posicao:
            posicao + TAMANHO_PLACA
        ]

        corrigido = aplicar_mascara_placa(
            trecho,
            mascara,
        )

        distancia_raw = distancia_placas(
            placa_esperada,
            trecho,
        )

        distancia_corrigida = distancia_placas(
            placa_esperada,
            corrigido,
        )

        # Prioridade:
        # 0 = trecho bruto exato
        # 1 = trecho corrigido exato
        # 2 = menor distância

        if trecho == placa_esperada:
            prioridade = 0

        elif corrigido == placa_esperada:
            prioridade = 1

        else:
            prioridade = 2

        candidatos.append(
            (
                prioridade,
                min(
                    distancia_raw,
                    distancia_corrigida,
                ),
                posicao,
                trecho,
                corrigido,
            )
        )

    candidatos.sort(
        key=lambda item: (
            item[0],
            item[1],
            item[2],
        )
    )

    (
        _prioridade,
        distancia,
        posicao,
        trecho,
        corrigido,
    ) = candidatos[0]

    return {
        "ocr_bruto_trecho": trecho,
        "ocr_limpo": corrigido,
        "distancia": distancia,
        "posicao": posicao,
    }


# ============================================================
# EXTRAI PLACA DO OCR
# ============================================================

def extrair_placa_ocr(
    texto_ocr: str,
    placa_esperada: str,
):
    resultado = encontrar_melhor_trecho(
        placa_esperada,
        texto_ocr,
    )

    return (
        resultado["ocr_limpo"],
        resultado["ocr_bruto_trecho"],
        resultado["distancia"],
    )


# ============================================================
# DESCREVE DIFERENÇAS DE CARACTERES
# ============================================================

def descrever_diferencas(
    placa_esperada,
    placa_ocr,
) -> str:
    esperado = re.sub(
        r"[^A-Z0-9]",
        "",
        str(placa_esperada or "").upper(),
    )

    recebido = re.sub(
        r"[^A-Z0-9]",
        "",
        str(placa_ocr or "").upper(),
    )

    if not esperado or not recebido:
        return "Sem comparação"

    linhas = len(esperado) + 1
    colunas = len(recebido) + 1

    matriz = [
        [0.0 for _ in range(colunas)]
        for _ in range(linhas)
    ]

    operacoes = [
        [None for _ in range(colunas)]
        for _ in range(linhas)
    ]

    for i in range(linhas):
        matriz[i][0] = float(i)

    for j in range(colunas):
        matriz[0][j] = float(j)

    for i in range(1, linhas):
        for j in range(1, colunas):
            ca = esperado[i - 1]
            cb = recebido[j - 1]

            if ca == cb:
                custo = 0.0

            elif (
                ca,
                cb,
            ) in _CONFUNDIVEIS:
                custo = 0.35

            else:
                custo = 1.0

            substituicao = (
                matriz[i - 1][j - 1]
                + custo
            )

            insercao = (
                matriz[i][j - 1]
                + 1.0
            )

            remocao = (
                matriz[i - 1][j]
                + 1.0
            )

            menor = min(
                substituicao,
                insercao,
                remocao,
            )

            matriz[i][j] = menor

            if menor == substituicao:
                if custo == 0:
                    operacoes[i][j] = (
                        "igual",
                        ca,
                        cb,
                    )
                else:
                    operacoes[i][j] = (
                        "substituicao",
                        ca,
                        cb,
                    )

            elif menor == insercao:
                operacoes[i][j] = (
                    "insercao",
                    "",
                    cb,
                )

            else:
                operacoes[i][j] = (
                    "remocao",
                    ca,
                    "",
                )

    # ========================================================
    # BACKTRACE
    # ========================================================

    i = len(esperado)
    j = len(recebido)
    diferencas = []

    while i > 0 or j > 0:
        if i > 0 and j > 0:
            tipo = operacoes[i][j][0]

        elif i > 0:
            tipo = "remocao"

        else:
            tipo = "insercao"

        if tipo == "igual":
            i -= 1
            j -= 1

        elif tipo == "substituicao":
            diferencas.append(
                (
                    i,
                    recebido[j - 1],
                    esperado[i - 1],
                )
            )

            i -= 1
            j -= 1

        elif tipo == "insercao":
            diferencas.append(
                (
                    j,
                    recebido[j - 1],
                    "—",
                )
            )

            j -= 1

        else:
            diferencas.append(
                (
                    i,
                    "—",
                    esperado[i - 1],
                )
            )

            i -= 1

    diferencas.reverse()

    if not diferencas:
        return "Nenhuma"

    resultado = []

    for posicao, ocr, sistema in diferencas:
        if sistema == "—":
            resultado.append(
                f"Posição {posicao}: "
                f"OCR={ocr} "
                f"(caractere extra)"
            )

        elif ocr == "—":
            resultado.append(
                f"Posição {posicao}: "
                f"OCR ausente "
                f"(esperado={sistema})"
            )

        else:
            resultado.append(
                f"Posição {posicao}: "
                f"OCR={ocr} "
                f"≠ Sistema={sistema}"
            )

    return "; ".join(resultado)


# ============================================================
# NORMALIZA DATA/HORA
# ============================================================

def normalizar_data_hora(
    texto: str,
) -> str:
    texto = " ".join(
        texto.split()
    )

    formatos = (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %H",
    )

    for formato in formatos:
        try:
            dt = pd.to_datetime(
                texto,
                format=formato,
            )

            return dt.strftime(
                "%d/%m/%Y %H:%M:%S"
            )

        except Exception:
            continue

    return texto


# ============================================================
# COMUNICAÇÃO
# ============================================================

def enviar_json(
    processo,
    dados,
):
    if processo.poll() is not None:
        raise RuntimeError(
            "Processo já encerrou. "
            f"Código: {processo.returncode}"
        )

    processo.stdin.write(
        json.dumps(
            dados,
            ensure_ascii=False,
        )
        + "\n"
    )

    processo.stdin.flush()


def ler_json(
    processo,
    nome,
):
    linha = processo.stdout.readline()

    if not linha:
        raise RuntimeError(
            f"Processo {nome} "
            "encerrou antes de responder."
        )

    try:
        return json.loads(
            linha
        )

    except json.JSONDecodeError as e:
        raise RuntimeError(
            f"Resposta inválida do processo "
            f"{nome}: {linha[:500]}"
        ) from e


def encerrar_processo(
    processo,
    nome,
):
    if processo is None:
        return

    if processo.poll() is not None:
        return

    try:
        enviar_json(
            processo,
            {
                "command": "exit"
            },
        )

        processo.wait(
            timeout=10
        )

    except Exception:
        try:
            processo.terminate()

            processo.wait(
                timeout=5
            )

        except Exception:
            try:
                processo.kill()

            except Exception:
                pass


# ============================================================
# INICIA YOLO
# ============================================================

def iniciar_yolo():
    print(
        "\n[INIT] Iniciando YOLO..."
    )

    print(
        f"[YOLO] Python: {PYTHON_YOLO}"
    )

    print(
        f"[YOLO] Script: {DETECTOR_PY}"
    )

    return subprocess.Popen(
        [
            str(PYTHON_YOLO),
            str(DETECTOR_PY),
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
        bufsize=1,
        cwd=str(BASE_DIR),
    )


# ============================================================
# INICIA OCR
# ============================================================

def iniciar_ocr():
    print(
        "\n[INIT] Iniciando PaddleOCR..."
    )

    print(
        f"[OCR] Python: {PYTHON_OCR}"
    )

    print(
        f"[OCR] Script: {OCR_PY}"
    )

    return subprocess.Popen(
        [
            str(PYTHON_OCR),
            str(OCR_PY),
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
        bufsize=1,
        cwd=str(BASE_DIR),
    )


# ============================================================
# LOCALIZA DATAS
# ============================================================

def extrair_datas_com_posicao(
    page: fitz.Page,
):
    palavras = page.get_text(
        "words"
    )

    encontrados = []

    for i, word in enumerate(
        palavras
    ):
        texto = word[4].strip()

        match = REGEX_DATA.fullmatch(
            texto
        )

        if match:
            rect = fitz.Rect(
                word[0],
                word[1],
                word[2],
                word[3],
            )

            encontrados.append(
                (
                    match.group(0),
                    rect,
                )
            )

            continue

        if not re.search(
            r"\d{1,2}/\d{1,2}/\d{4}",
            texto,
        ):
            continue

        partes = [texto]

        rect = fitz.Rect(
            word[0],
            word[1],
            word[2],
            word[3],
        )

        centro_atual = (
            word[1]
            + word[3]
        ) / 2.0

        for j in range(
            i + 1,
            min(
                i + 4,
                len(palavras),
            ),
        ):
            proxima = palavras[j]

            centro_proxima = (
                proxima[1]
                + proxima[3]
            ) / 2.0

            if abs(
                centro_atual
                - centro_proxima
            ) > 8:
                break

            partes.append(
                proxima[4].strip()
            )

            rect |= fitz.Rect(
                proxima[0],
                proxima[1],
                proxima[2],
                proxima[3],
            )

            combinado = " ".join(
                partes
            )

            match = REGEX_DATA.search(
                combinado
            )

            if match:
                encontrados.append(
                    (
                        match.group(0),
                        rect,
                    )
                )

                break

    unicos = []
    vistos = set()

    for data_str, rect in encontrados:
        chave = (
            data_str,
            round(rect.x0, 2),
            round(rect.y0, 2),
            round(rect.x1, 2),
            round(rect.y1, 2),
        )

        if chave in vistos:
            continue

        vistos.add(chave)

        unicos.append(
            (
                data_str,
                rect,
            )
        )

    unicos.sort(
        key=lambda item: (
            item[1].y0,
            item[1].x0,
        )
    )

    return unicos


# ============================================================
# CABEÇALHO
# ============================================================

def obter_regiao_cabecalho(
    page: fitz.Page,
):
    periodo_rects = page.search_for(
        "Período"
    )

    if not periodo_rects:
        periodo_rects = page.search_for(
            "Periodo"
        )

    if not periodo_rects:
        return 0.0, 0.0

    periodo_rect = periodo_rects[0]

    words = page.get_text(
        "words"
    )

    centro_periodo = (
        periodo_rect.y0
        + periodo_rect.y1
    ) / 2.0

    mesma_linha = []

    for word in words:
        centro = (
            word[1]
            + word[3]
        ) / 2.0

        if abs(
            centro
            - centro_periodo
        ) <= 5:
            mesma_linha.append(
                word
            )

    if mesma_linha:
        return (
            min(
                w[1]
                for w in mesma_linha
            ),
            max(
                w[3]
                for w in mesma_linha
            ),
        )

    return (
        periodo_rect.y0,
        periodo_rect.y1,
    )


# ============================================================
# IDENTIFICA PASSAGENS
# ============================================================

def identificar_passagens(
    page: fitz.Page,
):
    w_pagina = page.rect.width

    words = page.get_text(
        "words"
    )

    ocorrencias = (
        extrair_datas_com_posicao(
            page
        )
    )

    cabecalho_y_top, cabecalho_y_bottom = (
        obter_regiao_cabecalho(
            page
        )
    )

    datas = []

    for data_str, rect in ocorrencias:
        if (
            cabecalho_y_top > 0
            and rect.y0
            <= cabecalho_y_bottom + 5
        ):
            continue

        datas.append(
            (
                data_str,
                rect,
            )
        )

    # ========================================================
    # FIM DA ÁREA DE REGISTROS
    # ========================================================

    y_limite = page.rect.height

    # ========================================================
    # PASSAGENS
    # ========================================================

    passagens = []

    for i, (
        data_str,
        data_rect,
    ) in enumerate(datas):

        if data_rect.y0 >= y_limite:
            continue

        y_top = max(
            0,
            data_rect.y0 - 30,
        )

        if i + 1 < len(datas):
            y_bottom = (
                datas[i + 1][1].y0
                - 5
            )

        else:
            y_bottom = y_limite

        if y_bottom <= y_top:
            continue

        row_words = [
            word
            for word in words
            if (
                y_top
                <=
                (
                    word[1]
                    + word[3]
                ) / 2
                <=
                y_bottom
            )
        ]

        categoria = ""
        placa = ""

        for word in row_words:
            x_center = (
                word[0]
                + word[2]
            ) / 2.0

            texto = word[4].strip()

            # =================================================
            # CATEGORIA
            # =================================================

            if (
                0.20 * w_pagina
                <= x_center
                <= 0.40 * w_pagina
            ):
                if (
                    texto.isdigit()
                    and
                    len(texto) <= 2
                ):
                    categoria = texto

            # =================================================
            # PLACA
            # =================================================

            if (
                0.35 * w_pagina
                <= x_center
                <= 0.55 * w_pagina
            ):
                match = REGEX_PLACA.search(
                    texto
                )

                if match:
                    placa = match.group(0)

        # ====================================================
        # FALLBACK PLACA
        # ====================================================

        if not placa:
            texto_linha = " ".join(
                word[4]
                for word in sorted(
                    row_words,
                    key=lambda x: (
                        x[1],
                        x[0],
                    ),
                )
            )

            match = REGEX_PLACA.search(
                texto_linha
            )

            if match:
                placa = match.group(0)

        passagens.append(
            {
                "Data/Hora":
                    normalizar_data_hora(
                        data_str
                    ),

                "Categoria":
                    categoria,

                "Placa (Texto)":
                    placa,

                "y_top":
                    y_top,

                "y_bottom":
                    y_bottom,
            }
        )

    return passagens


# ============================================================
# EXTRAI IMAGENS ORIGINAIS
# ============================================================

def extrair_imagens_pagina(
    doc,
    page,
):
    imagens = []
    vistos = set()

    for img in page.get_images(
        full=True
    ):
        xref = img[0]
        largura = img[2]
        altura = img[3]

        if (
            largura < 30
            or altura < 15
        ):
            continue

        if xref in vistos:
            continue

        vistos.add(xref)

        rects = page.get_image_rects(
            xref
        )

        if not rects:
            continue

        try:
            info = doc.extract_image(
                xref
            )

        except Exception:
            continue

        if not info:
            continue

        dados = info.get(
            "image"
        )

        if not dados:
            continue

        for rect in rects:
            if (
                rect.width < 10
                or rect.height < 10
            ):
                continue

            if (
                rect.width
                >= page.rect.width * 0.85
                and
                rect.height
                >= page.rect.height * 0.85
            ):
                continue

            imagens.append(
                {
                    "xref":
                        xref,

                    "rect":
                        rect,

                    "bytes":
                        dados,

                    "centro_x":
                        (
                            rect.x0
                            + rect.x1
                        ) / 2.0,

                    "centro_y":
                        (
                            rect.y0
                            + rect.y1
                        ) / 2.0,
                }
            )

    return imagens


# ============================================================
# ASSOCIA IMAGENS ÀS PASSAGENS
# ============================================================

def selecionar_imagens_por_faixa(
    passagens,
    imagens,
):
    selecoes = []

    for passagem in passagens:
        y_top = passagem[
            "y_top"
        ]

        y_bottom = passagem[
            "y_bottom"
        ]

        centro_faixa = (
            y_top + y_bottom
        ) / 2.0

        candidatos = []

        for imagem in imagens:
            rect = imagem[
                "rect"
            ]

            if (
                rect.y1 < y_top
                or
                rect.y0 > y_bottom
            ):
                continue

            centro_y = (
                rect.y0
                + rect.y1
            ) / 2.0

            distancia_y = abs(
                centro_y
                - centro_faixa
            )

            candidatos.append(
                {
                    **imagem,
                    "_distancia_y":
                        distancia_y,
                }
            )

        candidatos.sort(
            key=lambda item: (
                item["_distancia_y"],
                item["rect"].x0,
            )
        )

        selecionadas = candidatos[
            :QTD_IMAGENS_POR_PASSAGEM
        ]

        selecoes.append(
            [
                item["bytes"]
                for item in selecionadas
                if item.get(
                    "bytes"
                )
            ]
        )

    return selecoes


# ============================================================
# WORKER PDF
# ============================================================

def extrair_pagina_worker(
    args
):
    pdf_path, page_num = args

    try:
        doc = fitz.open(
            pdf_path
        )

        page = doc[
            page_num
        ]

        imagens = (
            extrair_imagens_pagina(
                doc,
                page
            )
        )

        passagens = (
            identificar_passagens(
                page
            )
        )

        imagens_por_passagem = (
            selecionar_imagens_por_faixa(
                passagens,
                imagens,
            )
        )

        resultado = []

        for passagem, imagens_candidatas in zip(
            passagens,
            imagens_por_passagem,
        ):
            resultado.append(
                {
                    "Data/Hora":
                        passagem[
                            "Data/Hora"
                        ],

                    "Categoria":
                        passagem[
                            "Categoria"
                        ],

                    "Placa (Texto)":
                        passagem[
                            "Placa (Texto)"
                        ],

                    "imagens_candidatas":
                        imagens_candidatas,

                    "Pagina":
                        page_num + 1,
                }
            )

        doc.close()

        return resultado

    except Exception as e:
        print(
            f"[ERRO] "
            f"Página {page_num + 1}: "
            f"{e}",
            file=sys.stderr,
            flush=True,
        )

        return []


# ============================================================
# EXTRAI PDF
# ============================================================

def extrair_dados_pdf(
    pdf_path,
    max_workers=None,
):
    pdf_path = str(
        pdf_path
    )

    doc = fitz.open(
        pdf_path
    )

    quantidade_paginas = len(
        doc
    )

    doc.close()

    if max_workers is None:
        max_workers = max(
            1,
            multiprocessing.cpu_count() - 1
        )

    print(
        f"[PDF] "
        f"{quantidade_paginas} páginas "
        f"/ {max_workers} workers"
    )

    argumentos = [
        (
            pdf_path,
            pagina,
        )
        for pagina in range(
            quantidade_paginas
        )
    ]

    resultados = [
        None
    ] * quantidade_paginas

    with ProcessPoolExecutor(
        max_workers=max_workers
    ) as executor:

        futures = {
            executor.submit(
                extrair_pagina_worker,
                argumento,
            ):
                argumento[1]
            for argumento in argumentos
        }

        for future in tqdm(
            as_completed(futures),
            total=quantidade_paginas,
            desc="Páginas",
        ):
            page_num = futures[
                future
            ]

            try:
                resultados[
                    page_num
                ] = future.result()

            except Exception as e:
                print(
                    f"[ERRO] Worker "
                    f"página {page_num + 1}: "
                    f"{e}"
                )

                resultados[
                    page_num
                ] = []

    passagens = []

    for pagina in resultados:
        if pagina:
            passagens.extend(
                pagina
            )

    return passagens


# ============================================================
# YOLO
# ============================================================

def yolo_batch(
    processo,
    items,
):
    if not items:
        return {}

    enviar_json(
        processo,
        {
            "command":
                "detect_batch",

            "images":
                items,
        },
    )

    resposta = ler_json(
        processo,
        "YOLO",
    )

    resultados = {}

    for resultado in resposta.get(
        "results",
        []
    ):
        task_id = resultado.get(
            "task_id"
        )

        if task_id:
            resultados[
                task_id
            ] = resultado

    return resultados


# ============================================================
# OCR
# ============================================================

def ocr_batch(
    processo,
    yolo_results,
):
    if not yolo_results:
        return {}

    items = []

    for task_id, resultado in (
        yolo_results.items()
    ):
        if not resultado.get(
            "detected",
            False
        ):
            continue

        crop_b64 = resultado.get(
            "crop_b64"
        )

        if not crop_b64:
            continue

        items.append(
            {
                "task_id":
                    task_id,

                "crop_b64":
                    crop_b64,
            }
        )

    if not items:
        return {}

    enviar_json(
        processo,
        {
            "command":
                "ocr_batch",

            "images":
                items,
        },
    )

    resposta = ler_json(
        processo,
        "OCR",
    )

    resultados = {}

    for resultado in resposta.get(
        "results",
        []
    ):
        task_id = resultado.get(
            "task_id"
        )

        if task_id:
            resultados[
                task_id
            ] = resultado

    return resultados


# ============================================================
# AVALIA PASSAGEM
# ============================================================

def avaliar_passagem(
    placa_texto,
    ocr_por_imagem,
):
    melhor_ocr = ""
    melhor_ocr_bruto = ""
    melhor_distancia = 999.0
    melhor_img = None
    melhor_resultado = {}

    placa_esperada = re.sub(
        r"[^A-Z0-9]",
        "",
        str(placa_texto or "").upper(),
    )

    if not placa_esperada:
        return {
            "Placa (OCR)": "",
            "OCR Bruto": "",
            "Distancia OCR": 999.0,
            "Status OCR": "Revisar",
            "Imagem Usada": "",
            "OCR Cubic": "",
            "Conf Cubic": 0,
            "OCR CLAHE": "",
            "Conf CLAHE": 0,
            "OCR Best Normal": "",
            "Conf Best Normal": 0,
            "SR Usada": False,
            "SR OCR": "",
            "SR Conf": 0,
            "Resultado Graças SR": False,
            "Conf Best Final": 0,
        }

    for img_idx, resultado in sorted(
        ocr_por_imagem.items()
    ):
        texto = resultado.get(
            "ocr_best",
            "",
        )

        if not texto:
            continue

        # ====================================================
        # RECUPERA A MELHOR PLACA DE 7 CARACTERES
        # ====================================================

        (
            ocr_limpo,
            _ocr_bruto_trecho,
            distancia,
        ) = extrair_placa_ocr(
            texto,
            placa_esperada,
        )

        if distancia < melhor_distancia:
            melhor_distancia = distancia
            melhor_ocr = ocr_limpo
            melhor_ocr_bruto = texto
            melhor_img = img_idx
            melhor_resultado = resultado

        # ====================================================
        # DISTÂNCIA ZERO
        # ====================================================

        if distancia == 0:
            break

    aprovado = (
        melhor_distancia
        <= MAX_DIFERENCAS_APROVACAO
    )

    return {
        # ====================================================
        # PLACA RECUPERADA/CORRIGIDA
        # ====================================================

        "Placa (OCR)":
            melhor_ocr,

        # ====================================================
        # OCR ORIGINAL DO PADDLE
        # ====================================================

        "OCR Bruto":
            melhor_ocr_bruto,

        "Distancia OCR":
            round(
                melhor_distancia,
                2,
            ),

        "Status OCR":
            (
                "Aprovado"
                if aprovado
                else "Revisar"
            ),

        "Imagem Usada":
            (
                melhor_img + 1
                if melhor_img is not None
                else ""
            ),

        "OCR Cubic":
            melhor_resultado.get(
                "ocr_cubic",
                "",
            ),

        "Conf Cubic":
            round(
                float(
                    melhor_resultado.get(
                        "confidence_cubic",
                        0
                    )
                    or 0
                ),
                4,
            ),

        "OCR CLAHE":
            melhor_resultado.get(
                "ocr_clahe",
                "",
            ),

        "Conf CLAHE":
            round(
                float(
                    melhor_resultado.get(
                        "confidence_clahe",
                        0
                    )
                    or 0
                ),
                4,
            ),

        "OCR Best Normal":
            melhor_resultado.get(
                "ocr_best_normal",
                "",
            ),

        "Conf Best Normal":
            round(
                float(
                    melhor_resultado.get(
                        "confidence_best_normal",
                        0
                    )
                    or 0
                ),
                4,
            ),

        "SR Usada":
            melhor_resultado.get(
                "sr_used",
                False,
            ),

        "SR OCR":
            melhor_resultado.get(
                "sr_ocr",
                "",
            ),

        "SR Conf":
            round(
                float(
                    melhor_resultado.get(
                        "sr_confidence",
                        0
                    )
                    or 0
                ),
                4,
            ),

        "Resultado Graças SR":
            melhor_resultado.get(
                "resultado_gracas_sr",
                False,
            ),

        "Conf Best Final":
            round(
                float(
                    melhor_resultado.get(
                        "confidence_best",
                        0
                    )
                    or 0
                ),
                4,
            ),
    }


# ============================================================
# PROCESSA UM PDF
# ============================================================

def processar_pdf(
    pdf_path,
    processo_yolo,
    processo_ocr,
):
    print(
        "\n"
        + "=" * 75
    )

    print(
        f"PDF: {pdf_path}"
    )

    print(
        "=" * 75
    )

    inicio = time.perf_counter()

    passagens = extrair_dados_pdf(
        pdf_path
    )

    print(
        f"[PDF] "
        f"Passagens encontradas: "
        f"{len(passagens)}"
    )

    if not passagens:
        print(
            "[AVISO] "
            "Nenhuma passagem encontrada."
        )
        return None

    tmpdir = Path(
        tempfile.mkdtemp(
            prefix="autoban_"
        )
    )

    dados_finais = []

    # ========================================================
    # ESTATÍSTICAS
    # ========================================================

    total_yolo_enviadas = 0
    total_yolo_detectadas = 0
    total_ocr = 0

    try:
        total = len(
            passagens
        )

        n_lotes = (
            total
            + BATCH_SIZE
            - 1
        ) // BATCH_SIZE

        # ====================================================
        # BARRA ÚNICA
        # ====================================================

        barra = tqdm(
            total=total,
            desc="YOLO + OCR",
            unit="pass",
            dynamic_ncols=True,
            leave=True,
        )

        try:
            for lote_idx in range(
                n_lotes
            ):
                inicio_lote = (
                    lote_idx
                    * BATCH_SIZE
                )

                lote = passagens[
                    inicio_lote:
                    inicio_lote
                    + BATCH_SIZE
                ]

                items = []
                mapa_task = {}

                # ============================================
                # PREPARA IMAGENS
                # ============================================

                for p_idx, passagem in enumerate(
                    lote
                ):
                    imagens = passagem.get(
                        "imagens_candidatas",
                        []
                    )

                    for img_idx, img_bytes in enumerate(
                        imagens
                    ):
                        if not img_bytes:
                            continue

                        task_id = (
                            f"l{lote_idx}"
                            f"_p{p_idx}"
                            f"_i{img_idx}"
                        )

                        if img_bytes[:3] == (
                            b"\xff\xd8\xff"
                        ):
                            ext = "jpg"

                        elif img_bytes[:8] == (
                            b"\x89PNG\r\n\x1a\n"
                        ):
                            ext = "png"

                        elif img_bytes[:4] == (
                            b"RIFF"
                        ):
                            ext = "webp"

                        else:
                            ext = "img"

                        caminho = (
                            tmpdir
                            /
                            f"{task_id}.{ext}"
                        )

                        caminho.write_bytes(
                            img_bytes
                        )

                        items.append(
                            {
                                "task_id":
                                    task_id,

                                "image_path":
                                    str(
                                        caminho
                                    ),
                            }
                        )

                        mapa_task[
                            task_id
                        ] = (
                            p_idx,
                            img_idx,
                        )

                # ============================================
                # YOLO
                # ============================================

                yolo_results = yolo_batch(
                    processo_yolo,
                    items,
                )

                total_yolo_enviadas += (
                    len(items)
                )

                total_yolo_detectadas += sum(
                    1
                    for result
                    in yolo_results.values()
                    if result.get(
                        "detected",
                        False
                    )
                )

                # ============================================
                # OCR
                # ============================================

                ocr_results = ocr_batch(
                    processo_ocr,
                    yolo_results,
                )

                total_ocr += len(
                    ocr_results
                )

                # ============================================
                # AGRUPA OCR
                # ============================================

                ocr_por_passagem = {}

                for task_id, resultado in (
                    ocr_results.items()
                ):
                    if task_id not in mapa_task:
                        continue

                    p_idx, img_idx = (
                        mapa_task[
                            task_id
                        ]
                    )

                    ocr_por_passagem.setdefault(
                        p_idx,
                        {}
                    )

                    ocr_por_passagem[
                        p_idx
                    ][
                        img_idx
                    ] = resultado

                # ============================================
                # PASSAGENS
                # ============================================

                for p_idx, passagem in enumerate(
                    lote
                ):
                    avaliacao = avaliar_passagem(
                        passagem.get(
                            "Placa (Texto)",
                            ""
                        ),
                        ocr_por_passagem.get(
                            p_idx,
                            {}
                        ),
                    )

                    dados_finais.append(
                        {
                            "Pagina":
                                passagem.get(
                                    "Pagina",
                                    ""
                                ),

                            "Data/Hora":
                                passagem.get(
                                    "Data/Hora",
                                    ""
                                ),

                            "Categoria":
                                passagem.get(
                                    "Categoria",
                                    ""
                                ),

                            "Placa (Texto)":
                                passagem.get(
                                    "Placa (Texto)",
                                    ""
                                ),

                            **avaliacao,

                            "Qtd Imagens":
                                len(
                                    passagem.get(
                                        "imagens_candidatas",
                                        []
                                    )
                                ),

                            "Qtd Detecções YOLO":
                                sum(
                                    1
                                    for task_id, result
                                    in yolo_results.items()
                                    if (
                                        task_id in mapa_task
                                        and
                                        mapa_task[
                                            task_id
                                        ][0] == p_idx
                                        and
                                        result.get(
                                            "detected",
                                            False
                                        )
                                    )
                                ),

                            "Qtd OCRs":
                                len(
                                    ocr_por_passagem.get(
                                        p_idx,
                                        {}
                                    )
                                ),
                        }
                    )

                # ============================================
                # ATUALIZA BARRA
                # ============================================

                barra.update(
                    len(lote)
                )

                aprovados_atuais = sum(
                    1
                    for item in dados_finais
                    if item.get(
                        "Status OCR"
                    )
                    == "Aprovado"
                )

                revisar_atuais = sum(
                    1
                    for item in dados_finais
                    if item.get(
                        "Status OCR"
                    )
                    == "Revisar"
                )

                sem_deteccao_atuais = sum(
                    1
                    for item in dados_finais
                    if (
                        item.get(
                            "Status OCR"
                        )
                        == "Revisar"
                        and
                        not item.get(
                            "Placa (OCR)"
                        )
                    )
                )

                barra.set_postfix(
                    aprovados=aprovados_atuais,
                    revisar=revisar_atuais,
                    sem_det=sem_deteccao_atuais,
                    ocr=total_ocr,
                )

        finally:
            barra.close()

    finally:
        shutil.rmtree(
            tmpdir,
            ignore_errors=True
        )

    # ========================================================
    # DATAFRAME
    # ========================================================

    df = pd.DataFrame(
        dados_finais
    )

    return df, {
        "total_yolo_enviadas":
            total_yolo_enviadas,

        "total_yolo_detectadas":
            total_yolo_detectadas,

        "total_ocr":
            total_ocr,

        "tempo":
            time.perf_counter()
            - inicio,
    }


# ============================================================
# PROCESSA PASTA INTEIRA
# ============================================================

def processar_pasta(
    pasta
):
    pasta = Path(
        pasta
    ).resolve()

    # ========================================================
    # PROCURA TODOS OS PDFs
    # ========================================================

    arquivos = sorted(
        [
            arquivo
            for arquivo in pasta.rglob(
                "*.pdf"
            )
            if arquivo.is_file()
        ]
    )

    if not arquivos:
        print(
            "[ERRO] "
            "Nenhum PDF encontrado."
        )
        return

    print(
        "\n"
        + "=" * 75
    )

    print(
        "PROCESSAMENTO AUTOBAN"
    )

    print(
        "=" * 75
    )

    print(
        f"Pasta: "
        f"{pasta}"
    )

    print(
        f"PDFs encontrados: "
        f"{len(arquivos)}"
    )

    print(
        "=" * 75
    )

    # ========================================================
    # VERIFICA ARQUIVOS
    # ========================================================

    obrigatorios = [
        PYTHON_YOLO,
        PYTHON_OCR,
        DETECTOR_PY,
        OCR_PY,
    ]

    for arquivo in obrigatorios:
        if not Path(
            arquivo
        ).exists():
            raise FileNotFoundError(
                f"Arquivo não encontrado:\n"
                f"{arquivo}"
            )

    # ========================================================
    # INICIA PROCESSOS
    # ========================================================

    processo_yolo = iniciar_yolo()
    processo_ocr = iniciar_ocr()

    try:
        print(
            "\n[INIT] "
            "Aguardando carregamento dos modelos..."
        )

        time.sleep(
            15
        )

        arquivos_ok = 0
        arquivos_erro = 0

        total_passagens = 0
        total_aprovados = 0
        total_revisar = 0

        total_yolo_enviadas = 0
        total_yolo_detectadas = 0
        total_ocr = 0

        tempo_inicio_total = (
            time.perf_counter()
        )

        # ====================================================
        # PROCESSA CADA PDF
        # ====================================================

        for indice, pdf_path in enumerate(
            arquivos,
            start=1
        ):
            print(
                "\n"
                + "#" * 75
            )

            print(
                f"ARQUIVO "
                f"{indice}/"
                f"{len(arquivos)}"
            )

            print(
                f"{pdf_path}"
            )

            print(
                "#" * 75
            )

            try:
                resultado_processamento = (
                    processar_pdf(
                        pdf_path,
                        processo_yolo,
                        processo_ocr,
                    )
                )

                if resultado_processamento is None:
                    arquivos_erro += 1
                    continue

                df, metricas = (
                    resultado_processamento
                )

                # =================================================
                # NOME DA SAÍDA
                # =================================================

                nome_saida = (
                    pdf_path.stem
                    + "_resultado.xlsx"
                )

                saida = (
                    pdf_path.parent
                    / nome_saida
                )

                df.to_excel(
                    saida,
                    index=False
                )

                # =================================================
                # ESTATÍSTICAS
                # =================================================

                qtd = len(
                    df
                )

                aprovados = int(
                    (
                        df[
                            "Status OCR"
                        ]
                        == "Aprovado"
                    ).sum()
                )

                revisar = int(
                    (
                        df[
                            "Status OCR"
                        ]
                        == "Revisar"
                    ).sum()
                )

                total_passagens += qtd
                total_aprovados += aprovados
                total_revisar += revisar

                total_yolo_enviadas += (
                    metricas[
                        "total_yolo_enviadas"
                    ]
                )

                total_yolo_detectadas += (
                    metricas[
                        "total_yolo_detectadas"
                    ]
                )

                total_ocr += (
                    metricas[
                        "total_ocr"
                    ]
                )

                arquivos_ok += 1

                print(
                    "\n[OK] "
                    f"{pdf_path.name}"
                )

                print(
                    f"Passagens: "
                    f"{qtd}"
                )

                print(
                    f"Aprovados: "
                    f"{aprovados}"
                )

                print(
                    f"Revisar: "
                    f"{revisar}"
                )

                print(
                    f"Imagens enviadas ao YOLO: "
                    f"{metricas['total_yolo_enviadas']}"
                )

                print(
                    f"Detecções YOLO: "
                    f"{metricas['total_yolo_detectadas']}"
                )

                print(
                    f"OCRs processados: "
                    f"{metricas['total_ocr']}"
                )

                print(
                    f"Tempo do PDF: "
                    f"{metricas['tempo']:.2f}s"
                )

                print(
                    f"Saída: "
                    f"{saida}"
                )

            except KeyboardInterrupt:
                raise

            except Exception as e:
                arquivos_erro += 1

                print(
                    "\n[ERRO] "
                    f"{pdf_path.name}"
                )

                print(
                    f"{type(e).__name__}: "
                    f"{e}"
                )

    finally:
        encerrar_processo(
            processo_yolo,
            "YOLO"
        )

        encerrar_processo(
            processo_ocr,
            "OCR"
        )

    # ========================================================
    # RESUMO
    # ========================================================

    tempo_total = (
        time.perf_counter()
        - tempo_inicio_total
    )

    if total_passagens > 0:
        taxa_aprovacao = (
            total_aprovados
            /
            total_passagens
        ) * 100

    else:
        taxa_aprovacao = 0.0

    print(
        "\n"
        + "=" * 75
    )

    print(
        "RESUMO FINAL AUTOBAN"
    )

    print(
        "=" * 75
    )

    print(
        f"PDFs encontrados: "
        f"{len(arquivos)}"
    )

    print(
        f"PDFs processados: "
        f"{arquivos_ok}"
    )

    print(
        f"PDFs com erro: "
        f"{arquivos_erro}"
    )

    print(
        f"Passagens analisadas: "
        f"{total_passagens}"
    )

    print(
        f"Aprovados: "
        f"{total_aprovados}"
    )

    print(
        f"Revisar: "
        f"{total_revisar}"
    )

    print(
        f"Taxa de aprovação: "
        f"{taxa_aprovacao:.2f}%"
    )

    print(
        f"Imagens enviadas ao YOLO: "
        f"{total_yolo_enviadas}"
    )

    print(
        f"Detecções YOLO: "
        f"{total_yolo_detectadas}"
    )

    print(
        f"OCRs processados: "
        f"{total_ocr}"
    )

    print(
        f"Tempo total: "
        f"{tempo_total:.2f}s"
    )

    print(
        "=" * 75
    )


# ============================================================
# MAIN
# ============================================================

def main():
    # ========================================================
    # SEM ARGUMENTO
    # ========================================================

    if len(sys.argv) < 2:
        pasta = (
            DOWNLOADS_DIR
        )

    else:
        pasta = Path(
            sys.argv[1]
        ).resolve()

    # ========================================================
    # VALIDA
    # ========================================================

    if not pasta.exists():
        print(
            "[ERRO] "
            f"Pasta não encontrada:\n"
            f"{pasta}"
        )

        sys.exit(1)

    if not pasta.is_dir():
        print(
            "[ERRO] "
            "O caminho informado não é uma pasta:\n"
            f"{pasta}"
        )

        sys.exit(1)

    # ========================================================
    # SAÍDA
    # ========================================================

    saida = ensure_parent(
        OUTPUT_DIR
        /
        "resultado_autoban.xlsx"
    )

    # ========================================================
    # INFORMAÇÃO SOBRE CONVERSOR
    # ========================================================

    magick = (
        _encontrar_imagemagick()
    )

    if magick:
        print(
            f"[VECTOR] ImageMagick encontrado: "
            f"{magick}"
        )

    else:
        print(
            "[VECTOR] ImageMagick não encontrado."
        )

        print(
            "[VECTOR] "
            "EMF/WMF poderão depender do suporte "
            "do Pillow."
        )

    # ========================================================
    # EXECUTA
    # ========================================================

    # Mantido para compatibilidade
    # com a configuração existente.

    processar_pasta(
        pasta
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()