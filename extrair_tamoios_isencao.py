from __future__ import annotations
from placas import (
    aplicar_mascara as _aplicar_mascara_comum,
    descobrir_mascara as _descobrir_mascara_comum,
    distancia_placas as _distancia_placas_comum,
    distancia_ponderada as _distancia_ponderada,
    normalizar_ocr as _normalizar_ocr_comum,
    normalizar_placa as _normalizar_placa,
    CONFUNDIVEIS as _CONFUNDIVEIS,
    MAPA_LETRAS as MAPA_LETRAS,
    MAPA_DIGITOS as MAPA_DIGITOS,
)

from common import enviar_json

import functools
import contextlib
import io
import json
import re
import subprocess
import sys
import tempfile
import time
import unicodedata
from pathlib import Path

import pandas as pd
from tqdm import tqdm


# ============================================================
# CONFIG
# ============================================================

from config import (
    BASE_DIR,
    caminho_resultado,
    PYTHON_YOLO,
    PYTHON_OCR,
    DETECTOR_PY,
    OCR_PY,
    ensure_parent,
)


# ============================================================
# CONFIGURAÇÕES
# ============================================================

BATCH_SIZE = 64

MAX_DIFERENCAS_APROVACAO = 3

TAMANHO_PLACA = 7

# Máscara Mercosul:
#
# posição 0 = letra
# posição 1 = letra
# posição 2 = letra
# posição 3 = dígito
# posição 4 = letra
# posição 5 = dígito
# posição 6 = dígito
#
MASCARA_PLACA = "LLLDLDD"


EXTENSOES_IMAGEM = {
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
}


# ============================================================
# REGEX DO ID
# ============================================================

REGEX_ID_ARQUIVO = re.compile(
    r"^([A-Za-z0-9]+)(?:-|$)"
)


# ============================================================
# CONFUNDÍVEIS
# ============================================================



# ============================================================
# MAPA PARA LETRAS
# ============================================================



# ============================================================
# MAPA PARA DÍGITOS
# ============================================================



# ============================================================
# LEVENSHTEIN
# ============================================================

def _levenshtein(a: str, b: str) -> float:
    return _distancia_ponderada(a, b, _CONFUNDIVEIS)


# ============================================================
# LIMPA TEXTO
# ============================================================

def limpar_texto(texto) -> str:
    return _normalizar_placa(texto)


# ============================================================
# APLICA MÁSCARA LLLDLDD
# ============================================================

def aplicar_mascara_placa(candidato):
    return _aplicar_mascara_comum(candidato, MASCARA_PLACA, MAPA_LETRAS, MAPA_DIGITOS, tamanho=TAMANHO_PLACA, comprimento_invalido="vazio", preservar_indice=4)


# ============================================================
# PROCURA A MELHOR SEQUÊNCIA DE 7
# ============================================================

def encontrar_melhor_trecho(
    placa_esperada,
    placa_ocr,
):

    esperada = limpar_texto(
        placa_esperada
    )

    ocr = limpar_texto(
        placa_ocr
    )

    if not esperada or not ocr:

        return {
            "candidato_bruto":
                "",

            "candidato_limpo":
                "",

            "distancia":
                999.0,

            "inicio":
                None,
        }

    # ========================================================
    # PLACA DA PLANILHA NÃO TEM 7 CARACTERES
    # ========================================================

    if len(esperada) != TAMANHO_PLACA:

        return {
            "candidato_bruto":
                ocr,

            "candidato_limpo":
                ocr,

            "distancia":
                distancia_placas(
                    esperada,
                    ocr,
                ),

            "inicio":
                0,
        }

    # ========================================================
    # OCR MENOR QUE 7
    # ========================================================

    if len(ocr) < TAMANHO_PLACA:

        candidato_limpo = (
            aplicar_mascara_placa(
                ocr
            )
        )

        if not candidato_limpo:

            candidato_limpo = ocr

        return {
            "candidato_bruto":
                ocr,

            "candidato_limpo":
                candidato_limpo,

            "distancia":
                distancia_placas(
                    esperada,
                    candidato_limpo,
                ),

            "inicio":
                0,
        }

    # ========================================================
    # OCR TEM EXATAMENTE 7
    # ========================================================

    if len(ocr) == TAMANHO_PLACA:

        candidato_limpo = (
            aplicar_mascara_placa(
                ocr
            )
        )

        if not candidato_limpo:

            candidato_limpo = ocr

        distancia = distancia_placas(
            esperada,
            candidato_limpo,
        )

        # Bruto exatamente igual à placa
        # sempre tem prioridade máxima.
        prioridade = (
            0
            if ocr == esperada
            else
            1
            if candidato_limpo == esperada
            else
            2
        )

        return {
            "candidato_bruto":
                ocr,

            "candidato_limpo":
                candidato_limpo,

            "distancia":
                distancia,

            "inicio":
                0,

            "prioridade":
                prioridade,
        }

    # ========================================================
    # OCR MAIOR QUE 7
    #
    # Testa todas as janelas de 7 caracteres.
    # ========================================================

    melhor = None

    for inicio in range(
        0,
        len(ocr)
        -
        TAMANHO_PLACA
        +
        1,
    ):

        candidato_bruto = ocr[
            inicio:
            inicio
            +
            TAMANHO_PLACA
        ]

        candidato_limpo = (
            aplicar_mascara_placa(
                candidato_bruto
            )
        )

        if not candidato_limpo:
            continue

        distancia = distancia_placas(
            esperada,
            candidato_limpo,
        )

        match_bruto = (
            candidato_bruto
            ==
            esperada
        )

        match_limpo = (
            candidato_limpo
            ==
            esperada
        )

        prioridade = (
            0
            if match_bruto
            else
            1
            if match_limpo
            else
            2
        )

        chave = (
            prioridade,
            distancia,
            inicio,
        )

        candidato = {
            "candidato_bruto":
                candidato_bruto,

            "candidato_limpo":
                candidato_limpo,

            "distancia":
                distancia,

            "inicio":
                inicio,

            "prioridade":
                prioridade,

            "chave":
                chave,
        }

        if (
            melhor is None
            or
            chave
            <
            melhor["chave"]
        ):

            melhor = candidato

    if melhor is None:

        return {
            "candidato_bruto":
                ocr,

            "candidato_limpo":
                ocr,

            "distancia":
                distancia_placas(
                    esperada,
                    ocr,
                ),

            "inicio":
                None,
        }

    return melhor


# ============================================================
# DISTÂNCIA
# ============================================================

def distancia_placas(placa_a: str, placa_b: str) -> float:
    return _distancia_placas_comum(placa_a, placa_b, _CONFUNDIVEIS)


# ============================================================
# DESCREVE DIFERENÇAS DE CARACTERES
# ============================================================

def descrever_diferencas(
    placa_esperada: str,
    placa_ocr: str,
) -> str:

    esperado = limpar_texto(
        placa_esperada
    )

    recebido = limpar_texto(
        placa_ocr
    )

    if not esperado or not recebido:

        return "Sem comparação"

    linhas = len(
        esperado
    ) + 1

    colunas = len(
        recebido
    ) + 1

    matriz = [
        [
            0.0
            for _ in range(
                colunas
            )
        ]
        for _ in range(
            linhas
        )
    ]

    operacoes = [
        [
            None
            for _ in range(
                colunas
            )
        ]
        for _ in range(
            linhas
        )
    ]

    for i in range(
        linhas
    ):

        matriz[i][0] = float(i)

    for j in range(
        colunas
    ):

        matriz[0][j] = float(j)

    for i in range(
        1,
        linhas,
    ):

        for j in range(
            1,
            colunas,
        ):

            ca = esperado[
                i - 1
            ]

            cb = recebido[
                j - 1
            ]

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
                matriz[
                    i - 1
                ][
                    j - 1
                ]
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

    i = len(
        esperado
    )

    j = len(
        recebido
    )

    diferencas = []

    while (
        i > 0
        or
        j > 0
    ):

        if (
            i > 0
            and
            j > 0
        ):

            tipo = operacoes[
                i
            ][
                j
            ][0]

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
                    recebido[
                        j - 1
                    ],
                    esperado[
                        i - 1
                    ],
                )
            )

            i -= 1
            j -= 1

        elif tipo == "insercao":

            diferencas.append(
                (
                    j,
                    recebido[
                        j - 1
                    ],
                    "—",
                )
            )

            j -= 1

        else:

            diferencas.append(
                (
                    i,
                    "—",
                    esperado[
                        i - 1
                    ],
                )
            )

            i -= 1

    diferencas.reverse()

    if not diferencas:

        return "Nenhuma"

    resultado = []

    for (
        posicao,
        ocr,
        sistema,
    ) in diferencas:

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

    return "; ".join(
        resultado
    )


# ============================================================
# NORMALIZA COLUNA
# ============================================================

def normalizar_coluna(
    nome,
) -> str:

    nome = str(
        nome
    ).strip().lower()

    nome = unicodedata.normalize(
        "NFKD",
        nome,
    )

    nome = "".join(
        c
        for c in nome
        if not unicodedata.combining(c)
    )

    nome = re.sub(
        r"\s+",
        " ",
        nome,
    )

    return nome


# ============================================================
# ENVIA JSON
# ============================================================



# ============================================================
# LÊ JSON
# ============================================================

def ler_json(
    processo,
    nome,
):

    linha = (
        processo.stdout.readline()
    )

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


# ============================================================
# ENCERRA
# ============================================================

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
                "command":
                    "exit"
            }
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
#
# stderr vai para DEVNULL para não poluir terminal.
# ============================================================

def iniciar_yolo():

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
#
# stderr vai para DEVNULL para não poluir terminal.
# ============================================================

def iniciar_ocr():

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
# CARREGA EXCEL
# ============================================================

def carregar_planilha(
    caminho_excel,
):

    print(
        "\n[EXCEL] Carregando:"
    )

    print(
        caminho_excel
    )

    df = pd.read_excel(
        caminho_excel,
        header=0,
        dtype=str,
    )

    if df.empty:

        print(
            "[AVISO] "
            "Planilha está vazia."
        )

    # ========================================================
    # NORMALIZA CABEÇALHOS
    # ========================================================

    df.columns = [
        normalizar_coluna(
            coluna
        )
        for coluna in df.columns
    ]

    print(
        "[EXCEL] Colunas:"
    )

    print(
        list(df.columns)
    )

    # ========================================================
    # SINÔNIMOS
    # ========================================================

    sinonimos = {

        "id": [
            "id transacao",
            "id",
            "passagem",
            "numero",
            "numeroid",
            "idpassagem",
            "idsistema",
            "id sistema",
            "id passagem",
            "id_passagem",
        ],

        "placa": [
            "placa",
            "placaveiculo",
            "placa veiculo",
        ],

        "data_hora": [
            "data/hora",
            "data hora",
            "data_hora",
            "data e hora",
        ],

        "data": [
            "data",
            "data passagem",
            "datapassagem",
        ],

        "hora": [
            "hora",
            "hora passagem",
            "horapassagem",
        ],

        "categoria": [
            "categoria cca",
            "categoria",
            "cat",
            "classeveiculo",
            "classe veiculo",
            "tipoveiculo",
        ],
    }

    mapa = {}

    for campo, nomes in sinonimos.items():

        for nome in nomes:

            nome = normalizar_coluna(
                nome
            )

            if nome in df.columns:

                mapa[
                    campo
                ] = nome

                break

    print(
        "[EXCEL] Mapeamento:"
    )

    print(
        mapa
    )

    # ========================================================
    # RENOMEIA
    # ========================================================

    renomear = {}

    for campo, coluna in mapa.items():

        renomear[
            coluna
        ] = campo

    if renomear:

        df = df.rename(
            columns=renomear
        )

    # ========================================================
    # CAMPOS INTERNOS
    # ========================================================

    for campo in [
        "id",
        "placa",
        "data_hora",
        "data",
        "hora",
        "categoria",
    ]:

        if campo not in df.columns:

            df[
                campo
            ] = ""

    # ========================================================
    # LIMPEZA
    # ========================================================

    for coluna in [
        "id",
        "placa",
        "data_hora",
        "data",
        "hora",
        "categoria",
    ]:

        df[
            coluna
        ] = (
            df[
                coluna
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # ========================================================
    # DATA/HORA
    # ========================================================

    for indice in df.index:

        dh = str(
            df.at[
                indice,
                "data_hora",
            ]
        ).strip()

        data = str(
            df.at[
                indice,
                "data",
            ]
        ).strip()

        hora = str(
            df.at[
                indice,
                "hora",
            ]
        ).strip()

        if (
            dh
            and
            (
                not data
                or
                data.lower() == "nan"
                or
                not hora
                or
                hora.lower() == "nan"
            )
        ):

            try:

                dt = pd.to_datetime(
                    dh,
                    errors="coerce",
                )

                if pd.notna(dt):

                    df.at[
                        indice,
                        "data",
                    ] = dt.strftime(
                        "%d/%m/%Y"
                    )

                    df.at[
                        indice,
                        "hora",
                    ] = dt.strftime(
                        "%H:%M:%S"
                    )

            except Exception:

                pass

    # ========================================================
    # ÍNDICE ID
    # ========================================================

    indice_id = {}

    for _, row in df.iterrows():

        id_valor = limpar_texto(
            row["id"]
        )

        if not id_valor:
            continue

        if id_valor not in indice_id:

            indice_id[
                id_valor
            ] = row.to_dict()

    print(
        f"[EXCEL] "
        f"Registros: {len(df)}"
    )

    print(
        f"[EXCEL] "
        f"IDs únicos: {len(indice_id)}"
    )

    return (
        df,
        indice_id,
    )


# ============================================================
# ID DO ARQUIVO
# ============================================================

def extrair_id_arquivo(
    nome_arquivo,
):

    stem = Path(
        nome_arquivo
    ).stem.strip()

    match = (
        REGEX_ID_ARQUIVO.match(
            stem
        )
    )

    if not match:

        return (
            "",
            "ID fora do padrão"
        )

    id_extraido = limpar_texto(
        match.group(1)
    )

    if not id_extraido:

        return (
            "",
            "ID fora do padrão"
        )

    return (
        id_extraido,
        "OK"
    )


# ============================================================
# LISTA IMAGENS
# ============================================================

def listar_imagens(
    pasta_imagens,
):

    pasta = Path(
        pasta_imagens
    ).resolve()

    imagens = [
        arquivo
        for arquivo in pasta.rglob("*")
        if (
            arquivo.is_file()
            and
            arquivo.suffix.lower()
            in EXTENSOES_IMAGEM
        )
    ]

    return sorted(
        imagens
    )


# ============================================================
# YOLO
# ============================================================

def executar_yolo_batch(
    processo_yolo,
    lote,
):

    if not lote:
        return {}

    items = []

    for item in lote:

        caminho = str(
            Path(
                item["caminho"]
            ).resolve()
        )

        arquivo = Path(
            caminho
        )

        if (
            not caminho
            or
            not arquivo.exists()
            or
            arquivo.stat().st_size == 0
        ):

            continue

        items.append(
            {
                "task_id":
                    item[
                        "task_id"
                    ],

                "image_path":
                    caminho,
            }
        )

    if not items:
        return {}

    enviar_json(
        processo_yolo,
        {
            "command":
                "detect_batch",

            "images":
                items,
        }
    )

    resposta = ler_json(
        processo_yolo,
        "YOLO"
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

def executar_ocr_batch(
    processo_ocr,
    yolo_results,
):

    if not yolo_results:
        return {}

    crops = []

    for (
        task_id,
        resultado,
    ) in yolo_results.items():

        if not resultado.get(
            "detected",
            False
        ):

            continue

        crop_b64 = (
            resultado.get(
                "crop_b64"
            )
        )

        if not crop_b64:
            continue

        crops.append(
            {
                "task_id":
                    task_id,

                "crop_b64":
                    crop_b64,
            }
        )

    if not crops:
        return {}

    enviar_json(
        processo_ocr,
        {
            "command":
                "ocr_batch",

            "images":
                crops,
        }
    )

    resposta = ler_json(
        processo_ocr,
        "OCR"
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
# AVALIA OCR
# ============================================================

def avaliar_ocr(
    placa_planilha,
    resultado_ocr,
):

    placa_planilha = limpar_texto(
        placa_planilha
    )

    # ========================================================
    # TEXTO PREFERENCIAL:
    # OCR BEST DO PADDLE
    # ========================================================

    ocr_paddle = (
        resultado_ocr.get(
            "ocr_best",
            ""
        )
        or
        ""
    )

    ocr_paddle = str(
        ocr_paddle
    ).strip()

    # ========================================================
    # SEM OCR
    # ========================================================

    if not limpar_texto(
        ocr_paddle
    ):

        try:

            confianca = float(
                resultado_ocr.get(
                    "confidence_best",
                    0.0,
                )
                or
                0.0
            )

        except Exception:

            confianca = 0.0

        return {
            "placa_ocr":
                "",

            "ocr_bruto":
                resultado_ocr.get(
                    "ocr_best_original",
                    "",
                ),

            "confianca":
                confianca,

            "diferenca":
                999.0,

            "status":
                "Revisar - sem OCR",
        }

    # ========================================================
    # EXTRAÇÃO INTELIGENTE
    # ========================================================

    melhor = (
        encontrar_melhor_trecho(
            placa_planilha,
            ocr_paddle,
        )
    )

    placa_ocr_limpa = (
        melhor[
            "candidato_limpo"
        ]
    )

    diferenca = (
        melhor[
            "distancia"
        ]
    )

    try:

        confianca = float(
            resultado_ocr.get(
                "confidence_best",
                0.0,
            )
            or
            0.0
        )

    except Exception:

        confianca = 0.0

    # ========================================================
    # STATUS
    # ========================================================

    if (
        diferenca
        <=
        MAX_DIFERENCAS_APROVACAO
    ):

        status = (
            "Aprovado - busca nas imagens"
        )

    else:

        status = "Revisar"

    return {
        "placa_ocr":
            placa_ocr_limpa,

        "ocr_bruto":
            resultado_ocr.get(
                "ocr_best_original",
                "",
            ),

        "confianca":
            confianca,

        "diferenca":
            round(
                diferenca,
                2,
            ),

        "status":
            status,

        "candidato":
            melhor.get(
                "candidato_bruto",
                "",
            ),
    }


# ============================================================
# CRIA REGISTRO SEM IMAGEM
# ============================================================

def registro_sem_imagem(
    nome,
    id_extraido,
    status_id,
):

    return {
        "Arquivo":
            nome,

        "ID Extraído":
            id_extraido,

        "Status ID":
            status_id,

        "Placa Planilha":
            "",

        "Data":
            "",

        "Hora":
            "",

        "Categoria":
            "",

        "Placa OCR":
            "",

        "OCR Bruto":
            "",

        "OCR Confidence":
            0.0,

        "YOLO Confidence":
            0.0,

        "Diferença OCR":
            999.0,

        "Status OCR":
            "Revisar - sem imagem",

        "Imagem Usada":
            "",
    }


# ============================================================
# PROCESSA UM ARQUIVO
# ============================================================

def processar_excel(
    caminho_excel,
    pasta_imagens,
    caminho_saida,
    processo_yolo,
    processo_ocr,
):

    inicio = time.perf_counter()

    print(
        "\n"
        + "=" * 75
    )

    print(
        f"ARQUIVO EXCEL:\n"
        f"{caminho_excel}"
    )

    print(
        f"PASTA IMAGENS:\n"
        f"{pasta_imagens}"
    )

    print(
        "=" * 75
    )

    # ========================================================
    # EXCEL
    # ========================================================

    (
        df_ref,
        indice_id,
    ) = carregar_planilha(
        caminho_excel
    )

    # ========================================================
    # IMAGENS
    # ========================================================

    imagens = listar_imagens(
        pasta_imagens
    )

    print(
        f"\n[IMAGENS] "
        f"{len(imagens)} encontradas."
    )

    if not imagens:

        print(
            "[ERRO] "
            "Nenhuma imagem encontrada."
        )

        return None

    # ========================================================
    # REGISTROS
    # ========================================================

    registros = []

    itens_validos = []

    # ========================================================
    # CRUZAMENTO DOS IDS
    # ========================================================

    for caminho_img in tqdm(
        imagens,
        desc="Cruzando IDs",
        unit="img",
    ):

        nome = caminho_img.name

        (
            id_extraido,
            status_id,
        ) = (
            extrair_id_arquivo(
                nome
            )
        )

        # ----------------------------------------------------
        # ID INVÁLIDO
        # ----------------------------------------------------

        if status_id != "OK":

            registros.append(
                registro_sem_imagem(
                    nome,
                    id_extraido,
                    status_id,
                )
            )

            continue

        # ----------------------------------------------------
        # PROCURA ID
        # ----------------------------------------------------

        linha_ref = indice_id.get(
            id_extraido
        )

        if linha_ref is None:

            registros.append(
                registro_sem_imagem(
                    nome,
                    id_extraido,
                    "ID não encontrado na planilha",
                )
            )

            continue

        # ----------------------------------------------------
        # ID ENCONTRADO
        # ----------------------------------------------------

        item = {
            "caminho":
                str(
                    caminho_img.resolve()
                ),

            "nome":
                nome,

            "task_id":
                f"img_{len(itens_validos)}",

            "id":
                id_extraido,

            "placa":
                str(
                    linha_ref.get(
                        "placa",
                        "",
                    )
                ).strip().upper(),

            "data":
                str(
                    linha_ref.get(
                        "data",
                        "",
                    )
                ).strip(),

            "hora":
                str(
                    linha_ref.get(
                        "hora",
                        "",
                    )
                ).strip(),

            "categoria":
                str(
                    linha_ref.get(
                        "categoria",
                        "",
                    )
                ).strip(),
        }

        itens_validos.append(
            item
        )

    print(
        f"\n[VALIDAÇÃO] "
        f"{len(itens_validos)} imagens "
        f"com ID encontrado."
    )

    # ========================================================
    # AGRUPA POR ID
    # ========================================================

    imagens_por_id = {}

    for item in itens_validos:

        imagens_por_id.setdefault(
            item["id"],
            [],
        ).append(
            item
        )

    # ========================================================
    # PRESERVA ORDEM
    # ========================================================

    for id_valor in imagens_por_id:

        imagens_por_id[
            id_valor
        ].sort(
            key=lambda item:
                item["nome"].lower()
        )

    # ========================================================
    # STATUS
    # ========================================================

    ids_resolvidos = set()

    ids_esgotados = set()

    ids_com_erro = set()

    processados = set()

    # ========================================================
    # RESULTADOS
    # ========================================================

    resultados_yolo = {}

    resultados_ocr = {}

    # ========================================================
    # ESTATÍSTICAS
    # ========================================================

    total_ids = len(
        imagens_por_id
    )

    total_imagens_enviadas = 0

    total_yolo_detectadas = 0

    total_ocr_executados = 0

    total_aprovados = 0

    total_sem_deteccao = 0

    # ========================================================
    # MELHOR RESULTADO POR ID
    # ========================================================

    melhor_imagem_por_id = {}

    melhor_resultado_por_id = {}

    # ========================================================
    # BARRA
    # ========================================================

    maior_quantidade_imagens = max(
        (
            len(lista)
            for lista in (
                imagens_por_id.values()
            )
        ),
        default=0,
    )

    barra = tqdm(
        total=total_ids,
        desc="YOLO + OCR",
        unit="ID",
        dynamic_ncols=True,
        leave=True,
    )

    try:

        # ====================================================
        # RODADAS
        # ====================================================

        for rodada in range(
            maior_quantidade_imagens
        ):

            lote_rodada = []

            ids_lote = []

            # =================================================
            # UMA IMAGEM POR ID
            # =================================================

            for (
                id_valor,
                lista_imagens,
            ) in imagens_por_id.items():

                if id_valor in ids_resolvidos:
                    continue

                if id_valor in ids_esgotados:
                    continue

                if rodada >= len(
                    lista_imagens
                ):

                    ids_esgotados.add(
                        id_valor
                    )

                    continue

                item = lista_imagens[
                    rodada
                ]

                lote_rodada.append(
                    item
                )

                ids_lote.append(
                    id_valor
                )

            if not lote_rodada:
                break

            # =================================================
            # YOLO EM SUBLOTES
            # =================================================

            for inicio_lote in range(
                0,
                len(lote_rodada),
                BATCH_SIZE,
            ):

                lote = lote_rodada[
                    inicio_lote:
                    inicio_lote
                    +
                    BATCH_SIZE
                ]

                # =============================================
                # YOLO
                # =============================================

                yolo_lote = executar_yolo_batch(
                    processo_yolo,
                    lote,
                )

                resultados_yolo.update(
                    yolo_lote
                )

                total_imagens_enviadas += (
                    len(lote)
                )

                detectadas = 0

                for (
                    task_id,
                    resultado,
                ) in yolo_lote.items():

                    if resultado.get(
                        "detected",
                        False,
                    ):

                        detectadas += 1

                total_yolo_detectadas += (
                    detectadas
                )

                # =============================================
                # OCR
                # =============================================

                ocr_lote = executar_ocr_batch(
                    processo_ocr,
                    yolo_lote,
                )

                resultados_ocr.update(
                    ocr_lote
                )

                total_ocr_executados += (
                    len(ocr_lote)
                )

                # =============================================
                # MARCA PROCESSADOS
                # =============================================

                for item in lote:

                    processados.add(
                        item[
                            "task_id"
                        ]
                    )

                # =============================================
                # AVALIA CADA IMAGEM
                # =============================================

                for item in lote:

                    task_id = item[
                        "task_id"
                    ]

                    id_valor = item[
                        "id"
                    ]

                    # -----------------------------------------
                    # YOLO
                    # -----------------------------------------

                    yolo_result = (
                        yolo_lote.get(
                            task_id,
                            {}
                        )
                    )

                    if not yolo_result.get(
                        "detected",
                        False
                    ):

                        continue

                    # -----------------------------------------
                    # OCR
                    # -----------------------------------------

                    ocr_result = (
                        ocr_lote.get(
                            task_id
                        )
                    )

                    if not ocr_result:

                        continue

                    # -----------------------------------------
                    # AVALIA
                    # -----------------------------------------

                    avaliacao = avaliar_ocr(
                        item[
                            "placa"
                        ],
                        ocr_result,
                    )

                    distancia = (
                        avaliacao[
                            "diferenca"
                        ]
                    )

                    # =========================================
                    # GUARDA MELHOR RESULTADO
                    # =========================================

                    resultado_anterior = (
                        melhor_resultado_por_id.get(
                            id_valor
                        )
                    )

                    if (
                        resultado_anterior is None
                        or
                        distancia
                        <
                        resultado_anterior[
                            "diferenca"
                        ]
                    ):

                        melhor_resultado_por_id[
                            id_valor
                        ] = {
                            "diferenca":
                                distancia,

                            "avaliacao":
                                avaliacao,

                            "item":
                                item,

                            "yolo_result":
                                yolo_result,

                            "ocr_result":
                                ocr_result,
                        }

                    # =========================================
                    # MATCH
                    # =========================================

                    if (
                        avaliacao[
                            "status"
                        ]
                        ==
                        "Aprovado - busca nas imagens"
                    ):

                        if (
                            id_valor
                            not in ids_resolvidos
                        ):

                            ids_resolvidos.add(
                                id_valor
                            )

                            melhor_imagem_por_id[
                                id_valor
                            ] = item

                            total_aprovados += 1

                            barra.update(
                                1
                            )

            # =================================================
            # IDS ESGOTADOS
            # =================================================

            for (
                id_valor,
                lista_imagens,
            ) in imagens_por_id.items():

                if id_valor in ids_resolvidos:
                    continue

                if id_valor in ids_esgotados:
                    continue

                if rodada + 1 >= len(
                    lista_imagens
                ):

                    ids_esgotados.add(
                        id_valor
                    )

                    total_sem_deteccao += 1

                    barra.update(
                        1
                    )

            # =================================================
            # BARRA
            # =================================================

            concluidos = (
                len(
                    ids_resolvidos
                )
                +
                len(
                    ids_esgotados
                )
            )

            pendentes = max(
                0,
                total_ids
                -
                concluidos
            )

            barra.set_postfix(
                resolvidos=
                    len(
                        ids_resolvidos
                    ),

                pendentes=
                    pendentes,

                ocr=
                    total_ocr_executados,
            )

    finally:

        barra.close()

    # ========================================================
    # IDS RESTANTES COM ERRO
    # ========================================================

    for id_valor in imagens_por_id:

        if id_valor in ids_resolvidos:
            continue

        if id_valor in ids_esgotados:
            continue

        ids_com_erro.add(
            id_valor
        )

    # ========================================================
    # GERA RESULTADOS
    # ========================================================

    for item in itens_validos:

        task_id = item[
            "task_id"
        ]

        id_valor = item[
            "id"
        ]

        # ====================================================
        # ID RESOLVIDO
        # ====================================================

        if id_valor in ids_resolvidos:

            imagem_usada = (
                melhor_imagem_por_id.get(
                    id_valor
                )
            )

            # -----------------------------------------------
            # É A IMAGEM QUE RESOLVEU?
            # -----------------------------------------------

            if (
                imagem_usada is not None
                and
                imagem_usada[
                    "task_id"
                ]
                ==
                task_id
            ):

                yolo_result = (
                    resultados_yolo.get(
                        task_id,
                        {}
                    )
                )

                ocr_result = (
                    resultados_ocr.get(
                        task_id,
                        {}
                    )
                )

                if ocr_result:

                    avaliacao = avaliar_ocr(
                        item[
                            "placa"
                        ],
                        ocr_result,
                    )

                    validacao = {
                        "Placa OCR":
                            avaliacao[
                                "placa_ocr"
                            ],

                        "OCR Bruto":
                            avaliacao[
                                "ocr_bruto"
                            ],

                        "OCR Confidence":
                            avaliacao[
                                "confianca"
                            ],

                        "YOLO Confidence":
                            float(
                                yolo_result.get(
                                    "confidence",
                                    0.0,
                                )
                                or
                                0.0
                            ),

                        "Diferença OCR":
                            avaliacao[
                                "diferenca"
                            ],

                        "Status OCR":
                            avaliacao[
                                "status"
                            ],

                        "Imagem Usada":
                            item[
                                "nome"
                            ],
                    }

                else:

                    validacao = {
                        "Placa OCR":
                            "",

                        "OCR Bruto":
                            "",

                        "OCR Confidence":
                            0.0,

                        "YOLO Confidence":
                            float(
                                yolo_result.get(
                                    "confidence",
                                    0.0,
                                )
                                or
                                0.0
                            ),

                        "Diferença OCR":
                            999.0,

                        "Status OCR":
                            "Revisar - sem OCR",

                        "Imagem Usada":
                            item[
                                "nome"
                            ],
                    }

            else:

                # -------------------------------------------
                # OUTRA IMAGEM
                # -------------------------------------------

                validacao = {
                    "Placa OCR":
                        "",

                    "OCR Bruto":
                        "",

                    "OCR Confidence":
                        0.0,

                    "YOLO Confidence":
                        0.0,

                    "Diferença OCR":
                        999.0,

                    "Status OCR":
                        "Ignorada - passagem já resolvida",

                    "Imagem Usada":
                        "",
                }

        # ====================================================
        # ID NÃO RESOLVIDO
        # ====================================================

        else:

            yolo_result = (
                resultados_yolo.get(
                    task_id,
                    {}
                )
            )

            ocr_result = (
                resultados_ocr.get(
                    task_id,
                    {}
                )
            )

            if not yolo_result:

                validacao = {
                    "Placa OCR":
                        "",

                    "OCR Bruto":
                        "",

                    "OCR Confidence":
                        0.0,

                    "YOLO Confidence":
                        0.0,

                    "Diferença OCR":
                        999.0,

                    "Status OCR":
                        "Revisar - sem detecção",

                    "Imagem Usada":
                        "",
                }

            elif not yolo_result.get(
                "detected",
                False
            ):

                validacao = {
                    "Placa OCR":
                        "",

                    "OCR Bruto":
                        "",

                    "OCR Confidence":
                        0.0,

                    "YOLO Confidence":
                        0.0,

                    "Diferença OCR":
                        999.0,

                    "Status OCR":
                        "Revisar - sem detecção",

                    "Imagem Usada":
                        item[
                            "nome"
                        ],
                }

            elif not ocr_result:

                validacao = {
                    "Placa OCR":
                        "",

                    "OCR Bruto":
                        "",

                    "OCR Confidence":
                        0.0,

                    "YOLO Confidence":
                        float(
                            yolo_result.get(
                                "confidence",
                                0.0,
                            )
                            or
                            0.0
                        ),

                    "Diferença OCR":
                        999.0,

                    "Status OCR":
                        "Revisar - sem OCR",

                    "Imagem Usada":
                        item[
                            "nome"
                        ],
                }

            else:

                avaliacao = avaliar_ocr(
                    item[
                        "placa"
                    ],
                    ocr_result,
                )

                validacao = {
                    "Placa OCR":
                        avaliacao[
                            "placa_ocr"
                        ],

                    "OCR Bruto":
                        avaliacao[
                            "ocr_bruto"
                        ],

                    "OCR Confidence":
                        avaliacao[
                            "confianca"
                        ],

                    "YOLO Confidence":
                        float(
                            yolo_result.get(
                                "confidence",
                                0.0,
                            )
                            or
                            0.0
                        ),

                    "Diferença OCR":
                        avaliacao[
                            "diferenca"
                        ],

                    "Status OCR":
                        avaliacao[
                            "status"
                        ],

                    "Imagem Usada":
                        item[
                            "nome"
                        ],
                }

        # ====================================================
        # REGISTRO
        # ====================================================

        registros.append(
            {
                "Arquivo":
                    item[
                        "nome"
                    ],

                "ID Extraído":
                    item[
                        "id"
                    ],

                "Status ID":
                    "ID - OK",

                "Placa Planilha":
                    item[
                        "placa"
                    ],

                "Data":
                    item[
                        "data"
                    ],

                "Hora":
                    item[
                        "hora"
                    ],

                "Categoria":
                    item[
                        "categoria"
                    ],

                **validacao,
            }
        )

    # ========================================================
    # DATAFRAME
    # ========================================================

    df = pd.DataFrame(
        registros
    )

    if df.empty:

        print(
            "[ERRO] "
            "Nenhum registro gerado."
        )

        return None

    # ========================================================
    # MELHOR RESULTADO POR ID
    # ========================================================

    def score_linha(
        row
    ):

        status = str(
            row.get(
                "Status OCR",
                ""
            )
        )

        try:

            diferenca = float(
                row.get(
                    "Diferença OCR",
                    999
                )
            )

        except Exception:

            diferenca = 999.0

        try:

            confianca = float(
                row.get(
                    "OCR Confidence",
                    0
                )
            )

        except Exception:

            confianca = 0.0

        if status == (
            "Aprovado - busca nas imagens"
        ):

            peso_status = 0

        elif status == "Revisar":

            peso_status = 100

        elif status == (
            "Revisar - sem detecção"
        ):

            peso_status = 200

        elif status == (
            "Revisar - sem OCR"
        ):

            peso_status = 250

        elif status == (
            "Revisar - sem imagem"
        ):

            peso_status = 300

        else:

            peso_status = 400

        return (
            peso_status
            +
            diferenca
            -
            confianca
        )

    df_melhores = (
        df.copy()
    )

    df_melhores[
        "_score"
    ] = df_melhores.apply(
        score_linha,
        axis=1
    )

    df_melhores = (
        df_melhores
        .sort_values(
            by="_score"
        )
        .drop_duplicates(
            subset=[
                "ID Extraído"
            ],
            keep="first"
        )
        .drop(
            columns=[
                "_score"
            ]
        )
    )

    # ========================================================
    # SALVA EXCEL
    # ========================================================

    caminho_saida = ensure_parent(
        Path(
            caminho_saida
        )
    )

    with pd.ExcelWriter(
        caminho_saida,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Resultado Completo",
            index=False,
        )

        df_melhores.to_excel(
            writer,
            sheet_name="Melhor Resultado ID",
            index=False,
        )

    # ========================================================
    # ESTATÍSTICAS
    # ========================================================

    total_linhas = len(
        df
    )

    id_ok = int(
        (
            df[
                "Status ID"
            ]
            ==
            "ID - OK"
        ).sum()
    )

    id_nao = int(
        (
            df[
                "Status ID"
            ]
            ==
            "ID não encontrado na planilha"
        ).sum()
    )

    aprovados = int(
        (
            df[
                "Status OCR"
            ]
            ==
            "Aprovado - busca nas imagens"
        ).sum()
    )

    revisar = int(
        (
            df[
                "Status OCR"
            ]
            ==
            "Revisar"
        ).sum()
    )

    ignoradas = int(
        (
            df[
                "Status OCR"
            ]
            ==
            "Ignorada - passagem já resolvida"
        ).sum()
    )

    sem_deteccao = int(
        (
            df[
                "Status OCR"
            ]
            ==
            "Revisar - sem detecção"
        ).sum()
    )

    sem_ocr = int(
        (
            df[
                "Status OCR"
            ]
            ==
            "Revisar - sem OCR"
        ).sum()
    )

    sem_imagem = int(
        (
            df[
                "Status OCR"
            ]
            ==
            "Revisar - sem imagem"
        ).sum()
    )

    tempo_total = (
        time.perf_counter()
        -
        inicio
    )

    # ========================================================
    # RELATÓRIO
    # ========================================================

    print(
        "\n"
        + "=" * 75
    )

    print(
        "RESULTADO FINAL - TAMOIOS"
    )

    print(
        "=" * 75
    )

    print(
        f"Imagens encontradas: "
        f"{len(imagens)}"
    )

    print(
        f"Imagens com ID encontrado: "
        f"{len(itens_validos)}"
    )

    print(
        f"IDs únicos: "
        f"{total_ids}"
    )

    print(
        f"IDs resolvidos: "
        f"{len(ids_resolvidos)}"
    )

    print(
        f"IDs sem correspondência: "
        f"{len(ids_esgotados)}"
    )

    print(
        f"Imagens enviadas ao YOLO: "
        f"{total_imagens_enviadas}"
    )

    print(
        f"Detecções YOLO: "
        f"{total_yolo_detectadas}"
    )

    print(
        f"OCR executados: "
        f"{total_ocr_executados}"
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
        f"Sem detecção YOLO: "
        f"{sem_deteccao}"
    )

    print(
        f"Sem OCR: "
        f"{sem_ocr}"
    )

    print(
        f"Ignoradas após resolução: "
        f"{ignoradas}"
    )

    print(
        f"Sem imagem: "
        f"{sem_imagem}"
    )

    print(
        f"Linhas do resultado: "
        f"{total_linhas}"
    )

    print(
        f"Tempo total: "
        f"{tempo_total:.1f}s"
    )

    print(
        f"Arquivo: "
        f"{caminho_saida}"
    )

    print(
        "=" * 75
    )

    return df


# ============================================================
# MAIN
# ============================================================

def main():

    print(
        "\n"
        + "=" * 75
    )

    print(
        "VALIDADOR IMG TAMOIOS"
    )

    print(
        "=" * 75
    )

    print(
        f"[CONFIG] "
        f"Máximo para aprovação: "
        f"{MAX_DIFERENCAS_APROVACAO} diferenças"
    )

    # ========================================================
    # EXCEL
    # ========================================================

    caminho_excel = input(
        "\nCaminho do Excel: "
    ).strip().strip('"')

    if not caminho_excel:

        print(
            "[ERRO] "
            "Nenhum Excel informado."
        )

        sys.exit(1)

    caminho_entrada = Path(caminho_excel).resolve()

    if not caminho_entrada.exists():

        print(
            "[ERRO] "
            "Excel não encontrado:"
        )

        print(caminho_entrada)

        sys.exit(1)

    recursivo = "--recursivo" in sys.argv
    if caminho_entrada.is_dir():
        iterator = caminho_entrada.rglob("*.xlsx") if recursivo else caminho_entrada.glob("*.xlsx")
        arquivos_excel = sorted(
            (
                p for p in iterator
                if not p.name.startswith("~$")
                and not p.stem.endswith(("_analisado", "_validado", "_resultado"))
            ),
            key=lambda p: str(p).lower(),
        )
    elif caminho_entrada.is_file() and caminho_entrada.suffix.lower() == ".xlsx":
        arquivos_excel = [caminho_entrada]
    else:
        print("[ERRO] Informe um arquivo .xlsx ou uma pasta com arquivos Excel.")
        sys.exit(1)

    if not arquivos_excel:
        print("[ERRO] Nenhum arquivo .xlsx válido foi encontrado.")
        sys.exit(1)

    # ========================================================
    # PASTA
    # ========================================================

    pasta_imagens = input(
        "Caminho da pasta de imagens: "
    ).strip().strip('"')

    if not pasta_imagens:

        print(
            "[ERRO] "
            "Nenhuma pasta informada."
        )

        sys.exit(1)

    pasta_imagens = Path(
        pasta_imagens
    ).resolve()

    if not pasta_imagens.exists():

        print(
            "[ERRO] "
            "Pasta não encontrada:"
        )

        print(
            pasta_imagens
        )

        sys.exit(1)

    if not pasta_imagens.is_dir():

        print(
            "[ERRO] "
            "O caminho informado "
            "não é uma pasta."
        )

        sys.exit(1)

    # ========================================================
    # VERIFICA AMBIENTES
    # ========================================================

    arquivos_obrigatorios = [

        PYTHON_YOLO,
        PYTHON_OCR,
        DETECTOR_PY,
        OCR_PY,
    ]

    print(
        "\n[CHECK] Verificando arquivos..."
    )

    for caminho in (
        arquivos_obrigatorios
    ):

        caminho = Path(
            caminho
        )

        if not caminho.exists():

            print(
                "[ERRO] Arquivo não encontrado:"
            )

            print(
                caminho
            )

            sys.exit(1)

        print(
            f"[OK] {caminho}"
        )

    # ========================================================
    # SAÍDA
    # ========================================================

    # ========================================================
    # INICIA YOLO
    # ========================================================

    processo_yolo = iniciar_yolo()

    # ========================================================
    # INICIA OCR
    # ========================================================

    processo_ocr = iniciar_ocr()

    try:

        print(
            "\n[INIT] "
            "Aguardando carregamento dos modelos..."
        )

        time.sleep(
            15
        )

        # ====================================================
        # PROCESSAMENTO
        # ====================================================

        resultados_completos = []
        melhores_resultados = []
        erros = []

        with tempfile.TemporaryDirectory(prefix="tamoios_isencao_lote_") as pasta_temp:
            for indice, arquivo_excel in enumerate(arquivos_excel, start=1):
                print(f"[PROGRESSO] Excel {indice}/{len(arquivos_excel)}: {arquivo_excel.name}")
                saida_temporaria = Path(pasta_temp) / f"resultado_{indice}.xlsx"
                try:
                    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                        df = processar_excel(
                            arquivo_excel,
                            pasta_imagens,
                            saida_temporaria,
                            processo_yolo,
                            processo_ocr,
                        )
                    df_melhores = pd.read_excel(
                        saida_temporaria,
                        sheet_name="Melhor Resultado ID",
                    )
                    origem = (
                        str(arquivo_excel.relative_to(caminho_entrada))
                        if caminho_entrada.is_dir()
                        else arquivo_excel.name
                    )
                    df.insert(0, "Arquivo de Origem", origem)
                    df_melhores.insert(0, "Arquivo de Origem", origem)
                    resultados_completos.append(df)
                    melhores_resultados.append(df_melhores)
                except Exception as erro:
                    erros.append((arquivo_excel.name, f"{type(erro).__name__}: {erro}"))

        caminho_saida = None
        if resultados_completos:
            caminho_saida = caminho_resultado("tamoios_isencao", "resultado_tamoios.xlsx")
            df_completo = pd.concat(resultados_completos, ignore_index=True)
            df_melhores = pd.concat(melhores_resultados, ignore_index=True)
            with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
                df_completo.to_excel(writer, sheet_name="Resultado Completo", index=False)
                df_melhores.to_excel(writer, sheet_name="Melhor Resultado ID", index=False)

            aprovados = int((df_completo.get("Status OCR", pd.Series(dtype=str)) == "Aprovado - busca nas imagens").sum())
            revisar = int((df_completo.get("Status OCR", pd.Series(dtype=str)) == "Revisar").sum())
            print("\n" + "=" * 75)
            print("RESULTADO FINAL - TAMOIOS ISENÇÃO")
            print("=" * 75)
            print(f"Planilhas processadas: {len(resultados_completos)} de {len(arquivos_excel)}")
            print(f"Registros: {len(df_completo)}")
            print(f"Aprovados: {aprovados}")
            print(f"Revisar: {revisar}")
            print(f"Excel consolidado: {caminho_saida}")
        else:
            print("\n[ERRO] Nenhuma planilha foi processada com sucesso; não foi criado Excel.")

        if erros:
            print("\nArquivos com erro:")
            for nome, mensagem in erros:
                print(f"- {nome}: {mensagem}")

    except KeyboardInterrupt:

        print(
            "\n[INFO] "
            "Processamento interrompido."
        )

    except Exception as e:

        print(
            "\n[ERRO]"
        )

        print(
            f"{type(e).__name__}: {e}"
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


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    main()
