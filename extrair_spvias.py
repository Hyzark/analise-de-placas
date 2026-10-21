from __future__ import annotations
from placas import (
    aplicar_mascara as _aplicar_mascara_comum,
    descobrir_mascara as _descobrir_mascara_comum,
    distancia_placas as _distancia_placas_comum,
    distancia_ponderada as _distancia_ponderada,
    normalizar_ocr as _normalizar_ocr_comum,
    normalizar_placa as _normalizar_placa,
    CONFUNDIVEIS as _CONFUNDIVEIS,
    MAPA_LETRAS as _MAPA_LETRAS,
    MAPA_DIGITOS as _MAPA_DIGITOS,
)

from common import escrever_json


# ============================================================
# CONTROLE DE THREADS
# ============================================================

import os

os.environ["OMP_NUM_THREADS"] = "1"
os.environ["OPENBLAS_NUM_THREADS"] = "1"
os.environ["MKL_NUM_THREADS"] = "1"
os.environ["VECLIB_MAXIMUM_THREADS"] = "1"
os.environ["NUMEXPR_NUM_THREADS"] = "1"


# ============================================================
# IMPORTS
# ============================================================

import functools
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET

from concurrent.futures import (
    ThreadPoolExecutor,
    as_completed,
)

from io import BytesIO
from pathlib import Path

import cv2
import numpy as np
import pandas as pd

from PIL import Image
from openpyxl import load_workbook
from tqdm import tqdm


# ============================================================
# CONFIG
# ============================================================

from config import (
    BASE_DIR,
    caminho_resultado,
    ensure_parent,
)


# ============================================================
# CAMINHOS
# ============================================================

PYTHON_YOLO = (
    BASE_DIR
    / ".venv_yolo"
    / "Scripts"
    / "python.exe"
)

PYTHON_OCR = (
    BASE_DIR
    / ".venv_ocr"
    / "Scripts"
    / "python.exe"
)

DETECTOR_PY = (
    BASE_DIR
    / "detector_process.py"
)

OCR_PY = (
    BASE_DIR
    / "ocr_process.py"
)


# ============================================================
# CONFIGURAÇÕES
# ============================================================

BATCH_SIZE = 64

MAX_DIFERENCAS_APROVACAO = 3

NOME_ABA = "Envio DER"

LOGO_FILENAME = "image1.png"

EXTENSOES_XLSX = {
    ".xlsx",
}

EXTENSOES_RASTER = {
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
}

EXTENSOES_VETORIAIS = {
    ".emf",
    ".wmf",
}

IMAGEM_VECTOR_DPI = 300

TAMANHO_PLACA = 7


# ============================================================
# XML NAMESPACES
# ============================================================

_NS = {
    "xdr":
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",

    "a":
        "http://schemas.openxmlformats.org/drawingml/2006/main",

    "r":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",

    "rel":
        "http://schemas.openxmlformats.org/package/2006/relationships",

    "main":
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",

    "dgm":
        "http://schemas.openxmlformats.org/drawingml/2006/diagram",
}


# ============================================================
# CARACTERES CONFUNDÍVEIS
# ============================================================



# ============================================================
# MAPA DE LETRAS
# ============================================================



# ============================================================
# MAPA DE DÍGITOS
# ============================================================



# ============================================================
# LEVENSHTEIN
# ============================================================

def _levenshtein(a: str, b: str) -> float:
    if not a or not b:
        return 999.0
    return float(_distancia_ponderada(a, b, _CONFUNDIVEIS))


# ============================================================
# LIMPA TEXTO
# ============================================================

def _limpar_texto(texto) -> str:
    return _normalizar_placa(texto)


# ============================================================
# DISTÂNCIA
# ============================================================

def distancia_placas(placa_a: str, placa_b: str) -> float:
    return _distancia_placas_comum(placa_a, placa_b, _CONFUNDIVEIS)


# ============================================================
# DESCREVE DIFERENÇAS DE CARACTERES
# ============================================================

def _descrever_diferencas(
    esperado,
    recebido,
):

    esperado = _limpar_texto(
        esperado
    )

    recebido = _limpar_texto(
        recebido
    )

    if not esperado or not recebido:
        return "Sem comparação"

    linhas = (
        len(esperado)
        + 1
    )

    colunas = (
        len(recebido)
        + 1
    )

    matriz = [
        [0.0] * colunas
        for _ in range(
            linhas
        )
    ]

    operacoes = [
        [None] * colunas
        for _ in range(
            linhas
        )
    ]

    for i in range(linhas):

        matriz[i][0] = float(i)

    for j in range(colunas):

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

    i = len(esperado)
    j = len(recebido)

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

            operacao = operacoes[i][j]
            tipo = operacao[0]

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
# DESCOBRE MÁSCARA DA PLACA
# ============================================================

def descobrir_mascara(placa):
    return _descobrir_mascara_comum(placa, tamanho=TAMANHO_PLACA, mascara_invalida=None)


# ============================================================
# APLICA MÁSCARA
# ============================================================

def corrigir_mascara(texto, mascara):
    return _aplicar_mascara_comum(texto, mascara, _MAPA_LETRAS, _MAPA_DIGITOS, tamanho=TAMANHO_PLACA, comprimento_invalido="vazio", preservar_indice=4)


# ============================================================
# NORMALIZA OCR DE MOTOS
#
# Preserva a lógica antiga do SPVIAS.
# ============================================================

def normalizar_ocr(texto):
    return _normalizar_ocr_comum(texto, reordenar_formatos=True)


# ============================================================
# ENCONTRA MELHOR TRECHO DE 7
#
# Exemplo:
#
# OCR = DI2EYP7D88
#
# testa:
#
# DI2EYP7
# I2EYP7D
# 2EYP7D8
# EYP7D88  ← melhor
# YP7D88
#
# Depois aplica máscara.
# ============================================================

def encontrar_melhor_trecho(
    placa_esperada,
    texto_ocr,
    mascara=None,
):

    esperada = _limpar_texto(
        placa_esperada
    )

    ocr = normalizar_ocr(
        texto_ocr
    )

    if not esperada or not ocr:

        return {
            "candidato_bruto": "",
            "candidato_limpo": "",
            "distancia": 999.0,
            "inicio": None,
        }

    if mascara is None:

        mascara = descobrir_mascara(
            esperada
        )

    # ========================================================
    # PLACA SEM 7 CARACTERES
    # ========================================================

    if len(esperada) != 7:

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

    if len(ocr) < 7:

        candidato = (
            corrigir_mascara(
                ocr,
                mascara,
            )
            if len(ocr) == 7
            else ocr
        )

        if not candidato:
            candidato = ocr

        return {
            "candidato_bruto":
                ocr,

            "candidato_limpo":
                candidato,

            "distancia":
                distancia_placas(
                    esperada,
                    candidato,
                ),

            "inicio":
                0,
        }

    # ========================================================
    # OCR TEM EXATAMENTE 7
    # ========================================================

    if len(ocr) == 7:

        candidato_limpo = (
            corrigir_mascara(
                ocr,
                mascara,
            )
        )

        if not candidato_limpo:

            candidato_limpo = ocr

        distancia = distancia_placas(
            esperada,
            candidato_limpo,
        )

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
            corrigir_mascara(
                candidato_bruto,
                mascara,
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
            melhor[
                "chave"
            ]
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
# EXTRAÇÃO INTELIGENTE
# ============================================================

def extrair_placa_ocr(
    placa_esperada,
    texto_ocr,
    mascara=None,
):

    texto_original = str(
        texto_ocr
        or
        ""
    ).strip()

    if not texto_original:

        return {
            "ocr_bruto":
                texto_original,

            "placa_ocr":
                "",

            "candidato":
                "",

            "diferenca":
                999.0,

            "inicio":
                None,
        }

    resultado = (
        encontrar_melhor_trecho(
            placa_esperada,
            texto_original,
            mascara,
        )
    )

    return {
        "ocr_bruto":
            texto_original,

        "placa_ocr":
            resultado.get(
                "candidato_limpo",
                "",
            ),

        "candidato":
            resultado.get(
                "candidato_bruto",
                "",
            ),

        "diferenca":
            float(
                resultado.get(
                    "distancia",
                    999.0,
                )
            ),

        "inicio":
            resultado.get(
                "inicio"
            ),
    }


# ============================================================
# DECODIFICA RASTER
# ============================================================

def _decodificar_raster(
    img_bytes: bytes,
):

    if not img_bytes:
        return None

    try:

        array = np.frombuffer(
            img_bytes,
            dtype=np.uint8,
        )

        imagem = cv2.imdecode(
            array,
            cv2.IMREAD_COLOR,
        )

        if (
            imagem is not None
            and
            imagem.size > 0
        ):

            return imagem

    except Exception:

        pass

    try:

        pil_img = (
            Image
            .open(
                BytesIO(
                    img_bytes
                )
            )
            .convert(
                "RGB"
            )
        )

        return cv2.cvtColor(
            np.array(
                pil_img
            ),
            cv2.COLOR_RGB2BGR,
        )

    except Exception:

        return None


# ============================================================
# LOCALIZA IMAGEMAGICK
# ============================================================

def _encontrar_imagemagick():

    candidatos = [
        "magick",
        "magick.exe",
        r"C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe",
        r"C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe",
        r"C:\Program Files\ImageMagick-7.0.11-Q16\magick.exe",
    ]

    for candidato in candidatos:

        try:

            resultado = subprocess.run(
                [
                    candidato,
                    "-version",
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=10,
            )

            if resultado.returncode == 0:

                return candidato

        except Exception:

            continue

    return None


# ============================================================
# CONVERSÃO EMF/WMF → PNG
# ============================================================

def _converter_vetor_para_png(
    img_bytes,
    extensao,
    temp_dir,
    identificador,
):

    if not img_bytes:
        return None

    extensao = (
        extensao
        or
        ""
    ).lower()

    if (
        extensao
        not in
        EXTENSOES_VETORIAIS
    ):

        return None

    arquivo_entrada = (
        temp_dir
        /
        f"{identificador}{extensao}"
    )

    arquivo_saida = (
        temp_dir
        /
        f"{identificador}.png"
    )

    try:

        arquivo_entrada.write_bytes(
            img_bytes
        )

    except Exception as e:

        print(
            f"[VECTOR] "
            f"Erro gravando "
            f"{arquivo_entrada.name}: "
            f"{e}",
            flush=True,
        )

        return None

    # ========================================================
    # IMAGEMAGICK
    # ========================================================

    magick = (
        _encontrar_imagemagick()
    )

    if magick:

        try:

            comando = [
                magick,

                "-density",
                str(
                    IMAGEM_VECTOR_DPI
                ),

                str(
                    arquivo_entrada
                ),

                "-background",
                "white",

                "-alpha",
                "remove",

                "-alpha",
                "off",

                str(
                    arquivo_saida
                ),
            ]

            resultado = subprocess.run(
                comando,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=60,
            )

            if (
                resultado.returncode == 0
                and
                arquivo_saida.exists()
            ):

                png_bytes = (
                    arquivo_saida.read_bytes()
                )

                if png_bytes:

                    return {
                        "bytes":
                            png_bytes,

                        "nome":
                            (
                                arquivo_entrada.stem
                                +
                                ".png"
                            ),

                        "ext":
                            ".png",

                        "convertida":
                            True,

                        "formato_original":
                            extensao,
                    }

            erro = (
                resultado.stderr
                or
                resultado.stdout
                or
                ""
            )

            if erro:

                print(
                    f"[VECTOR] "
                    f"ImageMagick falhou em "
                    f"{arquivo_entrada.name}: "
                    f"{erro[:300]}",
                    flush=True,
                )

        except Exception as e:

            print(
                f"[VECTOR] "
                f"Erro ImageMagick em "
                f"{arquivo_entrada.name}: "
                f"{e}",
                flush=True,
            )

    # ========================================================
    # FALLBACK PILLOW
    # ========================================================

    try:

        with Image.open(
            arquivo_entrada
        ) as img:

            if img.mode not in (
                "RGB",
                "RGBA",
            ):

                img = img.convert(
                    "RGB"
                )

            buffer = BytesIO()

            img.save(
                buffer,
                format="PNG",
            )

            png_bytes = (
                buffer.getvalue()
            )

            if png_bytes:

                return {
                    "bytes":
                        png_bytes,

                    "nome":
                        (
                            arquivo_entrada.stem
                            +
                            ".png"
                        ),

                    "ext":
                        ".png",

                    "convertida":
                        True,

                    "formato_original":
                        extensao,
                }

    except Exception:

        pass

    return None


# ============================================================
# COMUNICAÇÃO
# ============================================================



def enviar_json(
    processo,
    dados,
):
    if processo is None:
        raise RuntimeError("Processo inexistente.")
    if processo.poll() is not None:
        raise RuntimeError(
            "Processo já encerrou. "
            f"Código={processo.returncode}"
        )
    if processo.stdin is None:
        raise RuntimeError("stdin indisponível.")
    escrever_json(processo, dados)


def ler_json(
    processo,
    nome,
):

    if processo is None:

        raise RuntimeError(
            f"Processo {nome} inexistente."
        )

    if processo.stdout is None:

        raise RuntimeError(
            f"stdout de {nome} indisponível."
        )

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
            f"Resposta inválida de {nome}: "
            f"{linha[:500]}"
        ) from e


# ============================================================
# INICIA PROCESSO
# ============================================================

def iniciar_processo(
    python_path,
    script_path,
    nome,
):

    python_path = Path(
        python_path
    )

    script_path = Path(
        script_path
    )

    if not python_path.exists():

        raise FileNotFoundError(
            f"Python de {nome} não encontrado:\n"
            f"{python_path}"
        )

    if not script_path.exists():

        raise FileNotFoundError(
            f"Script de {nome} não encontrado:\n"
            f"{script_path}"
        )

    print(
        f"[INIT] Iniciando {nome}..."
    )

    return subprocess.Popen(
        [
            str(python_path),
            str(script_path),
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
        bufsize=1,
        cwd=str(BASE_DIR),
    )


# ============================================================
# ENCERRA PROCESSO
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

        return

    except Exception:

        pass

    try:

        processo.terminate()

        processo.wait(
            timeout=5
        )

        return

    except Exception:

        pass

    try:

        processo.kill()

    except Exception:

        pass


# ============================================================
# CLASSE
# ============================================================

class ValidadorSPVias:

    def __init__(
        self,
        batch_size=BATCH_SIZE,
    ):

        self.batch_size = (
            batch_size
        )

        self.num_workers_cpu = max(
            1,
            os.cpu_count() or 4,
        )

        self.metrics = {
            "tempo_total":
                0.0,

            "tempo_leitura_excel":
                0.0,

            "tempo_decodificacao":
                0.0,

            "tempo_yolo":
                0.0,

            "qtd_imagens_yolo":
                0,

            "tempo_ocr":
                0.0,

            "qtd_chamadas_ocr":
                0,

            "tempo_avaliacoes":
                0.0,

            "total_blocos":
                0,

            "vetores_convertidos":
                0,

            "vetores_falharam":
                0,

            "rodadas_processadas":
                0,
        }

    # ========================================================
    # HASH
    # ========================================================

    def _hash_bytes(
        self,
        img_bytes,
    ):

        return hashlib.md5(
            img_bytes
        ).hexdigest()

    # ========================================================
    # NORMALIZA TEXTO
    # ========================================================

    def _limpar_texto(
        self,
        texto,
    ):

        return _limpar_texto(
            texto
        )

    # ========================================================
    # NORMALIZA OCR
    #
    # Mantém tratamento de motos.
    # ========================================================

    def _normalizar_ocr(
        self,
        texto,
    ):

        return normalizar_ocr(
            texto
        )

    # ========================================================
    # DESCOBRE MÁSCARA
    # ========================================================

    def _descobrir_mascara(
        self,
        placa,
    ):

        return descobrir_mascara(
            placa
        )

    # ========================================================
    # CORRIGE MÁSCARA
    # ========================================================

    def _corrigir_mascara(
        self,
        texto,
        mascara,
    ):

        return corrigir_mascara(
            texto,
            mascara,
        )

    # ========================================================
    # EXTRAÇÃO INTELIGENTE
    # ========================================================

    def _extrair_placa_ocr(
        self,
        placa_esperada,
        texto_ocr,
        mascara=None,
    ):

        return extrair_placa_ocr(
            placa_esperada,
            texto_ocr,
            mascara,
        )

    # ========================================================
    # NORMALIZA PLACA
    #
    # Compatibilidade com chamadas existentes.
    #
    # IMPORTANTE:
    # Quando recebe uma placa esperada + um texto OCR,
    # tenta recuperar a placa dentro do texto.
    # ========================================================

    def _normalizar_placa(
        self,
        texto,
        mascara=None,
        placa_esperada=None,
    ):

        texto = self._normalizar_ocr(
            texto
        )

        if not texto:
            return ""

        # ====================================================
        # Se houver placa esperada, faz busca contextual.
        # ====================================================

        if placa_esperada:

            resultado = (
                self._extrair_placa_ocr(
                    placa_esperada,
                    texto,
                    mascara,
                )
            )

            if resultado[
                "placa_ocr"
            ]:

                return resultado[
                    "placa_ocr"
                ]

        # ====================================================
        # Sem placa esperada:
        # comportamento de normalização local.
        # ====================================================

        if len(texto) != 7:

            return texto

        if mascara:

            corrigido = (
                self._corrigir_mascara(
                    texto,
                    mascara,
                )
            )

            if corrigido:

                return corrigido

        return texto

    # ========================================================
    # DISTÂNCIA
    # ========================================================

    def _distancia(
        self,
        placa1,
        placa2,
    ):

        a = self._limpar_texto(
            placa1
        )

        b = self._limpar_texto(
            placa2
        )

        if not a or not b:
            return 999.0

        return _levenshtein(
            a,
            b,
        )

    # ========================================================
    # VERIFICA OCR
    #
    # Agora usa a MESMA extração inteligente.
    # ========================================================

    def _ocr_satisfaz_passagem(
        self,
        placa_esperada,
        resultado,
    ):

        if not resultado:
            return False

        texto = resultado.get(
            "ocr_best",
            "",
        )

        if not texto:
            return False

        mascara = (
            self._descobrir_mascara(
                placa_esperada
            )
        )

        extracao = (
            self._extrair_placa_ocr(
                placa_esperada,
                texto,
                mascara,
            )
        )

        placa_ocr = (
            extracao[
                "placa_ocr"
            ]
        )

        if not placa_ocr:
            return False

        distancia = (
            extracao[
                "diferenca"
            ]
        )

        return (
            distancia
            <=
            MAX_DIFERENCAS_APROVACAO
        )

    # ========================================================
    # LÊ PASSAGENS
    # ========================================================

    def _ler_blocos(
        self,
        ws,
    ):

        inicio = (
            time.perf_counter()
        )

        blocos = []

        for row_idx, row in enumerate(
            ws.iter_rows(
                values_only=True
            ),
            start=1,
        ):

            if not row:
                continue

            col_a = (
                row[0]
                if len(row) > 0
                else None
            )

            if not (
                isinstance(
                    col_a,
                    int,
                )
                and
                col_a > 0
            ):

                continue

            def valor(
                indice,
            ):

                if indice >= len(
                    row
                ):

                    return ""

                return row[
                    indice
                ]

            blocos.append(
                {
                    "seq":
                        col_a,

                    "linha":
                        row_idx,

                    "id":
                        str(
                            valor(1)
                            or
                            ""
                        ).strip(),

                    "data":
                        valor(2),

                    "hora":
                        valor(3),

                    "sp":
                        valor(4),

                    "km":
                        str(
                            valor(5)
                            or
                            ""
                        ),

                    "praca":
                        str(
                            valor(6)
                            or
                            ""
                        ),

                    "pista":
                        valor(7),

                    "sentido":
                        str(
                            valor(8)
                            or
                            ""
                        ),

                    "municipio":
                        str(
                            valor(9)
                            or
                            ""
                        ),

                    "codigo":
                        valor(10),

                    "placa":
                        str(
                            valor(11)
                            or
                            ""
                        ).strip().upper(),

                    "modelo":
                        str(
                            valor(12)
                            or
                            ""
                        ),

                    "categoria":
                        str(
                            valor(13)
                            or
                            ""
                        ),

                    "eixos":
                        valor(14),

                    "tarifa":
                        valor(15),

                    "responsavel":
                        str(
                            valor(16)
                            or
                            ""
                        ),

                    "linha_fim":
                        None,
                }
            )

        for i, bloco in enumerate(
            blocos
        ):

            if (
                i + 1
                <
                len(blocos)
            ):

                bloco[
                    "linha_fim"
                ] = (
                    blocos[
                        i + 1
                    ][
                        "linha"
                    ]
                )

            else:

                bloco[
                    "linha_fim"
                ] = 999999

        self.metrics[
            "tempo_leitura_excel"
        ] += (
            time.perf_counter()
            - inicio
        )

        return blocos

    # ========================================================
    # MAPEIA IMAGENS XML
    # ========================================================

    def _mapear_imagens_xml(
        self,
        caminho_excel,
        nome_planilha,
    ):

        inicio = (
            time.perf_counter()
        )

        mapa = {}

        try:

            with zipfile.ZipFile(
                caminho_excel,
                "r",
            ) as z:

                nomes = set(
                    z.namelist()
                )

                workbook_xml = (
                    "xl/workbook.xml"
                )

                if workbook_xml not in nomes:
                    return mapa

                workbook_root = (
                    ET.fromstring(
                        z.read(
                            workbook_xml
                        )
                    )
                )

                sheet_rid = None

                for sheet in workbook_root.findall(
                    ".//main:sheet",
                    _NS,
                ):

                    if (
                        sheet.get(
                            "name"
                        )
                        ==
                        nome_planilha
                    ):

                        sheet_rid = (
                            sheet.get(
                                f"{{{_NS['r']}}}id"
                            )
                        )

                        break

                if not sheet_rid:
                    return mapa

                # ==========================================
                # RELAÇÕES DO WORKBOOK
                # ==========================================

                wb_rels_path = (
                    "xl/_rels/workbook.xml.rels"
                )

                if wb_rels_path not in nomes:
                    return mapa

                wb_rels_root = ET.fromstring(
                    z.read(
                        wb_rels_path
                    )
                )

                sheet_target = None

                for rel in wb_rels_root.findall(
                    f"{{{_NS['rel']}}}Relationship"
                ):

                    if (
                        rel.get(
                            "Id"
                        )
                        ==
                        sheet_rid
                    ):

                        sheet_target = (
                            rel.get(
                                "Target",
                                "",
                            )
                        )

                        break

                if not sheet_target:
                    return mapa

                sheet_target = (
                    sheet_target
                    .replace(
                        "\\",
                        "/",
                    )
                    .lstrip("/")
                )

                if not sheet_target.startswith(
                    "xl/"
                ):

                    sheet_target = (
                        "xl/worksheets/"
                        +
                        Path(
                            sheet_target
                        ).name
                    )

                sheet_filename = (
                    Path(
                        sheet_target
                    ).name
                )

                # ==========================================
                # RELAÇÕES DA PLANILHA
                # ==========================================

                sheet_rels = (
                    "xl/worksheets/_rels/"
                    +
                    f"{sheet_filename}.rels"
                )

                if sheet_rels not in nomes:
                    return mapa

                sheet_rels_root = ET.fromstring(
                    z.read(
                        sheet_rels
                    )
                )

                drawing_target = None

                for rel in sheet_rels_root.findall(
                    f"{{{_NS['rel']}}}Relationship"
                ):

                    tipo = (
                        rel.get(
                            "Type",
                            "",
                        )
                        .lower()
                    )

                    if "drawing" in tipo:

                        drawing_target = (
                            rel.get(
                                "Target",
                                "",
                            )
                        )

                        break

                if not drawing_target:
                    return mapa

                drawing_filename = (
                    Path(
                        drawing_target.replace(
                            "\\",
                            "/",
                        )
                    ).name
                )

                drawing_path = (
                    "xl/drawings/"
                    +
                    drawing_filename
                )

                drawing_rels_path = (
                    "xl/drawings/_rels/"
                    +
                    drawing_filename
                    +
                    ".rels"
                )

                if (
                    drawing_path not in nomes
                    or
                    drawing_rels_path not in nomes
                ):

                    return mapa

                # ==========================================
                # RELAÇÕES DRAWING
                # ==========================================

                drawing_rels_root = ET.fromstring(
                    z.read(
                        drawing_rels_path
                    )
                )

                rid_target = {}

                for rel in drawing_rels_root.findall(
                    f"{{{_NS['rel']}}}Relationship"
                ):

                    rid_target[
                        rel.get(
                            "Id"
                        )
                    ] = rel.get(
                        "Target",
                        "",
                    )

                # ==========================================
                # MEDIA
                # ==========================================

                media = {}

                for nome in nomes:

                    if not nome.startswith(
                        "xl/media/"
                    ):

                        continue

                    nome_base = Path(
                        nome
                    ).name

                    try:

                        media[
                            nome_base
                        ] = z.read(
                            nome
                        )

                    except Exception:

                        pass

                # ==========================================
                # DRAWING
                # ==========================================

                drawing_root = ET.fromstring(
                    z.read(
                        drawing_path
                    )
                )

                # ==========================================
                # ÂNCORAS
                # ==========================================

                for anchor in drawing_root:

                    tipo_anchor = (
                        anchor.tag
                        .split(
                            "}"
                        )[-1]
                    )

                    if tipo_anchor not in (
                        "twoCellAnchor",
                        "oneCellAnchor",
                    ):

                        continue

                    from_elem = anchor.find(
                        "xdr:from",
                        _NS,
                    )

                    if from_elem is None:
                        continue

                    row_elem = from_elem.find(
                        "xdr:row",
                        _NS,
                    )

                    if row_elem is None:
                        continue

                    try:

                        row_inicio = (
                            int(
                                row_elem.text
                            )
                            + 1
                        )

                    except Exception:

                        continue

                    # ======================================
                    # PIC
                    # ======================================

                    pic = anchor.find(
                        "xdr:pic",
                        _NS,
                    )

                    if pic is not None:

                        blip = pic.find(
                            ".//a:blip",
                            _NS,
                        )

                        if blip is None:
                            continue

                        rid = blip.get(
                            f"{{{_NS['r']}}}embed",
                            "",
                        )

                        target = (
                            rid_target.get(
                                rid,
                                "",
                            )
                        )

                        nome_midia = Path(
                            target.replace(
                                "\\",
                                "/",
                            )
                        ).name

                        if not nome_midia:
                            continue

                        if (
                            nome_midia
                            ==
                            LOGO_FILENAME
                        ):

                            continue

                        ext = (
                            Path(
                                nome_midia
                            )
                            .suffix
                            .lower()
                        )

                        if (
                            ext
                            not in
                            (
                                EXTENSOES_RASTER
                                |
                                EXTENSOES_VETORIAIS
                            )
                        ):

                            continue

                        img_bytes = media.get(
                            nome_midia
                        )

                        if not img_bytes:
                            continue

                        mapa.setdefault(
                            row_inicio,
                            []
                        ).append(
                            {
                                "bytes":
                                    img_bytes,

                                "nome":
                                    nome_midia,

                                "ext":
                                    ext,

                                "convertida":
                                    False,

                                "formato_original":
                                    ext,
                            }
                        )

                        continue

                    # ======================================
                    # GRAPHIC FRAME
                    # ======================================

                    graphic = (
                        anchor.find(
                            "xdr:graphicFrame",
                            _NS,
                        )
                    )

                    if graphic is None:
                        continue

                    rel_ids = (
                        graphic.find(
                            ".//dgm:relIds",
                            _NS,
                        )
                    )

                    if rel_ids is None:
                        continue

                    dm_rid = rel_ids.get(
                        f"{{{_NS['r']}}}dm",
                        "",
                    )

                    if not dm_rid:
                        continue

                    data_target = (
                        rid_target.get(
                            dm_rid,
                            "",
                        )
                    )

                    data_filename = Path(
                        data_target.replace(
                            "\\",
                            "/",
                        )
                    ).name

                    data_num = (
                        data_filename
                        .replace(
                            "data",
                            "",
                        )
                        .replace(
                            ".xml",
                            "",
                        )
                    )

                    diag_rels_path = (
                        "xl/diagrams/_rels/"
                        f"drawing{data_num}.xml.rels"
                    )

                    if (
                        diag_rels_path
                        not in nomes
                    ):

                        continue

                    diag_root = ET.fromstring(
                        z.read(
                            diag_rels_path
                        )
                    )

                    for rel in diag_root.findall(
                        f"{{{_NS['rel']}}}Relationship"
                    ):

                        tipo = (
                            rel.get(
                                "Type",
                                "",
                            )
                            .lower()
                        )

                        if "image" not in tipo:
                            continue

                        nome_midia = Path(
                            rel.get(
                                "Target",
                                "",
                            ).replace(
                                "\\",
                                "/",
                            )
                        ).name

                        if not nome_midia:
                            continue

                        if (
                            nome_midia
                            ==
                            LOGO_FILENAME
                        ):

                            continue

                        ext = (
                            Path(
                                nome_midia
                            )
                            .suffix
                            .lower()
                        )

                        if (
                            ext
                            not in
                            (
                                EXTENSOES_RASTER
                                |
                                EXTENSOES_VETORIAIS
                            )
                        ):

                            continue

                        img_bytes = media.get(
                            nome_midia
                        )

                        if img_bytes:

                            mapa.setdefault(
                                row_inicio,
                                []
                            ).append(
                                {
                                    "bytes":
                                        img_bytes,

                                    "nome":
                                        nome_midia,

                                    "ext":
                                        ext,

                                    "convertida":
                                        False,

                                    "formato_original":
                                        ext,
                                }
                            )

                        break

        except Exception as e:

            print(
                "[XML] "
                f"{type(e).__name__}: {e}",
                flush=True,
            )

        self.metrics[
            "tempo_leitura_excel"
        ] += (
            time.perf_counter()
            - inicio
        )

        return mapa

    # ========================================================
    # PROCESSA LOTE
    # ========================================================

    def _processar_lote(
        self,
        lote,
        nome_arquivo,
        processo_yolo,
        processo_ocr,
        lote_idx,
        temp_dir,
    ):

        registros = []

        # ====================================================
        # MAPA DE IMAGENS
        # ====================================================

        mapa_imagens = {}

        for bloco in lote:

            for info in bloco.get(
                "_imagens_bytes",
                [],
            ):

                if not isinstance(
                    info,
                    dict,
                ):

                    continue

                img_bytes = info.get(
                    "bytes"
                )

                if not img_bytes:
                    continue

                h = (
                    self._hash_bytes(
                        img_bytes
                    )
                )

                if h not in mapa_imagens:

                    mapa_imagens[
                        h
                    ] = info.copy()

        # ====================================================
        # CONVERSÃO DE VETORES
        # ====================================================

        for h, info in list(
            mapa_imagens.items()
        ):

            ext = (
                info.get(
                    "ext",
                    ""
                )
                .lower()
            )

            if ext in EXTENSOES_VETORIAIS:

                convertido = (
                    _converter_vetor_para_png(
                        info[
                            "bytes"
                        ],
                        ext,
                        temp_dir,
                        f"vector_{h[:12]}",
                    )
                )

                if convertido:

                    mapa_imagens[h] = convertido

                    self.metrics[
                        "vetores_convertidos"
                    ] += 1

                else:

                    self.metrics[
                        "vetores_falharam"
                    ] += 1

        # ====================================================
        # DECODIFICAÇÃO
        # ====================================================

        pares_cv2 = []

        inicio_dec = (
            time.perf_counter()
        )

        def decodificar_hash(
            img_hash
        ):

            info = (
                mapa_imagens[
                    img_hash
                ]
            )

            return (
                img_hash,
                _decodificar_raster(
                    info[
                        "bytes"
                    ]
                ),
            )

        hashes = list(
            mapa_imagens.keys()
        )

        if hashes:

            with ThreadPoolExecutor(
                max_workers=
                    self.num_workers_cpu
            ) as executor:

                futures = [
                    executor.submit(
                        decodificar_hash,
                        h,
                    )
                    for h in hashes
                ]

                for future in as_completed(
                    futures
                ):

                    try:

                        h, imagem = (
                            future.result()
                        )

                    except Exception:

                        continue

                    if (
                        imagem is not None
                        and
                        imagem.size > 0
                    ):

                        pares_cv2.append(
                            (
                                h,
                                imagem,
                            )
                        )

        self.metrics[
            "tempo_decodificacao"
        ] += (
            time.perf_counter()
            - inicio_dec
        )

        # ====================================================
        # TASK IDS
        # ====================================================

        hash_to_task = {}
        hash_to_path = {}

        for idx, (
            img_hash,
            imagem_cv2,
        ) in enumerate(
            pares_cv2
        ):

            info = (
                mapa_imagens[
                    img_hash
                ]
            )

            dados = info[
                "bytes"
            ]

            ext = (
                info.get(
                    "ext",
                    ".png",
                )
                .lower()
            )

            if ext not in EXTENSOES_RASTER:

                ext = ".png"

            task_id = (
                f"l{lote_idx}"
                f"_img{idx}"
            )

            arquivo = (
                temp_dir
                /
                f"{task_id}{ext}"
            )

            try:

                arquivo.write_bytes(
                    dados
                )

            except Exception as e:

                print(
                    f"[ARQUIVO] "
                    f"Erro salvando "
                    f"{arquivo}: "
                    f"{e}"
                )

                continue

            hash_to_task[
                img_hash
            ] = task_id

            hash_to_path[
                img_hash
            ] = str(
                arquivo
            )

        # ====================================================
        # ORDENA IMAGENS POR PASSAGEM
        # ====================================================

        imagens_por_bloco = {}

        for bloco_idx, bloco in enumerate(
            lote
        ):

            imagens_bloco = (
                bloco.get(
                    "_imagens_bytes",
                    [],
                )
            )

            lista_hashes = []

            hashes_vistos = set()

            for info in imagens_bloco:

                if not isinstance(
                    info,
                    dict,
                ):

                    continue

                img_bytes = info.get(
                    "bytes"
                )

                if not img_bytes:
                    continue

                img_hash = (
                    self._hash_bytes(
                        img_bytes
                    )
                )

                if img_hash in hashes_vistos:
                    continue

                hashes_vistos.add(
                    img_hash
                )

                if img_hash not in hash_to_task:
                    continue

                lista_hashes.append(
                    img_hash
                )

            imagens_por_bloco[
                bloco_idx
            ] = lista_hashes

        # ====================================================
        # RESULTADOS OCR POR PASSAGEM
        # ====================================================

        resultados_por_bloco = {
            bloco_idx: {}
            for bloco_idx in range(
                len(lote)
            )
        }

        # ====================================================
        # PENDENTES
        # ====================================================

        pendentes = set(
            range(
                len(lote)
            )
        )

        # ====================================================
        # MAIOR NÚMERO DE IMAGENS
        # ====================================================

        maior_quantidade_imagens = max(
            (
                len(
                    imagens_por_bloco.get(
                        bloco_idx,
                        []
                    )
                )

                for bloco_idx in range(
                    len(lote)
                )
            ),
            default=0,
        )

        # ====================================================
        # PROCESSAMENTO PROGRESSIVO
        # ====================================================

        for rodada in range(
            maior_quantidade_imagens
        ):

            if not pendentes:
                break

            self.metrics[
                "rodadas_processadas"
            ] += 1

            # =================================================
            # MONTA RODADA
            # =================================================

            itens_rodada = []

            task_para_blocos = {}

            for bloco_idx in sorted(
                pendentes
            ):

                hashes = (
                    imagens_por_bloco.get(
                        bloco_idx,
                        []
                    )
                )

                if rodada >= len(
                    hashes
                ):

                    continue

                img_hash = hashes[
                    rodada
                ]

                task_id = (
                    hash_to_task.get(
                        img_hash
                    )
                )

                image_path = (
                    hash_to_path.get(
                        img_hash
                    )
                )

                if not task_id:
                    continue

                if not image_path:
                    continue

                itens_rodada.append(
                    {
                        "task_id":
                            task_id,

                        "image_path":
                            image_path,
                    }
                )

                task_para_blocos.setdefault(
                    task_id,
                    []
                ).append(
                    bloco_idx
                )

            if not itens_rodada:
                continue

            # =================================================
            # YOLO
            # =================================================

            yolo_results = (
                self._yolo_batch(
                    processo_yolo,
                    itens_rodada,
                )
            )

            # =================================================
            # OCR
            # =================================================

            ocr_results = (
                self._ocr_batch(
                    processo_ocr,
                    yolo_results,
                )
            )

            # =================================================
            # DISTRIBUI
            # =================================================

            resolvidas_nesta_rodada = set()

            for (
                task_id,
                blocos_indices,
            ) in task_para_blocos.items():

                resultado_ocr = (
                    ocr_results.get(
                        task_id
                    )
                )

                if resultado_ocr is None:

                    for bloco_idx in (
                        blocos_indices
                    ):

                        resultados_por_bloco[
                            bloco_idx
                        ][
                            rodada
                        ] = None

                    continue

                for bloco_idx in (
                    blocos_indices
                ):

                    resultados_por_bloco[
                        bloco_idx
                    ][
                        rodada
                    ] = resultado_ocr

                    placa_esperada = (
                        lote[
                            bloco_idx
                        ][
                            "placa"
                        ]
                    )

                    if self._ocr_satisfaz_passagem(
                        placa_esperada,
                        resultado_ocr,
                    ):

                        resolvidas_nesta_rodada.add(
                            bloco_idx
                        )

            # =================================================
            # REMOVE RESOLVIDAS
            # =================================================

            for bloco_idx in (
                resolvidas_nesta_rodada
            ):

                pendentes.discard(
                    bloco_idx
                )

        # ====================================================
        # AVALIA PASSAGENS
        # ====================================================

        for bloco_idx, bloco in enumerate(
            lote
        ):

            resultados_passagem = {}

            resultados_rodadas = (
                resultados_por_bloco.get(
                    bloco_idx,
                    {}
                )
            )

            for rodada, resultado in sorted(
                resultados_rodadas.items()
            ):

                if resultado is None:
                    continue

                if not resultado:
                    continue

                resultados_passagem[
                    rodada
                ] = resultado

            # =================================================
            # AVALIAÇÃO FINAL
            # =================================================

            resultado = (
                self._avaliar_passagem(
                    bloco[
                        "placa"
                    ],
                    resultados_passagem,
                    len(
                        bloco.get(
                            "_imagens_bytes",
                            []
                        )
                    ),
                )
            )

            # =================================================
            # REGISTRO
            # =================================================

            registros.append(
                {
                    "Arquivo":
                        nome_arquivo,

                    "Seq":
                        bloco[
                            "seq"
                        ],

                    "ID Sistema":
                        bloco[
                            "id"
                        ],

                    "Data":
                        bloco[
                            "data"
                        ],

                    "Hora":
                        bloco[
                            "hora"
                        ],

                    "Praça":
                        bloco[
                            "praca"
                        ],

                    "Sentido":
                        bloco[
                            "sentido"
                        ],

                    "Pista":
                        bloco[
                            "pista"
                        ],

                    "Placa Esperada":
                        bloco[
                            "placa"
                        ],

                    "Categoria":
                        bloco[
                            "categoria"
                        ],

                    "Eixos":
                        bloco[
                            "eixos"
                        ],

                    "Placa OCR":
                        resultado[
                            "placa_ocr"
                        ],

                    "OCR Bruto":
                        resultado[
                            "ocr_bruto"
                        ],

                    "Status OCR":
                        resultado[
                            "status"
                        ],

                    "Diferença OCR":
                        resultado[
                            "diferenca"
                        ],

                    "Confiança OCR":
                        resultado[
                            "confidence"
                        ],

                    "Imagem Usada":
                        resultado[
                            "imagem_usada"
                        ],

                    "Total Imagens":
                        resultado[
                            "total_imgs"
                        ],

                    "Imagens c/ OCR":
                        resultado[
                            "imgs_com_ocr"
                        ],

                    "Linha Excel":
                        bloco[
                            "linha"
                        ],
                }
            )

        return registros

    # ========================================================
    # BATCH YOLO
    # ========================================================

    def _yolo_batch(
        self,
        processo,
        items,
    ):

        if not items:
            return {}

        inicio = (
            time.perf_counter()
        )

        enviar_json(
            processo,
            {
                "command":
                    "detect_batch",

                "images":
                    items,
            }
        )

        resposta = ler_json(
            processo,
            "YOLO",
        )

        self.metrics[
            "tempo_yolo"
        ] += (
            time.perf_counter()
            - inicio
        )

        self.metrics[
            "qtd_imagens_yolo"
        ] += len(
            items
        )

        resultados = {}

        for resultado in resposta.get(
            "results",
            [],
        ):

            task_id = resultado.get(
                "task_id"
            )

            if task_id:

                resultados[
                    task_id
                ] = resultado

        return resultados

    # ========================================================
    # BATCH OCR
    # ========================================================

    def _ocr_batch(
        self,
        processo,
        yolo_results,
    ):

        if not yolo_results:
            return {}

        items = []

        for (
            task_id,
            resultado,
        ) in yolo_results.items():

            if not resultado.get(
                "detected",
                False,
            ):

                continue

            crop_b64 = (
                resultado.get(
                    "crop_b64"
                )
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

        inicio = (
            time.perf_counter()
        )

        enviar_json(
            processo,
            {
                "command":
                    "ocr_batch",

                "images":
                    items,
            }
        )

        resposta = ler_json(
            processo,
            "OCR",
        )

        self.metrics[
            "tempo_ocr"
        ] += (
            time.perf_counter()
            - inicio
        )

        self.metrics[
            "qtd_chamadas_ocr"
        ] += len(
            items
        )

        resultados = {}

        for resultado in resposta.get(
            "results",
            [],
        ):

            task_id = resultado.get(
                "task_id"
            )

            if task_id:

                resultados[
                    task_id
                ] = resultado

        return resultados

    # ========================================================
    # AVALIA PASSAGEM
    #
    # AQUI ESTÁ A PRINCIPAL MUDANÇA.
    #
    # O OCR pode ser maior que 7 caracteres.
    # Procuramos dentro dele a melhor sequência de 7.
    # ========================================================

    def _avaliar_passagem(
        self,
        placa_esperada,
        resultados,
        total_imgs,
    ):

        inicio = (
            time.perf_counter()
        )

        placa_esperada = (
            self._limpar_texto(
                placa_esperada
            )
        )

        mascara = (
            self._descobrir_mascara(
                placa_esperada
            )
        )

        melhor_placa = ""
        melhor_distancia = 999.0

        melhor_indice = None

        melhor_ocr = ""

        melhor_conf = 0.0

        melhor_candidato = ""

        imagens_com_ocr = 0

        # ====================================================
        # TODAS AS RODADAS
        # ====================================================

        for img_idx, resultado in sorted(
            resultados.items()
        ):

            texto = resultado.get(
                "ocr_best",
                "",
            )

            if not texto:
                continue

            imagens_com_ocr += 1

            # =================================================
            # EXTRAÇÃO INTELIGENTE
            # =================================================

            extracao = (
                self._extrair_placa_ocr(
                    placa_esperada,
                    texto,
                    mascara,
                )
            )

            placa_ocr = (
                extracao[
                    "placa_ocr"
                ]
            )

            distancia = (
                extracao[
                    "diferenca"
                ]
            )

            candidato = (
                extracao[
                    "candidato"
                ]
            )

            try:

                confianca = float(
                    resultado.get(
                        "confidence_best",
                        0.0,
                    )
                    or
                    0.0
                )

            except Exception:

                confianca = 0.0

            # =================================================
            # DEBUG SOMENTE QUANDO HOUVE LIMPEZA
            # =================================================

            if (
                _limpar_texto(
                    texto
                )
                !=
                _limpar_texto(
                    placa_ocr
                )
                and
                placa_ocr
            ):

                print(
                    "\n[SPVIAS OCR]"
                )

                print(
                    f"Placa sistema: "
                    f"{placa_esperada}"
                )

                print(
                    f"OCR bruto: "
                    f"{texto}"
                )

                print(
                    f"Candidato: "
                    f"{candidato}"
                )

                print(
                    f"OCR limpo: "
                    f"{placa_ocr}"
                )

                print(
                    f"Distância: "
                    f"{distancia}"
                )

            # =================================================
            # MELHOR RESULTADO
            # =================================================

            if (
                distancia
                <
                melhor_distancia
                or
                (
                    distancia
                    ==
                    melhor_distancia
                    and
                    confianca
                    >
                    melhor_conf
                )
            ):

                melhor_distancia = (
                    distancia
                )

                melhor_placa = (
                    placa_ocr
                )

                melhor_indice = (
                    img_idx
                )

                melhor_ocr = (
                    texto
                )

                melhor_conf = (
                    confianca
                )

                melhor_candidato = (
                    candidato
                )

            # =================================================
            # MATCH EXATO
            # =================================================

            if distancia == 0:

                break

        # ====================================================
        # STATUS
        # ====================================================

        if (
            melhor_placa
            and
            melhor_distancia
            <=
            MAX_DIFERENCAS_APROVACAO
        ):

            status = (
                "Aprovado - busca nas imagens"
            )

            # =================================================
            # MUITO IMPORTANTE:
            #
            # O valor final é a placa recuperada,
            # não simplesmente o OCR bruto.
            # =================================================

            placa_final = (
                melhor_placa
            )

        elif melhor_placa:

            status = "Revisar"

            placa_final = (
                melhor_placa
            )

        elif total_imgs == 0:

            status = (
                "Revisar - sem imagem"
            )

            placa_final = (
                "Nao foi possivel "
                "identificar os caracteres"
            )

        else:

            status = (
                "Revisar - sem detecção"
            )

            placa_final = (
                "Nao foi possivel "
                "identificar os caracteres"
            )

        self.metrics[
            "tempo_avaliacoes"
        ] += (
            time.perf_counter()
            - inicio
        )

        return {
            "placa_ocr":
                placa_final,

            "ocr_bruto":
                melhor_ocr,

            "status":
                status,

            "diferenca":
                round(
                    melhor_distancia,
                    2,
                ),

            "confidence":
                round(
                    melhor_conf,
                    4,
                ),

            "imagem_usada":
                (
                    melhor_indice + 1
                    if
                    melhor_indice is not None
                    else ""
                ),

            "total_imgs":
                total_imgs,

            "imgs_com_ocr":
                imagens_com_ocr,

            "candidato":
                melhor_candidato,
        }

    # ========================================================
    # CRIA LOTES
    # ========================================================

    def criar_lotes(
        self,
        itens,
    ):

        return [
            itens[
                i:
                i
                +
                self.batch_size
            ]

            for i in range(
                0,
                len(itens),
                self.batch_size,
            )
        ]

    # ========================================================
    # PROCESSA PASTA
    # ========================================================

    def processar(
        self,
        pasta,
        caminho_saida,
    ):

        inicio_total = (
            time.perf_counter()
        )

        pasta = Path(
            pasta
        ).resolve()

        # ====================================================
        # XLSX
        # ====================================================

        arquivos = sorted(
            [
                arquivo

                for arquivo
                in pasta.rglob("*")

                if (
                    arquivo.is_file()
                    and
                    arquivo.suffix.lower()
                    in EXTENSOES_XLSX
                    and
                    not arquivo.name.startswith(
                        "~$"
                    )
                    and
                    arquivo.name
                    !=
                    Path(
                        caminho_saida
                    ).name
                )
            ]
        )

        if not arquivos:

            print(
                "[ERRO] "
                "Nenhum XLSX encontrado."
            )

            return

        print(
            "\n"
            + "=" * 75
        )

        print(
            "ANÁLISE SPVIAS"
        )

        print(
            "=" * 75
        )

        print(
            f"Pasta: {pasta}"
        )

        print(
            f"Arquivos encontrados: "
            f"{len(arquivos)}"
        )

        print(
            "=" * 75
        )

        processo_yolo = None
        processo_ocr = None

        temp_dir = Path(
            tempfile.mkdtemp(
                prefix="spvias_"
            )
        )

        registros = []

        arquivos_processados = 0
        arquivos_erro = 0

        try:

            # =================================================
            # PROCESSOS
            # =================================================

            processo_yolo = iniciar_processo(
                PYTHON_YOLO,
                DETECTOR_PY,
                "YOLO",
            )

            processo_ocr = iniciar_processo(
                PYTHON_OCR,
                OCR_PY,
                "OCR",
            )

            print(
                "\n[INIT] "
                "Aguardando inicialização "
                "dos modelos..."
            )

            time.sleep(
                15
            )

            # =================================================
            # CADA EXCEL
            # =================================================

            for numero, caminho_excel in enumerate(
                arquivos,
                start=1,
            ):

                print(
                    "\n"
                    + "#" * 75
                )

                print(
                    f"ARQUIVO "
                    f"{numero}/"
                    f"{len(arquivos)}"
                )

                print(
                    caminho_excel
                )

                print(
                    "#" * 75
                )

                wb = None

                try:

                    # =========================================
                    # WORKBOOK
                    # =========================================

                    wb = load_workbook(
                        caminho_excel,
                        data_only=True,
                    )

                    if (
                        NOME_ABA
                        not in
                        wb.sheetnames
                    ):

                        print(
                            f"[AVISO] "
                            f"Aba '{NOME_ABA}' "
                            f"não encontrada."
                        )

                        arquivos_erro += 1

                        wb.close()

                        continue

                    ws = wb[
                        NOME_ABA
                    ]

                    # =========================================
                    # PASSAGENS
                    # =========================================

                    blocos = (
                        self._ler_blocos(
                            ws
                        )
                    )

                    print(
                        f"[EXCEL] "
                        f"{len(blocos)} "
                        f"passagens encontradas."
                    )

                    # =========================================
                    # IMAGENS
                    # =========================================

                    mapa_imagens = (
                        self._mapear_imagens_xml(
                            str(
                                caminho_excel
                            ),
                            ws.title,
                        )
                    )

                    qtd_imagens = sum(
                        len(
                            lista
                        )
                        for lista
                        in mapa_imagens.values()
                    )

                    print(
                        f"[IMAGENS XML] "
                        f"{qtd_imagens} "
                        f"mídias associadas encontradas."
                    )

                    # =========================================
                    # FORMATOS
                    # =========================================

                    qtd_raster = 0
                    qtd_vetorial = 0

                    for lista in (
                        mapa_imagens.values()
                    ):

                        for info in lista:

                            ext = (
                                info.get(
                                    "ext",
                                    ""
                                )
                                .lower()
                            )

                            if (
                                ext
                                in
                                EXTENSOES_RASTER
                            ):

                                qtd_raster += 1

                            elif (
                                ext
                                in
                                EXTENSOES_VETORIAIS
                            ):

                                qtd_vetorial += 1

                    print(
                        f"[IMAGENS] "
                        f"Raster: {qtd_raster} | "
                        f"EMF/WMF: {qtd_vetorial}"
                    )

                    # =========================================
                    # ASSOCIAÇÃO
                    # =========================================

                    com_imagem = 0
                    sem_imagem = 0

                    for bloco in blocos:

                        imagens_bloco = []

                        for (
                            linha_ancora,
                            lista_imagens,
                        ) in mapa_imagens.items():

                            if (
                                bloco[
                                    "linha"
                                ]
                                <=
                                linha_ancora
                                <
                                bloco[
                                    "linha_fim"
                                ]
                            ):

                                imagens_bloco.extend(
                                    lista_imagens
                                )

                        bloco[
                            "_imagens_bytes"
                        ] = imagens_bloco

                        if imagens_bloco:

                            com_imagem += 1

                        else:

                            sem_imagem += 1

                    print(
                        f"[ASSOCIAÇÃO] "
                        f"{com_imagem} com imagem | "
                        f"{sem_imagem} sem imagem"
                    )

                    # =========================================
                    # LOTES
                    # =========================================

                    lotes = (
                        self.criar_lotes(
                            blocos
                        )
                    )

                    # =========================================
                    # BARRA
                    # =========================================

                    barra = tqdm(
                        total=len(blocos),
                        desc="YOLO + OCR",
                        unit="pass",
                        dynamic_ncols=True,
                    )

                    try:

                        for lote_idx, lote in enumerate(
                            lotes
                        ):

                            regs = (
                                self._processar_lote(
                                    lote,
                                    caminho_excel.name,
                                    processo_yolo,
                                    processo_ocr,
                                    lote_idx,
                                    temp_dir,
                                )
                            )

                            registros.extend(
                                regs
                            )

                            novos_aprovados = sum(
                                1
                                for r in regs
                                if r.get(
                                    "Status OCR"
                                )
                                ==
                                "Aprovado - busca nas imagens"
                            )

                            novos_revisar = sum(
                                1
                                for r in regs
                                if r.get(
                                    "Status OCR"
                                )
                                ==
                                "Revisar"
                            )

                            novos_sem_deteccao = sum(
                                1
                                for r in regs
                                if r.get(
                                    "Status OCR"
                                )
                                ==
                                "Revisar - sem detecção"
                            )

                            barra.update(
                                len(lote)
                            )

                            barra.set_postfix(
                                aprovados=
                                    novos_aprovados,

                                revisar=
                                    novos_revisar,

                                sem_det=
                                    novos_sem_deteccao,
                            )

                    finally:

                        barra.close()

                    arquivos_processados += 1

                    wb.close()

                except Exception as e:

                    arquivos_erro += 1

                    print(
                        f"[ERRO] "
                        f"{caminho_excel.name}: "
                        f"{type(e).__name__}: "
                        f"{e}",
                        flush=True,
                    )

                    if wb is not None:

                        try:
                            wb.close()
                        except Exception:
                            pass

        finally:

            # =================================================
            # ENCERRA
            # =================================================

            encerrar_processo(
                processo_yolo,
                "YOLO",
            )

            encerrar_processo(
                processo_ocr,
                "OCR",
            )

            shutil.rmtree(
                temp_dir,
                ignore_errors=True,
            )

        # ====================================================
        # SEM RESULTADOS
        # ====================================================

        if not registros:

            print(
                "\n[ERRO] "
                "Nenhum registro gerado."
            )

            return

        # ====================================================
        # DATAFRAME
        # ====================================================

        df = pd.DataFrame(
            registros
        )

        # ====================================================
        # SALVA
        # ====================================================

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

        # ====================================================
        # MÉTRICAS
        # ====================================================

        self.metrics[
            "tempo_total"
        ] = (
            time.perf_counter()
            -
            inicio_total
        )

        total = len(
            df
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

        sem_deteccao = int(
            (
                df[
                    "Status OCR"
                ]
                ==
                "Revisar - sem detecção"
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

        sem_ocr = int(
            (
                df[
                    "Status OCR"
                ]
                ==
                "Revisar - sem OCR"
            ).sum()
        )

        if total > 0:

            taxa_aprovacao = (
                aprovados
                /
                total
            ) * 100

        else:

            taxa_aprovacao = 0.0

        # ====================================================
        # RELATÓRIO
        # ====================================================

        print(
            "\n"
            + "=" * 75
        )

        print(
            "RESULTADO FINAL - SPVIAS"
        )

        print(
            "=" * 75
        )

        print(
            f"Arquivos encontrados    : "
            f"{len(arquivos)}"
        )

        print(
            f"Arquivos processados    : "
            f"{arquivos_processados}"
        )

        print(
            f"Arquivos com erro       : "
            f"{arquivos_erro}"
        )

        print(
            f"Passagens analisadas    : "
            f"{total}"
        )

        print(
            f"Aprovados               : "
            f"{aprovados}"
        )

        print(
            f"Revisar                 : "
            f"{revisar}"
        )

        print(
            f"Sem detecção YOLO       : "
            f"{sem_deteccao}"
        )

        print(
            f"Sem OCR                 : "
            f"{sem_ocr}"
        )

        print(
            f"Sem imagem              : "
            f"{sem_imagem}"
        )

        print(
            f"Taxa de aprovação       : "
            f"{taxa_aprovacao:.2f}%"
        )

        print(
            f"Imagens enviadas YOLO   : "
            f"{self.metrics['qtd_imagens_yolo']}"
        )

        print(
            f"OCRs processados        : "
            f"{self.metrics['qtd_chamadas_ocr']}"
        )

        print(
            f"Rodadas processadas     : "
            f"{self.metrics['rodadas_processadas']}"
        )

        print(
            f"Vetores convertidos     : "
            f"{self.metrics['vetores_convertidos']}"
        )

        print(
            f"Vetores sem conversão   : "
            f"{self.metrics['vetores_falharam']}"
        )

        print(
            f"Tempo total             : "
            f"{self.metrics['tempo_total']:.2f}s"
        )

        print(
            f"Arquivo gerado          : "
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
        "ANÁLISE SPVIAS"
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
    # ARGUMENTO / INPUT
    # ========================================================

    if len(
        sys.argv
    ) >= 2:

        caminho = (
            sys.argv[1]
            .strip()
            .strip('"')
        )

    else:

        caminho = input(
            "\nCaminho da pasta com os Excel: "
        ).strip().strip('"')

    if not caminho:

        print(
            "[ERRO] "
            "Nenhuma pasta informada."
        )

        return

    pasta = Path(
        caminho
    ).resolve()

    if not pasta.exists():

        print(
            "[ERRO] "
            "Pasta não encontrada:"
        )

        print(
            pasta
        )

        return

    if not pasta.is_dir():

        print(
            "[ERRO] "
            "O caminho informado "
            "não é uma pasta."
        )

        return

    # ========================================================
    # SAÍDA
    # ========================================================

    saida = caminho_resultado("spvias", "resultado_spvias.xlsx")

    # ========================================================
    # IMAGEMAGICK
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
            "EMF/WMF poderão depender "
            "do suporte do Pillow."
        )

    # ========================================================
    # EXECUTA
    # ========================================================

    validador = ValidadorSPVias(
        batch_size=BATCH_SIZE
    )

    validador.processar(
        pasta,
        saida,
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    main()