from __future__ import annotations


# ============================================================
# IMPORTS
# ============================================================

import base64
import json
import re
import subprocess
import sys
import unicodedata
from pathlib import Path

import pandas as pd


# ============================================================
# IMPORT DO CONFIG
# ============================================================

# Este ler_imagem.py está dentro da pasta SAU.
# O config.py está na pasta principal do projeto.

BASE_DIR_LOCAL = Path(
    __file__
).resolve().parent.parent

if str(BASE_DIR_LOCAL) not in sys.path:
    sys.path.insert(
        0,
        str(BASE_DIR_LOCAL)
    )

from config import (
    BASE_DIR,
    PYTHON_OCR,
    OCR_SAU_PY,
    OCR_CONF_MINIMA,
)


# ============================================================
# CONFIGURAÇÕES
# ============================================================

RESULTADOS_DIR = (
    BASE_DIR
    / "resultados_ocr"
)


# ============================================================
# MÁSCARAS OCR
# ============================================================

# ------------------------------------------------------------
# Quando a posição deveria ser LETRA
# ------------------------------------------------------------

MAPA_LETRAS = {
    "0": "O",
    "1": "I",
    "2": "Z",
    "4": "A",
    "5": "S",
    "6": "G",
    "7": "Z",
    "8": "B",
}


# ------------------------------------------------------------
# Quando a posição deveria ser DÍGITO
# ------------------------------------------------------------

MAPA_DIGITOS = {
    "O": "0",
    "Q": "0",
    "D": "0",

    "I": "1",
    "L": "1",
    "T": "7",

    "Z": "2",

    "A": "4",

    "S": "5",

    "G": "6",

    "B": "8",
}


# ============================================================
# CONFIGURAÇÃO DA MÁSCARA DA PLACA
# ============================================================

# LLLDLDD
#
# Posição:
#
# 1 = L
# 2 = L
# 3 = L
# 4 = D
# 5 = NÃO ALTERAR
# 6 = D
# 7 = D
#
# A 5ª posição pode ser:
#
# Padrão antigo:
#     ABC-1234
#
# Mercosul:
#     ABC-1D23
#
# Portanto ela nunca passa pela máscara.
# ============================================================

MASCARA_PLACA = (
    "LLL?D??"
)


# ============================================================
# CADASTRO DE PREFIXO + PLACA
# ============================================================

def limpar_caminho_digitado(valor):
    """
    Remove aspas que normalmente aparecem quando o usuário
    arrasta um arquivo/pasta para o terminal do Windows.
    """

    return str(
        valor
    ).strip().strip(
        '"'
    ).strip(
        "'"
    )


def normalizar_chave(valor):
    """
    Normaliza valores usados para comparação:

    - maiúsculas;
    - remove acentos;
    - remove espaços;
    - remove hífens;
    - remove pontos;
    - remove demais separadores.
    """

    if valor is None or pd.isna(valor):
        return ""

    valor = str(
        valor
    ).strip().replace(
        "\ufeff",
        ""
    )

    valor = unicodedata.normalize(
        "NFKD",
        valor
    )

    valor = "".join(
        c
        for c in valor
        if not unicodedata.combining(c)
    )

    return re.sub(
        r"[^A-Z0-9]",
        "",
        valor.upper()
    )


def normalizar_nome_coluna(valor):
    """
    Aceita, por exemplo:

    Prefixo / prefixo / PREfIXo / P R E F I X O
    """

    return normalizar_chave(
        valor
    )


def formatar_placa(placa):
    """
    Mantém a placa em um formato visual padronizado.

    Ex.: ABC1234 -> ABC-1234
         ABC1D23 -> ABC-1D23
    """

    placa = normalizar_chave(
        placa
    )

    if len(placa) == 7:

        return (
            f"{placa[:3]}-"
            f"{placa[3:]}"
        )

    return placa


def carregar_cadastro_excel():
    """
    Solicita e lê a planilha-base com as colunas Prefixo e Placa.
    """

    print(
        "\n"
        + "=" * 70
    )

    print(
        "CADASTRO DE PREFIXOS E PLACAS"
    )

    print(
        "=" * 70
    )

    while True:

        caminho_digitado = input(
            "\nInforme o caminho do Excel contendo as colunas "
            "'Prefixo' e 'Placa':\n> "
        )

        caminho_excel = Path(
            limpar_caminho_digitado(
                caminho_digitado
            )
        )

        if not caminho_excel.is_file():

            print(
                "\nERRO: o arquivo informado não existe:\n"
                f"{caminho_excel}\n"
            )

            continue

        try:

            df_cadastro = pd.read_excel(
                caminho_excel,
                dtype=str
            )

        except Exception as erro:

            print(
                "\nERRO ao abrir o Excel:\n"
                f"{erro}\n"
            )

            continue

        mapa_colunas = {
            normalizar_nome_coluna(
                coluna
            ): coluna
            for coluna in df_cadastro.columns
        }

        coluna_prefixo = mapa_colunas.get(
            "PREFIXO"
        )

        coluna_placa = mapa_colunas.get(
            "PLACA"
        )

        if not coluna_prefixo or not coluna_placa:

            print(
                "\nERRO: a planilha precisa conter as colunas "
                "'Prefixo' e 'Placa'.\n"
                "A comparação de nomes ignora maiúsculas/minúsculas "
                "e separadores.\n"
            )

            print(
                "Colunas encontradas:",
                ", ".join(
                    str(c)
                    for c in df_cadastro.columns
                )
            )

            continue

        cadastro = {}

        ignorados = 0
        duplicados = 0

        for _, linha in df_cadastro.iterrows():

            prefixo = normalizar_chave(
                linha.get(
                    coluna_prefixo
                )
            )

            placa = normalizar_chave(
                linha.get(
                    coluna_placa
                )
            )

            if not prefixo or not placa:

                ignorados += 1
                continue

            if prefixo not in cadastro:

                cadastro[prefixo] = {
                    "prefixo": prefixo,
                    "placas": []
                }

            if placa not in cadastro[
                prefixo
            ][
                "placas"
            ]:

                cadastro[
                    prefixo
                ][
                    "placas"
                ].append(
                    placa
                )

                if len(
                    cadastro[
                        prefixo
                    ][
                        "placas"
                    ]
                ) > 1:

                    duplicados += 1

        if not cadastro:

            print(
                "\nERRO: nenhuma combinação válida de "
                "Prefixo + Placa foi encontrada no Excel.\n"
            )

            continue

        print(
            "\nCadastro carregado com sucesso:"
        )

        print(
            f"  Prefixos cadastrados: "
            f"{len(cadastro)}"
        )

        print(
            f"  Linhas ignoradas: "
            f"{ignorados}"
        )

        if duplicados:

            print(
                "  ATENÇÃO: existem prefixos associados "
                f"a mais de uma placa ({duplicados} ocorrências)."
            )

        print(
            f"  Arquivo: "
            f"{caminho_excel}\n"
        )

        return cadastro


# ============================================================
# MÁSCARA DO PREFIXO
# ============================================================

def aplicar_mascara_prefixo(
    texto
):
    """
    Procura um prefixo no formato:

        AA-0000
        AA-00000
        AAA-0000
        AAA-00000

    O hífen é obrigatório.

    Antes do hífen:
        somente letras.

    Depois do hífen:
        somente números.

    A máscara corrige confusões comuns do OCR.
    """

    if texto is None:
        return ""

    texto = str(
        texto
    ).strip().upper()

    # --------------------------------------------------------
    # HÍFEN OBRIGATÓRIO
    # --------------------------------------------------------

    padrao = re.compile(
        r"\b([A-Z0-9]{2,3})-([A-Z0-9]{4,5})\b",
        re.IGNORECASE
    )

    match = padrao.search(
        texto
    )

    if not match:
        return ""

    parte_letras = (
        match.group(1)
        .upper()
    )

    parte_numeros = (
        match.group(2)
        .upper()
    )

    # --------------------------------------------------------
    # PARTE DAS LETRAS
    # --------------------------------------------------------

    letras_corrigidas = []

    for char in parte_letras:

        if char.isalpha():

            letras_corrigidas.append(
                char
            )

        elif char in MAPA_LETRAS:

            letras_corrigidas.append(
                MAPA_LETRAS[
                    char
                ]
            )

        else:

            # Não deveria acontecer por causa do regex,
            # mas mantém o caractere caso apareça algo inesperado.
            letras_corrigidas.append(
                char
            )

    # --------------------------------------------------------
    # PARTE DOS NÚMEROS
    # --------------------------------------------------------

    numeros_corrigidos = []

    for char in parte_numeros:

        if char.isdigit():

            numeros_corrigidos.append(
                char
            )

        elif char in MAPA_DIGITOS:

            numeros_corrigidos.append(
                MAPA_DIGITOS[
                    char
                ]
            )

        else:

            numeros_corrigidos.append(
                char
            )

    return (
        "".join(
            letras_corrigidas
        )
        + "-"
        +
        "".join(
            numeros_corrigidos
        )
    )


# ============================================================
# MÁSCARA DA PLACA
# ============================================================

def aplicar_mascara_placa(
    candidato,
    indices_originais=None,
):
    """
    Aplica a máscara LLLDLDD.

    A 5ª posição da placa NUNCA é alterada.

    Isso permite lidar com:

        ABC-1234
        ABC-1D23

    sem obrigar a posição 5 a ser letra ou número.
    """

    candidato = normalizar_chave(
        candidato
    )

    if not candidato:
        return ""

    if indices_originais is None:

        indices_originais = list(
            range(
                len(candidato)
            )
        )

    resultado = []

    for indice_candidato, char in enumerate(
        candidato
    ):

        if indice_candidato >= len(
            indices_originais
        ):

            resultado.append(
                char
            )

            continue

        indice_original = indices_originais[
            indice_candidato
        ]

        # ====================================================
        # 5ª POSIÇÃO
        # ====================================================
        # Nunca alterar.
        # ====================================================

        if indice_original == 4:

            resultado.append(
                char
            )

            continue

        # ====================================================
        # LETRAS
        # ====================================================

        if indice_original in {
            0,
            1,
            2,
        }:

            if char.isalpha():

                resultado.append(
                    char
                )

            elif char in MAPA_LETRAS:

                resultado.append(
                    MAPA_LETRAS[
                        char
                    ]
                )

            else:

                resultado.append(
                    char
                )

        # ====================================================
        # NÚMEROS
        # ====================================================

        else:

            if char.isdigit():

                resultado.append(
                    char
                )

            elif char in MAPA_DIGITOS:

                resultado.append(
                    MAPA_DIGITOS[
                        char
                    ]
                )

            else:

                resultado.append(
                    char
                )

    return "".join(
        resultado
    )


# ============================================================
# GERA CANDIDATOS DE PLACA
# ============================================================

def gerar_candidatos_placa(
    texto
):
    """
    Gera candidatos de 7 e 6 caracteres.

    A janela de 6 caracteres permite cobrir o caso em que
    o OCR perdeu exatamente um caractere da placa.
    """

    normalizado = normalizar_chave(
        texto
    )

    candidatos = []

    for tamanho in (
        7,
        6,
    ):

        if len(
            normalizado
        ) < tamanho:

            continue

        for inicio in range(
            0,
            len(
                normalizado
            )
            -
            tamanho
            +
            1
        ):

            candidato = normalizado[
                inicio:
                inicio
                +
                tamanho
            ]

            if candidato not in candidatos:

                candidatos.append(
                    candidato
                )

    return candidatos


# ============================================================
# PONTUA O CANDIDATO DA PLACA
# ============================================================

def pontuar_candidato_placa(
    candidato,
    esperado,
):
    """
    Compara uma placa OCR com a placa esperada.

    7 caracteres:
        aplica máscara e compara posição por posição.

    6 caracteres:
        considera a perda de um caractere e testa todas as
        possíveis posições ausentes.

    A 5ª posição ORIGINAL da placa nunca é alterada pela máscara.
    """

    candidato = normalizar_chave(
        candidato
    )

    esperado = normalizar_chave(
        esperado
    )

    if len(
        esperado
    ) != 7:

        return None

    # ========================================================
    # PLACA COMPLETA
    # ========================================================

    if len(
        candidato
    ) == 7:

        candidato_mascarado = (
            aplicar_mascara_placa(
                candidato
            )
        )

        iguais = sum(
            c1 == c2
            for c1, c2 in zip(
                candidato_mascarado,
                esperado
            )
        )

        if iguais >= 6:

            return {
                "pontuacao":
                    iguais,

                "candidato":
                    candidato_mascarado,
            }

        return None

    # ========================================================
    # OCR PERDEU UM CARACTERE
    # ========================================================

    if len(
        candidato
    ) == 6:

        melhor = None

        for remover in range(
            7
        ):

            # ------------------------------------------------
            # Remove uma posição da placa esperada
            # ------------------------------------------------

            esperado_reduzido = (
                esperado[:remover]
                +
                esperado[remover + 1:]
            )

            # ------------------------------------------------
            # Mantém a relação entre cada caractere OCR
            # e sua posição original na placa.
            # ------------------------------------------------

            indices_originais = [
                indice
                for indice in range(
                    7
                )
                if indice != remover
            ]

            # ------------------------------------------------
            # Aplica máscara respeitando a posição original.
            # ------------------------------------------------

            candidato_mascarado = (
                aplicar_mascara_placa(
                    candidato,
                    indices_originais,
                )
            )

            iguais = sum(
                c1 == c2
                for c1, c2 in zip(
                    candidato_mascarado,
                    esperado_reduzido
                )
            )

            if iguais == 6:

                resultado = {
                    "pontuacao":
                        6,

                    "candidato":
                        candidato_mascarado,
                }

                if (
                    melhor is None
                    or
                    resultado[
                        "pontuacao"
                    ]
                    >
                    melhor[
                        "pontuacao"
                    ]
                ):

                    melhor = resultado

        return melhor

    return None


# ============================================================
# VALIDA PLACA OCR
# ============================================================

def validar_placa_ocr(
    textos,
    placas_esperadas,
):
    """
    Procura em todos os textos OCR uma placa que corresponda
    a alguma das placas cadastradas para o prefixo.
    """

    if not placas_esperadas:
        return None

    melhor_resultado = None

    for item in textos:

        texto = str(
            item.get(
                "texto",
                ""
            )
        ).strip()

        if not texto:
            continue

        candidatos = (
            gerar_candidatos_placa(
                texto
            )
        )

        try:

            confianca = float(
                item.get(
                    "confianca",
                    0.0
                )
                or
                0.0
            )

        except Exception:

            confianca = 0.0

        for candidato in candidatos:

            for placa_esperada in placas_esperadas:

                resultado = (
                    pontuar_candidato_placa(
                        candidato,
                        placa_esperada
                    )
                )

                if resultado is None:
                    continue

                atual = {
                    "placa_esperada":
                        placa_esperada,

                    "placa_ocr":
                        resultado[
                            "candidato"
                        ],

                    "pontuacao":
                        resultado[
                            "pontuacao"
                        ],

                    "texto_ocr":
                        texto,

                    "confianca":
                        confianca,
                }

                if (
                    melhor_resultado is None
                    or
                    atual[
                        "pontuacao"
                    ]
                    >
                    melhor_resultado[
                        "pontuacao"
                    ]
                    or
                    (
                        atual[
                            "pontuacao"
                        ]
                        ==
                        melhor_resultado[
                            "pontuacao"
                        ]
                        and
                        atual[
                            "confianca"
                        ]
                        >
                        melhor_resultado[
                            "confianca"
                        ]
                    )
                ):

                    melhor_resultado = atual

    return melhor_resultado


# ============================================================
# PROCESSO OCR
# ============================================================

def enviar_json(
    processo,
    dados,
):
    """
    Envia uma requisição JSON para o
    SAU/ocr_sau_process.py.
    """

    if processo.poll() is not None:

        raise RuntimeError(
            "Processo OCR já encerrou. "
            f"Código: {processo.returncode}"
        )

    processo.stdin.write(
        json.dumps(
            dados,
            ensure_ascii=False
        )
        +
        "\n"
    )

    processo.stdin.flush()


def ler_json(
    processo,
):
    """
    Lê uma resposta JSON do
    SAU/ocr_sau_process.py.
    """

    linha = (
        processo.stdout.readline()
    )

    if not linha:

        raise RuntimeError(
            "O processo OCR encerrou antes de responder."
        )

    try:

        return json.loads(
            linha
        )

    except json.JSONDecodeError as erro:

        raise RuntimeError(
            "Resposta inválida do processo OCR: "
            f"{linha[:500]}"
        ) from erro


def iniciar_ocr():
    """
    Inicia o OCR específico da análise de veículos.

    IMPORTANTE:
    não usa OCR_PY / ocr_process.py.
    """

    if not PYTHON_OCR.is_file():

        raise FileNotFoundError(
            "Python do ambiente OCR não encontrado:\n"
            f"{PYTHON_OCR}\n\n"
            "Verifique se a pasta .venv_ocr existe."
        )

    if not OCR_SAU_PY.is_file():

        raise FileNotFoundError(
            "OCR específico da análise não encontrado:\n"
            f"{OCR_SAU_PY}\n\n"
            "Verifique se "
            "SAU/ocr_sau_process.py existe."
        )

    print(
        "\n"
        + "=" * 70
    )

    print(
        "INICIANDO PROCESSO OCR"
    )

    print(
        "=" * 70
    )

    print(
        f"Python OCR: "
        f"{PYTHON_OCR}"
    )

    print(
        f"Script OCR:  "
        f"{OCR_SAU_PY}\n"
    )

    return subprocess.Popen(
        [
            str(
                PYTHON_OCR
            ),
            str(
                OCR_SAU_PY
            ),
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=None,
        text=True,
        bufsize=1,
        cwd=str(
            BASE_DIR
        ),
    )


def encerrar_ocr(
    processo,
):
    """
    Encerra o processo OCR de forma limpa.
    """

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
# EXTRAÇÃO DOS TEXTOS DO OCR
# ============================================================

def extrair_textos_resultado_ocr(
    resultado,
):
    """
    Prioriza o novo formato:

        ocr_items = [
            {
                "texto": ...,
                "confianca": ...,
                "box": ...
            }
        ]

    O fallback mantém compatibilidade caso o OCR retorne
    os campos antigos.
    """

    if not resultado:
        return []

    # ========================================================
    # NOVO FORMATO: OCR ITEMS
    # ========================================================

    ocr_items = resultado.get(
        "ocr_items"
    )

    if isinstance(
        ocr_items,
        list
    ):

        textos = []

        for item in ocr_items:

            if not isinstance(
                item,
                dict
            ):

                continue

            texto = str(
                item.get(
                    "texto",
                    ""
                )
                or
                ""
            ).strip()

            if not texto:
                continue

            try:

                confianca = float(
                    item.get(
                        "confianca",
                        0.0
                    )
                    or
                    0.0
                )

            except Exception:

                confianca = 0.0

            textos.append(
                {
                    "texto":
                        texto,

                    "confianca":
                        confianca,

                    "box":
                        item.get(
                            "box"
                        ),

                    "variante":
                        "ocr_item",
                }
            )

        if textos:

            return textos

    # ========================================================
    # FALLBACK ANTIGO
    # ========================================================

    textos = []

    vistos = set()

    campos = [
        (
            "ocr_best",
            "best",
        ),
        (
            "ocr_best_original",
            "best_original",
        ),
        (
            "ocr_best_normal",
            "best_normal",
        ),
        (
            "ocr_best_sem_primeiro",
            "best_sem_primeiro",
        ),
        (
            "sr_ocr",
            "sr",
        ),
        (
            "sr_ocr_auxiliar",
            "sr_auxiliar",
        ),
        (
            "sr_ocr_sem_primeiro",
            "sr_sem_primeiro",
        ),
        (
            "ocr_cubic",
            "cubic",
        ),
        (
            "ocr_cubic_auxiliar",
            "cubic_auxiliar",
        ),
        (
            "ocr_cubic_sem_primeiro",
            "cubic_sem_primeiro",
        ),
        (
            "ocr_clahe",
            "clahe",
        ),
        (
            "ocr_clahe_auxiliar",
            "clahe_auxiliar",
        ),
        (
            "ocr_clahe_sem_primeiro",
            "clahe_sem_primeiro",
        ),
    ]

    for campo, variante in campos:

        valor = str(
            resultado.get(
                campo,
                ""
            )
            or
            ""
        ).strip()

        if not valor:
            continue

        chave = (
            variante,
            valor.upper()
        )

        if chave in vistos:
            continue

        vistos.add(
            chave
        )

        if campo.startswith(
            "ocr_cubic"
        ):

            confianca = (
                resultado.get(
                    "confidence_cubic",
                    0.0
                )
            )

        elif campo.startswith(
            "ocr_clahe"
        ):

            confianca = (
                resultado.get(
                    "confidence_clahe",
                    0.0
                )
            )

        elif campo.startswith(
            "sr_"
        ):

            confianca = (
                resultado.get(
                    "sr_confidence",
                    0.0
                )
            )

        elif (
            campo
            ==
            "ocr_best_normal"
        ):

            confianca = (
                resultado.get(
                    "confidence_best_normal",
                    0.0
                )
            )

        else:

            confianca = (
                resultado.get(
                    "confidence_best",
                    0.0
                )
            )

        try:

            confianca = float(
                confianca
                or
                0.0
            )

        except Exception:

            confianca = 0.0

        textos.append(
            {
                "texto":
                    valor,

                "confianca":
                    confianca,

                "box":
                    None,

                "variante":
                    variante,
            }
        )

    return textos


def ler_imagem(
    processo_ocr,
    imagem: Path,
):
    """
    Envia uma imagem ao OCR específico de veículos.
    """

    try:

        dados_imagem = (
            imagem.read_bytes()
        )

        imagem_b64 = (
            base64.b64encode(
                dados_imagem
            ).decode(
                "ascii"
            )
        )

    except Exception as erro:

        raise RuntimeError(
            f"Erro lendo a imagem "
            f"{imagem.name}: {erro}"
        ) from erro

    task_id = (
        imagem.stem
    )

    print(
        "\n"
        + "-" * 80
    )

    print(
        f"OCR DA IMAGEM: "
        f"{imagem.name}"
    )

    print(
        "-" * 80
    )

    enviar_json(
        processo_ocr,
        {
            "command":
                "ocr_batch",

            "images": [
                {
                    "task_id":
                        task_id,

                    "crop_b64":
                        imagem_b64,
                }
            ],
        },
    )

    resposta = ler_json(
        processo_ocr
    )

    if resposta.get(
        "error"
    ):

        raise RuntimeError(
            "OCR retornou erro: "
            f"{resposta['error']}"
        )

    resultados = resposta.get(
        "results",
        []
    )

    if not resultados:

        print(
            "[OCR] Nenhum resultado retornado."
        )

        return []

    resultado = (
        resultados[0]
    )

    if resultado.get(
        "error"
    ):

        raise RuntimeError(
            f"OCR retornou erro para "
            f"{imagem.name}: "
            f"{resultado['error']}"
        )

    textos = (
        extrair_textos_resultado_ocr(
            resultado
        )
    )

    if not textos:

        print(
            "[OCR] Nenhum texto encontrado."
        )

    for item in textos:

        print(
            f"[TEXTO] "
            f"{item['texto']!r} | "
            f"CONF: "
            f"{item['confianca']:.4f} | "
            f"VARIANTE: "
            f"{item['variante']}"
        )

    print(
        "-" * 80
        +
        "\n"
    )

    return textos


# ============================================================
# ORDENAÇÃO DOS TEXTOS
# ============================================================

def ordenar_textos(
    textos,
):
    """
    Ordena os textos de cima para baixo e da esquerda
    para a direita quando existe box.
    """

    def chave(item):

        box = item.get(
            "box"
        )

        if not box:

            return (
                0,
                0
            )

        try:

            # ------------------------------------------------
            # BOX:
            # [x_min, y_min, x_max, y_max]
            # ------------------------------------------------

            if (
                isinstance(
                    box,
                    list
                )
                and
                len(box) == 4
                and
                all(
                    isinstance(
                        valor,
                        (int, float)
                    )
                    for valor in box
                )
            ):

                return (
                    box[1],
                    box[0]
                )

            # ------------------------------------------------
            # POLYGON:
            # [[x,y], [x,y], ...]
            # ------------------------------------------------

            pontos = [
                ponto
                for ponto in box
                if (
                    isinstance(
                        ponto,
                        list
                    )
                    and
                    len(ponto) >= 2
                )
            ]

            if pontos:

                y = min(
                    ponto[1]
                    for ponto in pontos
                )

                x = min(
                    ponto[0]
                    for ponto in pontos
                )

                return (
                    y,
                    x
                )

        except Exception:

            pass

        return (
            0,
            0
        )

    return sorted(
        textos,
        key=chave
    )


# ============================================================
# BUSCA TELEFONE 0800
# ============================================================

def encontrar_telefone_0800(
    textos,
):
    """
    Procura telefones 0800 entre os textos OCR.
    """

    partes = []

    for item in textos:

        texto = str(
            item.get(
                "texto",
                ""
            )
        ).strip()

        if texto:

            partes.append(
                texto
            )

    if not partes:
        return None

    texto_completo = " ".join(
        partes
    )

    texto_completo = re.sub(
        r"\s+",
        " ",
        texto_completo
    ).strip()

    # ========================================================
    # 1) TELEFONE COMPLETO
    # ========================================================

    padroes_completos = [

        re.compile(
            r"\b0800[\s.\-]+055[\s.\-]+\d{4}\b",
            re.IGNORECASE
        ),

        re.compile(
            r"\b0800[\s.\-]+055\d{4}\b",
            re.IGNORECASE
        ),

        re.compile(
            r"\b0800055\d{4}\b",
            re.IGNORECASE
        ),
    ]

    for padrao in padroes_completos:

        match = padrao.search(
            texto_completo
        )

        if not match:
            continue

        numero = re.sub(
            r"\D",
            "",
            match.group(0)
        )

        if len(numero) == 11:

            return (
                f"{numero[:4]} "
                f"{numero[4:7]} "
                f"{numero[7:]}"
            )

    # ========================================================
    # 2) OCR SEPAROU EM DOIS BLOCOS
    # ========================================================

    encontrou_0800 = False
    bloco_055 = None

    for parte in partes:

        numero_parte = re.sub(
            r"\D",
            "",
            parte
        )

        if numero_parte == "0800":

            encontrou_0800 = True

        if re.fullmatch(
            r"055\d{4}",
            numero_parte
        ):

            bloco_055 = (
                numero_parte
            )

    if (
        encontrou_0800
        and
        bloco_055
    ):

        numero = (
            "0800"
            +
            bloco_055
        )

        return (
            f"{numero[:4]} "
            f"{numero[4:7]} "
            f"{numero[7:]}"
        )

    # ========================================================
    # 3) PROCURA NÚMEROS DENTRO DO TEXTO
    # ========================================================

    numeros = re.findall(
        r"\d+",
        texto_completo
    )

    encontrou_0800 = False
    bloco_055 = None

    for numero in numeros:

        if numero == "0800":

            encontrou_0800 = True

        elif re.fullmatch(
            r"055\d{4}",
            numero
        ):

            bloco_055 = numero

    if (
        encontrou_0800
        and
        bloco_055
    ):

        numero = (
            "0800"
            +
            bloco_055
        )

        return (
            f"{numero[:4]} "
            f"{numero[4:7]} "
            f"{numero[7:]}"
        )

    return None


# ============================================================
# EXTRAÇÃO DE DADOS DO VEÍCULO
# ============================================================

def extrair_dados_veiculo(
    textos,
    prefixo_esperado,
    lado,
    placas_esperadas=None,
):

    dados = {
        "servico_gratuito":
            False,

        "prefixo_lido":
            None,

        "telefone":
            None,

        "placa_ocr":
            None,

        "placa_esperada":
            None,

        "placa_validada":
            False,

        "placa_pontuacao":
            None,
    }

    # ========================================================
    # PADRÕES
    # ========================================================

    padrao_gratuito = re.compile(
        r"gratuito",
        re.IGNORECASE
    )

    # Prefixo:
    #
    # 2 ou 3 letras
    # hífen obrigatório
    # 4 ou 5 números
    #
    padrao_ocr_prefixo = re.compile(
        r"\b[A-Z0-9]{2,3}-[A-Z0-9]{4,5}\b",
        re.IGNORECASE
    )

    # ========================================================
    # TELEFONE
    # ========================================================

    dados[
        "telefone"
    ] = encontrar_telefone_0800(
        textos
    )

    # ========================================================
    # PLACA
    # ========================================================

    if placas_esperadas:

        resultado_placa = (
            validar_placa_ocr(
                textos,
                placas_esperadas
            )
        )

        if resultado_placa:

            dados[
                "placa_ocr"
            ] = formatar_placa(
                resultado_placa[
                    "placa_ocr"
                ]
            )

            dados[
                "placa_esperada"
            ] = formatar_placa(
                resultado_placa[
                    "placa_esperada"
                ]
            )

            dados[
                "placa_validada"
            ] = True

            dados[
                "placa_pontuacao"
            ] = (
                resultado_placa[
                    "pontuacao"
                ]
            )

    # ========================================================
    # DEMAIS CAMPOS
    # ========================================================

    for item in textos:

        texto = str(
            item.get(
                "texto",
                ""
            )
        ).strip()

        # ====================================================
        # GRATUITO
        # ====================================================

        if padrao_gratuito.search(
            texto
        ):

            dados[
                "servico_gratuito"
            ] = True

        # ====================================================
        # PREFIXO
        # ====================================================

        if (
            not dados[
                "prefixo_lido"
            ]
            and
            prefixo_esperado
        ):

            # ------------------------------------------------
            # Primeiro verifica se existe um prefixo
            # visualmente válido no texto OCR.
            # ------------------------------------------------

            match_prefixo = (
                padrao_ocr_prefixo.search(
                    texto
                )
            )

            if match_prefixo:

                prefixo_ocr_bruto = (
                    match_prefixo.group(0)
                )

                prefixo_ocr = (
                    aplicar_mascara_prefixo(
                        prefixo_ocr_bruto
                    )
                )

                if prefixo_ocr:

                    if (
                        normalizar_chave(
                            prefixo_ocr
                        )
                        ==
                        normalizar_chave(
                            prefixo_esperado
                        )
                    ):

                        dados[
                            "prefixo_lido"
                        ] = prefixo_esperado

    return dados


# ============================================================
# FORMATAÇÃO
# ============================================================

def formatar_celula(
    dados_lado,
):

    if not dados_lado:
        return "Sem Imagem"

    grat = (
        "Gratuito"
        if dados_lado.get(
            "servico_gratuito"
        )
        else
        "Não encontrado"
    )

    tel = (
        dados_lado.get(
            "telefone"
        )
        if dados_lado.get(
            "telefone"
        )
        else
        "Não encontrado"
    )

    pref = (
        dados_lado.get(
            "prefixo_lido"
        )
        if dados_lado.get(
            "prefixo_lido"
        )
        else
        "Não encontrado"
    )

    return (
        f"{grat} | "
        f"{tel} | "
        f"{pref}"
    )


# ============================================================
# CORES
# ============================================================

def aplicar_cores(
    val,
):

    if not isinstance(
        val,
        str
    ):

        return ""

    if val in [
        "Sem Imagem",
        "Não Identificada",
    ]:

        return ""

    if " | " in val:

        partes = val.split(
            " | "
        )

        sucessos = sum(
            1
            for p in partes
            if p.strip()
            !=
            "Não encontrado"
        )

        if sucessos == 3:

            return (
                "background-color: #c6efce; "
                "color: #006100;"
            )

        elif sucessos > 0:

            return (
                "background-color: #ffeb9c; "
                "color: #9c5700;"
            )

        else:

            return (
                "background-color: #ffc7ce; "
                "color: #9c0006;"
            )

    return ""


# ============================================================
# MAIN
# ============================================================

def main():

    # ========================================================
    # ARGUMENTO
    # ========================================================

    if len(
        sys.argv
    ) < 2:

        print(
            "\nUso:\n"
            '    py ler_imagem.py '
            '"C:\\caminho\\da\\pasta\\de\\imagens"\n'
        )

        return

    pasta_imagens = Path(
        limpar_caminho_digitado(
            sys.argv[1]
        )
    )

    if not pasta_imagens.is_dir():

        print(
            "\nERRO: O caminho informado "
            "não é uma pasta válida:\n"
            f"{pasta_imagens}\n"
        )

        return

    # ========================================================
    # CARREGA CADASTRO
    # ========================================================

    cadastro = (
        carregar_cadastro_excel()
    )

    # ========================================================
    # INICIA OCR
    # ========================================================

    processo_ocr = (
        iniciar_ocr()
    )

    # ========================================================
    # LOCALIZA IMAGENS
    # ========================================================

    arquivos = sorted(
        [
            arquivo
            for arquivo in (
                pasta_imagens.iterdir()
            )
            if (
                arquivo.is_file()
                and
                arquivo.suffix.lower()
                in {
                    ".jpg",
                    ".jpeg",
                    ".png",
                }
            )
        ],
        key=lambda p:
            p.name.lower()
    )

    total_arquivos = len(
        arquivos
    )

    print(
        f"[{total_arquivos}] imagens encontradas. "
        "Iniciando processamento...\n"
    )

    # ========================================================
    # VEÍCULOS
    # ========================================================

    veiculos = {}

    # ========================================================
    # PADRÃO DO NOME DO ARQUIVO
    # ========================================================

    # Formatos aceitos:
    #
    # ADM-11000-FR-2S25.JPG
    # AA-1234-LD-2S25.JPG
    # AAA-12345-LE-2S25.JPG
    #
    # Prefixo:
    # 2 ou 3 letras
    # hífen
    # 4 ou 5 números
    #
    padrao_arquivo = re.compile(
        r"^("
        r"[A-Z]{2,3}"
        r"-"
        r"\d{4,5}"
        r")"
        r"[_-]"
        r"(LD|LE|FR|TR)"
        r"(?:[_-]|\.[^.]*$|$)",
        re.IGNORECASE
    )

    # ========================================================
    # PROCESSAMENTO
    # ========================================================

    try:

        for idx, imagem in enumerate(
            arquivos,
            start=1
        ):

            nome_arquivo = (
                imagem.name
            )

            match_arquivo = (
                padrao_arquivo.search(
                    nome_arquivo
                )
            )

            if not match_arquivo:

                print(
                    f"[{idx}/{total_arquivos}] "
                    f"Ignorado (Fora do padrão): "
                    f"{nome_arquivo}"
                )

                continue

            prefixo = (
                match_arquivo.group(1)
                .upper()
            )

            lado = (
                match_arquivo.group(2)
                .upper()
            )

            registro = cadastro.get(
                normalizar_chave(
                    prefixo
                )
            )

            if registro:

                placas_esperadas = (
                    registro[
                        "placas"
                    ]
                )

                placas_texto = ", ".join(
                    formatar_placa(
                        placa
                    )
                    for placa
                    in placas_esperadas
                )

            else:

                placas_esperadas = []

                placas_texto = (
                    "Não cadastrado"
                )

            print(
                f"[{idx}/{total_arquivos}] "
                f"Analisando "
                f"{prefixo} "
                f"(Lado: {lado})...",
                end=" ",
                flush=True
            )

            if placas_esperadas:

                print(
                    f"Placa(s) esperada(s): "
                    f"{placas_texto}",
                    end=" ",
                    flush=True
                )

            else:

                print(
                    "Placa esperada: "
                    "Não cadastrada",
                    end=" ",
                    flush=True
                )

            # ====================================================
            # OCR
            # ====================================================

            textos_brutos = (
                ler_imagem(
                    processo_ocr,
                    imagem
                )
            )

            # ====================================================
            # ORDENAÇÃO
            # ====================================================

            textos_ordenados = (
                ordenar_textos(
                    textos_brutos
                )
            )

            # ====================================================
            # EXTRAÇÃO + VALIDAÇÃO
            # ====================================================

            dados_extraidos = (
                extrair_dados_veiculo(
                    textos_ordenados,
                    prefixo,
                    lado,
                    placas_esperadas
                )
            )

            if dados_extraidos[
                "placa_validada"
            ]:

                print(
                    f"PLACA OK "
                    f"({dados_extraidos['placa_ocr']} | "
                    f"{dados_extraidos['placa_pontuacao']}/7)"
                )

            elif placas_esperadas:

                print(
                    "PLACA NÃO VALIDADA"
                )

            else:

                print(
                    "SEM CADASTRO DE PLACA"
                )

            # ====================================================
            # SALVA
            # ====================================================

            if prefixo not in veiculos:

                veiculos[
                    prefixo
                ] = {}

            veiculos[
                prefixo
            ][
                lado
            ] = dados_extraidos

    finally:

        # ========================================================
        # ENCERRA OCR
        # ========================================================

        encerrar_ocr(
            processo_ocr
        )

    # ========================================================
    # CONSOLIDAÇÃO
    # ========================================================

    print(
        "\n"
        + "=" * 70
    )

    print(
        "CONSOLIDANDO DADOS E GERANDO PLANILHA"
    )

    print(
        "=" * 70
    )

    linhas_tabela = []

    for prefixo, lados in veiculos.items():

        # ----------------------------------------------------
        # PLACA CADASTRADA
        # ----------------------------------------------------

        registro = cadastro.get(
            normalizar_chave(
                prefixo
            )
        )

        if registro:

            placas_cadastradas = [
                formatar_placa(
                    placa
                )
                for placa
                in registro[
                    "placas"
                ]
            ]

            placa_informada = (
                " / ".join(
                    placas_cadastradas
                )
            )

        else:

            placa_informada = (
                "Não cadastrada"
            )

        # ----------------------------------------------------
        # PLACA OCR
        # ----------------------------------------------------

        placa_ocr = (
            "Não validada"
        )

        placa_status = (
            "Não encontrada"
        )

        # Prioridade:
        #
        # TR → FR → LD → LE

        ordem_lados = [
            "TR",
            "FR",
            "LD",
            "LE",
        ]

        for lado_preferencial in (
            ordem_lados
        ):

            dados_lado = lados.get(
                lado_preferencial
            )

            if (
                dados_lado
                and
                dados_lado.get(
                    "placa_validada"
                )
            ):

                placa_ocr = (
                    dados_lado.get(
                        "placa_ocr"
                    )
                )

                placa_status = (
                    f"OK "
                    f"({dados_lado.get('placa_pontuacao')}/7)"
                )

                break

        # ----------------------------------------------------
        # LINHA
        # ----------------------------------------------------

        linha = {

            "Prefixo":
                prefixo,

            "Frente (FR)":
                formatar_celula(
                    lados.get(
                        "FR"
                    )
                ),

            "Direita (LD)":
                formatar_celula(
                    lados.get(
                        "LD"
                    )
                ),

            "Esquerda (LE)":
                formatar_celula(
                    lados.get(
                        "LE"
                    )
                ),

            "Trás (TR)":
                formatar_celula(
                    lados.get(
                        "TR"
                    )
                ),

            "Placa informada":
                placa_informada,

            "Placa OCR":
                placa_ocr,

            "Status placa":
                placa_status,
        }

        linhas_tabela.append(
            linha
        )

    # ========================================================
    # DATAFRAME
    # ========================================================

    colunas_relatorio = [

        "Prefixo",

        "Frente (FR)",

        "Direita (LD)",

        "Esquerda (LE)",

        "Trás (TR)",

        "Placa informada",

        "Placa OCR",

        "Status placa",
    ]

    df = pd.DataFrame(
        linhas_tabela,
        columns=colunas_relatorio
    )

    # ========================================================
    # DIRETÓRIO DE RESULTADOS
    # ========================================================

    RESULTADOS_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    caminho_excel = (
        RESULTADOS_DIR
        /
        "relatorio_evasao.xlsx"
    )

    # ========================================================
    # COLUNAS DOS LADOS
    # ========================================================

    colunas_lados = [

        "Frente (FR)",

        "Direita (LD)",

        "Esquerda (LE)",

        "Trás (TR)",
    ]

    # ========================================================
    # ESTILIZAÇÃO
    # ========================================================

    styler = df.style

    if hasattr(
        styler,
        "map"
    ):

        styled_df = styler.map(
            aplicar_cores,
            subset=colunas_lados
        )

    else:

        styled_df = styler.applymap(
            aplicar_cores,
            subset=colunas_lados
        )

    # ========================================================
    # EXPORTAÇÃO
    # ========================================================

    styled_df.to_excel(
        caminho_excel,
        index=False,
        engine="openpyxl"
    )

    # ========================================================
    # FINAL
    # ========================================================

    print(
        "Processamento em lote "
        "finalizado com sucesso!"
    )

    print(
        f"Total de veículos processados: "
        f"{len(veiculos)}"
    )

    print(
        f"Planilha gerada em: "
        f"{caminho_excel}\n"
    )


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    main()
