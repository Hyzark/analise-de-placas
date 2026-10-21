from pathlib import Path


# ============================================================
# DIRETÓRIO BASE DO PROJETO
# ============================================================

BASE_DIR = Path(__file__).resolve().parent


# ============================================================
# AMBIENTES VIRTUAIS
# ============================================================

VENV_YOLO_DIR = BASE_DIR / ".venv_yolo"
VENV_OCR_DIR = BASE_DIR / ".venv_ocr"


PYTHON_YOLO = (
    VENV_YOLO_DIR
    / "Scripts"
    / "python.exe"
)

PYTHON_OCR = (
    VENV_OCR_DIR
    / "Scripts"
    / "python.exe"
)


# ============================================================
# SCRIPTS
# ============================================================

DETECTOR_PY = (
    BASE_DIR
    / "detector_process.py"
)

OCR_PY = (
    BASE_DIR
    / "ocr_process.py"
)

OCR_SAU_PY = (
    BASE_DIR
    / "SAU"
    / "ocr_sau_process.py"
)

APP_PDF_PY = (
    BASE_DIR
    / "app_pdf_paddle.py"
)

VALIDADOR_EXCEL_PY = (
    BASE_DIR
    / "validador_excel.py"
)


# ============================================================
# MODELOS
# ============================================================

MODELS_DIR = (
    BASE_DIR
    / "models"
)


# YOLO ONNX 1024
YOLO_MODEL = (
    MODELS_DIR
    / "bestnew1024.onnx"
)


# Super Resolution FSRCNN
SR_MODEL = (
    MODELS_DIR
    / "FSRCNN_x2.pb"
)


# ============================================================
# SAÍDA
# ============================================================

OUTPUT_DIR = BASE_DIR / "resultados"


# ============================================================
# CONFIGURAÇÕES DO PIPELINE
# ============================================================

BATCH_SIZE = 64

QTD_IMAGENS_POR_PASSAGEM = 2

MAX_DIFERENCAS_APROVACAO = 3


# ============================================================
# CONFIGURAÇÕES YOLO
# ============================================================

YOLO_IMG_SIZE = 1024

YOLO_CONF_THRESHOLD = 0.35

YOLO_IOU_THRESHOLD = 0.45

YOLO_MARGEM_X = 0.15

YOLO_MARGEM_Y = 0.20

YOLO_ZOOM_SCALE = 1.5


# ============================================================
# CONFIGURAÇÕES OCR
# ============================================================

OCR_CONF_MINIMA = 0.70

OCR_LARGURA_MINIMA = 600

OCR_SUB_BATCH_SIZE = 8

OCR_LANG = "en"


# ============================================================
# SUPER RESOLUTION
# ============================================================

USE_SUPER_RESOLUTION = True

SR_ALGORITHM = "fsrcnn"

SR_SCALE = 2


# ============================================================
# GARANTE PASTA PAI DE UM ARQUIVO
# ============================================================

def ensure_parent(path: Path) -> Path:

    path.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    return path


# ============================================================
# RESULTADOS POR SEGMENTO, SEM SOBRESCRITA
# ============================================================

SEGMENTOS_RESULTADOS = {
    "fraudes": "Análise de Fraudes",
    "autoban": "Autoban — Isenção Judicial",
    "ecovias": "Ecovias - Evasão de Pedágio",
    "spvias": "SPVias — Evasão de Pedágio",
    "tamoios": "Tamoios — Evasão Free Flow",
    "tamoios_isencao": "Tamoios - Rota das Bandeiras — Isenção",
    "cnl_isencao": "CNL — Isenção Judicial",
    "via_colinas": "Via Colinas — Evasão de Pedágio",
}

def caminho_resultado(segmento: str, nome_arquivo: str) -> Path:
    """Retorna caminho exclusivo em resultados/<segmento>/, sem substituir arquivo."""
    pasta_segmento = OUTPUT_DIR / SEGMENTOS_RESULTADOS.get(segmento, segmento)
    pasta_segmento.mkdir(parents=True, exist_ok=True)

    nome = Path(nome_arquivo).name
    original = pasta_segmento / nome
    if not original.exists():
        return original

    caminho = Path(nome)
    indice = 2
    while True:
        candidato = pasta_segmento / f"{caminho.stem}{indice}{caminho.suffix}"
        if not candidato.exists():
            return candidato
        indice += 1

