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
# DIRETÓRIOS DE ENTRADA
# ============================================================

DOWNLOADS_DIR = (
    BASE_DIR
    / "downloads"
)


PDF_DIR = (
    DOWNLOADS_DIR
    / "pdf"
)


EXCEL_DIR = (
    DOWNLOADS_DIR
    / "excel"
)


IMAGENS_DIR = (
    DOWNLOADS_DIR
    / "imagens"
)


# ============================================================
# DIRETÓRIOS DE SAÍDA
# ============================================================

OUTPUT_DIR = (
    BASE_DIR
    / "resultados"
)


PDF_RESULTS_DIR = (
    OUTPUT_DIR
    / "pdf"
)


EXCEL_RESULTS_DIR = (
    OUTPUT_DIR
    / "excel"
)


# ============================================================
# DIRETÓRIOS DE DEBUG
# ============================================================

DEBUG_DIR = (
    BASE_DIR
    / "debug"
)


DEBUG_YOLO_DIR = (
    DEBUG_DIR
    / "yolo"
)


DEBUG_OCR_DIR = (
    DEBUG_DIR
    / "ocr"
)


DEBUG_SR_DIR = (
    DEBUG_DIR
    / "ocr_sr"
)


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
# FUNÇÃO PARA GARANTIR DIRETÓRIO
# ============================================================

def ensure_dir(path: Path) -> Path:
    """
    Cria o diretório caso ele não exista.
    """

    path.mkdir(
        parents=True,
        exist_ok=True
    )

    return path


# ============================================================
# CRIA ESTRUTURA DE DIRETÓRIOS
# ============================================================

def inicializar_diretorios():

    diretorios = [
        DOWNLOADS_DIR,
        PDF_DIR,
        EXCEL_DIR,
        IMAGENS_DIR,

        OUTPUT_DIR,
        PDF_RESULTS_DIR,
        EXCEL_RESULTS_DIR,

        DEBUG_DIR,
        DEBUG_YOLO_DIR,
        DEBUG_OCR_DIR,
        DEBUG_SR_DIR,
    ]

    for diretorio in diretorios:
        ensure_dir(diretorio)


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
# INICIALIZA
# ============================================================

inicializar_diretorios()