import sys
import json
import base64

import numpy as np
import cv2
from paddleocr import PaddleOCR


# ============================================================
# CONFIGURAÇÕES
# ============================================================

LARGURA_MINIMA = 600
MAX_SIDE_LIMIT = 4000

OCR_CONF_MINIMA = 0.70


# ============================================================
# MAPAS OCR AUXILIAR
# ============================================================

NUMERO_PARA_LETRA = {
    "0": "O",
    "1": "I",
    "2": "Z",
    "3": "B",
    "4": "A",
    "5": "S",
    "6": "G",
    "7": "T",
    "8": "B",
    "9": "G",
}


LETRA_PARA_NUMERO = {
    "O": "0",
    "Q": "0",
    "D": "0",
    "I": "1",
    "L": "1",
    "Z": "2",
    "B": "8",
    "A": "4",
    "S": "5",
    "G": "6",
    "T": "7",
}


# ============================================================
# LIMPA TEXTO
# ============================================================

def limpar_texto_ocr(texto):

    if texto is None:
        return ""

    texto = str(
        texto
    ).upper()

    return "".join(
        char
        for char in texto
        if (
            "A" <= char <= "Z"
            or "0" <= char <= "9"
        )
    )


# ============================================================
# PREPARA TEXTO PARA AUXILIAR
# ============================================================

def preparar_para_auxiliar(texto):

    texto = limpar_texto_ocr(
        texto
    )

    if len(texto) == 8:
        return texto[1:]

    return texto


# ============================================================
# OCR AUXILIAR
# ============================================================

def OCR_auxiliar(texto):

    if not texto:
        return ""

    texto = limpar_texto_ocr(
        texto
    )

    if len(texto) != 7:
        return texto

    chars = list(texto)

    # --------------------------------------------------------
    # POSIÇÕES 1-3
    # DEVEM SER LETRAS
    # --------------------------------------------------------

    for i in range(0, 3):

        char = chars[i]

        if char.isdigit():

            chars[i] = (
                NUMERO_PARA_LETRA.get(
                    char,
                    char
                )
            )

    # --------------------------------------------------------
    # POSIÇÃO 4
    # DEVE SER NÚMERO
    # --------------------------------------------------------

    char = chars[3]

    if char.isalpha():

        chars[3] = (
            LETRA_PARA_NUMERO.get(
                char,
                char
            )
        )

    # --------------------------------------------------------
    # POSIÇÃO 5
    # NÃO ALTERAR
    # --------------------------------------------------------

    # intencionalmente vazio

    # --------------------------------------------------------
    # POSIÇÃO 6
    # DEVE SER NÚMERO
    # --------------------------------------------------------

    char = chars[5]

    if char.isalpha():

        chars[5] = (
            LETRA_PARA_NUMERO.get(
                char,
                char
            )
        )

    # --------------------------------------------------------
    # POSIÇÃO 7
    # DEVE SER NÚMERO
    # --------------------------------------------------------

    char = chars[6]

    if char.isalpha():

        chars[6] = (
            LETRA_PARA_NUMERO.get(
                char,
                char
            )
        )

    return "".join(chars)


# ============================================================
# VALIDA PLACA
# ============================================================

def placa_valida(texto):

    if not texto:
        return False

    if len(texto) != 7:
        return False

    if not (
        texto[0].isalpha()
        and texto[1].isalpha()
        and texto[2].isalpha()
    ):
        return False

    if not texto[3].isdigit():
        return False

    if not texto[4].isalnum():
        return False

    if not texto[5].isdigit():
        return False

    if not texto[6].isdigit():
        return False

    return True


# ============================================================
# RESULTADO SATISFATÓRIO
# ============================================================

def resultado_satisfatorio(
    texto,
    confianca
):

    if not texto:
        return False

    if not placa_valida(texto):
        return False

    if confianca < OCR_CONF_MINIMA:
        return False

    return True


# ============================================================
# LIMITA TAMANHO
# ============================================================

def limitar_tamanho_imagem(
    image,
    max_side=MAX_SIDE_LIMIT
):

    if (
        image is None
        or image.size == 0
    ):
        return image

    h, w = image.shape[:2]

    maior_lado = max(
        h,
        w
    )

    if maior_lado <= max_side:
        return image

    escala = (
        max_side
        / float(maior_lado)
    )

    nova_w = max(
        1,
        int(
            round(
                w * escala
            )
        )
    )

    nova_h = max(
        1,
        int(
            round(
                h * escala
            )
        )
    )

    return cv2.resize(
        image,
        (
            nova_w,
            nova_h
        ),
        interpolation=cv2.INTER_AREA
    )


# ============================================================
# PREPARA VARIANTES
# ============================================================

def preparar_variantes(crop):

    if (
        crop is None
        or crop.size == 0
    ):

        return None, None

    altura, largura = (
        crop.shape[:2]
    )

    escala = max(
        3.0,
        LARGURA_MINIMA
        / max(
            largura,
            1
        )
    )

    maior_dimensao = max(
        altura * escala,
        largura * escala
    )

    if (
        maior_dimensao
        > MAX_SIDE_LIMIT
    ):

        escala = (
            MAX_SIDE_LIMIT
            / float(
                max(
                    altura,
                    largura
                )
            )
        )

    crop_cubic = cv2.resize(
        crop,
        None,
        fx=escala,
        fy=escala,
        interpolation=cv2.INTER_CUBIC
    )

    crop_cubic = limitar_tamanho_imagem(
        crop_cubic
    )

    # ========================================================
    # CLAHE
    # ========================================================

    gray = cv2.cvtColor(
        crop_cubic,
        cv2.COLOR_BGR2GRAY
    )

    clahe = cv2.createCLAHE(
        clipLimit=1.2,
        tileGridSize=(8, 8)
    )

    gray_clahe = clahe.apply(
        gray
    )

    crop_clahe = cv2.cvtColor(
        gray_clahe,
        cv2.COLOR_GRAY2BGR
    )

    crop_clahe = limitar_tamanho_imagem(
        crop_clahe
    )

    return (
        crop_cubic,
        crop_clahe
    )


# ============================================================
# PARSE PADDLEOCR
# ============================================================

def parse_ocr_result(
    resultado
):

    textos = []
    confiancas = []

    if resultado is None:
        return "", 0.0

    try:

        dados = resultado.json.get(
            "res",
            {}
        )

        for texto in dados.get(
            "rec_texts",
            []
        ):

            if texto:
                textos.append(
                    str(texto)
                )

        for score in dados.get(
            "rec_scores",
            []
        ):

            try:

                confiancas.append(
                    float(score)
                )

            except Exception:
                pass

    except Exception:

        return "", 0.0

    texto_bruto = "".join(
        textos
    )

    texto_limpo = (
        limpar_texto_ocr(
            texto_bruto
        )
    )

    confianca = (
        float(
            np.mean(
                confiancas
            )
        )
        if confiancas
        else 0.0
    )

    return (
        texto_limpo,
        confianca
    )


# ============================================================
# EXECUTA OCR
#
# Sem tqdm.
# Sem prints por imagem.
# ============================================================

def executar_ocr_variante(
    ocr,
    imagens,
    indices
):

    resultados = {}

    if not indices:
        return resultados

    for start in range(
        0,
        len(indices),
        8
    ):

        sub_indices = indices[
            start:start + 8
        ]

        sub_imagens = [
            limitar_tamanho_imagem(
                imagens[i]
            )
            for i in sub_indices
        ]

        try:

            outputs = ocr.predict(
                sub_imagens
            )

            outputs = list(
                outputs
            )

            for indice, output in zip(
                sub_indices,
                outputs
            ):

                resultados[
                    indice
                ] = parse_ocr_result(
                    output
                )

        except Exception:

            # ------------------------------------------------
            # FALLBACK INDIVIDUAL
            # ------------------------------------------------

            for indice, imagem in zip(
                sub_indices,
                sub_imagens
            ):

                try:

                    outputs = list(
                        ocr.predict(
                            limitar_tamanho_imagem(
                                imagem
                            )
                        )
                    )

                    output = (
                        outputs[0]
                        if outputs
                        else None
                    )

                    resultados[
                        indice
                    ] = parse_ocr_result(
                        output
                    )

                except Exception:

                    resultados[
                        indice
                    ] = (
                        "",
                        0.0
                    )

    return resultados


# ============================================================
# PREPARA BATCH
# ============================================================

def preparar_batch(
    items
):

    imagens = []
    metadados = []
    erros = []

    for item in items:

        task_id = item.get(
            "task_id",
            "sem_id"
        )

        try:

            img_bytes = (
                base64.b64decode(
                    item[
                        "crop_b64"
                    ]
                )
            )

            array = np.frombuffer(
                img_bytes,
                np.uint8
            )

            crop = cv2.imdecode(
                array,
                cv2.IMREAD_COLOR
            )

            if (
                crop is None
                or crop.size == 0
            ):

                erros.append(
                    {
                        "task_id":
                            task_id,

                        "error":
                            "Crop inválido"
                    }
                )

                continue

            crop = (
                limitar_tamanho_imagem(
                    crop
                )
            )

            (
                crop_cubic,
                crop_clahe
            ) = preparar_variantes(
                crop
            )

            if (
                crop_cubic is None
                or crop_clahe is None
            ):

                erros.append(
                    {
                        "task_id":
                            task_id,

                        "error":
                            "Falha preparando variantes"
                    }
                )

                continue

            imagens.append(
                crop_cubic
            )

            metadados.append(
                {
                    "task_id":
                        task_id,

                    "variant":
                        "CUBIC"
                }
            )

            imagens.append(
                crop_clahe
            )

            metadados.append(
                {
                    "task_id":
                        task_id,

                    "variant":
                        "CUBIC + CLAHE"
                }
            )

        except Exception as e:

            erros.append(
                {
                    "task_id":
                        task_id,

                    "error":
                        str(e)
                }
            )

    return (
        imagens,
        metadados,
        erros
    )


# ============================================================
# PROCESSA BATCH
# ============================================================

def processar_batch(
    ocr,
    items
):

    (
        imagens,
        metadados,
        erros
    ) = preparar_batch(
        items
    )

    resultados = list(
        erros
    )

    if not imagens:
        return resultados

    # ========================================================
    # ÍNDICES
    # ========================================================

    indices_cubic = [
        i
        for i, meta
        in enumerate(
            metadados
        )
        if meta["variant"]
        == "CUBIC"
    ]

    indices_clahe = [
        i
        for i, meta
        in enumerate(
            metadados
        )
        if meta["variant"]
        == "CUBIC + CLAHE"
    ]

    # ========================================================
    # OCR CUBIC
    # ========================================================

    resultados_cubic = (
        executar_ocr_variante(
            ocr,
            imagens,
            indices_cubic
        )
    )

    # ========================================================
    # OCR CLAHE
    # ========================================================

    resultados_clahe = (
        executar_ocr_variante(
            ocr,
            imagens,
            indices_clahe
        )
    )

    # ========================================================
    # AGRUPA
    # ========================================================

    agrupado = {}

    for indice, resultado in (
        resultados_cubic.items()
    ):

        task_id = (
            metadados[indice][
                "task_id"
            ]
        )

        if task_id not in agrupado:
            agrupado[
                task_id
            ] = {}

        agrupado[
            task_id
        ]["CUBIC"] = resultado

    for indice, resultado in (
        resultados_clahe.items()
    ):

        task_id = (
            metadados[indice][
                "task_id"
            ]
        )

        if task_id not in agrupado:
            agrupado[
                task_id
            ] = {}

        agrupado[
            task_id
        ]["CLAHE"] = resultado

    # ========================================================
    # CONSOLIDA
    # ========================================================

    for task_id, variantes in (
        agrupado.items()
    ):

        cubic = variantes.get(
            "CUBIC",
            (
                "",
                0.0
            )
        )

        clahe = variantes.get(
            "CLAHE",
            (
                "",
                0.0
            )
        )

        (
            cubic_original,
            cubic_conf
        ) = cubic

        (
            clahe_original,
            clahe_conf
        ) = clahe

        # ====================================================
        # AUXILIAR
        # ====================================================

        cubic_sem_primeiro = (
            preparar_para_auxiliar(
                cubic_original
            )
        )

        clahe_sem_primeiro = (
            preparar_para_auxiliar(
                clahe_original
            )
        )

        cubic_auxiliar = (
            OCR_auxiliar(
                cubic_sem_primeiro
            )
        )

        clahe_auxiliar = (
            OCR_auxiliar(
                clahe_sem_primeiro
            )
        )

        # ====================================================
        # ESCOLHE MELHOR NORMAL
        # ====================================================

        if clahe_conf > cubic_conf:

            normal_original = (
                clahe_original
            )

            normal_sem_primeiro = (
                clahe_sem_primeiro
            )

            normal_aux = (
                clahe_auxiliar
            )

            normal_conf = (
                clahe_conf
            )

            normal_variant = (
                "CUBIC + CLAHE"
            )

        else:

            normal_original = (
                cubic_original
            )

            normal_sem_primeiro = (
                cubic_sem_primeiro
            )

            normal_aux = (
                cubic_auxiliar
            )

            normal_conf = (
                cubic_conf
            )

            normal_variant = (
                "CUBIC"
            )

        # ====================================================
        # RESULTADO FINAL
        #
        # SR DESATIVADO
        # ====================================================

        ocr_best_original = (
            normal_original
        )

        ocr_best_sem_primeiro = (
            normal_sem_primeiro
        )

        ocr_best_normal = (
            normal_aux
        )

        confidence_best_normal = (
            normal_conf
        )

        ocr_best = (
            ocr_best_normal
        )

        confidence_best = (
            confidence_best_normal
        )

        best_variant = (
            normal_variant
        )

        auxiliar_alterou = (
            ocr_best_original
            != ocr_best_normal
        )

        # ====================================================
        # RESULTADO
        # ====================================================

        resultados.append(
            {
                "task_id":
                    task_id,

                # --------------------------------------------
                # CUBIC
                # --------------------------------------------

                "ocr_cubic":
                    cubic_original,

                "confidence_cubic":
                    cubic_conf,

                "ocr_cubic_sem_primeiro":
                    cubic_sem_primeiro,

                "ocr_cubic_auxiliar":
                    cubic_auxiliar,

                # --------------------------------------------
                # CLAHE
                # --------------------------------------------

                "ocr_clahe":
                    clahe_original,

                "confidence_clahe":
                    clahe_conf,

                "ocr_clahe_sem_primeiro":
                    clahe_sem_primeiro,

                "ocr_clahe_auxiliar":
                    clahe_auxiliar,

                # --------------------------------------------
                # NORMAL
                # --------------------------------------------

                "ocr_best_original":
                    ocr_best_original,

                "ocr_best_sem_primeiro":
                    ocr_best_sem_primeiro,

                "ocr_best_normal":
                    ocr_best_normal,

                "confidence_best_normal":
                    confidence_best_normal,

                # --------------------------------------------
                # SR
                #
                # Mantido somente por compatibilidade.
                # Não executado.
                # --------------------------------------------

                "sr_used":
                    False,

                "sr_ocr":
                    "",

                "sr_ocr_sem_primeiro":
                    "",

                "sr_ocr_auxiliar":
                    "",

                "sr_confidence":
                    0.0,

                "sr_variant":
                    "",

                "resultado_gracas_sr":
                    False,

                # --------------------------------------------
                # FINAL
                # --------------------------------------------

                "ocr_best":
                    ocr_best,

                "confidence_best":
                    confidence_best,

                "best_variant":
                    best_variant,

                "auxiliar_alterou":
                    auxiliar_alterou,
            }
        )

    return resultados


# ============================================================
# MAIN
# ============================================================

def main():

    # ========================================================
    # INICIALIZA OCR
    # ========================================================

    try:

        ocr = PaddleOCR(
            lang="en",
            device="gpu:0",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
        )

    except Exception as e:

        print(
            f"[OCR] Erro inicializando: {e}",
            file=sys.stderr,
            flush=True
        )

        sys.exit(1)

    # ========================================================
    # MENSAGENS INICIAIS
    #
    # Apenas uma vez.
    # ========================================================

    print(
        "[OCR] PaddleOCR pronto.",
        file=sys.stderr,
        flush=True
    )

    print(
        "[OCR] CUBIC + CLAHE | SR desativado.",
        file=sys.stderr,
        flush=True
    )

    # ========================================================
    # STDIN
    # ========================================================

    for line in sys.stdin:

        line = line.strip()

        if not line:
            continue

        try:

            req = json.loads(
                line
            )

            # =================================================
            # EXIT
            # =================================================

            if req.get(
                "command"
            ) == "exit":

                break

            # =================================================
            # BATCH
            # =================================================

            if req.get(
                "command"
            ) == "ocr_batch":

                items = req.get(
                    "images",
                    []
                )

                response = {
                    "results":
                        processar_batch(
                            ocr,
                            items
                        )
                }

            # =================================================
            # INDIVIDUAL
            # =================================================

            else:

                response = {
                    "results":
                        processar_batch(
                            ocr,
                            [req]
                        )
                }

            # =================================================
            # RETORNO JSON
            # =================================================

            sys.stdout.write(
                json.dumps(
                    response,
                    ensure_ascii=False
                )
                + "\n"
            )

            sys.stdout.flush()

        except Exception as e:

            print(
                f"[OCR] Erro: {e}",
                file=sys.stderr,
                flush=True
            )

            sys.stdout.write(
                json.dumps(
                    {
                        "results": [],
                        "error": str(e)
                    },
                    ensure_ascii=False
                )
                + "\n"
            )

            sys.stdout.flush()


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()