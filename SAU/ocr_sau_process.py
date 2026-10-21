from __future__ import annotations

# ============================================================
# IMPORTS
# ============================================================

import base64
import json
import sys

import cv2
import numpy as np

from paddleocr import PaddleOCR


# ============================================================
# CONFIGURAÇÃO
# ============================================================

DEVICE = "gpu:0"

TEXT_REC_SCORE_THRESH = 0.0


# ============================================================
# INICIALIZAÇÃO DO PADDLEOCR
# ============================================================

def criar_ocr():
    """
    Cria uma única instância do PaddleOCR.

    O modelo permanece carregado durante toda a execução,
    evitando reinicialização a cada imagem.
    """

    return PaddleOCR(
        device=DEVICE,

        # Para esta análise queremos OCR geral de texto.
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,

        # Aceita textos latinos/veículos.
        lang="en",

        # Não descarta detecções por confiança aqui.
        # A filtragem será feita posteriormente.
        text_rec_score_thresh=TEXT_REC_SCORE_THRESH,
    )


# ============================================================
# DECODIFICAÇÃO DA IMAGEM
# ============================================================

def decodificar_imagem(crop_b64: str):
    """
    Converte base64 em imagem OpenCV.
    """

    try:
        dados = base64.b64decode(crop_b64)

        array = np.frombuffer(
            dados,
            dtype=np.uint8
        )

        imagem = cv2.imdecode(
            array,
            cv2.IMREAD_COLOR
        )

    except Exception as erro:
        raise RuntimeError(
            f"Erro ao decodificar imagem: {erro}"
        ) from erro

    if imagem is None:
        raise RuntimeError(
            "Não foi possível decodificar a imagem."
        )

    return imagem


# ============================================================
# CONVERSÃO SEGURA PARA LISTA
# ============================================================

def converter_para_lista(valor):
    """
    Converte numpy arrays e estruturas semelhantes em listas
    comuns para poderem ser serializadas em JSON.
    """

    if valor is None:
        return None

    if hasattr(valor, "tolist"):
        return valor.tolist()

    if isinstance(valor, (list, tuple)):
        return [
            converter_para_lista(item)
            for item in valor
        ]

    return valor


# ============================================================
# EXTRAÇÃO DO RESULTADO
# ============================================================

def obter_dado(resultado, chave, padrao=None):
    """
    Obtém dados de um resultado PaddleOCR.

    Compatível com os objetos retornados pelas versões recentes
    do PaddleOCR, que permitem acesso por chave.
    """

    # --------------------------------------------------------
    # Tentativa 1: acesso direto
    # --------------------------------------------------------

    try:
        valor = resultado.get(
            chave,
            padrao
        )

        if valor is not None:
            return valor

    except Exception:
        pass

    # --------------------------------------------------------
    # Tentativa 2: acesso por []
    # --------------------------------------------------------

    try:
        return resultado[chave]

    except Exception:
        pass

    # --------------------------------------------------------
    # Tentativa 3: .json
    # --------------------------------------------------------

    try:
        json_resultado = resultado.json

        if isinstance(
            json_resultado,
            str
        ):
            json_resultado = json.loads(
                json_resultado
            )

        if isinstance(
            json_resultado,
            dict
        ):

            # Algumas versões colocam os dados dentro
            # de "res".
            if isinstance(
                json_resultado.get("res"),
                dict
            ):
                json_resultado = (
                    json_resultado["res"]
                )

            return json_resultado.get(
                chave,
                padrao
            )

    except Exception:
        pass

    return padrao


# ============================================================
# NORMALIZAÇÃO DE BOX
# ============================================================

def normalizar_box(box):
    """
    Converte um box para uma lista JSON.
    """

    if box is None:
        return None

    convertido = converter_para_lista(box)

    if isinstance(
        convertido,
        list
    ):
        return convertido

    return None


# ============================================================
# ORDENAÇÃO DAS DETECÇÕES
# ============================================================

def ordenar_deteccoes(deteccoes):
    """
    Ordena os textos aproximadamente de cima para baixo
    e da esquerda para a direita.

    Isso facilita a visualização e também deixa o resultado
    mais previsível para as etapas seguintes.
    """

    def chave(item):

        box = item.get("box")

        if not box:
            return (
                0,
                0
            )

        try:

            # rec_box:
            # [x_min, y_min, x_max, y_max]

            if (
                isinstance(box, list)
                and len(box) == 4
                and all(
                    isinstance(v, (int, float))
                    for v in box
                )
            ):
                x = box[0]
                y = box[1]

                return (
                    y,
                    x
                )

            # rec_poly:
            # [[x,y], [x,y], ...]

            if isinstance(
                box,
                list
            ) and box:

                pontos = [
                    ponto
                    for ponto in box
                    if (
                        isinstance(ponto, list)
                        and len(ponto) >= 2
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
        deteccoes,
        key=chave
    )


# ============================================================
# PROCESSAMENTO OCR
# ============================================================

def processar_imagem(ocr, imagem):
    """
    Executa OCR e retorna as detecções individualmente.
    """

    try:

        resultados = ocr.predict(
            imagem
        )

    except Exception as erro:

        raise RuntimeError(
            f"Erro durante PaddleOCR: {erro}"
        ) from erro

    deteccoes = []

    for resultado in resultados:

        rec_texts = obter_dado(
            resultado,
            "rec_texts",
            []
        )

        rec_scores = obter_dado(
            resultado,
            "rec_scores",
            []
        )

        rec_boxes = obter_dado(
            resultado,
            "rec_boxes",
            []
        )

        rec_polys = obter_dado(
            resultado,
            "rec_polys",
            []
        )

        if rec_texts is None:
            rec_texts = []

        if rec_scores is None:
            rec_scores = []

        if rec_boxes is None:
            rec_boxes = []

        if rec_polys is None:
            rec_polys = []

        for indice, texto in enumerate(
            rec_texts
        ):

            if texto is None:
                continue

            texto = str(
                texto
            ).strip()

            if not texto:
                continue

            # ------------------------------------------------
            # CONFIANÇA
            # ------------------------------------------------

            try:

                if indice < len(
                    rec_scores
                ):
                    confianca = float(
                        rec_scores[indice]
                    )
                else:
                    confianca = 0.0

            except Exception:

                confianca = 0.0

            # ------------------------------------------------
            # BOX
            # ------------------------------------------------

            box = None

            if indice < len(
                rec_boxes
            ):

                box = normalizar_box(
                    rec_boxes[indice]
                )

            # ------------------------------------------------
            # POLYGON FALLBACK
            # ------------------------------------------------

            if box is None:

                if indice < len(
                    rec_polys
                ):

                    box = normalizar_box(
                        rec_polys[indice]
                    )

            # ------------------------------------------------
            # RESULTADO INDIVIDUAL
            # ------------------------------------------------

            deteccoes.append(
                {
                    "texto": texto,
                    "confianca": confianca,
                    "box": box,
                }
            )

    return ordenar_deteccoes(
        deteccoes
    )


# ============================================================
# PROCESSAMENTO DE UMA REQUISIÇÃO
# ============================================================

def processar_requisicao(
    ocr,
    requisicao
):
    """
    Processa o comando recebido via stdin.
    """

    command = requisicao.get(
        "command"
    )

    # ========================================================
    # EXIT
    # ========================================================

    if command == "exit":

        return {
            "exit": True
        }

    # ========================================================
    # OCR BATCH
    # ========================================================

    if command != "ocr_batch":

        return {
            "error": (
                f"Comando desconhecido: "
                f"{command!r}"
            )
        }

    imagens = requisicao.get(
        "images",
        []
    )

    if not isinstance(
        imagens,
        list
    ):

        return {
            "error": (
                "'images' precisa ser uma lista."
            )
        }

    resultados = []

    for imagem_info in imagens:

        task_id = imagem_info.get(
            "task_id"
        )

        crop_b64 = imagem_info.get(
            "crop_b64"
        )

        if not crop_b64:

            resultados.append(
                {
                    "task_id": task_id,
                    "error": (
                        "Campo 'crop_b64' não informado."
                    )
                }
            )

            continue

        try:

            imagem = decodificar_imagem(
                crop_b64
            )

            deteccoes = processar_imagem(
                ocr,
                imagem
            )

            # ------------------------------------------------
            # TEXTO COMPLETO PARA COMPATIBILIDADE
            # ------------------------------------------------
            #
            # IMPORTANTE:
            # usamos espaço entre detecções.
            #
            # O detalhamento individual fica em "ocr_items".
            # ------------------------------------------------

            ocr_best = " ".join(
                item["texto"]
                for item in deteccoes
            )

            if deteccoes:

                confidence_best = max(
                    item["confianca"]
                    for item in deteccoes
                )

            else:

                confidence_best = 0.0

            resultados.append(
                {
                    "task_id": task_id,

                    # ------------------------------------------------
                    # DETECÇÕES INDIVIDUAIS
                    # ------------------------------------------------

                    "ocr_items": deteccoes,

                    # ------------------------------------------------
                    # CAMPOS DE COMPATIBILIDADE
                    # ------------------------------------------------

                    "ocr_best": ocr_best,
                    "confidence_best": confidence_best,

                    # ------------------------------------------------
                    # CAMPOS VAZIOS PARA NÃO QUEBRAR ESTRUTURA
                    # ------------------------------------------------

                    "ocr_best_original": ocr_best,
                    "confidence_best_original": confidence_best,

                    "ocr_best_normal": ocr_best,
                    "confidence_best_normal": confidence_best,

                    "ocr_best_sem_primeiro": ocr_best,
                    "confidence_best_sem_primeiro": confidence_best,

                    "sr_ocr": "",
                    "sr_confidence": 0.0,

                    "sr_ocr_auxiliar": "",
                    "sr_ocr_sem_primeiro": "",

                    "ocr_cubic": ocr_best,
                    "confidence_cubic": confidence_best,

                    "ocr_cubic_auxiliar": ocr_best,
                    "ocr_cubic_sem_primeiro": ocr_best,

                    "ocr_clahe": ocr_best,
                    "confidence_clahe": confidence_best,

                    "ocr_clahe_auxiliar": ocr_best,
                    "ocr_clahe_sem_primeiro": ocr_best,
                }
            )

        except Exception as erro:

            resultados.append(
                {
                    "task_id": task_id,
                    "error": str(erro)
                }
            )

    return {
        "results": resultados
    }


# ============================================================
# LOOP PRINCIPAL
# ============================================================

def main():

    # ========================================================
    # CRIA OCR UMA ÚNICA VEZ
    # ========================================================

    try:

        ocr = criar_ocr()

    except Exception as erro:

        resposta = {
            "error": (
                "Falha ao inicializar PaddleOCR: "
                f"{erro}"
            )
        }

        print(
            json.dumps(
                resposta,
                ensure_ascii=False
            ),
            flush=True
        )

        return

    # ========================================================
    # LOOP JSON
    # ========================================================

    for linha in sys.stdin:

        linha = linha.strip()

        if not linha:
            continue

        try:

            requisicao = json.loads(
                linha
            )

        except json.JSONDecodeError as erro:

            print(
                json.dumps(
                    {
                        "error": (
                            "JSON inválido: "
                            f"{erro}"
                        )
                    },
                    ensure_ascii=False
                ),
                flush=True
            )

            continue

        try:

            resposta = processar_requisicao(
                ocr,
                requisicao
            )

        except Exception as erro:

            resposta = {
                "error": str(erro)
            }

        print(
            json.dumps(
                resposta,
                ensure_ascii=False
            ),
            flush=True
        )

        # ====================================================
        # ENCERRAMENTO
        # ====================================================

        if resposta.get(
            "exit"
        ):

            break


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    main()