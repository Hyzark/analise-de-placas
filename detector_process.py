import sys
import json
import base64
from pathlib import Path

import numpy as np
import cv2
import onnxruntime as ort


# ============================================================
# CONFIG
# ============================================================

try:
    from config import (
        YOLO_MODEL,
        YOLO_CONF_THRESHOLD,
        YOLO_MARGEM_X,
        YOLO_MARGEM_Y,
    )
except ImportError:
    YOLO_MODEL = Path("models/bestnew1024.onnx")
    YOLO_CONF_THRESHOLD = 0.35
    YOLO_MARGEM_X = 0.15
    YOLO_MARGEM_Y = 0.20


BATCH_SIZE = 64


# ============================================================
# LETTERBOX
# ============================================================

def letterbox(
    img,
    target_shape=(1024, 1024),
    color=(114, 114, 114)
):

    h, w = img.shape[:2]

    tw, th = target_shape

    r = min(
        tw / w,
        th / h
    )

    new_unpad = (
        int(round(w * r)),
        int(round(h * r))
    )

    dw = (
        tw - new_unpad[0]
    ) / 2.0

    dh = (
        th - new_unpad[1]
    ) / 2.0

    if (w, h) != new_unpad:

        resized = cv2.resize(
            img,
            new_unpad,
            interpolation=cv2.INTER_CUBIC
        )

    else:

        resized = img.copy()

    top = int(round(dh - 0.1))
    bottom = int(round(dh + 0.1))
    left = int(round(dw - 0.1))
    right = int(round(dw + 0.1))

    padded = cv2.copyMakeBorder(
        resized,
        top,
        bottom,
        left,
        right,
        cv2.BORDER_CONSTANT,
        value=color
    )

    return padded, r, dw, dh


# ============================================================
# TENSOR
# ============================================================

def imagem_para_tensor(image):

    input_img, r, dw, dh = letterbox(
        image,
        (1024, 1024)
    )

    blob = cv2.cvtColor(
        input_img,
        cv2.COLOR_BGR2RGB
    )

    blob = (
        blob.astype(np.float32)
        / 255.0
    )

    blob = np.transpose(
        blob,
        (2, 0, 1)
    )

    meta = {
        "r": r,
        "dw": dw,
        "dh": dh
    }

    return blob, meta


# ============================================================
# EXPANDE BOX
# ============================================================

def expandir_box(
    x1,
    y1,
    x2,
    y2,
    largura,
    altura
):

    bw = x2 - x1
    bh = y2 - y1

    extra_x = bw * YOLO_MARGEM_X
    extra_y = bh * YOLO_MARGEM_Y

    x1 = int(round(x1 - extra_x))
    y1 = int(round(y1 - extra_y))
    x2 = int(round(x2 + extra_x))
    y2 = int(round(y2 + extra_y))

    x1 = max(
        0,
        min(
            x1,
            largura - 1
        )
    )

    y1 = max(
        0,
        min(
            y1,
            altura - 1
        )
    )

    x2 = max(
        1,
        min(
            x2,
            largura
        )
    )

    y2 = max(
        1,
        min(
            y2,
            altura
        )
    )

    return x1, y1, x2, y2


# ============================================================
# PROCESSA SAÍDA YOLO
# ============================================================

def processar_saida(
    output,
    meta,
    original,
    image_path,
    task_id
):

    if output is None:

        return {
            "detected": False,
            "detections": []
        }

    try:

        if output.ndim == 3:
            preds = output[0]

        elif output.ndim == 2:
            preds = output

        else:

            return {
                "detected": False,
                "detections": []
            }

        if (
            preds.ndim != 2
            or preds.shape[1] < 6
        ):

            return {
                "detected": False,
                "detections": []
            }

        scores = preds[:, 4]

        mask = (
            scores
            >= YOLO_CONF_THRESHOLD
        )

        preds = preds[mask]
        scores = scores[mask]

        if len(preds) == 0:

            return {
                "detected": False,
                "detections": []
            }

        boxes = preds[:, :4].copy()

        h_orig, w_orig = (
            original.shape[:2]
        )

        candidates = []

        # ====================================================
        # CONVERTE DETECÇÕES
        # ====================================================

        for idx in range(len(boxes)):

            box = boxes[idx]

            confidence = float(
                scores[idx]
            )

            x1 = (
                box[0]
                - meta["dw"]
            ) / meta["r"]

            y1 = (
                box[1]
                - meta["dh"]
            ) / meta["r"]

            x2 = (
                box[2]
                - meta["dw"]
            ) / meta["r"]

            y2 = (
                box[3]
                - meta["dh"]
            ) / meta["r"]

            x1 = max(
                0,
                min(
                    int(round(x1)),
                    w_orig - 1
                )
            )

            y1 = max(
                0,
                min(
                    int(round(y1)),
                    h_orig - 1
                )
            )

            x2 = max(
                1,
                min(
                    int(round(x2)),
                    w_orig
                )
            )

            y2 = max(
                1,
                min(
                    int(round(y2)),
                    h_orig
                )
            )

            bw = x2 - x1
            bh = y2 - y1

            if (
                bw <= 0
                or bh <= 0
            ):
                continue

            aspect_ratio = (
                bw / bh
            )

            if (
                bw < 20
                or bh < 10
                or aspect_ratio < 0.8
                or aspect_ratio > 6.0
            ):
                continue

            candidates.append(
                {
                    "x1": x1,
                    "y1": y1,
                    "x2": x2,
                    "y2": y2,
                    "confidence": confidence
                }
            )

        if not candidates:

            return {
                "detected": False,
                "detections": []
            }

        # ====================================================
        # NMS
        # ====================================================

        nms_boxes = []
        nms_scores = []

        for candidate in candidates:

            x1 = candidate["x1"]
            y1 = candidate["y1"]
            x2 = candidate["x2"]
            y2 = candidate["y2"]

            nms_boxes.append(
                [
                    x1,
                    y1,
                    x2 - x1,
                    y2 - y1
                ]
            )

            nms_scores.append(
                candidate["confidence"]
            )

        indices = cv2.dnn.NMSBoxes(
            nms_boxes,
            nms_scores,
            YOLO_CONF_THRESHOLD,
            0.45
        )

        if (
            indices is None
            or len(indices) == 0
        ):

            return {
                "detected": False,
                "detections": []
            }

        indices = (
            np.array(indices)
            .reshape(-1)
        )

        candidates_nms = [
            candidates[int(i)]
            for i in indices
        ]

        # ====================================================
        # MAIOR CONFIANÇA PRIMEIRO
        # ====================================================

        candidates_nms.sort(
            key=lambda c: (
                c["confidence"]
            ),
            reverse=True
        )

        detections = []

        # ====================================================
        # GERA CROPS
        # ====================================================

        for numero, candidate in enumerate(
            candidates_nms,
            start=1
        ):

            (
                cx1,
                cy1,
                cx2,
                cy2
            ) = expandir_box(
                candidate["x1"],
                candidate["y1"],
                candidate["x2"],
                candidate["y2"],
                w_orig,
                h_orig
            )

            crop = original[
                cy1:cy2,
                cx1:cx2
            ].copy()

            if (
                crop is None
                or crop.size == 0
            ):
                continue

            success, encoded = (
                cv2.imencode(
                    ".png",
                    crop
                )
            )

            if not success:
                continue

            crop_b64 = (
                base64.b64encode(
                    encoded
                ).decode("utf-8")
            )

            detections.append(
                {
                    "detection_index":
                        numero,

                    "box":
                        [
                            cx1,
                            cy1,
                            cx2,
                            cy2
                        ],

                    "raw_box":
                        [
                            candidate["x1"],
                            candidate["y1"],
                            candidate["x2"],
                            candidate["y2"]
                        ],

                    "confidence":
                        candidate[
                            "confidence"
                        ],

                    "crop_b64":
                        crop_b64
                }
            )

        if not detections:

            return {
                "detected": False,
                "detections": []
            }

        # ====================================================
        # PRIMEIRO CROP
        # ====================================================

        primeira = detections[0]

        return {
            "detected": True,

            "detections":
                detections,

            "box":
                primeira["box"],

            "raw_box":
                primeira["raw_box"],

            "confidence":
                primeira["confidence"],

            "crop_b64":
                primeira["crop_b64"],

            "strategy":
                "YOLO 1024"
        }

    except Exception as e:

        return {
            "detected": False,
            "detections": [],
            "error": str(e)
        }


# ============================================================
# PREPARA BATCH
# ============================================================

def preparar_batch(items):

    tensors = []
    metas = []
    originais = []
    valid_indices = []

    for index, item in enumerate(
        items
    ):

        image_path = item.get(
            "image_path"
        )

        if not image_path:
            continue

        image_path = str(
            Path(
                image_path
            ).resolve()
        )

        if not Path(
            image_path
        ).exists():

            continue

        image = cv2.imread(
            image_path
        )

        if image is None:
            continue

        original = image.copy()

        blob, meta = (
            imagem_para_tensor(
                original
            )
        )

        tensors.append(blob)
        metas.append(meta)
        originais.append(original)
        valid_indices.append(index)

    if not tensors:

        return (
            None,
            [],
            [],
            []
        )

    batch = np.stack(
        tensors,
        axis=0
    )

    return (
        batch,
        metas,
        originais,
        valid_indices
    )


# ============================================================
# COMPLETA BATCH
# ============================================================

def completar_batch(batch):

    quantidade = batch.shape[0]

    if quantidade == BATCH_SIZE:
        return batch

    if quantidade == 0:
        return None

    padding = np.repeat(
        batch[-1:],
        BATCH_SIZE - quantidade,
        axis=0
    )

    return np.concatenate(
        [
            batch,
            padding
        ],
        axis=0
    )


# ============================================================
# INFERÊNCIA
# ============================================================

def inferencia(
    session,
    batch
):

    input_name = (
        session
        .get_inputs()[0]
        .name
    )

    batch_completo = (
        completar_batch(batch)
    )

    outputs = session.run(
        None,
        {
            input_name:
                batch_completo
        }
    )[0]

    return outputs


# ============================================================
# EXECUTA BATCH
#
# UMA ÚNICA TENTATIVA.
# SEM ZOOM.
# ============================================================

def executar_batch(
    session,
    items
):

    results = [
        {
            "detected": False,
            "detections": []
        }
        for _ in items
    ]

    total = len(items)

    if total == 0:
        return results

    for start in range(
        0,
        total,
        BATCH_SIZE
    ):

        sub_items = items[
            start:start + BATCH_SIZE
        ]

        (
            batch,
            metas,
            originais,
            valid_indices
        ) = preparar_batch(
            sub_items
        )

        if batch is None:
            continue

        try:

            outputs = inferencia(
                session,
                batch
            )

        except Exception as e:

            # erro continua sendo enviado
            # somente no stderr, nunca no
            # stdout usado pelo protocolo JSON

            print(
                f"[YOLO] Erro batch: {e}",
                file=sys.stderr,
                flush=True
            )

            continue

        for local_idx, original_idx in enumerate(
            valid_indices
        ):

            item = sub_items[
                original_idx
            ]

            result = processar_saida(
                outputs[
                    local_idx:
                    local_idx + 1
                ],
                metas[local_idx],
                originais[local_idx],
                item["image_path"],
                item["task_id"]
            )

            results[
                start + original_idx
            ] = result

    return results


# ============================================================
# MAIN
# ============================================================

def main():

    # ========================================================
    # ORT
    # ========================================================

    try:

        ort.preload_dlls(
            directory=""
        )

    except Exception:
        pass

    providers = [
        "CUDAExecutionProvider",
        "CPUExecutionProvider"
    ]

    try:

        session = (
            ort.InferenceSession(
                str(YOLO_MODEL),
                providers=providers
            )
        )

    except Exception as e:

        print(
            f"[YOLO] Erro carregando modelo: {e}",
            file=sys.stderr,
            flush=True
        )

        sys.exit(1)

    # ========================================================
    # INFORMAÇÕES INICIAIS
    #
    # Apenas uma vez.
    # ========================================================

    print(
        "[YOLO] Modelo carregado.",
        file=sys.stderr,
        flush=True
    )

    print(
        "[YOLO] Providers: "
        f"{session.get_providers()}",
        file=sys.stderr,
        flush=True
    )

    print(
        "[YOLO] Input: "
        f"{session.get_inputs()[0].shape}",
        file=sys.stderr,
        flush=True
    )

    print(
        f"[YOLO] Confiança mínima: "
        f"{YOLO_CONF_THRESHOLD}",
        file=sys.stderr,
        flush=True
    )

    print(
        "[YOLO] Estratégia: 1024 único, sem zoom.",
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

            req = json.loads(line)

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

            items = req.get(
                "images",
                []
            )

            resultados = executar_batch(
                session,
                items
            )

            # =================================================
            # RESPOSTA
            # =================================================

            final_results = []

            for item, result in zip(
                items,
                resultados
            ):

                result[
                    "task_id"
                ] = item.get(
                    "task_id"
                )

                result[
                    "image_path"
                ] = item.get(
                    "image_path"
                )

                final_results.append(
                    result
                )

            response = {
                "results":
                    final_results
            }

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
                f"[YOLO] Erro: {e}",
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