"""Helpers compartilhados pelos scripts de análise de passagens."""

import json

def escrever_json(processo, dados):
    processo.stdin.write(json.dumps(dados, ensure_ascii=False) + "\n")
    processo.stdin.flush()

def enviar_json(processo, dados):
    if processo.poll() is not None:
        raise RuntimeError(
            "Processo já encerrou. "
            f"Código: {processo.returncode}"
        )
    escrever_json(processo, dados)

def extrair_bytes_imagem(img):
    try:
        return img._data()
    except Exception:
        return None
