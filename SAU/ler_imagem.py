from __future__ import annotations

import subprocess
import sys
from pathlib import Path


# ============================================================
# CAMINHO DOS SCRIPTS
# ============================================================

BASE_DIR = Path(
    __file__
).resolve().parent


SCRIPT_OPERACIONAL = (
    BASE_DIR
    / "ler_imagem_operacionais.py"
)


SCRIPT_ADMINISTRATIVOS = (
    BASE_DIR
    / "ler_imagem_administrativos.py"
)


# ============================================================
# MENU
# ============================================================

def escolher_tipo_analise():

    print("\n")
    print("=" * 70)
    print("ANÁLISE DE VEÍCULOS")
    print("=" * 70)

    print()
    print("1 - Veículos Administrativos")
    print("2 - Veículos Operacionais")
    print()

    while True:

        opcao = input(
            "Escolha o tipo de análise [1/2]: "
        ).strip()

        if opcao == "1":
            return "administrativos"

        if opcao == "2":
            return "operacionais"

        print(
            "\nOpção inválida. "
            "Digite 1 ou 2.\n"
        )


# ============================================================
# EXECUTA SCRIPT
# ============================================================

def executar_script(
    script,
    pasta
):

    if not script.is_file():

        print(
            "\nERRO: script não encontrado:\n"
            f"{script}\n"
        )

        return 1

    comando = [
        sys.executable,
        str(script),
        str(pasta),
    ]

    print(
        "\n"
        + "=" * 70
    )

    print(
        "INICIANDO ANÁLISE"
    )

    print(
        "=" * 70
    )

    print(
        f"Script: {script}"
    )

    print(
        f"Pasta:  {pasta}"
    )

    print(
        "=" * 70
    )

    resultado = subprocess.run(
        comando
    )

    return resultado.returncode


# ============================================================
# MAIN
# ============================================================

def main():

    if len(
        sys.argv
    ) < 2:

        print(
            "\nUso:\n"
            '    py ler_imagem.py '
            '"C:\\caminho\\da\\pasta\\de\\imagens"\n'
        )

        return

    pasta = Path(
        sys.argv[1]
    ).resolve()

    if not pasta.is_dir():

        print(
            "\nERRO: a pasta informada "
            "não existe ou não é válida:\n"
            f"{pasta}\n"
        )

        return

    tipo = (
        escolher_tipo_analise()
    )

    if tipo == "administrativos":

        executar_script(
            SCRIPT_ADMINISTRATIVOS,
            pasta
        )

    else:

        executar_script(
            SCRIPT_OPERACIONAL,
            pasta
        )


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    main()