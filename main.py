#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
main.py — Gerenciador central de sistemas.

Funciona exclusivamente como launcher: localiza arquivos, coleta
entradas do usuário e lança cada script como subprocesso com os
argumentos corretos.  Não toca na lógica interna de nenhum script.

Como adicionar um novo sistema:
  1. Crie uma instância de Sistema com seus metadados.
  2. Acrescente-a à lista SISTEMAS.
  Nenhuma outra alteração é necessária.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path
from typing import List, Optional


# ============================================================
# DIRETÓRIO BASE DO PROJETO
# ============================================================

BASE_DIR = Path(__file__).resolve().parent


# ============================================================
# TIPOS DE ENTRADA SUPORTADOS
# ============================================================

class TipoEntrada(Enum):
    """
    Define como cada script recebe seus dados de entrada.

    PASTA_ARGV
        O script espera um caminho de pasta via sys.argv[1].
        Internamente ele localiza os arquivos (normalmente via rglob).
        Exemplos: extrair_autoban.py, analises_fraudes.py.

    ARQUIVO_STDIN
        O script solicita o caminho de um arquivo via input().
        O main.py injeta o caminho através do stdin do subprocesso.
        Exemplo: isencao_judicial.py.

    EXCEL_PASTA_STDIN
        O script solicita o caminho de um Excel e depois de uma pasta
        de imagens, ambos via input().  O main.py injeta ambos via stdin.
        Exemplo: extrair_tamoios_isencao.py.
    """
    PASTA_ARGV        = auto()
    ARQUIVO_STDIN     = auto()
    EXCEL_PASTA_STDIN = auto()


# ============================================================
# DEFINIÇÃO DE SISTEMA
# ============================================================

@dataclass
class Sistema:
    """
    Descreve um sistema disponível no menu.

    chave            Tecla digitada pelo usuário no menu principal.
    nome             Nome exibido no menu.
    script           Nome do arquivo .py (relativo a BASE_DIR).
    tipo_entrada     Como o script recebe seus dados (ver TipoEntrada).
    extensoes        Extensões dos arquivos principais aceitos pelo script.
    extensoes_img    Extensões de imagens externas (só para EXCEL_PASTA_STDIN).
    saida_junto      True quando o script salva o resultado na mesma pasta
                     do arquivo de entrada (necessário para realocar saídas
                     em modo de arquivo único com pasta temporária).
    aceita_pasta     True quando ARQUIVO_STDIN pode receber uma pasta inteira
                     em uma única execução.
    """
    chave          : str
    nome           : str
    script         : str
    tipo_entrada   : TipoEntrada
    extensoes      : List[str]
    extensoes_img  : List[str] = field(default_factory=list)
    saida_junto    : bool = False
    aceita_pasta   : bool = False


# ============================================================
# REGISTRO DE SISTEMAS
# ============================================================

SISTEMAS: list[Sistema] = [

    Sistema(
        chave         = "1",
        nome          = "Análise de Fraudes",
        script        = "analises_fraudes.py",
        tipo_entrada  = TipoEntrada.PASTA_ARGV,
        extensoes     = [".xlsx"],
        saida_junto   = True,
    ),

    Sistema(
        chave         = "2",
        nome          = "Autoban — Isenção Judicial",
        script        = "extrair_autoban.py",
        tipo_entrada  = TipoEntrada.PASTA_ARGV,
        extensoes     = [".pdf"],
        saida_junto   = False,
    ),

    Sistema(
        chave         = "3",
        nome          = "SPVias — Evasão de Pedágio",
        script        = "extrair_spvias.py",
        tipo_entrada  = TipoEntrada.PASTA_ARGV,
        extensoes     = [".xlsx"],
        saida_junto   = False,
    ),

    Sistema(
        chave         = "4",
        nome          = "Tamoios — Evasão Free Flow",
        script        = "extrair_tamoios.py",
        tipo_entrada  = TipoEntrada.PASTA_ARGV,
        extensoes     = [".xlsx"],
        saida_junto   = True,
    ),

    Sistema(
        chave         = "5",
        nome          = "Tamoios/Rota das Bandeiras — Isenção",
        script        = "extrair_tamoios_isencao.py",
        tipo_entrada  = TipoEntrada.EXCEL_PASTA_STDIN,
        extensoes     = [".xlsx"],
        extensoes_img = [".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"],
        saida_junto   = False,
    ),

    Sistema(
        chave         = "6",
        nome          = "CNL — Isenção Judicial",
        script        = "cnl_isencao_judicial.py",
        tipo_entrada  = TipoEntrada.ARQUIVO_STDIN,
        extensoes     = [".xlsx"],
        saida_junto   = True,

        # O isencao_judicial.py aceita uma pasta inteira
        # e localiza os Excels internamente.
        aceita_pasta  = True,
    ),

    Sistema(
        chave         = "7",
        nome          = "Via Colinas — Evasão de Pedágio",
        script        = "via_colinas.py",
        tipo_entrada  = TipoEntrada.PASTA_ARGV,
        extensoes     = [".xlsx"],
        saida_junto   = True,
    ),

]

# Mapa rápido: chave → Sistema
_MAPA: dict[str, Sistema] = {s.chave: s for s in SISTEMAS}


# ============================================================
# CONSTANTES VISUAIS
# ============================================================

_LINHA       = "=" * 75
_LINHA_FINA  = "-" * 75
_HASH        = "#" * 75


# ============================================================
# UTILITÁRIOS DE TERMINAL
# ============================================================

def _ler(prompt: str) -> str:
    """Lê entrada do usuário com tratamento de Ctrl+C / EOF."""
    try:
        return input(prompt).strip()
    except (KeyboardInterrupt, EOFError):
        print("\n\n[SAÍDA] Operação cancelada pelo usuário.")
        sys.exit(0)


def _limpar_caminho(texto: str) -> str:
    """Remove espaços, aspas simples e duplas das extremidades."""
    return texto.strip().strip('"').strip("'").strip()


def _cabecalho(titulo: str) -> None:
    print(f"\n{_LINHA}")
    print(f"  {titulo}")
    print(_LINHA)


def _ext_desc(extensoes: List[str]) -> str:
    """'[.xlsx]' ou '[.pdf / .xlsx]' etc."""
    return " / ".join(e.upper().lstrip(".") for e in extensoes)


# ============================================================
# COLETAR CAMINHO COM VALIDAÇÃO
# ============================================================

def _pedir_caminho(
    prompt:       str,
    deve_existir: bool = True,
    tipo:         str  = "qualquer",
) -> Optional[Path]:
    """
    Solicita um caminho ao usuário com validação e opção de re-tentativa.
    Retorna None quando o usuário cancela.
    """
    while True:
        texto = _limpar_caminho(_ler(f"\n{prompt}\n> "))

        if not texto:
            print("  [!] Nenhum caminho informado.")
            if _ler("      Tentar novamente? [s/n]: ").lower() != "s":
                return None
            continue

        caminho = Path(texto).resolve()

        if deve_existir and not caminho.exists():
            print(f"  [!] Caminho não encontrado:\n      {caminho}")
            if _ler("      Tentar novamente? [s/n]: ").lower() != "s":
                return None
            continue

        if tipo == "arquivo" and caminho.exists() and not caminho.is_file():
            print(f"  [!] O caminho indicado não é um arquivo:\n      {caminho}")
            if _ler("      Tentar novamente? [s/n]: ").lower() != "s":
                return None
            continue

        if tipo == "pasta" and caminho.exists() and not caminho.is_dir():
            print(f"  [!] O caminho indicado não é uma pasta:\n      {caminho}")
            if _ler("      Tentar novamente? [s/n]: ").lower() != "s":
                return None
            continue

        return caminho


# ============================================================
# LOCALIZAÇÃO DE ARQUIVOS
# ============================================================

_SUFIXOS_SAIDA = ("_analisado", "_validado", "_resultado")

_PREFIXOS_TEMP = ("~$",)


def _listar_arquivos(
    pasta:     Path,
    extensoes: List[str],
    recursivo: bool,
) -> list[Path]:
    """
    Localiza arquivos compatíveis em *pasta*.
    Ignora temporários do Office e arquivos de saída já gerados.
    """
    padrao = "**/*" if recursivo else "*"
    encontrados: list[Path] = []

    for ext in extensoes:
        for f in sorted(pasta.glob(f"{padrao}{ext}")):
            if any(f.name.startswith(p) for p in _PREFIXOS_TEMP):
                continue
            if any(f.stem.endswith(s) for s in _SUFIXOS_SAIDA):
                continue
            encontrados.append(f)

    vistos: set[Path] = set()
    unicos: list[Path] = []

    for f in encontrados:
        if f not in vistos:
            vistos.add(f)
            unicos.append(f)

    return unicos


def _exibir_lista_arquivos(
    arquivos: list[Path],
    pasta:    Path,
    limite:   int = 15,
) -> None:
    """Exibe até *limite* arquivos com caminho relativo à *pasta*."""
    for f in arquivos[:limite]:
        try:
            rel = f.relative_to(pasta)
        except ValueError:
            rel = f
        print(f"    • {rel}")

    if len(arquivos) > limite:
        print(f"    ... e mais {len(arquivos) - limite} arquivo(s).")


# ============================================================
# LANÇADOR DE SCRIPTS
# ============================================================

def _executar(
    sistema:    Sistema,
    args:       Optional[List[str]] = None,
    stdin_data: Optional[str]       = None,
) -> int:
    """
    Lança o script como subprocesso.

    - stdout e stderr são herdados: a saída do script aparece no terminal.
    - stdin é herdado normalmente, exceto quando *stdin_data* é fornecido.
    """
    caminho_script = BASE_DIR / sistema.script

    if not caminho_script.exists():
        print(f"\n  [ERRO] Script não encontrado: {caminho_script}")
        return 1

    cmd = [sys.executable, str(caminho_script)] + (args or [])

    print(f"\n[→] Iniciando: {sistema.script}")
    print(_LINHA_FINA)

    resultado = subprocess.run(
        cmd,
        cwd   = str(BASE_DIR),
        input = stdin_data,
        text  = True,
    )

    print(_LINHA_FINA)

    if resultado.returncode != 0:
        print(f"  [AVISO] O script encerrou com código {resultado.returncode}.")

    return resultado.returncode


# ============================================================
# MODO 1: ARQUIVO ÚNICO — com pasta temporária (PASTA_ARGV)
# ============================================================

def _executar_arquivo_unico_pasta_argv(sistema: Sistema) -> None:
    """
    Para scripts PASTA_ARGV em modo de arquivo único:
      1. Cria pasta temporária.
      2. Copia o arquivo selecionado para ela.
      3. Passa a pasta temporária ao script via sys.argv[1].
      4. Se o script gravar a saída junto ao arquivo de entrada
         (sistema.saida_junto = True), move os arquivos gerados
         de volta para a pasta original do arquivo.
      5. Remove a pasta temporária.
    """
    ext = _ext_desc(sistema.extensoes)

    arquivo = _pedir_caminho(
        f"Caminho do arquivo {ext}:",
        deve_existir=True,
        tipo="arquivo",
    )

    if arquivo is None:
        return

    if arquivo.suffix.lower() not in sistema.extensoes:
        print(f"\n  [!] Extensão incompatível. Esperado: {ext}")
        return

    pasta_original = arquivo.parent
    tmpdir = Path(tempfile.mkdtemp(prefix="main_launcher_"))

    try:
        destino = tmpdir / arquivo.name
        shutil.copy2(arquivo, destino)

        print(f"\n  [INFO] Arquivo selecionado: {arquivo.name}")
        print(f"  [INFO] Pasta temporária:    {tmpdir}")

        _executar(sistema, args=[str(tmpdir)])

        if sistema.saida_junto:
            _realocar_saidas(
                tmpdir,
                pasta_original,
                ignorar={destino},
            )

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _realocar_saidas(
    origem:    Path,
    destino:   Path,
    ignorar:   set,
) -> None:
    """
    Move todos os arquivos gerados em *origem* (exceto os listados em
    *ignorar*) para *destino*.
    """
    gerados = [
        f for f in origem.iterdir()
        if f.is_file() and f not in ignorar
    ]

    if not gerados:
        return

    print(f"\n  [INFO] Realocando saída(s) para: {destino}")

    for f in gerados:
        alvo = destino / f.name

        if alvo.exists():
            alvo.unlink()

        shutil.move(str(f), str(alvo))
        print(f"    ✓ {f.name}")


# ============================================================
# EXECUÇÃO POR TIPO DE ENTRADA
# ============================================================

def _executar_pasta_argv(sistema: Sistema, modo: int) -> None:
    """
    Gerencia scripts do tipo PASTA_ARGV nos três modos de entrada.
    """
    ext = _ext_desc(sistema.extensoes)

    if modo == 1:
        _executar_arquivo_unico_pasta_argv(sistema)
        return

    pasta = _pedir_caminho(
        "Caminho da pasta:",
        deve_existir=True,
        tipo="pasta",
    )

    if pasta is None:
        return

    arquivos = _listar_arquivos(
        pasta,
        sistema.extensoes,
        recursivo=True,
    )

    if not arquivos:
        print(
            f"\n  [!] Nenhum arquivo {ext} encontrado em:\n"
            f"      {pasta}"
        )
        return

    _exibir_arvore_pasta(
        pasta,
        sistema,
        arquivos,
    )

    if _ler("\n  Confirmar e iniciar? [s/n]: ").lower() != "s":
        print("  [INFO] Operação cancelada.")
        return

    _executar(
        sistema,
        args=[str(pasta)],
    )


def _executar_arquivo_stdin(sistema: Sistema, modo: int) -> None:
    """
    Gerencia scripts do tipo ARQUIVO_STDIN.

    Modo 1:
        Envia um arquivo único.

    Modos 2 e 3:
        Normalmente processam arquivo por arquivo.

        Porém, quando sistema.aceita_pasta=True, a pasta inteira
        é enviada UMA ÚNICA VEZ para o script.
    """
    ext = _ext_desc(sistema.extensoes)

    # ----------------------------------------------------------
    # MODO 1 — ARQUIVO ÚNICO
    # ----------------------------------------------------------
    if modo == 1:

        arquivo = _pedir_caminho(
            f"Caminho do arquivo {ext}:",
            deve_existir=True,
            tipo="arquivo",
        )

        if arquivo is None:
            return

        if arquivo.suffix.lower() not in sistema.extensoes:
            print(
                f"\n  [!] Extensão incompatível. Esperado: {ext}"
            )
            return

        _executar(
            sistema,
            stdin_data=f"{arquivo}\n",
        )

        return

    # ----------------------------------------------------------
    # MODOS 2 E 3 — PASTA
    # ----------------------------------------------------------
    pasta = _pedir_caminho(
        "Caminho da pasta:",
        deve_existir=True,
        tipo="pasta",
    )

    if pasta is None:
        return

    recursivo = (modo == 3)

    arquivos = _listar_arquivos(
        pasta,
        sistema.extensoes,
        recursivo=recursivo,
    )

    if not arquivos:
        print(
            f"\n  [!] Nenhum arquivo {ext} encontrado."
        )
        return

    _exibir_arvore_pasta(
        pasta,
        sistema,
        arquivos,
    )

    if _ler("\n  Confirmar e iniciar? [s/n]: ").lower() != "s":
        print("  [INFO] Operação cancelada.")
        return

    # ========================================================
    # ALTERAÇÃO:
    # SISTEMAS QUE ACEITAM PASTA RECEBEM A PASTA UMA ÚNICA VEZ
    # ========================================================

    if sistema.aceita_pasta:

        _executar(
            sistema,
            stdin_data=f"{pasta}\n",
        )

        return

    # ----------------------------------------------------------
    # COMPORTAMENTO ORIGINAL PARA OS DEMAIS SISTEMAS
    # ----------------------------------------------------------
    total = len(arquivos)

    for idx, arquivo in enumerate(
        arquivos,
        start=1,
    ):

        print(f"\n{_HASH}")
        print(f"  ARQUIVO {idx}/{total}: {arquivo.name}")
        print(_HASH)

        _executar(
            sistema,
            stdin_data=f"{arquivo}\n",
        )


def _executar_excel_pasta_stdin(sistema: Sistema, modo: int) -> None:
    """
    Gerencia scripts do tipo EXCEL_PASTA_STDIN nos três modos de entrada.

    O script espera dois inputs() em sequência:
      1. Caminho do arquivo Excel.
      2. Caminho da pasta de imagens.
    """
    ext = _ext_desc(sistema.extensoes)

    # ----------------------------------------------------------
    if modo == 1:
        excel = _pedir_caminho(
            f"Caminho do arquivo {ext}:",
            deve_existir=True,
            tipo="arquivo",
        )

        if excel is None:
            return

        if excel.suffix.lower() not in sistema.extensoes:
            print(f"\n  [!] Extensão incompatível. Esperado: {ext}")
            return

        pasta_imgs = _pedir_caminho(
            "Caminho da pasta de imagens:",
            deve_existir=True,
            tipo="pasta",
        )

        if pasta_imgs is None:
            return

        stdin_data = f"{excel}\n{pasta_imgs}\n"
        _executar(
            sistema,
            stdin_data=stdin_data,
        )

        return

    # ----------------------------------------------------------
    # MODOS 2 E 3
    # ----------------------------------------------------------
    pasta_excels = _pedir_caminho(
        f"Caminho da pasta com arquivos {ext}:",
        deve_existir=True,
        tipo="pasta",
    )

    if pasta_excels is None:
        return

    recursivo = (modo == 3)

    excels = _listar_arquivos(
        pasta_excels,
        sistema.extensoes,
        recursivo=recursivo,
    )

    if not excels:
        print(f"\n  [!] Nenhum arquivo {ext} encontrado.")
        return

    _exibir_arvore_pasta(
        pasta_excels,
        sistema,
        excels,
    )

    pasta_imgs = _pedir_caminho(
        "Caminho da pasta de imagens\n"
        "  (será usada para todos os arquivos acima):",
        deve_existir=True,
        tipo="pasta",
    )

    if pasta_imgs is None:
        return

    if sistema.extensoes_img:
        imgs = _listar_arquivos(
            pasta_imgs,
            sistema.extensoes_img,
            recursivo=True,
        )

        print(
            f"\n  [INFO] {len(imgs)} imagem(ns) encontrada(s) "
            f"em: {pasta_imgs}"
        )

    if _ler("\n  Confirmar e iniciar? [s/n]: ").lower() != "s":
        print("  [INFO] Operação cancelada.")
        return

    total = len(excels)

    for idx, excel in enumerate(
        excels,
        start=1,
    ):

        print(f"\n{_HASH}")
        print(f"  ARQUIVO {idx}/{total}: {excel.name}")
        print(_HASH)

        stdin_data = f"{excel}\n{pasta_imgs}\n"

        _executar(
            sistema,
            stdin_data=stdin_data,
        )


# ============================================================
# EXIBIÇÃO DE ÁRVORE DE ARQUIVOS ENCONTRADOS
# ============================================================

def _exibir_arvore_pasta(
    pasta:    Path,
    sistema:  Sistema,
    arquivos: list[Path],
) -> None:
    """
    Exibe um resumo dos arquivos encontrados antes de iniciar o processamento.
    """
    ext = _ext_desc(sistema.extensoes)

    print(f"\n  {_LINHA_FINA}")
    print(f"  Pasta  : {pasta}")
    print(f"  Sistema: {sistema.nome}")
    print(f"  Tipo   : {ext}")
    print(f"  Total  : {len(arquivos)} arquivo(s)")
    print(f"  {_LINHA_FINA}")

    _exibir_lista_arquivos(
        arquivos,
        pasta,
    )

    print(f"  {_LINHA_FINA}")


# ============================================================
# MENU: MODO DE ENTRADA
# ============================================================

def _menu_modo(sistema: Sistema) -> Optional[int]:
    """
    Exibe o submenu de seleção de modo de entrada e retorna
    1, 2, 3 ou None (voltar).
    """
    _cabecalho(
        f"MODO DE ENTRADA  ›  {sistema.nome}"
    )

    print(
        f"  Tipo      : {_ext_desc(sistema.extensoes)}"
    )

    print(
        f"  Script    : {sistema.script}"
    )

    print()

    print(
        "  [1] Arquivo único"
    )

    print(
        "  [2] Pasta"
    )

    print(
        "  [3] Pasta + Subpastas  (recursivo)"
    )

    print()

    print(
        "  [0] Voltar"
    )

    opcao = _ler(
        "\n  Modo: "
    )

    if opcao == "0":
        return None

    if opcao not in ("1", "2", "3"):
        print("  [!] Opção inválida.")
        return None

    return int(opcao)


# ============================================================
# MENU PRINCIPAL
# ============================================================

def _menu_principal() -> None:
    """Loop principal do gerenciador."""

    while True:

        _cabecalho(
            "GERENCIADOR DE SISTEMAS  —  MENU PRINCIPAL"
        )

        print()

        for s in SISTEMAS:

            print(
                f"  [{s.chave}]  {s.nome}"
            )

            print()

        print(
            "  [0]  Sair"
        )

        opcao = _ler(
            "\n  Sistema: "
        )

        if opcao == "0":

            print(
                "\n  [INFO] Encerrando.\n"
            )

            sys.exit(0)

        if opcao not in _MAPA:

            print(
                "\n  [!] Opção inválida. Tente novamente."
            )

            _ler(
                "  Pressione Enter para continuar..."
            )

            continue

        sistema = _MAPA[opcao]

        modo = _menu_modo(
            sistema
        )

        if modo is None:
            continue

        try:

            if sistema.tipo_entrada is TipoEntrada.PASTA_ARGV:

                _executar_pasta_argv(
                    sistema,
                    modo,
                )

            elif sistema.tipo_entrada is TipoEntrada.ARQUIVO_STDIN:

                _executar_arquivo_stdin(
                    sistema,
                    modo,
                )

            elif sistema.tipo_entrada is TipoEntrada.EXCEL_PASTA_STDIN:

                _executar_excel_pasta_stdin(
                    sistema,
                    modo,
                )

            else:

                print(
                    f"\n  [ERRO] Tipo de entrada não suportado: "
                    f"{sistema.tipo_entrada}"
                )

        except KeyboardInterrupt:

            print(
                "\n\n  [INFO] Operação interrompida pelo usuário."
            )

        _ler(
            "\n\n  Pressione Enter para voltar ao menu..."
        )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    _menu_principal()