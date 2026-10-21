"""Funções comuns para normalizar e comparar placas reconhecidas por OCR."""

from functools import lru_cache
import re

TAMANHO_PLACA = 7
MASCARA_MERCOSUL = "LLLDLDD"
MASCARA_ANTIGA = "LLLDDDD"

# Confusões usadas no cálculo ponderado da distância OCR/placa.
CONFUNDIVEIS = frozenset({
    ("0", "O"), ("O", "0"), ("1", "I"), ("I", "1"), ("1", "L"), ("L", "1"),
    ("T", "I"), ("I", "T"), ("2", "Z"), ("Z", "2"), ("Z", "7"), ("7", "Z"),
    ("4", "A"), ("A", "4"), ("5", "S"), ("S", "5"), ("6", "G"), ("G", "6"),
    ("8", "B"), ("B", "8"), ("B", "E"), ("E", "B"), ("D", "0"), ("0", "D"),
    ("D", "O"), ("O", "D"),
})

# Conversões quando o caractere da posição da placa deve ser letra ou dígito.
MAPA_LETRAS = {"0": "O", "1": "I", "2": "Z", "4": "A", "5": "S", "6": "G", "7": "Z", "8": "B"}
MAPA_DIGITOS = {"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "T": "7", "J": "1", "Z": "2", "A": "4", "S": "5", "G": "6", "B": "8", "E": "8", "U": "0"}


def normalizar_placa(valor):
    if valor is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(valor).upper())

def normalizar_ocr(valor, *, reordenar_formatos=False):
    texto = normalizar_placa(valor)
    if not reordenar_formatos or not texto:
        return texto
    if re.fullmatch(r"\d{4}[A-Z]{3}", texto):
        return texto[4:] + texto[:4]
    if re.fullmatch(r"\d[A-Z]\d{2}[A-Z]{3}", texto):
        return texto[4:] + texto[:4]
    return texto

@lru_cache(maxsize=32768)
def _distancia_ponderada_cache(a, b, confundiveis):
    anterior = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        atual = [i]
        for j, cb in enumerate(b, start=1):
            if ca == cb:
                custo = 0.0
            elif (ca, cb) in confundiveis:
                custo = 0.35
            else:
                custo = 1.0
            atual.append(min(anterior[j] + 1, atual[j - 1] + 1, anterior[j - 1] + custo))
        anterior = atual
    return anterior[-1]

def distancia_ponderada(a, b, confundiveis):
    pares = frozenset(confundiveis)
    return _distancia_ponderada_cache(str(a), str(b), pares)

def distancia_placas(placa_a, placa_b, confundiveis):
    a = normalizar_placa(placa_a)
    b = normalizar_placa(placa_b)
    if not a or not b:
        return 999.0
    return distancia_ponderada(a, b, confundiveis)

def descobrir_mascara(placa, *, tamanho=TAMANHO_PLACA, mascara_invalida=MASCARA_MERCOSUL):
    texto = normalizar_placa(placa)
    if len(texto) != tamanho:
        return mascara_invalida
    return MASCARA_MERCOSUL if texto[4].isalpha() else MASCARA_ANTIGA

def aplicar_mascara(candidato, mascara, mapa_letras, mapa_digitos, *, tamanho=TAMANHO_PLACA, comprimento_invalido="manter", preservar_indice=None):
    texto = normalizar_placa(candidato)
    if len(texto) != tamanho:
        return "" if comprimento_invalido == "vazio" else texto
    resultado = []
    for indice, (caractere, tipo) in enumerate(zip(texto, mascara)):
        if indice == preservar_indice:
            resultado.append(caractere)
        elif tipo == "L":
            resultado.append(caractere if caractere.isalpha() else mapa_letras.get(caractere, caractere))
        else:
            resultado.append(caractere if caractere.isdigit() else mapa_digitos.get(caractere, caractere))
    return "".join(resultado)
