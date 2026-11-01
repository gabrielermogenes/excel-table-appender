"""Constantes derivadas do config.py (não edite este arquivo).

Convenção usada nos comentários e nomes do código:
  - "A..O" (COLUNAS_A_O)  = colunas de DADOS, preenchidas a partir das planilhas de entrada
  - "P..U" (COLUNAS_P_U)  = colunas de FÓRMULA, calculadas pelo Excel e nunca tocadas pelo script
Com o seu layout as letras podem ser outras: elas são calculadas aqui a partir do config.py.
"""

import string

import config

N_COLUNAS_FONTE = len(config.CABECALHO_ESPERADO)

if not (1 <= N_COLUNAS_FONTE <= config.N_COLUNAS_DESTINO <= 26):
    raise SystemExit(
        "config.py inválido: precisa valer 1 <= colunas de entrada <= N_COLUNAS_DESTINO <= 26."
    )
if len(config.DESTINO_CABECALHO_ESPERADO) != N_COLUNAS_FONTE:
    raise SystemExit("config.py inválido: DESTINO_CABECALHO_ESPERADO deve ter o mesmo tamanho de CABECALHO_ESPERADO.")

LETRAS_COLUNAS = list(string.ascii_uppercase[: config.N_COLUNAS_DESTINO])
COLUNAS_A_O = set(LETRAS_COLUNAS[:N_COLUNAS_FONTE])   # o Python escreve
COLUNAS_P_U = set(LETRAS_COLUNAS[N_COLUNAS_FONTE:])   # o Python NUNCA toca

# A coluna usada como filtro é sempre a ÚLTIMA coluna de dados (no exemplo, "Motivo")
INDICE_FILTRO = N_COLUNAS_FONTE - 1
LETRA_FILTRO = LETRAS_COLUNAS[INDICE_FILTRO]
