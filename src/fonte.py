"""Leitura das planilhas de ENTRADA.

Abre cada arquivo em modo leitura (leve), acha a aba cujo cabeçalho bate com o
esperado e devolve só as linhas em que a coluna de filtro está preenchida.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl

import config
from src.layout import INDICE_FILTRO, N_COLUNAS_FONTE


def encontrar_aba_fonte(caminho_arquivo: Path) -> str | None:
    """Procura, entre todas as abas do arquivo, a primeira cujo cabeçalho
    (linha 1) bate com config.CABECALHO_ESPERADO nas primeiras 15 colunas."""
    wb = openpyxl.load_workbook(caminho_arquivo, read_only=True, data_only=True)
    try:
        for nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            primeira_linha = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if not primeira_linha:
                continue
            cabecalho = [str(c).strip() if c is not None else "" for c in primeira_linha[:N_COLUNAS_FONTE]]
            if cabecalho == config.CABECALHO_ESPERADO:
                return nome_aba
        return None
    finally:
        wb.close()


class AbaNaoEncontrada(Exception):
    """Nenhuma aba do arquivo bateu com o cabeçalho esperado."""


def extrair_linhas_com_motivo(caminho_arquivo: Path) -> list[tuple]:
    """Abre o arquivo em modo leitura (leve) e devolve só as linhas onde a
    coluna de filtro (a última coluna do cabeçalho, "Motivo" no exemplo) tem algo escrito."""
    nome_aba = encontrar_aba_fonte(caminho_arquivo)
    if nome_aba is None:
        raise AbaNaoEncontrada(
            f'Não encontrei, em "{caminho_arquivo.name}", nenhuma aba com o cabeçalho esperado: '
            f"{', '.join(config.CABECALHO_ESPERADO)}"
        )

    wb = openpyxl.load_workbook(caminho_arquivo, read_only=True, data_only=True)
    try:
        ws = wb[nome_aba]
        linhas_filtradas = []
        for linha in ws.iter_rows(min_row=2, values_only=True):
            if len(linha) < N_COLUNAS_FONTE:
                continue
            motivo = linha[INDICE_FILTRO]
            if motivo is not None and str(motivo).strip() != "":
                linhas_filtradas.append(tuple(linha[:N_COLUNAS_FONTE]))
        return linhas_filtradas
    finally:
        wb.close()
