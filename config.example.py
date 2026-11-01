"""
Configuração de exemplo.

Como usar:
    1. Copie este arquivo para config.py  (cp config.example.py config.py)
    2. Troque os valores "COLE_AQUI" e ajuste o layout para o SEU arquivo mestre
    3. NUNCA publique o config.py: ele já está no .gitignore

Dica: no Windows, use o "r" antes das aspas nos caminhos: Path(r"C:\\pasta\\arquivo.xlsx").
"""

from pathlib import Path

# ==================== CAMINHOS ====================

# Pasta onde você joga as planilhas de ENTRADA (a subpasta "Processados" é criada sozinha)
PASTA_ENTRADA = Path(r"C:\COLE_AQUI\pasta_de_entrada")

# Arquivo MESTRE (.xlsx) que recebe as linhas. O script faz backup antes de alterar.
ARQUIVO_MESTRE = Path(r"C:\COLE_AQUI\arquivo_mestre.xlsx")

# Nome da aba do arquivo mestre que contém a Tabela do Excel de destino
ABA_DESTINO = "Banco de Dados"

# Quais arquivos da pasta de entrada serão lidos
PADROES_ARQUIVO = ("*.xlsx", "*.xlsm")

# ==================== LAYOUT DAS PLANILHAS ====================
# Exemplo com 6 colunas de dados (A..F) e 2 colunas de fórmula no mestre (G..H).
# Ajuste para o seu caso.

# Cabeçalho esperado nas planilhas de ENTRADA (linha 1). A ÚLTIMA coluna é o filtro:
# só entram no mestre as linhas em que ela estiver preenchida.
CABECALHO_ESPERADO = [
    "Equipamento", "Ordem", "Início", "Fim", "Duração (min)", "Motivo",
]

# Cabeçalho esperado na aba de DESTINO (mesmas colunas, nomes podem diferir)
DESTINO_CABECALHO_ESPERADO = [
    "Equipamento", "Ordem", "Início", "Fim", "Duração (min)", "MOTIVOS",
]

# Total de colunas da tabela de destino: colunas de dados + colunas de fórmula (máx. 26).
# As colunas de fórmula (as que sobram) já devem existir no mestre, com a fórmula pronta
# em todas as linhas reservadas. O script nunca escreve nelas.
N_COLUNAS_DESTINO = 8

# Tipo de cada coluna de dados: "texto" ou "numero" (datas/horas contam como "numero")
TIPO_COLUNA_FONTE = {
    "A": "texto", "B": "texto", "C": "numero", "D": "numero", "E": "numero", "F": "texto",
}

# Estilo de fallback (índice de estilo do Excel) para colunas sem estilo na linha-modelo.
# Normalmente pode ficar vazio: o script reaproveita o estilo que a linha já tem.
ESTILOS_FALLBACK_A_O = {}
