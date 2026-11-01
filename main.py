"""
Ponto de entrada: lê as planilhas da pasta de entrada, filtra as linhas em que a
coluna de filtro está preenchida e anexa essas linhas na tabela do arquivo mestre.

Uso:
    python main.py
"""

import datetime as dt
import shutil
from pathlib import Path

try:
    import config
except ModuleNotFoundError:
    raise SystemExit(
        "Arquivo config.py não encontrado.\n"
        "Copie config.example.py para config.py e preencha os valores."
    )

from src.fonte import AbaNaoEncontrada, extrair_linhas_com_motivo
from src.mestre import anexar_linhas_no_master

PLACEHOLDER = "COLE_AQUI"


def validar_configuracao():
    """Avisa cedo se algum caminho de exemplo ainda não foi preenchido."""
    pendentes = [
        nome
        for nome in ("PASTA_ENTRADA", "ARQUIVO_MESTRE")
        if PLACEHOLDER in str(getattr(config, nome))
    ]
    if pendentes:
        raise SystemExit("Preencha no config.py: " + ", ".join(pendentes))


def mover_para_processados(caminho_arquivo: Path):
    pasta_processados = caminho_arquivo.parent / "Processados"
    pasta_processados.mkdir(exist_ok=True)
    destino = pasta_processados / caminho_arquivo.name
    if destino.exists():
        carimbo = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = pasta_processados / f"{caminho_arquivo.stem}_{carimbo}{caminho_arquivo.suffix}"
    shutil.move(str(caminho_arquivo), str(destino))


def main():
    validar_configuracao()
    if not config.ARQUIVO_MESTRE.exists():
        print(f"[ERRO] Arquivo mestre não encontrado: {config.ARQUIVO_MESTRE}")
        return
    if not config.PASTA_ENTRADA.exists():
        print(f"[ERRO] Pasta de entrada não encontrada: {config.PASTA_ENTRADA}")
        return

    arquivos = []
    for padrao in config.PADROES_ARQUIVO:
        arquivos.extend(sorted(config.PASTA_ENTRADA.glob(padrao)))
    arquivos = [a for a in arquivos if not a.name.startswith("~$")]

    if not arquivos:
        print("Nenhum arquivo novo encontrado na pasta de entrada.")
        return

    print("Fazendo backup de segurança do arquivo mestre...")
    carimbo = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = config.ARQUIVO_MESTRE.with_name(f"{config.ARQUIVO_MESTRE.stem}_backup_{carimbo}{config.ARQUIVO_MESTRE.suffix}")
    try:
        shutil.copy2(config.ARQUIVO_MESTRE, backup)
    except PermissionError:
        print("[ERRO] Não consegui fazer o backup. Feche o arquivo mestre no Excel e rode de novo.")
        return
    print(f"  Backup salvo em: {backup}")

    total_linhas = 0
    resumo = []

    for arquivo in arquivos:
        print(f"\nProcessando: {arquivo.name}")
        try:
            linhas_fonte = extrair_linhas_com_motivo(arquivo)
        except AbaNaoEncontrada as erro:
            print(f"  [ERRO] {erro}")
            print("  Esse arquivo NÃO foi movido pra 'Processados' — confira o cabeçalho e rode de novo.")
            resumo.append((arquivo.name, "ERRO"))
            continue
        print(f"  Linhas com Motivo preenchido: {len(linhas_fonte)}")

        if not linhas_fonte:
            mover_para_processados(arquivo)
            resumo.append((arquivo.name, 0))
            continue

        try:
            ultima_linha = anexar_linhas_no_master(config.ARQUIVO_MESTRE, linhas_fonte, config.ABA_DESTINO)
        except PermissionError:
            print("  [ERRO] Não consegui gravar o arquivo mestre. Feche-o no Excel e rode de novo.")
            print("  Esse arquivo NÃO foi movido pra 'Processados' — o arquivo mestre não foi alterado.")
            resumo.append((arquivo.name, "ERRO"))
            break
        except ValueError as erro:
            print(f"  [ERRO] {erro}")
            print("  Esse arquivo NÃO foi movido pra 'Processados' — o arquivo mestre não foi alterado.")
            resumo.append((arquivo.name, "ERRO"))
            continue

        total_linhas += len(linhas_fonte)
        resumo.append((arquivo.name, len(linhas_fonte)))
        print(f"  {len(linhas_fonte)} linha(s) adicionada(s) e conferida(s). Última linha agora: {ultima_linha}")

        mover_para_processados(arquivo)

    print("\n" + "=" * 60)
    print("RESUMO")
    print("=" * 60)
    for nome, qtd in resumo:
        print(f"  {nome}: {qtd} linha(s)")
    print(f"\nTotal de linhas adicionadas: {total_linhas}")
    print(f"Arquivo mestre atualizado: {config.ARQUIVO_MESTRE}")


if __name__ == "__main__":
    main()
