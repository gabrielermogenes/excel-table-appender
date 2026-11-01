"""Testes com planilhas sintéticas (nenhum dado real)."""

import re
import zipfile

import openpyxl
import pytest
from openpyxl.worksheet.table import Table

import config
import main
from src.fonte import AbaNaoEncontrada, extrair_linhas_com_motivo
from src.mestre import anexar_linhas_no_master

FORMULA_MES = "=MONTH(Tabela1[[#This Row],[Início]])"
FORMULA_ANO = "=YEAR(Tabela1[[#This Row],[Início]])"


NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def excelizar(caminho):
    """O openpyxl serializa alguns trechos do XML numa ordem diferente da do Excel.
    O script lê o XML no formato que o próprio Excel grava, então ajustamos o arquivo
    de teste para esse formato (workbook.xml, relacionamentos e tablePart)."""

    def rels(xml, tirar_prefixo):
        def um(m):
            alvo = m.group(2).lstrip("/")
            alvo = alvo[len("xl/"):] if alvo.startswith("xl/") and tirar_prefixo == "xl/" else alvo
            if tirar_prefixo == "../":
                alvo = "../" + alvo[len("xl/"):]
            return f'<Relationship Id="{m.group(3)}" Type="{m.group(1)}" Target="{alvo}"/>'
        return re.sub(r'<Relationship Type="([^"]*)" Target="([^"]*)" Id="([^"]*)"\s*/>', um, xml)

    with zipfile.ZipFile(caminho) as zin:
        itens = [(i, zin.read(i.filename)) for i in zin.infolist()]
    with zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, dados in itens:
            nome = info.filename
            if nome == "xl/workbook.xml":
                x = dados.decode("utf-8")
                x = x.replace("<workbook xmlns=", f'<workbook xmlns:r="{NS_R}" xmlns=', 1)
                x = re.sub(r'<sheet xmlns:r="[^"]*" name="([^"]*)" sheetId="(\d+)" state="visible" r:id="([^"]+)"\s*/>',
                           r'<sheet name="\1" sheetId="\2" r:id="\3"/>', x)
                dados = x.encode("utf-8")
            elif nome == "xl/_rels/workbook.xml.rels":
                dados = rels(dados.decode("utf-8"), "xl/").encode("utf-8")
            elif nome == "xl/worksheets/_rels/sheet1.xml.rels":
                dados = rels(dados.decode("utf-8"), "../").encode("utf-8")
            elif nome == "xl/worksheets/sheet1.xml":
                x = dados.decode("utf-8")
                x = x.replace("<worksheet xmlns=", f'<worksheet xmlns:r="{NS_R}" xmlns=', 1)
                x = re.sub(r'<tablePart xmlns:r="[^"]*" r:id="([^"]+)"\s*/>', r'<tablePart r:id="\1"/>', x)
                # o Excel grava acentos como UTF-8 puro; o openpyxl usa referências numéricas (&#237;)
                x = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), x)
                dados = x.encode("utf-8")
            zout.writestr(info, dados)


def criar_mestre(caminho, n_dados=3, n_reservadas=5, aba=None):
    """Mestre com Tabela1: linhas de dados + linhas reservadas (A..F vazias, G..H com fórmula)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = aba or config.ABA_DESTINO
    ws.append(config.DESTINO_CABECALHO_ESPERADO + ["Mês", "Ano"])
    for i in range(n_dados):
        ws.append([f"EQ{i}", f"OP{i}", 45000 + i, 45000.5 + i, 10 + i, f"motivo {i}", FORMULA_MES, FORMULA_ANO])
    for _ in range(n_reservadas):
        ws.append([None] * 6 + [FORMULA_MES, FORMULA_ANO])
    ultima = 1 + n_dados + n_reservadas
    ws.add_table(Table(displayName="Tabela1", ref=f"A1:H{ultima}"))
    wb.save(caminho)
    excelizar(caminho)


def criar_entrada(caminho, linhas, cabecalho=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(cabecalho or config.CABECALHO_ESPERADO)
    for linha in linhas:
        ws.append(linha)
    wb.save(caminho)


def linhas_novas(n, inicio=0):
    return [(f"NEW{i}", f"OP{i}", 46000 + i, 46000.5 + i, 5, f"novo {i}") for i in range(inicio, inicio + n)]


def ler_mestre(caminho):
    ws = openpyxl.load_workbook(caminho)[config.ABA_DESTINO]
    return ws, list(ws.iter_rows(values_only=True))


def test_extrai_so_linhas_com_filtro_preenchido(tmp_path):
    f = tmp_path / "entrada.xlsx"
    criar_entrada(f, [
        ("A", "1", 1, 2, 3, "tem motivo"),
        ("B", "2", 1, 2, 3, None),
        ("C", "3", 1, 2, 3, "   "),
        ("D", "4", 1, 2, 3, "outro motivo"),
    ])
    assert [l[0] for l in extrair_linhas_com_motivo(f)] == ["A", "D"]


def test_cabecalho_errado_levanta_erro(tmp_path):
    f = tmp_path / "entrada.xlsx"
    criar_entrada(f, [("A", "1", 1, 2, 3, "x")], cabecalho=["Outra", "Coisa", "X", "Y", "Z", "W"])
    with pytest.raises(AbaNaoEncontrada):
        extrair_linhas_com_motivo(f)


def test_reaproveita_linhas_reservadas_e_preserva_formulas(tmp_path):
    mestre = tmp_path / "mestre.xlsx"
    criar_mestre(mestre, n_dados=3, n_reservadas=5)
    ultima = anexar_linhas_no_master(mestre, linhas_novas(2), config.ABA_DESTINO)
    assert ultima == 6
    ws, linhas = ler_mestre(mestre)
    assert linhas[4][0] == "NEW0" and linhas[5][0] == "NEW1"          # linhas 5 e 6
    assert linhas[4][5] == "novo 0"
    for numero in (5, 6):                                              # fórmulas preservadas
        assert ws[f"G{numero}"].value == FORMULA_MES
        assert ws[f"H{numero}"].value == FORMULA_ANO
    assert linhas[1][0] == "EQ0"                                       # dados antigos intactos
    assert ws.tables["Tabela1"].ref == "A1:H9"                         # tabela não precisou crescer


def test_cria_linhas_novas_e_estende_a_tabela(tmp_path):
    mestre = tmp_path / "mestre.xlsx"
    criar_mestre(mestre, n_dados=3, n_reservadas=1)                    # linhas 2..5 (4 linhas)
    ultima = anexar_linhas_no_master(mestre, linhas_novas(4), config.ABA_DESTINO)
    assert ultima == 8
    ws, linhas = ler_mestre(mestre)
    assert [l[0] for l in linhas[4:8]] == ["NEW0", "NEW1", "NEW2", "NEW3"]
    assert ws["G8"].value == FORMULA_MES                               # molde de fórmula copiado
    assert ws.tables["Tabela1"].ref == "A1:H8"                         # tabela cresceu


def test_varias_rodadas_continuam_de_onde_parou(tmp_path):
    mestre = tmp_path / "mestre.xlsx"
    criar_mestre(mestre, n_dados=2, n_reservadas=10)
    anexar_linhas_no_master(mestre, linhas_novas(3, 0), config.ABA_DESTINO)
    anexar_linhas_no_master(mestre, linhas_novas(2, 3), config.ABA_DESTINO)
    _, linhas = ler_mestre(mestre)
    assert [l[0] for l in linhas[3:8]] == ["NEW0", "NEW1", "NEW2", "NEW3", "NEW4"]


def test_aba_inexistente_nao_altera_o_arquivo(tmp_path):
    mestre = tmp_path / "mestre.xlsx"
    criar_mestre(mestre)
    antes = mestre.read_bytes()
    with pytest.raises(ValueError):
        anexar_linhas_no_master(mestre, linhas_novas(1), "Aba que não existe")
    assert mestre.read_bytes() == antes


def test_mestre_sem_dados_e_cabecalho_errado_aborta(tmp_path):
    mestre = tmp_path / "mestre.xlsx"
    criar_mestre(mestre, n_dados=0, n_reservadas=3)
    with zipfile.ZipFile(mestre) as zf:
        assert "xl/worksheets/sheet1.xml" in zf.namelist()
    # cabeçalho correto + nenhuma linha com dado -> começa na linha 2
    assert anexar_linhas_no_master(mestre, linhas_novas(1), config.ABA_DESTINO) == 2

    errado = tmp_path / "errado.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = config.ABA_DESTINO
    ws.append(["X", "Y", "Z", "W", "V", "U", "G", "H"])
    ws.append([None] * 6 + [FORMULA_MES, FORMULA_ANO])
    ws.add_table(Table(displayName="Tabela1", ref="A1:H2"))
    wb.save(errado)
    excelizar(errado)
    antes = errado.read_bytes()
    with pytest.raises(ValueError):
        anexar_linhas_no_master(errado, linhas_novas(1), config.ABA_DESTINO)
    assert errado.read_bytes() == antes


def test_fluxo_completo(tmp_path, monkeypatch, capsys):
    entrada = tmp_path / "entrada"
    entrada.mkdir()
    mestre = tmp_path / "mestre.xlsx"
    criar_mestre(mestre, n_dados=1, n_reservadas=6)
    criar_entrada(entrada / "arq1.xlsx", [("A", "1", 1, 2, 3, "m1"), ("B", "2", 1, 2, 3, None)])
    criar_entrada(entrada / "arq2.xlsx", [("C", "3", 1, 2, 3, "m2"), ("D", "4", 1, 2, 3, "m3")])
    criar_entrada(entrada / "ruim.xlsx", [("X", "1", 1, 2, 3, "m")], cabecalho=["a", "b", "c", "d", "e", "f"])
    monkeypatch.setattr(config, "PASTA_ENTRADA", entrada)
    monkeypatch.setattr(config, "ARQUIVO_MESTRE", mestre)

    main.main()

    saida = capsys.readouterr().out
    assert "Total de linhas adicionadas: 3" in saida
    assert sorted(p.name for p in (entrada / "Processados").iterdir()) == ["arq1.xlsx", "arq2.xlsx"]
    assert (entrada / "ruim.xlsx").exists()                            # com erro: não é movido
    assert len(list(tmp_path.glob("mestre_backup_*.xlsx"))) == 1       # backup criado
    _, linhas = ler_mestre(mestre)
    assert [l[0] for l in linhas[2:5]] == ["A", "C", "D"]


def test_arquivo_mestre_aberto_mostra_mensagem_amigavel(tmp_path, monkeypatch, capsys):
    entrada = tmp_path / "entrada"
    entrada.mkdir()
    mestre = tmp_path / "mestre.xlsx"
    criar_mestre(mestre)
    criar_entrada(entrada / "arq1.xlsx", [("A", "1", 1, 2, 3, "m1")])
    monkeypatch.setattr(config, "PASTA_ENTRADA", entrada)
    monkeypatch.setattr(config, "ARQUIVO_MESTRE", mestre)

    def bloqueado(*a, **k):
        raise PermissionError("arquivo em uso")

    monkeypatch.setattr(main, "anexar_linhas_no_master", bloqueado)
    main.main()
    saida = capsys.readouterr().out
    assert "Feche-o no Excel" in saida
    assert (entrada / "arq1.xlsx").exists()                            # não foi movido
