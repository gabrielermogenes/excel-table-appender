"""Escrita no arquivo MESTRE (.xlsx) por dentro do ZIP, sem abrir o arquivo inteiro.

POR QUE NÃO USAMOS openpyxl PARA ABRIR O ARQUIVO MESTRE INTEIRO
-----------------------------------------------------------------
Um .xlsx é um .zip de arquivos XML. Se o arquivo mestre tem muitas Tabelas
Dinâmicas, o "cache" de uma delas pode ter dezenas de MB de XML. Carregar o
arquivo inteiro na memória (openpyxl ou Power Query) para editar fica pesado
e pode travar. Este módulo NUNCA abre o mestre inteiro: troca só o XML da aba
de destino e o XML da tabela, e copia todo o resto byte a byte.

SÓ AS COLUNAS DE DADOS SÃO ESCRITAS
-----------------------------------------------------------------
As colunas de fórmula já existem em todas as linhas reservadas da tabela. O
script nunca calcula nem escreve nada nelas: preenche as colunas de dados e
preserva as de fórmula (fórmula, estilo, tipo), reaproveitando a própria linha
reservada que já existia.
"""

from __future__ import annotations

import datetime as dt
import re
import zipfile
from pathlib import Path

import config
from src.layout import (
    COLUNAS_A_O,
    COLUNAS_P_U,
    INDICE_FILTRO,
    LETRA_FILTRO,
    LETRAS_COLUNAS,
    N_COLUNAS_FONTE,
)


def serial_excel(data: dt.datetime) -> float:
    """Converte um datetime do Python pro número de série que o Excel usa
    internamente para representar datas/horas."""
    epoca = dt.datetime(1899, 12, 30)
    delta = data - epoca
    return delta.days + delta.seconds / 86400 + delta.microseconds / 86400_000_000


def _escapar_xml(texto: str) -> str:
    return (
        texto.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _desescapar_xml(texto: str) -> str:
    return (
        texto.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
    )


def _celula_texto(ref: str, valor, estilo: str | None) -> str:
    valor = "" if valor is None else str(valor)
    valor = _escapar_xml(valor)
    s_attr = f' s="{estilo}"' if estilo else ""
    if valor == "":
        return f'<c r="{ref}"{s_attr}/>'
    return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{valor}</t></is></c>'


def _celula_numero(ref: str, valor, estilo: str | None) -> str:
    s_attr = f' s="{estilo}"' if estilo else ""
    if valor is None or valor == "":
        return f'<c r="{ref}"{s_attr}/>'
    if isinstance(valor, dt.datetime):
        valor = serial_excel(valor)
    return f'<c r="{ref}"{s_attr}><v>{valor}</v></c>'


def _iter_celulas(conteudo_linha: str):
    """Percorre todas as células <c> de uma linha e devolve
    (letra_da_coluna, atributos_da_tag, corpo_da_célula) para cada uma.
    `corpo_da_célula` é "" para células autofechadas (<c .../>)."""
    for m_cel in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', conteudo_linha, re.S):
        yield m_cel.group(1), m_cel.group(2) or "", m_cel.group(3) or ""


def _carregar_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """O Excel guarda textos repetidos (como o cabeçalho) numa tabela à
    parte (xl/sharedStrings.xml) — a célula só guarda o NÚMERO da posição
    nessa tabela, não o texto em si."""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    xml = zf.read("xl/sharedStrings.xml").decode("utf-8")
    textos = []
    for m_si in re.finditer(r"<si>(.*?)</si>", xml, re.S):
        partes = re.findall(r"<t[^>]*>(.*?)</t>", m_si.group(1), re.S)
        textos.append(_desescapar_xml("".join(partes)))
    return textos


def _valor_real_da_celula(atributos: str, corpo_celula: str, shared_strings: list[str] | None = None) -> str:
    """Devolve só o VALOR de verdade de uma célula, venha de onde vier:
    - fórmula (<f>...</f>) — ignora a fórmula, olha só o valor calculado
    - shared string (t="s") — resolve o número contra a tabela de textos
    - inline string (t="inlineStr") — lê o texto direto
    - número puro — lê o <v> direto"""
    shared_strings = shared_strings or []
    sem_formula = re.sub(r"<f[^>]*>.*?</f>", "", corpo_celula, flags=re.S)

    m_is = re.search(r"<is>.*?<t[^>]*>(.*?)</t>", sem_formula, re.S)
    if m_is:
        return _desescapar_xml(m_is.group(1))

    m_v = re.search(r"<v>(.*?)</v>", sem_formula, re.S)
    if not m_v:
        return ""
    valor_bruto = m_v.group(1)

    if 't="s"' in atributos:
        try:
            indice = int(valor_bruto)
        except ValueError:
            return ""
        return shared_strings[indice] if 0 <= indice < len(shared_strings) else ""

    return _desescapar_xml(valor_bruto)


def _atributos_da_linha(sheet_xml: str, numero: int) -> str:
    """Devolve os atributos da tag <row> em si (ex: f' spans="1:{config.N_COLUNAS_DESTINO}"'),
    sem incluir o r="N"."""
    m = re.search(rf'<row r="{numero}"([^>]*)>', sheet_xml)
    return m.group(1) if m else f' spans="1:{config.N_COLUNAS_DESTINO}"'


def _estilos_colunas_ao(conteudo_linha: str) -> dict[str, str | None]:
    """Lê, célula por célula, o estilo (s="...") que cada coluna A..O já
    tem HOJE nessa linha — pra reaproveitar exatamente o mesmo (mantém
    formatação de data, zebra, etc, sem depender de nenhum mapa fixo)."""
    estilos = {}
    for letra, atributos, _ in _iter_celulas(conteudo_linha):
        if letra in COLUNAS_A_O:
            m = re.search(r's="([^"]+)"', atributos)
            estilos[letra] = m.group(1) if m else None
    return estilos


def _celulas_pu_preservadas(conteudo_linha: str, numero_alvo: int) -> str:
    """Copia as células P..U EXATAMENTE como estão (fórmula, tipo, estilo —
    tudo), só ajustando o número da linha na referência (r="P94" ->
    r="P123") para o caso de a linha-modelo ser de outro número."""
    partes = []
    for letra, atributos, corpo in _iter_celulas(conteudo_linha):
        if letra not in COLUNAS_P_U:
            continue
        m_completo = re.search(rf'<c r="{letra}\d+"[^>]*?(?:/>|>.*?</c>)', conteudo_linha, re.S)
        texto_original = m_completo.group() if m_completo else ""
        texto_renumerado = re.sub(rf'r="{letra}\d+"', f'r="{letra}{numero_alvo}"', texto_original, count=1)
        partes.append(texto_renumerado)
    return "".join(partes)


def montar_linha_xml(numero_linha: int, linha_fonte: tuple, conteudo_modelo: str, atributos_linha: str) -> str:
    """Monta a linha nova: escreve o dado real SÓ nas colunas A..O
    (reaproveitando o estilo que a própria coluna já tinha na
    linha-modelo) e copia as colunas P..U (colunas de fórmula) exatamente como estavam —
    o script nunca calcula nem escreve nada nelas."""
    estilos_ao = _estilos_colunas_ao(conteudo_modelo)
    celulas_ao = []
    for letra, valor in zip(LETRAS_COLUNAS[:N_COLUNAS_FONTE], linha_fonte):
        ref = f"{letra}{numero_linha}"
        estilo = estilos_ao.get(letra) or config.ESTILOS_FALLBACK_A_O.get(letra)
        tipo = config.TIPO_COLUNA_FONTE.get(letra, "texto")
        if tipo == "numero":
            celulas_ao.append(_celula_numero(ref, valor, estilo))
        else:
            celulas_ao.append(_celula_texto(ref, valor, estilo))

    celulas_pu = _celulas_pu_preservadas(conteudo_modelo, numero_linha)
    return f'<row r="{numero_linha}"{atributos_linha}>' + "".join(celulas_ao) + celulas_pu + "</row>"


def _achar_ultima_linha_real(sheet_xml: str, shared_strings: list[str], n_colunas_conferir: int = N_COLUNAS_FONTE) -> int:
    """CRITÉRIO — nunca confiar só na tag <dimension> nem só em "a célula
    tem algo escrito" (uma fórmula que calcula vazio ainda tem texto no
    XML, e o cabeçalho sempre tem texto). Varre linha por linha (ignorando
    a linha 1, que é o cabeçalho) e acha a última linha que tem, de fato,
    um VALOR não vazio em alguma das colunas A até O."""
    ultima_linha_com_dado = 0

    for m_linha in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', sheet_xml, re.S):
        numero = int(m_linha.group(1))
        if numero <= 1:
            continue  # linha 1 é o cabeçalho — não conta como "dado"
        conteudo_linha = m_linha.group(2)
        for letra_coluna, atributos, corpo in _iter_celulas(conteudo_linha):
            if letra_coluna not in COLUNAS_A_O:
                continue
            if _valor_real_da_celula(atributos, corpo, shared_strings).strip():
                ultima_linha_com_dado = max(ultima_linha_com_dado, numero)
                break

    return ultima_linha_com_dado


def _cabecalho_destino_bate(sheet_xml: str, shared_strings: list[str]) -> bool:
    """Confere se a linha 1 da aba de destino tem o cabeçalho esperado.
    Usado só quando não sobra NENHUMA linha com dado real (ex: você apagou
    tudo pra recomeçar do zero) — nesse caso, em vez de abortar, o script
    confirma que está na aba certa pelo cabeçalho e começa a escrever a
    partir da linha 2."""
    achado = _extrair_linha_xml(sheet_xml, 1)
    if achado is None:
        return False
    _, _, conteudo = achado
    valores = {}
    for letra, atributos, corpo in _iter_celulas(conteudo):
        valores[letra] = _valor_real_da_celula(atributos, corpo, shared_strings).strip()
    cabecalho_lido = [valores.get(letra, "").strip() for letra in LETRAS_COLUNAS[:N_COLUNAS_FONTE]]
    esperado = [c.strip() for c in config.DESTINO_CABECALHO_ESPERADO]
    return cabecalho_lido == esperado


def _extrair_linha_xml(sheet_xml: str, numero: int) -> tuple[int, int, str] | None:
    """Acha o <row r="numero">...</row> inteiro no XML e devolve
    (posição inicial, posição final, conteúdo). None se a linha não existir
    ainda como elemento no XML."""
    m = re.search(rf'<row r="{numero}"[^>]*>.*?</row>', sheet_xml, re.S)
    if m:
        return m.start(), m.end(), m.group()
    return None


def _resolver_caminho_aba(zf: zipfile.ZipFile, nome_aba: str) -> tuple[str, str]:
    """Descobre o caminho interno (ex: xl/worksheets/sheet2.xml) da aba
    pedida e o caminho do arquivo de tabela (Tabela1) associado a ela."""
    workbook_xml = zf.read("xl/workbook.xml").decode("utf-8")
    m = re.search(
        r'<sheet name="' + re.escape(nome_aba) + r'"[^>]*r:id="([^"]+)"',
        workbook_xml,
    )
    if not m:
        raise ValueError(f'Não achei a aba "{nome_aba}" no workbook.xml do arquivo mestre.')
    rid = m.group(1)

    rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m2 = re.search(r'<Relationship Id="' + re.escape(rid) + r'"[^>]*Target="([^"]+)"', rels_xml)
    caminho_aba = "xl/" + m2.group(1)

    nome_sheet_rels = caminho_aba.replace("worksheets/", "worksheets/_rels/") + ".rels"
    caminho_tabela = None
    if nome_sheet_rels in zf.namelist():
        sheet_rels_xml = zf.read(nome_sheet_rels).decode("utf-8")
        sheet_xml = zf.read(caminho_aba).decode("utf-8")
        m3 = re.search(r'<tablePart r:id="([^"]+)"/>', sheet_xml)
        if m3:
            rid_tabela = m3.group(1)
            m4 = re.search(r'<Relationship Id="' + re.escape(rid_tabela) + r'"[^>]*Target="([^"]+)"', sheet_rels_xml)
            if m4:
                caminho_tabela = "xl/" + m4.group(1).replace("../", "")

    return caminho_aba, caminho_tabela


def anexar_linhas_no_master(caminho_mestre: Path, linhas_fonte: list[tuple], nome_aba: str) -> int:
    """Faz a cirurgia no .xlsx: escreve as linhas novas na aba de destino
    sem nunca carregar o arquivo inteiro na memória, e SEM NUNCA tocar nas
    colunas P..U (fórmulas). Devolve o número da última linha usada.

    CRITÉRIO — "jogar sempre na primeira linha [disponível]": a tabela de
    destino tem centenas de linhas "reservadas" com fórmulas em P..U que
    hoje calculam vazio (A..O ainda não foram preenchidas). Em vez de criar
    linha nova sempre no fim, a gente REAPROVEITA essas linhas reservadas
    na ordem: escreve o dado real em A..O e preserva P..U 100% intocado.
    Só cria linha de verdade nova (copiando o "molde" de P..U de uma linha
    existente) quando as linhas reservadas acabarem."""
    if not linhas_fonte:
        return 0

    with zipfile.ZipFile(caminho_mestre, "r") as zf:
        caminho_aba, caminho_tabela = _resolver_caminho_aba(zf, nome_aba)
        sheet_xml = zf.read(caminho_aba).decode("utf-8")
        tabela_xml = zf.read(caminho_tabela).decode("utf-8") if caminho_tabela else None
        shared_strings = _carregar_shared_strings(zf)

    # 1) CRITÉRIO — acha a última linha REAL com dado (nunca confia na tag
    #    <dimension> nem em "a célula tem algo escrito").
    ultima_linha_real = _achar_ultima_linha_real(sheet_xml, shared_strings)
    if ultima_linha_real == 0:
        if _cabecalho_destino_bate(sheet_xml, shared_strings):
            ultima_linha_real = 1  # linha 1 = cabeçalho; escreve a partir da 2
        else:
            raise ValueError(
                "Não encontrei nenhuma linha com dado na aba de destino, e o cabeçalho "
                "da linha 1 também não bate com o esperado — isso é sinal de que o "
                "arquivo ou a aba configurados estão errados. Abortando sem mexer no arquivo."
            )

    m_dim = re.search(r'<dimension ref="[A-Z]+\d+:([A-Z]+)\d+"/>', sheet_xml)
    ultima_coluna = m_dim.group(1) if m_dim else LETRAS_COLUNAS[-1]

    # maior linha que EXISTE de verdade como elemento <row> no XML hoje —
    # nunca tirado da tag <dimension> (que pode estar desatualizada). Serve
    # de "molde" pra copiar P..U quando for preciso criar linha nova. A
    # linha 1 (cabeçalho) NUNCA serve de molde — ela tem texto fixo em
    # P..U (os próprios nomes das colunas), não fórmula.
    maior_linha_existente = max(
        (int(m.group(1)) for m in re.finditer(r'<row r="(\d+)"', sheet_xml) if int(m.group(1)) > 1),
        default=0,
    )
    conteudo_modelo_pu = ""
    atributos_modelo = f' spans="1:{config.N_COLUNAS_DESTINO}"'
    if maior_linha_existente:
        achado_modelo = _extrair_linha_xml(sheet_xml, maior_linha_existente)
        if achado_modelo:
            conteudo_modelo_pu = achado_modelo[2]
            atributos_modelo = _atributos_da_linha(sheet_xml, maior_linha_existente)
    else:
        print(
            "  [AVISO] Não sobrou nenhuma linha de dado/reservada pra copiar o molde das "
            "fórmulas das colunas calculadas — essas colunas "
            "vão ficar vazias nas linhas novas desta execução. Depois, no Excel, arraste a "
            "fórmula de uma linha qualquer dessas colunas para baixo, uma vez só; as próximas "
            "execuções do script voltam a copiar o molde normalmente."
        )

    sheet_xml_novo = sheet_xml
    numero_linha = ultima_linha_real
    numeros_usados = []
    motivos_esperados = []

    for linha_fonte in linhas_fonte:
        numero_linha += 1
        motivos_esperados.append(linha_fonte[INDICE_FILTRO])

        # CRITÉRIO — decide "reaproveitar ou criar" perguntando direto pro
        # XML se a linha já existe como elemento, nunca por um número
        # guardado (tipo a tag <dimension>) que pode estar desatualizado.
        achado = _extrair_linha_xml(sheet_xml_novo, numero_linha)

        if achado is not None:
            # 2a) Linha já existe (hoje é uma "reservada" com A..O vazio e
            #     P..U com fórmula) — confere que A..O está mesmo vazio e
            #     substitui SÓ A..O, preservando P..U dessa mesma linha.
            inicio, fim, conteudo_atual = achado
            tem_dado_real = False
            for letra_col, atributos, corpo in _iter_celulas(conteudo_atual):
                if letra_col in COLUNAS_A_O and _valor_real_da_celula(atributos, corpo, shared_strings).strip():
                    tem_dado_real = True
                    break
            if tem_dado_real:
                raise ValueError(
                    f"A linha {numero_linha} deveria estar vazia em A..O (é uma linha 'reservada'), "
                    "mas já tem dado real preenchido nela. Abortando sem mexer no arquivo — "
                    "confira manualmente a aba de destino antes de rodar de novo."
                )
            atributos_linha = _atributos_da_linha(sheet_xml_novo, numero_linha)
            xml_linha_pronta = montar_linha_xml(numero_linha, linha_fonte, conteudo_atual, atributos_linha)
            sheet_xml_novo = sheet_xml_novo[:inicio] + xml_linha_pronta + sheet_xml_novo[fim:]
        else:
            # 2b) Linha não existe ainda — cria de verdade no fim da tabela,
            #     copiando o "molde" de P..U da última linha existente
            #     (mesma fórmula, estilo — só muda o número da linha).
            xml_linha_pronta = montar_linha_xml(numero_linha, linha_fonte, conteudo_modelo_pu, atributos_modelo)
            sheet_xml_novo = sheet_xml_novo.replace(
                "</sheetData>", xml_linha_pronta + "</sheetData>", 1
            )

        numeros_usados.append(numero_linha)

    nova_ultima_linha = numero_linha

    # 3) CRITÉRIO — revisa se pegou tudo mesmo: relê, no XML já montado, o
    #    valor de Motivo (coluna O) de cada linha escrita e confere se bate
    #    com o esperado, na ordem certa, antes de considerar válido.
    for numero, motivo_esperado in zip(numeros_usados, motivos_esperados):
        achado = _extrair_linha_xml(sheet_xml_novo, numero)
        motivo_escrito = ""
        if achado:
            _, _, conteudo = achado
            for letra, atributos, corpo in _iter_celulas(conteudo):
                if letra == LETRA_FILTRO:
                    motivo_escrito = _valor_real_da_celula(atributos, corpo, shared_strings)
                    break
        if motivo_escrito.strip() != str(motivo_esperado).strip():
            raise ValueError(
                f"Falha na conferência: a linha {numero} deveria ter Motivo "
                f'"{motivo_esperado}", mas ficou com "{motivo_escrito}". '
                "Abortando sem mexer no arquivo mestre."
            )

    # 4) Atualiza a <dimension> só se passamos das linhas que já existiam
    if nova_ultima_linha > maior_linha_existente:
        sheet_xml_novo = re.sub(
            r'<dimension ref="([A-Z]+\d+):[A-Z]+\d+"/>',
            rf'<dimension ref="\1:{ultima_coluna}{nova_ultima_linha}"/>',
            sheet_xml_novo,
            count=1,
        )

    # 5) Atualiza o range da Tabela1 (ref + autoFilter) só se precisou crescer
    if tabela_xml is not None and nova_ultima_linha > maior_linha_existente:
        tabela_xml_novo = re.sub(
            r'ref="([A-Z]+\d+):([A-Z]+)\d+"',
            lambda mm: f'ref="{mm.group(1)}:{mm.group(2)}{nova_ultima_linha}"',
            tabela_xml,
        )
    else:
        tabela_xml_novo = tabela_xml

    # 6) Reescreve o .xlsx: copia tudo igual, troca só a aba e a tabela.
    #    Remove o calcChain.xml (Excel reconstrói sozinho) pra evitar ele
    #    ficar dessincronizado com as fórmulas que passaram a ter valor.
    caminho_temp = caminho_mestre.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(caminho_mestre, "r") as zf_in, \
         zipfile.ZipFile(caminho_temp, "w", zipfile.ZIP_DEFLATED) as zf_out:
        for item in zf_in.infolist():
            if item.filename == "xl/calcChain.xml":
                continue
            if item.filename == caminho_aba:
                zf_out.writestr(item, sheet_xml_novo)
            elif caminho_tabela is not None and item.filename == caminho_tabela:
                zf_out.writestr(item, tabela_xml_novo)
            else:
                zf_out.writestr(item, zf_in.read(item.filename))

    caminho_temp.replace(caminho_mestre)
    return nova_ultima_linha
