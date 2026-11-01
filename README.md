# Excel Table Appender

> **EN summary:** Python tool that reads many daily spreadsheets, keeps only the rows where a key column (e.g. "Motivo"/reason) is filled, and appends them to a table in a master `.xlsx` workbook by editing the file's internal XML — so the (large, pivot-heavy) master is never loaded into memory. It reuses pre-built formula rows, never touches formula columns, verifies every written row, makes a backup first and aborts without changes on any inconsistency. No Excel installation needed.

Ferramenta em Python que **consolida em uma tabela mestre as linhas relevantes de dezenas de planilhas**, sem abrir o arquivo mestre inteiro.

## O problema

Chegam várias planilhas de análise. Só interessam as linhas em que a coluna **"Motivo"** está preenchida, e elas precisam ir para uma aba do arquivo mestre, que alimenta tabelas dinâmicas. O processo manual era:

1. abrir planilha por planilha;
2. filtrar as linhas com motivo;
3. conferir os dados;
4. copiar e colar no mestre, sem sobrescrever nem pular linhas.

Com dezenas de arquivos por rodada, era demorado e fácil de errar. Automatizar carregando o mestre inteiro (com `openpyxl` ou Power Query) também é pesado: quando o arquivo tem muitas tabelas dinâmicas, o cache delas pode ter dezenas de MB de XML e a memória pode estourar.

## A solução

Um `.xlsx` é um `.zip` de arquivos XML. Em vez de carregar o mestre, o script:

```
 pasta de entrada                       arquivo mestre (.xlsx = zip)
┌──────────────────┐   filtra Motivo   ┌───────────────────────────────┐
│ planilha 1       │ ───────────────►  │ troca SÓ o XML da aba         │
│ planilha 2       │   (modo leitura,  │ e o XML da tabela;            │
│ ...              │    leve)          │ copia o resto byte a byte     │
└──────────────────┘                   │ (inclusive caches pesados)    │
        │                              └───────────────────────────────┘
        └── arquivos processados vão para "Processados/"
```

1. Lê cada planilha de entrada em **modo leitura** e acha a aba cujo cabeçalho bate com o esperado.
2. Guarda só as linhas com a coluna de filtro preenchida.
3. Escreve nas **linhas reservadas** da tabela do mestre (as que já têm fórmula pronta e dados vazios), na ordem, sem tocar nas colunas de fórmula.
4. **Relê** cada linha escrita e confere o valor da coluna de filtro. Se algo não bater, aborta sem alterar o mestre.
5. Move o arquivo de entrada para `Processados/`.

## Resultado

Em uma rodada real, com **72 arquivos (146 linhas com motivo)**, o trabalho manual levava **cerca de 1 hora** e passou a levar **menos de 1 minuto** (medição aproximada).

## Recursos

- **Nunca abre o mestre inteiro**: trabalha direto no ZIP, então não estoura memória com tabelas dinâmicas pesadas.
- **Não mexe nas colunas de fórmula**: preserva fórmula, tipo e estilo byte a byte.
- **Sempre grava na primeira linha disponível**: reaproveita as linhas reservadas e só cria linhas novas (copiando o molde das fórmulas) quando elas acabam; atualiza o intervalo da tabela quando precisa.
- **Conferência pós-escrita**: relê cada linha gravada e compara com o esperado antes de salvar.
- **Backup automático** do mestre antes de qualquer alteração.
- **Escrita segura**: grava em arquivo temporário e só então substitui o mestre; em qualquer erro, o mestre fica como estava.
- **Cabeçalho validado** nas planilhas de entrada: arquivo fora do padrão não é processado nem movido.
- **Não precisa de Excel instalado** (usa só `openpyxl` e manipulação de ZIP/XML).

## Requisitos

- Python 3.9 ou superior
- `openpyxl`
- Arquivo mestre em `.xlsx`, salvo pelo Excel, com uma **Tabela do Excel** na aba de destino e **linhas reservadas** com as colunas de fórmula já preenchidas

## Instalação

```bash
git clone https://github.com/gabrielermogenes/excel-table-appender.git
cd excel-table-appender

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
```

## Configuração

```bash
copy config.example.py config.py
```

Abra o `config.py` e ajuste:

| Variável | O que informar |
|---|---|
| `PASTA_ENTRADA` | Pasta onde você joga as planilhas de entrada |
| `ARQUIVO_MESTRE` | Caminho do mestre `.xlsx` (comece apontando para uma **cópia**) |
| `ABA_DESTINO` | Nome da aba do mestre que contém a Tabela de destino |
| `PADROES_ARQUIVO` | Padrões de arquivo a ler (por padrão `*.xlsx` e `*.xlsm`) |
| `CABECALHO_ESPERADO` | Cabeçalho das planilhas de entrada. **A última coluna é o filtro** |
| `DESTINO_CABECALHO_ESPERADO` | Cabeçalho da tabela de destino (mesma quantidade de colunas) |
| `N_COLUNAS_DESTINO` | Total de colunas da tabela de destino: dados + fórmulas (máx. 26) |
| `TIPO_COLUNA_FONTE` | Tipo de cada coluna de dados: `"texto"` ou `"numero"` (datas contam como número) |
| `ESTILOS_FALLBACK_A_O` | Estilo de reserva por coluna (normalmente pode ficar vazio) |

> O `config.py` está no `.gitignore` e **não deve ser publicado**: ele contém caminhos e nomes do seu ambiente.

Nos comentários do código, "A..O" são as colunas de dados e "P..U" as colunas de fórmula do layout original. Com o seu layout, as letras são calculadas a partir do `config.py`.

## Como usar

1. Coloque as planilhas de entrada em `PASTA_ENTRADA`.
2. Rode:
   ```bash
   python main.py
   ```
3. Leia o resumo no terminal. Os arquivos processados vão para `Processados/` e os que falharam continuam na pasta, com a mensagem do motivo.

**Feche o arquivo mestre no Excel antes de rodar.** Se ele estiver aberto, o script avisa e não altera nada.

## Estrutura

```
excel-table-appender/
├── main.py                 # ponto de entrada (orquestração)
├── config.example.py       # modelo de configuração (copie para config.py)
├── requirements.txt
├── requirements-dev.txt    # pytest
├── src/
│   ├── layout.py           # constantes derivadas do config
│   ├── fonte.py            # leitura e filtro das planilhas de entrada
│   └── mestre.py           # escrita no mestre (ZIP/XML)
└── tests/                  # testes com planilhas sintéticas
```

## Testes

```bash
pip install -r requirements-dev.txt
python -m pytest
```

Os testes geram planilhas sintéticas (sem nenhum dado real) e cobrem: filtro das linhas, reaproveitamento das linhas reservadas, criação de linhas novas e extensão da tabela, várias rodadas seguidas, aborto sem alterar o arquivo e o fluxo completo (backup e `Processados/`).

## Segurança e boas práticas

- Nunca versione planilhas, backups ou o `config.py` (o `.gitignore` já bloqueia esses tipos).
- Teste sempre em uma **cópia** do arquivo mestre antes de usar o de produção.
- Cada rodada gera um backup ao lado do mestre; apague os antigos de vez em quando.

## Limitações

- Lê o XML no formato que o **próprio Excel grava**. Arquivos gerados por outras bibliotecas podem não ser reconhecidos.
- Uma Tabela do Excel por aba de destino; no máximo 26 colunas (letras simples).
- As fórmulas de linhas novas são copiadas da última linha; funciona bem com colunas calculadas de Tabela do Excel (referências estruturadas).
- Remove o `calcChain.xml` do mestre (o Excel recria sozinho ao abrir).
- Mestre em `.xlsx` (não `.xlsm`).

## Licença

MIT. Veja o arquivo [LICENSE](LICENSE).

## Autor

**Gabriel Ermogenes de Paula** — [LinkedIn](https://www.linkedin.com/in/devgabrielermogenes) · [GitHub](https://github.com/gabrielermogenes)
