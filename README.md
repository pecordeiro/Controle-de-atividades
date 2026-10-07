# Painel de Montagem SoBot

Painel web (estilo Power BI) para acompanhar a montagem de robôs, **sincronizado nos dois sentidos com uma planilha Excel**: o que você digita no painel aparece na planilha e o que você muda no Excel aparece no painel em até 2 segundos.

> **Repositório privado.** Contém a planilha de exemplo com os nomes das etapas da SoBot. Não torne público sem revisar o conteúdo e confirmar a autorização da empresa.

## Ideia central

**O painel não faz nenhuma conta.** Progresso, pesos, tempos e dias úteis são calculados pelas **fórmulas da própria planilha**; o painel apenas mostra o resultado e grava nas células amarelas (de entrada). Assim o número do painel é sempre igual ao do Excel.

Fluxo de trabalho que a planilha implementa:

1. Listar as **subatividades** de cada etapa e o **tempo ideal** de cada uma → o tempo define o **peso** (quanto cada uma representa).
2. Definir a **meta de robôs** → vira as "previstas" de todas as etapas.
3. Informar quantas **peças** de cada subatividade já estão concluídas.
4. Duas visões do progresso:
   * **Tabela 1 – pelas tarefas:** conta as parciais (todas em 90% = 90%).
   * **Tabela 2 – por unidades:** só conta unidade 100% pronta (3 de 10 = 30%, mesmo com tudo em 90%).
5. **Equipe:** pessoas × horas por dia → dias úteis restantes.

## Arquitetura

```
 Navegador                    Servidor (janela preta)              Planilha
┌──────────────────┐ pede    ┌──────────────────────────┐ lê/grava ┌────────────────────────────┐
│ sistema/         │ ──────► │ sistema/server.py        │ ───────► │ acompanhamento_            │
│   index.html     │ ◄────── │ Flask + COM do Excel     │ ◄─────── │   montagem_vN.xlsx         │
│ telas + gráficos │  JSON   │ valida e grava           │          │ (fórmulas fazem as contas) │
└──────────────────┘         └──────────────────────────┘          └────────────────────────────┘
```

| Arquivo | Papel |
|---|---|
| `sistema/server.py` | Servidor Flask. Lê e grava a planilha, valida o que a tela envia e expõe a API. |
| `sistema/index.html` | A tela inteira (HTML + CSS + JavaScript, gráficos com [ECharts](https://echarts.apache.org/)). |
| `sistema/iniciar_painel.bat` | Dois cliques para ligar. |
| `sistema/requirements.txt` | Dependências Python. |
| `acompanhamento_montagem_v4.xlsx` | **Planilha de exemplo** (3 abas: Acompanhamento, Equipe, Como funciona). |

### API

| Método | Rota | O que faz |
|---|---|---|
| `GET` | `/` | Entrega `index.html`. |
| `GET` | `/api/data` | Devolve todos os dados atuais da planilha (a tela consulta a cada 2 s). |
| `POST` | `/api/set` | Recebe `{"changes":[{"cell":"Acompanhamento!D15","value":5}]}`, valida, grava no Excel e devolve os dados recalculados. |

### Como o servidor fala com o Excel

* **Planilha aberta no Excel:** o servidor localiza a janela aberta (Running Object Table do Windows, inclusive abertas via OneDrive/SharePoint), lê e grava **ao vivo** e salva.
* **Planilha fechada:** lê o arquivo com `openpyxl` e, para gravar, abre um **Excel invisível** que fecha sozinho 12 s depois da última edição. Usar o Excel de verdade preserva gráficos e fórmulas e garante o recálculo.
* Todas as chamadas ao Excel passam por **uma única thread** (exigência do COM).

### Descoberta automática do layout

O servidor **não usa linhas fixas**: procura títulos na coluna A (`Subatividade`, `Unidades previstas`, `Indicadores`, `Tabela 1`, `Tabela 2`, `Separação de materiais`). Se uma versão nova da planilha mover as linhas, o painel continua funcionando. Ele usa sempre a **maior versão** `acompanhamento_montagem_vN.xlsx` da pasta.

## Requisitos

* Windows com **Microsoft Excel** instalado (o controle é via COM).
* **Python 3.10+**.
* Internet na primeira abertura (biblioteca de gráficos carregada de uma CDN).

## Como rodar

```bash
git clone <url-deste-repositorio>
cd painel-montagem-sobot/sistema
python -m pip install -r requirements.txt
python server.py
```

Ou dê dois cliques em `sistema/iniciar_painel.bat`. O painel abre em <http://localhost:8765>. **Deixe a janela preta aberta** enquanto usar; fechá-la desliga o painel.

Opções do `server.py`: `--no-browser` (não abre o navegador) e `--force` (ignora outro servidor já ativo na rede).

### Uso em rede

O primeiro PC que abre o painel vira o **servidor** e registra o endereço em `sistema/servidor.txt`. Quem abrir depois em outro PC é redirecionado para ele, de modo que só um servidor grava na planilha.

## Segurança — leia antes de usar em rede

* O servidor escuta em `0.0.0.0` e **não tem login**: qualquer pessoa que alcance a porta 8765 pode ver e **editar** a planilha. Use apenas em rede confiável ou VPN.
* Só as **células amarelas** (lista branca, calculada a cada pedido) podem ser gravadas; fórmulas nunca são sobrescritas.
* O servidor valida cada valor (sem negativos, inteiros onde exigido, peças ≤ meta) além da validação da tela.
* Todo texto vindo da planilha é escapado antes de entrar na página (`esc()`), evitando injeção de HTML.

## Limitações conhecidas

* Só funciona em **Windows + Excel** (não dá para hospedar em Linux/nuvem comum).
* Sem autenticação nem histórico de alterações.
* As 24 etapas e a estrutura da planilha são as da SoBot (não há tela para criar etapas novas).

## Ideias para evoluir

* Trocar o COM por **Microsoft Graph (Excel Online)** para hospedar fora do PC da empresa, mantendo a planilha como fonte.
* Ou substituir a planilha por **banco de dados** (Postgres), com importação/exportação de Excel.
* Login e perfis (só ver / editar), histórico de mudanças, etapas configuráveis pela tela.

## Estrutura e convenções do código

Os arquivos `server.py` e `index.html` são **bastante comentados** em português: cada função tem uma docstring explicando o que faz, e os trechos menos óbvios (COM, descoberta de layout, validações) têm comentários explicando o porquê.
