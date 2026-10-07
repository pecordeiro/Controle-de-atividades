# -*- coding: utf-8 -*-
"""
Painel de Montagem SoBot - servidor com sincronização bidirecional com o Excel.

Todas as contas são feitas pelas FÓRMULAS da planilha; o painel só lê os resultados
e grava nas células amarelas (de entrada).

- Usa sempre a versão mais nova da pasta: acompanhamento_montagem_vN.xlsx (maior N).
- O layout (onde ficam as etapas, indicadores e tabelas) é descoberto pelos títulos,
  então novas versões da planilha funcionam sem mexer no código.
- Planilha aberta no Excel: lê e grava AO VIVO na janela aberta e salva após cada edição.
- Planilha fechada: lê o arquivo e grava com um Excel invisível (preserva gráficos e
  fórmulas). O Excel invisível fecha sozinho alguns segundos depois da última edição.

Execute:  python server.py   ->  abre http://localhost:8765
"""
# ---- Biblioteca padrão do Python ----
import os                      # caminhos e arquivos
import re                      # expressões regulares (versão do arquivo, referências A1)
import sys                     # argumentos de linha de comando e saída do programa
import time                    # relógio, para fechar o Excel invisível quando ocioso
import threading               # temporizadores (timer de ociosidade, abrir o navegador)
import socket                  # nome/IP deste computador (modo servidor em rede)
import atexit                  # roda uma limpeza quando o programa termina
import urllib.parse            # decodifica endereços (%20 etc.) vindos do SharePoint
import urllib.request          # pergunta a outro PC se o painel dele está no ar
import urllib.error
import webbrowser              # abre o painel no navegador padrão
from concurrent.futures import ThreadPoolExecutor   # fila de tarefas em UMA thread só

# ---- Bibliotecas de terceiros (veja requirements.txt) ----
import pythoncom               # COM do Windows: a "ponte" para falar com o Excel
import win32com.client         # controla o Excel (abrir, ler e escrever células, salvar)
import openpyxl                # lê o .xlsx direto do disco quando o Excel não está aberto
from flask import Flask, jsonify, request, send_from_directory   # o servidor web

# ---------- configuração ----------
BASE = os.path.dirname(os.path.abspath(__file__))     # pasta deste arquivo (sistema/)
PASTA = os.path.abspath(os.path.join(BASE, ".."))     # pasta de cima: onde ficam as planilhas
PORT = 8765                                           # porta do painel -> http://localhost:8765
ARQ_SERVIDOR = os.path.join(BASE, "servidor.txt")   # endereço de quem está servindo o painel
HIDDEN_IDLE_S = 12    # segundos sem uso até fechar o Excel invisível usado para gravar

# Nomes das 3 abas da planilha (precisam ser exatamente estes).
MAIN, TEAM, HOW = "Acompanhamento", "Equipe", "Como funciona"
# Quantas (linhas, colunas) lemos de cada aba, a partir de A1.
# A aba principal é lida até a linha 700 para sobrar folga se a planilha crescer.
RANGES = {MAIN: (700, 16), TEAM: (18, 3), HOW: (20, 3)}


def _versao(nome):
    """Número N de um arquivo '..._vN.xlsx'. Sem sufixo _vN conta como versão 1."""
    m = re.search(r"_v(\d+)\.xlsx$", nome, re.I)
    return int(m.group(1)) if m else 1


def escolher_planilha():
    """Versão mais nova: acompanhamento_montagem_vN.xlsx com o maior N (sem _vN = versão 1)."""
    # Candidatas: acompanhamento_montagem*.xlsx. O arquivo "~$..." é o bloqueio que o
    # Excel cria enquanto a planilha está aberta; não é uma planilha de verdade.
    nomes = [f for f in os.listdir(PASTA)
             if f.lower().startswith("acompanhamento_montagem") and f.lower().endswith(".xlsx")
             and not f.startswith("~$")]
    if not nomes:
        return None
    return os.path.normcase(os.path.join(PASTA, max(nomes, key=_versao)))


# Planilha em uso (caminho em minúsculas, para poder comparar com o que o Excel informa).
XLSX = escolher_planilha()


# ---------- layout da planilha (descoberto pelos títulos) ----------
def txt(x):
    """Texto limpo de uma célula: vazio vira '', números viram texto, espaços nas pontas saem."""
    return "" if x is None else str(x).strip()


def detect_layout(M):
    """M = grade de valores da aba Acompanhamento (lista de linhas, 1ª linha = linha 1)."""
    # A planilha NÃO tem linhas fixas aqui: procuramos os títulos na coluna A. Assim,
    # se uma versão nova da planilha mudar as linhas de lugar, o painel continua achando tudo.
    # A(r), B(r), D(r) = texto das colunas A, B e D da linha r (1ª linha = linha 1).
    n = len(M)
    A = lambda r: txt(M[r - 1][0]) if 0 < r <= n else ""
    B = lambda r: txt(M[r - 1][1]) if 0 < r <= n else ""
    D = lambda r: txt(M[r - 1][3]) if 0 < r <= n else ""
    # lay guarda o que foi encontrado: linha da separação, linha de cada etapa, indicadores
    # (nome -> linha), primeira linha das Tabelas 1 e 2, e se a planilha usa "Peças" (v4+).
    lay = {"sep": None, "stages": [], "ind": {}, "t1": None, "t2": None, "pecas": True}
    for r in range(1, n - 1):
        # Quadro opcional de separação de materiais: título e, 2 linhas abaixo, o cabeçalho "Material".
        if A(r) == "Separação de materiais" and A(r + 2) == "Material":
            lay["sep"] = r
        # Início de uma etapa: título, "Unidades previstas" logo abaixo e o cabeçalho
        # "Subatividade | Tempo ideal ..." duas linhas abaixo do título.
        elif A(r + 1) == "Unidades previstas" and A(r + 2) == "Subatividade" and B(r + 2).startswith("Tempo ideal"):
            lay["stages"].append(r)
            # A partir da v4 a coluna D é "Peças concluídas". Se não for, a planilha é antiga
            # (% digitado) e este painel não consegue usá-la.
            if not D(r + 2).startswith("Peças"):
                lay["pecas"] = False      # planilha antiga (% digitado): este painel não serve
        # Tabelas do resumo final: os dados começam 2 linhas depois do título
        # (1 linha de cabeçalho no meio).
        elif A(r).startswith("Tabela 1"):
            lay["t1"] = r + 2
        elif A(r).startswith("Tabela 2"):
            lay["t2"] = r + 2
        # Quadro de indicadores: guarda "nome do indicador -> linha" (ex.: Robôs concluídos -> 377).
        elif A(r) == "Indicadores" and A(r + 1) == "Indicador":
            for k in range(r + 2, r + 16):
                if A(k):
                    lay["ind"][A(k)] = k
    # Cada tabela tem uma linha por bloco (etapas + separação, se houver) mais a linha "Total".
    nb = len(lay["stages"]) + (1 if lay["sep"] else 0)
    lay["t1_rows"] = list(range(lay["t1"], lay["t1"] + nb + 1)) if lay["t1"] else []   # + linha Total
    lay["t2_rows"] = list(range(lay["t2"], lay["t2"] + nb + 1)) if lay["t2"] else []
    return lay


def input_cells(lay):
    """Células amarelas (de entrada) + quais são inteiras, % (0 a 1) e peças concluídas (0 a meta)."""
    # SEGURANÇA: só estas células (as amarelas) podem ser gravadas pelo painel. Nunca uma fórmula.
    #   ok    = células permitidas, no formato "Aba!Célula" (ex.: Acompanhamento!B4)
    #   ints  = as que precisam ser número inteiro
    #   fracs = as que são % (valor de 0 a 1); vazio na v4, que usa peças em vez de %
    # Fixas: B4 = meta de robôs; Equipe!B6 = pessoas; Equipe!B7 = horas por dia.
    ok, ints, fracs = {f"{MAIN}!B4", f"{TEAM}!B6", f"{TEAM}!B7"}, {f"{MAIN}!B4"}, set()
    if lay["sep"]:
        # Separação: tempo por kit (1 linha abaixo do título) e 8 linhas de materiais
        # (colunas A = nome, B = quantidade total, D = kits separados).
        s0 = lay["sep"]
        ok.add(f"{MAIN}!D{s0+1}")
        for r in range(s0 + 3, s0 + 11):
            ok |= {f"{MAIN}!A{r}", f"{MAIN}!B{r}", f"{MAIN}!D{r}"}
            ints.add(f"{MAIN}!D{r}")
    # Etapas: o nome (coluna A do título) e 8 linhas de subatividades
    # (A = nome, B = tempo ideal em minutos, D = peças concluídas).
    pecas = set()   # células de "peças concluídas": não podem passar da meta
    for t in lay["stages"]:
        ok.add(f"{MAIN}!A{t}")
        for r in range(t + 3, t + 11):
            ok |= {f"{MAIN}!A{r}", f"{MAIN}!B{r}", f"{MAIN}!D{r}"}
            ints.add(f"{MAIN}!D{r}")
            pecas.add(f"{MAIN}!D{r}")
    return ok, ints, fracs, pecas


# ---------- acesso ao Excel (thread única COM) ----------
# O COM do Windows exige que todas as chamadas ao Excel venham da MESMA thread. Por isso toda
# leitura/gravação é enviada a este executor de 1 só thread (o Flask atende vários pedidos ao
# mesmo tempo, mas eles entram numa fila e o Excel só é tocado por um de cada vez).
executor = ThreadPoolExecutor(max_workers=1, initializer=pythoncom.CoInitialize)
# Estado compartilhado:
#   mode     "excel" (lendo a janela aberta) ou "arquivo" (lendo do disco)
#   grids    última leitura: {aba: [[valores...], ...]}
#   mtime    data de modificação do arquivo na última leitura (evita reler sem mudança)
#   error    mensagem de erro para mostrar na tela, se houver
#   hidden   (Excel, planilha) do Excel invisível aberto para gravar, ou None
#   last_use quando o Excel invisível foi usado pela última vez
state = {"mode": "?", "grids": None, "mtime": None, "error": None, "hidden": None, "last_use": 0}


def mesma_planilha(nome):
    """Compara o nome do ROT com a planilha. Aceita caminho local ou endereço do
    OneDrive/SharePoint (https://…/Pré-montagem/arquivo.xlsx): confere pasta + arquivo."""
    if os.path.normcase(nome) == XLSX:
        return True
    # Planilha aberta pelo OneDrive/SharePoint: o Excel informa uma URL (https://...) e não um
    # caminho do disco. Então comparamos só o nome da pasta + o nome do arquivo.
    if nome.lower().startswith("http"):
        partes = urllib.parse.unquote(nome).replace("\\", "/").rstrip("/").split("/")
        local = XLSX.replace("\\", "/").split("/")
        return [p.lower() for p in partes[-2:]] == [p.lower() for p in local[-2:]]
    return False


def find_open_workbook():
    """Procura a planilha aberta em qualquer instância do Excel (Running Object Table)."""
    try:
        # ROT = tabela do Windows com os objetos COM em uso. Cada Excel aberto registra ali
        # suas planilhas, então dá para "pegar" a planilha já aberta sem abrir outra cópia.
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        for mk in rot:
            try:
                name = mk.GetDisplayName(ctx, None)
            except Exception:
                continue
            if mesma_planilha(name):
                obj = rot.GetObject(mk)
                return win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
    except Exception:
        pass
    return None


def clean(v):
    """Troca os códigos de erro do Excel (#DIV/0!, #REF!...) por '#ERRO', que a tela entende."""
    if isinstance(v, int) and v < -2146820000:   # erro do Excel (#DIV/0!, #REF! …)
        return "#ERRO"
    return v


def read_com(wb):
    """Lê as 3 abas direto do Excel (planilha aberta). Pega valores já calculados, em um só pedido por aba."""
    grids = {}
    for sheet, (rows, cols) in RANGES.items():
        ws = wb.Worksheets(sheet)
        vals = ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols)).Value2
        grids[sheet] = [[clean(v) for v in row] for row in vals]
    return grids


def read_file():
    """Lê as 3 abas do arquivo no disco (Excel fechado). Reaproveita a leitura anterior se o arquivo não mudou."""
    mtime = os.path.getmtime(XLSX)
    if state["grids"] is not None and state["mode"] == "arquivo" and state["mtime"] == mtime:
        return state["grids"]
    # data_only=True: traz o ÚLTIMO RESULTADO calculado das fórmulas (o Excel o guarda no arquivo).
    wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=True)
    grids = {}
    for sheet, (rows, cols) in RANGES.items():
        grid = [list(r) for r in wb[sheet].iter_rows(min_row=1, max_row=rows, max_col=cols, values_only=True)]
        grid = [row + [None] * (cols - len(row)) for row in grid]
        grid += [[None] * cols for _ in range(rows - len(grid))]
        grids[sheet] = grid
    wb.close()
    state["mtime"] = mtime
    return grids


def close_hidden():
    """Salva e fecha o Excel invisível (se houver), liberando o arquivo para outras pessoas."""
    h = state["hidden"]
    if not h:
        return
    state["hidden"] = None
    xl, wb = h
    try:
        wb.Close(True)
    except Exception:
        pass
    try:
        xl.Quit()
    except Exception:
        pass
    state["mtime"] = None


def _idle_check():
    """Fecha o Excel invisível se ninguém gravou nada nos últimos HIDDEN_IDLE_S segundos."""
    if state["hidden"] and time.time() - state["last_use"] >= HIDDEN_IDLE_S:
        close_hidden()


def agendar_idle_check():
    """Pede ao executor para rodar _idle_check (chamada pelo timer, em outra thread)."""
    try:
        executor.submit(_idle_check)
    except RuntimeError:        # servidor já está desligando
        pass


def _read():
    """Atualiza state["grids"] com os dados mais recentes. Ordem de preferência:
    1) Excel invisível que nós mesmos abrimos; 2) planilha aberta pelo usuário; 3) arquivo no disco."""
    if state["hidden"]:
        try:
            state.update(grids=read_com(state["hidden"][1]), mode="arquivo", error=None)
            return
        except pythoncom.com_error:
            close_hidden()
    wb = find_open_workbook()
    if wb is not None:
        try:
            state.update(grids=read_com(wb), mode="excel", error=None)
        except pythoncom.com_error:
            state["error"] = "Excel ocupado (célula em edição). Tentando de novo…"
        return
    state.update(grids=read_file(), mode="arquivo", error=None)


def _write(changes):
    """Grava uma lista de ('Aba!Célula', valor) na planilha e salva.
    Se a planilha estiver aberta no Excel do usuário, escreve nela (ao vivo). Senão abre um Excel
    invisível (que fecha sozinho depois de HIDDEN_IDLE_S s). Usar o Excel de verdade garante que
    gráficos, formatação e fórmulas da planilha não se percam, e que as contas sejam recalculadas."""
    wb = find_open_workbook() if not state["hidden"] else None
    if wb is None:
        if not state["hidden"]:
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            state["hidden"] = (xl, xl.Workbooks.Open(XLSX))
        wb = state["hidden"][1]
        state["last_use"] = time.time()
        timer = threading.Timer(HIDDEN_IDLE_S + 0.5, agendar_idle_check)
        timer.daemon = True
        timer.start()
    # Valor vazio ("") apaga a célula.
    for addr, val in changes:
        sheet, cell = addr.split("!")
        wb.Worksheets(sheet).Range(cell).Value = "" if val is None else val
    wb.Save()
    _read()   # relê já com as fórmulas recalculadas, para devolver o resultado novo à tela


def coerce(addr, value, ints, fracs, pecas=(), meta=None):
    """Valida e converte o que veio da tela. Levanta ValueError (mensagem em português) se for inválido.
    Textos (coluna A) passam; números aceitam vírgula; sem negativos; inteiros onde exigido;
    peças concluídas não podem passar da meta."""
    if value is None or value == "":
        return None
    sheet, cell = addr.split("!")
    if sheet == MAIN and cell[0] == "A":
        return str(value).strip()
    v = float(str(value).replace(",", "."))
    if v < 0:
        raise ValueError("O valor não pode ser negativo.")
    if addr in fracs and v > 1:
        raise ValueError("O % concluído vai de 0% a 100%.")
    if addr in ints:
        if not v.is_integer():
            raise ValueError("Use um número inteiro.")
        if addr in pecas and meta is not None and v > meta:
            raise ValueError(f"As peças concluídas não podem passar das previstas ({int(meta)}).")
        return int(v)
    return v


# ---------- montagem do JSON ----------
def snapshot():
    """Monta o JSON que a tela consome, a partir da última leitura. Devolve None se não há dados
    ou se a planilha é de uma versão antiga. NENHUMA conta é feita aqui: tudo vem das fórmulas."""
    G = state["grids"]
    if G is None:
        return None
    M, T, H = G[MAIN], G[TEAM], G[HOW]

    # v(grade, "B4") -> valor da célula B4 (vazio vira ""). m(...) é o atalho para a aba principal.
    def v(grid, ref):
        mm = re.fullmatch(r"([A-Z])(\d+)", ref)
        x = grid[int(mm.group(2)) - 1][ord(mm.group(1)) - 65]
        return "" if x is None else x

    m = lambda ref: v(M, ref)
    # A planilha marca dados de demonstração com a palavra EXEMPLO; a tela mostra um aviso.
    exemplo = lambda r: "EXEMPLO" in str(m(f"A{r}")).upper()
    lay = detect_layout(M)
    if lay["stages"] and not lay["pecas"]:
        state["error"] = (f"A planilha {os.path.basename(XLSX)} é de uma versão antiga (% digitado). "
                          "Este painel precisa da acompanhamento_montagem_v4.xlsx ou mais nova.")
        return None

    # "blocks" = lista de quadros na ordem da planilha: primeiro a separação (se houver) e
    # depois cada etapa. As posições relativas (título+1, título+3...) seguem o desenho da planilha.
    blocks = []
    if lay["sep"]:
        s0 = lay["sep"]
        blocks.append({
            "kind": "sep", "titulo": m(f"A{s0}"), "titleRow": s0,
            "tempoKit": m(f"D{s0+1}"), "tempoKitCell": f"D{s0+1}", "previstas": m(f"B{s0+1}"),
            "rows": [{"row": r, "material": m(f"A{r}"), "qtdTotal": m(f"B{r}"), "qtdRobo": m(f"C{r}"),
                      "kits": m(f"D{r}"), "pct": m(f"E{r}"), "falta": m(f"F{r}")} for r in range(s0 + 3, s0 + 11)],
            "totalQtd": m(f"B{s0+11}"), "pctMedio": m(f"E{s0+11}"), "faltaTotal": m(f"F{s0+11}"),
            "kitsFrac": m(f"C{s0+12}"), "kitsPct": m(f"E{s0+12}"), "exemplo": exemplo(s0 + 13),
        })
    # Cada etapa: título (t), previstas (t+1), cabeçalho (t+2), 8 subatividades (t+3..t+10),
    # total (t+11) e "unidades concluídas ÷ previstas" (t+12). As colunas P/O à direita guardam
    # contadores auxiliares da planilha (tarefas e tarefas 100%).
    for t in lay["stages"]:
        blocks.append({
            "kind": "etapa", "titulo": m(f"A{t}"), "titleRow": t,
            "pesoRobo": m(f"F{t}"), "previstas": m(f"B{t+1}"), "concluidas": m(f"F{t+1}"),
            "rows": [{"row": r, "sub": m(f"A{r}"), "tempo": m(f"B{r}"), "peso": m(f"C{r}"),
                      "pecas": m(f"D{r}"), "pct": m(f"E{r}"), "tTotal": m(f"F{r}"), "tFalta": m(f"G{r}")}
                     for r in range(t + 3, t + 11)],
            "tempoEtapa": m(f"B{t+11}"), "progresso": m(f"E{t+11}"),
            "tTotal": m(f"F{t+11}"), "tFalta": m(f"G{t+11}"),
            "unidFrac": m(f"D{t+12}"), "unidPct": m(f"E{t+12}"),
            "tarefas": m(f"P{t+3}"), "tarefas100": m(f"P{t+4}"),
            "exemplo": exemplo(t + 13),
        })

    # Tabela 1 (progresso pelas tarefas) e Tabela 2 (unidades concluídas), já com a linha Total.
    t1 = [{"etapa": m(f"A{r}"), "tempoRobo": m(f"B{r}"), "peso": m(f"C{r}"), "previstas": m(f"D{r}"),
           "progresso": m(f"E{r}"), "hTotal": m(f"F{r}"), "hFalta": m(f"G{r}")}
          for r in lay["t1_rows"]]
    t2 = [{"etapa": m(f"A{r}"), "feitas": m(f"B{r}"), "previstas": m(f"C{r}"), "frac": m(f"D{r}"),
           "pct": m(f"E{r}"), "status": m(f"F{r}")} for r in lay["t2_rows"]]

    # Indicadores: localizados pelo NOME na coluna A. trio = (feito, total, %); one = valor único.
    I = lay["ind"]
    trio = lambda k: ({"feito": m(f"B{I[k]}"), "total": m(f"C{I[k]}"), "pct": m(f"D{I[k]}")}
                      if k in I else {"feito": "", "total": "", "pct": ""})
    one = lambda k: m(f"B{I[k]}") if k in I else ""
    ind = {"robos": trio("Robôs concluídos"),
           "etapas": trio("Etapas concluídas"), "tarefas": trio("Tarefas (subatividades) concluídas"),
           "unidades": trio("Unidades concluídas (soma das etapas)"),
           "tempo": trio("Tempo trabalhado (h) = progresso geral"),
           "tempoRobo": one("Tempo de 1 robô (h)"), "falta": one("Tempo que falta (h)"),
           "dias": one("Dias úteis restantes (aba Equipe)"), "semTempo": one("Etapas ainda sem tempo definido")}

    # Aba Equipe: pessoas, horas/dia, dias úteis e a tabela "e se mudar o tamanho da equipe?" (linhas 13-18).
    team = {"pessoas": v(T, "B6"), "horasDia": v(T, "B7"), "diasRest": v(T, "B8"), "diasMeta": v(T, "B9"),
            "cenarios": [{"pessoas": v(T, f"A{r}"), "horas": v(T, f"B{r}"), "dias": v(T, f"C{r}")}
                         for r in range(13, 19)]}
    # Aba "Como funciona": explicação de cada conta, com exemplo calculado pela própria planilha.
    how = [{"conta": v(H, f"A{r}"), "como": v(H, f"B{r}"), "exemplo": v(H, f"C{r}")}
           for r in range(5, 21) if v(H, f"A{r}") != ""]
    return {"mode": state["mode"], "error": state["error"], "arquivo": os.path.basename(XLSX),
            "titulo": m("A1"), "meta": m("B4"), "blocks": blocks,
            "t1": t1, "t2": t2, "ind": ind, "team": team, "how": how}


# ---------- API (o que a tela pede ao servidor) ----------
#   GET  /           -> a página (index.html)
#   GET  /api/data   -> todos os dados atuais da planilha (a tela pergunta a cada 2 s)
#   POST /api/set    -> grava células de entrada e devolve os dados já recalculados
app = Flask(__name__, static_folder=None)


@app.get("/api/data")
def api_data():
    """Lê a planilha (pela fila do Excel) e devolve o JSON. 503 se não houver dados."""
    try:
        executor.submit(_read).result(timeout=30)
    except Exception as e:
        state["error"] = f"Falha ao ler a planilha: {e}"
    s = snapshot()
    if s is None:
        return jsonify({"error": state["error"] or "Sem dados"}), 503
    return jsonify(s)


@app.post("/api/set")
def api_set():
    """Recebe {"changes": [{"cell": "Acompanhamento!D15", "value": 5}, ...]}.
    Passos: 1) confere se cada célula é de preenchimento; 2) valida o valor; 3) grava no Excel;
    4) devolve o JSON atualizado. Erros voltam como {"ok": false, "error": "mensagem"}."""
    changes = (request.get_json(force=True) or {}).get("changes", [])
    try:
        if state["grids"] is None:
            executor.submit(_read).result(timeout=30)
        # As regras do que pode ser gravado são recalculadas a cada pedido, a partir da
        # planilha atual (assim funcionam mesmo se o layout mudar).
        allowed, ints, fracs, pecas = input_cells(detect_layout(state["grids"][MAIN]))
        # B4 (linha 4, coluna 2) = meta de robôs; é o teto das "peças concluídas".
        meta = state["grids"][MAIN][3][1]
        meta = meta if isinstance(meta, (int, float)) else None
        clean_changes = []
        for ch in changes:
            addr = str(ch["cell"])
            if "!" not in addr:
                addr = f"{MAIN}!{addr}"
            sheet, cell = addr.split("!")
            addr = f"{sheet}!{cell.upper()}"
            # Lista branca: qualquer célula fora das amarelas é recusada (protege as fórmulas).
            if addr not in allowed:
                return jsonify({"ok": False, "error": f"A célula {addr} não é de preenchimento."}), 400
            clean_changes.append((addr, coerce(addr, ch.get("value"), ints, fracs, pecas, meta)))
        executor.submit(_write, clean_changes).result(timeout=120)
    # 409 = o Excel recusou (normalmente porque o usuário está com uma célula em edição).
    except pythoncom.com_error:
        return jsonify({"ok": False, "error": "O Excel está ocupado: saia da célula que está editando no Excel e tente de novo."}), 409
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"Erro ao salvar: {e}"}), 500
    return jsonify({"ok": True, "data": snapshot()})


@app.get("/")
def index():
    """Entrega a página do painel."""
    return send_from_directory(BASE, "index.html")


def servidor_ativo():
    """Se outro PC já está servindo o painel, devolve a URL dele."""
    try:
        url = open(ARQ_SERVIDOR, encoding="utf-8").read().strip()
        urllib.request.urlopen(url + "/api/data", timeout=3)
        return url
    except urllib.error.HTTPError:
        return url          # respondeu (mesmo com erro) = está no ar
    except Exception:
        return None


def remover_registro():
    """Ao encerrar, apaga o servidor.txt se ele apontava para ESTE computador."""
    try:
        if open(ARQ_SERVIDOR, encoding="utf-8").read().strip() == f"http://{socket.gethostname()}:{PORT}":
            os.remove(ARQ_SERVIDOR)
    except Exception:
        pass


# ---------- início do programa ----------
# Modo rede: o primeiro PC que abre o painel vira o SERVIDOR e registra seu endereço em
# servidor.txt. Quem abrir depois só é redirecionado para esse endereço, para existir um
# único servidor gravando na planilha. "--force" ignora isso; "--no-browser" não abre o navegador.
if __name__ == "__main__":
    if not XLSX:
        sys.exit("Nenhuma planilha acompanhamento_montagem*.xlsx encontrada na pasta Pré-montagem.")
    abrir = "--no-browser" not in sys.argv
    outro = servidor_ativo()
    if outro and "--force" not in sys.argv:
        print(f"Painel já está rodando em {outro} - conectando a ele.")
        if abrir:
            webbrowser.open(outro)
        sys.exit(0)
    host = socket.gethostname()
    ip = socket.gethostbyname(host)
    with open(ARQ_SERVIDOR, "w", encoding="utf-8") as f:
        f.write(f"http://{host}:{PORT}")
    atexit.register(remover_registro)
    print(f"Planilha: {XLSX}")
    print(f"Este PC e o SERVIDOR do painel. Os outros PCs abrem o mesmo iniciar_painel.bat")
    print(f"ou acessam http://{host}:{PORT}  /  http://{ip}:{PORT}  (feche esta janela para parar)")
    if abrir:
        threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    try:
        # ATENÇÃO: host="0.0.0.0" abre o painel para a rede inteira e NÃO há login.
        # Quem alcançar a porta pode ver e editar a planilha. Use só em rede confiável/VPN.
        app.run(host="0.0.0.0", port=PORT, threaded=True)
    finally:
        executor.submit(close_hidden).result(timeout=30)
