"""
Gera conversao.js (painel "Triagem dos retirados") rastreando cada equipamento pelo N° de Série.

Uso (pares movimentação + aplicados/retirados do mesmo mês, uma ou mais operações):
    python atualizar_conversao.py MOV_NORDESTE.xlsx APLIC_NORDESTE.xlsx [MOV_SUL.xlsx APLIC_SUL.xlsx ...]

Entradas (SINAPSE):
  1. Conferência / movimentação de materiais: "Cod.Compl." (N° de Série), "Tipo Movimento"
     (RM = entrega ao técnico, DM = devolução, CM = conversão), "Entrada", "Saída".
  2. Materiais Aplicados/Retirados (completo ou só família de retirados).

Saída por almoxarifado (coluna "Projeto"):
  hubs   = todas as conversões de RET no mês: [REAP, SUC, DET, DEF]
  status = cada RET convertido em REAP no mês: [instalado, com o técnico, no almoxarifado]
  ret    = cada equipamento RETIRADO no mês (pela série), situação no fim do mês:
           [REAP, SUC, DET, DEF, estornado, devolvido aguardando triagem, não devolvido]
           (só conta conversão feita depois da data da retirada)
Vão para o site apenas totais por almoxarifado (nenhuma série, cliente ou técnico).
"""
import json, re, sys, unicodedata
from datetime import datetime
from pathlib import Path
import pandas as pd

RAIZ = Path(__file__).resolve().parent
MESES = "jan fev mar abr mai jun jul ago set out nov dez".split()
DESTINOS = ["REAP", "SUC", "DET", "DEF"]
STATUS = ["instalado", "tecnico", "almox"]
RET_K = ["REAP", "SUC", "DET", "DEF", "estorno", "aguardando", "naodev"]
APELIDOS = {"sj rio do peixe": "sjrp", "inst vt picui": "picui", "sto ant platina": "sap", "sto ant": "sap",
            "ourofino": "ouro fino", "laranj paulista": "laranjal"}

fold = lambda x: re.sub(r"[^a-z0-9]+", " ", unicodedata.normalize("NFD", str(x)).encode("ascii", "ignore").decode().lower()).strip()
num = lambda s: pd.to_numeric(s.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce").fillna(0)
serie = lambda s: s.astype(str).str.strip().str.upper()
VAZIO = {"-", "NAN", ""}

def carregar_js(nome):
    txt = (RAIZ / nome).read_text(encoding="utf-8")
    return json.loads(txt[txt.index("=") + 1:].strip().rstrip(";"))

def ler(xlsx):
    topo = pd.read_excel(xlsx, header=None, nrows=12)
    lin = next(i for i in range(len(topo)) if str(topo.iloc[i, 0]).strip() == "Contrato")
    df = pd.read_excel(xlsx, header=lin, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def processar(mov_xlsx, apl_xlsx, hub_de, hub_cid):
    m = ler(mov_xlsx)
    m["ser"] = serie(m["Cod.Compl."]); m = m[~m["ser"].isin(VAZIO)].copy()
    m["E"], m["S"] = num(m["Entrada"]), num(m["Saída"])
    m["tm"] = m["Tipo Movimento"].fillna("").str.strip()
    m["tp"] = m["Descrição Material"].str.extract(r"^\s*([A-Z]+)\s*-")[0].fillna("NOVO")
    m["dt"] = pd.to_datetime(m["Data Movimento"], format="%d/%m/%Y", errors="coerce")

    a = ler(apl_xlsx)
    a["ser"] = serie(a["N° Serie"]); a = a[~a["ser"].isin(VAZIO)].copy()
    a["dt"] = pd.to_datetime(a["Data Aplicação"], format="%d/%m/%Y", errors="coerce")
    inst = a[pd.to_numeric(a["Qtde. Aplic."], errors="coerce").fillna(0) > 0].groupby("ser")["dt"].max()
    retir = a[a["Material"].astype(str).str.startswith("RET-")].sort_values("dt").drop_duplicates("ser", keep="last")

    # eventos de conversão (um por documento CM e série): de -> para
    cm = m[m["tm"] == "CM"]
    sai = cm[cm["S"] > 0].groupby(["Nº. Doc", "ser"]).agg(de=("tp", "first"), dt=("dt", "min"), proj=("Projeto", "first"),
                                                          hora=("Hora Criação", "min"))
    ent = cm[cm["E"] > 0].groupby(["Nº. Doc", "ser"]).agg(para=("tp", "first"))
    ev = sai.join(ent, how="left").reset_index().sort_values(["dt", "hora", "Nº. Doc"])
    ev["para"] = ev["para"].fillna("OUTRO")

    rm = m[m["tm"] == "RM"].groupby("ser")["dt"].max()
    dm_all = m[m["tm"] == "DM"][["ser", "dt"]]

    # 1) todas as conversões de RET no mês
    cv = ev[ev["de"] == "RET"].drop_duplicates("ser", keep="last").copy()
    cv["hub"] = cv["proj"].map(hub_de)
    def st_reap(r):
        s, d = r["ser"], r["dt"]
        if s in inst and inst[s] >= d: return "instalado"
        dm_s = dm_all[dm_all["ser"] == s]["dt"]
        if s in rm and rm[s] >= d and not (len(dm_s) and dm_s.max() >= rm[s]): return "tecnico"
        return "almox"
    reap = cv[cv["para"] == "REAP"].copy(); reap["st"] = reap.apply(st_reap, axis=1)

    # 2) retirados do mês: situação por série
    rows = []
    for _, x in retir.iterrows():
        s, d = x["ser"], x["dt"]
        e = ev[(ev["ser"] == s) & (ev["dt"] >= d)]
        fin = e["para"].iloc[-1] if len(e) else None
        if fin in DESTINOS: k = fin
        elif fin is not None: k = "estorno"
        elif len(dm_all[(dm_all["ser"] == s) & (dm_all["dt"] >= d)]): k = "aguardando"
        else: k = "naodev"
        rows.append((hub_de(x["Projeto"]) or hub_cid(x["Município"]), k))
    rt = pd.DataFrame(rows, columns=["hub", "k"])
    return cv, reap, rt, m["dt"].min(), m["dt"].max()

def main(args):
    if len(args) < 2 or len(args) % 2:
        sys.exit("Uso: python atualizar_conversao.py MOV.xlsx APLIC.xlsx [MOV2.xlsx APLIC2.xlsx ...]")
    HUBS = carregar_js("cobertura.js")["hubs"]
    hubs = {fold(h["nome"]): h["id"] for h in HUBS}
    regiao = {h["id"]: ("ss" if h["uf"] in ("SP", "PR", "MG") else "ne") for h in HUBS}
    def hub_de(projeto):
        p = fold(projeto)
        p = re.sub(r"^(op|lab|fer|estrut) ", "", p)
        p = re.sub(r" (pb|rn|pe|ba|ce|sp|mg|pr|pa)$", "", p)
        return hubs.get(APELIDOS.get(p, p))
    cid = {}
    for r in carregar_js("cobertura.js")["rows"]:
        cid.setdefault(fold(r["c"]), set()).add(r["h"])
    def hub_cid(municipio):  # cidade atendida por um único almoxarifado
        hs = cid.get(fold(municipio), set())
        return next(iter(hs)) if len(hs) == 1 else None

    cvs, reaps, rts, d0s, d1s = [], [], [], [], []
    for i in range(0, len(args), 2):
        cv, reap, rt, d0, d1 = processar(args[i], args[i + 1], hub_de, hub_cid)
        reg = rt["hub"].dropna().map(regiao).mode()
        rt["reg"] = reg.iloc[0] if len(reg) else "ne"
        cvs.append(cv); reaps.append(reap); rts.append(rt); d0s.append(d0); d1s.append(d1)
        print(f"   {Path(args[i]).name}: {len(cv)} RET convertidos · {len(rt)} retirados rastreados ({rt['hub'].isna().sum()} sem almoxarifado)")
    cv, reap, rt = pd.concat(cvs), pd.concat(reaps), pd.concat(rts)
    sem = sorted(set(cv.loc[cv["hub"].isna(), "proj"].dropna()))

    vec = lambda g, col, keys: [int((g[col] == k).sum()) for k in keys]
    saida = {
        "meta": {
            "periodo": f"{MESES[max(d1s).month - 1]}/{max(d1s).year}",
            "de": min(d0s).strftime("%d/%m/%Y"), "ate": max(d1s).strftime("%d/%m/%Y"),
            "fonte": "SINAPSE · Movimentação de materiais (CM/RM/DM) + Aplicados/Retirados, por N° de Série",
            "geradoEm": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "destinos": DESTINOS, "status": STATUS, "retK": RET_K, "semAlmox": sem,
        },
        "hubs": {h: vec(g, "para", DESTINOS) for h, g in cv.dropna(subset=["hub"]).groupby("hub")},
        "status": {h: vec(g, "st", STATUS) for h, g in reap.dropna(subset=["hub"]).groupby("hub")},
        "ret": {h: vec(g, "k", RET_K) for h, g in rt.dropna(subset=["hub"]).groupby("hub")},
        "retSem": {g: vec(x[x["hub"].isna()], "k", RET_K) for g, x in rt.groupby("reg")},
    }
    (RAIZ / "conversao.js").write_text("/* gerado automaticamente — não editar à mão */\nwindow.DADOS_CONVERSAO = "
                                       + json.dumps(saida, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
    tot = lambda d, n: [sum(v[i] for v in d.values()) for i in range(n)]
    print("OK · conversões RET:", dict(zip(DESTINOS, tot(saida["hubs"], 4))))
    print("   status REAP:", dict(zip(STATUS, tot(saida["status"], 3))))
    print("   retirados do mês:", dict(zip(RET_K, tot(saida["ret"], 7))), "· sem almoxarifado:", saida["retSem"])
    if sem: print("   ATENÇÃO, projetos sem almoxarifado no mapa:", ", ".join(sem))

if __name__ == "__main__":
    main(sys.argv[1:])
