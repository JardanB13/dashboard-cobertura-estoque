"""
Gera conversao.js (painel "Destino dos retirados") rastreando cada equipamento pelo N° de Série.

Uso:
    python atualizar_conversao.py movimentacao.xlsx aplicados_retirados.xlsx

Entradas (SINAPSE, mesmo mês):
  1. Movimentação de materiais: colunas "Cod.Compl." (N° de Série), "Tipo Movimento"
     (RM = entrega ao técnico, DM = devolução, CM = conversão), "Entrada", "Saída".
  2. Materiais Aplicados/Retirados: colunas "N° Serie", "Material", "Qtde. Aplic.", "Data Aplicação".

Regras:
  - Só itens com N° de Série.
  - Conversão: no mesmo documento CM, saída do material "RET-" e entrada do destino
    (REAP-, SUC-, DET-, DEF-) com a mesma série.
  - Status de cada RET convertido em REAP, até o fim do período:
      instalado  = aparece aplicado no cliente depois da conversão;
      tecnico    = entregue a um técnico (RM) depois da conversão e não instalado nem devolvido;
      almox      = continua no almoxarifado.
  - Almoxarifado pela coluna "Projeto" (ex.: CAMPINA-GRANDE-PB -> Campina Grande).
  - Vão para o site apenas totais por almoxarifado (nenhuma série, cliente ou técnico).
"""
import json, re, sys, unicodedata
from datetime import datetime
from pathlib import Path
import pandas as pd

RAIZ = Path(__file__).resolve().parent
MESES = "jan fev mar abr mai jun jul ago set out nov dez".split()
DESTINOS = ["REAP", "SUC", "DET", "DEF"]
STATUS = ["instalado", "tecnico", "almox"]
APELIDOS = {"sj rio do peixe": "sjrp", "op inst vt picui": "picui"}

fold = lambda x: re.sub(r"[^a-z0-9]+", " ", unicodedata.normalize("NFD", str(x)).encode("ascii", "ignore").decode().lower()).strip()
num = lambda s: pd.to_numeric(s.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce").fillna(0)
serie = lambda s: s.astype(str).str.strip().str.upper()
VAZIO = {"-", "NAN", ""}

def carregar_js(nome):
    txt = (RAIZ / nome).read_text(encoding="utf-8")
    return json.loads(txt[txt.index("=") + 1:].strip().rstrip(";"))

def main(mov_xlsx, apl_xlsx):
    m = pd.read_excel(mov_xlsx, dtype=str)
    m.columns = [str(c).strip() for c in m.columns]
    m["ser"] = serie(m["Cod.Compl."]); m = m[~m["ser"].isin(VAZIO)].copy()
    m["E"], m["S"] = num(m["Entrada"]), num(m["Saída"])
    m["tm"] = m["Tipo Movimento"].fillna("").str.strip()
    m["tp"] = m["Descrição Material"].str.extract(r"^([A-Z]+)-")[0].fillna("NOVO")
    m["dt"] = pd.to_datetime(m["Data Movimento"], format="%d/%m/%Y", errors="coerce")

    a = pd.read_excel(apl_xlsx, header=2, usecols=["Material", "N° Serie", "Qtde. Aplic.", "Data Aplicação"], dtype=str)
    a["ser"] = serie(a["N° Serie"]); a = a[~a["ser"].isin(VAZIO)].copy()
    a = a[pd.to_numeric(a["Qtde. Aplic."], errors="coerce").fillna(0) > 0]
    a["dt"] = pd.to_datetime(a["Data Aplicação"], format="%d/%m/%Y", errors="coerce")
    inst = a.groupby("ser")["dt"].max()

    cm = m[m["tm"] == "CM"]
    saida = cm[(cm["tp"] == "RET") & (cm["S"] > 0)][["Nº. Doc", "ser", "dt", "Projeto"]]
    entrada = cm[cm["E"] > 0][["Nº. Doc", "ser", "tp"]].rename(columns={"tp": "dest"})
    cv = saida.merge(entrada, on=["Nº. Doc", "ser"], how="left").sort_values("dt").drop_duplicates("ser", keep="last")
    cv["dest"] = cv["dest"].fillna("OUTRO")

    rm = m[m["tm"] == "RM"].groupby("ser")["dt"].max()
    dm = m[m["tm"] == "DM"].groupby("ser")["dt"].max()
    def status(r):
        s, d = r["ser"], r["dt"]
        if s in inst and inst[s] >= d: return "instalado"
        if s in rm and rm[s] >= d and not (s in dm and dm[s] >= rm[s]): return "tecnico"
        return "almox"
    reap = cv[cv["dest"] == "REAP"].copy()
    reap["st"] = reap.apply(status, axis=1)

    hubs = {fold(h["nome"]): h["id"] for h in carregar_js("cobertura.js")["hubs"]}
    def hub_de(projeto):
        p = fold(re.sub(r"-(PB|RN|PE|BA|CE|SP|MG|PR|PA)$", "", str(projeto)))
        return hubs.get(APELIDOS.get(p, p))
    cv["hub"] = cv["Projeto"].map(hub_de); reap["hub"] = reap["Projeto"].map(hub_de)
    sem = sorted(cv.loc[cv["hub"].isna(), "Projeto"].dropna().unique())

    destinos = DESTINOS + sorted(set(cv["dest"]) - set(DESTINOS))
    por_hub = {h: [int((g["dest"] == d).sum()) for d in destinos] for h, g in cv.dropna(subset=["hub"]).groupby("hub")}
    st_hub = {h: [int((g["st"] == s).sum()) for s in STATUS] for h, g in reap.dropna(subset=["hub"]).groupby("hub")}

    d0, d1 = cv["dt"].min(), cv["dt"].max()
    saida_js = {
        "meta": {
            "periodo": f"{MESES[d1.month - 1]}/{d1.year}",
            "de": d0.strftime("%d/%m/%Y"), "ate": d1.strftime("%d/%m/%Y"),
            "fonte": "SINAPSE · Movimentação de materiais (CM/RM/DM) + Aplicados/Retirados, por N° de Série",
            "geradoEm": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "destinos": destinos, "status": STATUS, "semAlmox": sem,
        },
        "hubs": por_hub,
        "status": st_hub,
    }
    (RAIZ / "conversao.js").write_text("/* gerado automaticamente — não editar à mão */\nwindow.DADOS_CONVERSAO = "
                                       + json.dumps(saida_js, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
    tot = [sum(v[i] for v in por_hub.values()) for i in range(len(destinos))]
    st = [sum(v[i] for v in st_hub.values()) for i in range(len(STATUS))]
    print(f"OK · {saida_js['meta']['de']} a {saida_js['meta']['ate']}: {sum(tot)} RET convertidos -> "
          + ", ".join(f"{d} {q}" for d, q in zip(destinos, tot)))
    print("   status dos REAP: " + ", ".join(f"{s} {q}" for s, q in zip(STATUS, st)))
    if sem: print("   ATENÇÃO, projetos sem almoxarifado no mapa:", ", ".join(sem))

if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("Uso: python atualizar_conversao.py movimentacao.xlsx aplicados_retirados.xlsx")
    main(sys.argv[1], sys.argv[2])
