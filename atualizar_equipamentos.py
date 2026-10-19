"""
Atualiza equipamentos.js a partir do relatório SINAPSE
(Relatório de Materiais Aplicados/Retirados).

Uso:
    python atualizar_equipamentos.py caminho/da/planilha.xlsx

Regras:
  - Só linhas com N° de Série (equipamentos).
  - Prefixo do material: "REAP-" = Reaproveitado, "RET-" = Retirado, sem prefixo = Novo.
  - Região pelo contrato: final 1 = PB, 2 = RN, 3 = PE, 4 = BA.
  - Só entram dados agregados por cidade (nenhum dado de cliente vai para o site).
"""
import json, re, sys, unicodedata
from pathlib import Path
import pandas as pd

RAIZ = Path(__file__).resolve().parent
UF_CONTRATO = {"1": "PB", "2": "RN", "3": "PE", "4": "BA"}
MESES = "jan fev mar abr mai jun jul ago set out nov dez".split()

fold = lambda x: unicodedata.normalize("NFD", str(x)).encode("ascii", "ignore").decode().lower().strip()

def carregar_js(nome):
    txt = (RAIZ / nome).read_text(encoding="utf-8")
    return json.loads(txt[txt.index("=") + 1:].strip().rstrip(";"))

def main(xlsx):
    bruto = pd.read_excel(xlsx, header=None, nrows=2)
    periodo_txt = str(bruto.iloc[1, 0])  # "Período da Aplicação: 01/09/2026 à 30/09/2026"
    datas = re.findall(r"(\d{2})/(\d{2})/(\d{4})", periodo_txt)
    periodo = f"{MESES[int(datas[0][1]) - 1]}/{datas[0][2]}" if datas else "período"

    df = pd.read_excel(xlsx, header=2)
    df.columns = [str(c).strip() for c in df.columns]
    df["UF"] = df["Contrato"].astype(str).str[-1].map(UF_CONTRATO)
    df["Natureza"] = df["Natureza"].fillna("SEM NATUREZA")
    df["k"] = df["Município"].map(fold) + "|" + df["UF"]
    for c in ["Qtde. Aplic.", "Qtde. Remov."]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    eq = df[df["N° Serie"].astype(str).str.strip() != "-"].copy()
    eq["T"] = eq["Material"].astype(str).str.extract(r"^(REAP|RET)-")[0].fillna("NOVO")
    eq["q"] = eq["Qtde. Aplic."].where(eq["T"] != "RET", eq["Qtde. Remov."])

    cid = {k: {"n": 0, "r": 0, "t": 0, "os": int(g["OS"].nunique()), "nat": {}} for k, g in df.groupby("k")}
    idx = {"NOVO": 0, "REAP": 1, "RET": 2}
    for (k, nat, t), q in eq.groupby(["k", "Natureza", "T"])["q"].sum().items():
        c = cid[k]; i = idx[t]
        c["nrt"[i]] += int(q)
        c["nat"].setdefault(nat, [0, 0, 0])[i] += int(q)

    mapa = {f"{fold(r['c'])}|{r['uf']}" for r in carregar_js("cobertura.js")["rows"]}
    no_mapa = {k: v for k, v in cid.items() if k in mapa}
    fora = [v for k, v in cid.items() if k not in mapa]
    tot = lambda tipo: int(eq.loc[eq["T"] == tipo, "q"].sum())
    saida = {
        "meta": {
            "periodo": periodo,
            "fonte": f"SINAPSE · Materiais Aplicados/Retirados ({periodo_txt.split(':', 1)[-1].strip()})",
            "tot": {"n": tot("NOVO"), "r": tot("REAP"), "t": tot("RET")},
            "fora": {"cid": len(fora), "n": sum(v["n"] for v in fora), "r": sum(v["r"] for v in fora), "t": sum(v["t"] for v in fora)},
        },
        "c": no_mapa,
    }
    destino = RAIZ / "equipamentos.js"
    destino.write_text("/* gerado automaticamente — não editar à mão */\nwindow.DADOS_EQUIPAMENTOS = "
                       + json.dumps(saida, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
    m = saida["meta"]
    print(f"OK · {periodo}: {m['tot']['n']} novos, {m['tot']['r']} reaproveitados, {m['tot']['t']} retirados")
    print(f"   {len(no_mapa)} cidades no mapa · {m['fora']['cid']} cidades sem vínculo com almoxarifado (fora do mapa)")
    print(f"   gravado em {destino.relative_to(RAIZ)}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Uso: python atualizar_equipamentos.py caminho/da/planilha.xlsx")
    main(sys.argv[1])
