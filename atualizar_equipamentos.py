"""
Atualiza equipamentos.js a partir de um ou mais relatórios SINAPSE
(Relatório de Materiais Aplicados/Retirados), do mesmo mês.

Uso:
    python atualizar_equipamentos.py planilha1.xlsx [planilha2.xlsx ...]

Exemplo (Nordeste completo + Sul/Sudeste só retirados):
    python atualizar_equipamentos.py aplicados_nordeste.xlsx retirados_sul.xlsx

Regras:
  - Só linhas com N° de Série (equipamentos serializados); conectores, cabos etc. ficam fora.
  - Prefixo do material: "REAP-" = Reaproveitado, "RET-" = Retirado, sem prefixo = Novo.
  - Estado (UF) da cidade:
      1. pelo almoxarifado da coluna "Projeto" (ex.: OP-OSASCO -> Osasco/SP), procurando a cidade
         entre as atendidas por ele no cobertura.js;
      2. senão, pelo contrato: final 1 = PB, 2 = RN, 3 = PE, 4 = BA (operação Nordeste).
  - O cabeçalho da planilha é localizado automaticamente (linha que começa com "Contrato").
  - Só entram dados agregados por cidade (nenhum dado de cliente vai para o site).
"""
import json, re, sys, unicodedata
from pathlib import Path
import pandas as pd

RAIZ = Path(__file__).resolve().parent
UF_CONTRATO = {"1": "PB", "2": "RN", "3": "PE", "4": "BA"}
MESES = "jan fev mar abr mai jun jul ago set out nov dez".split()
# Projetos cujo nome não bate direto com o almoxarifado do cobertura.js
APELIDO_PROJETO = {"sto ant platina": "sap", "ourofino": "ouro fino"}
UF_PROJETO = {"santa rita": "MG"}  # projeto sem almoxarifado no mapa: UF de referência

fold = lambda x: unicodedata.normalize("NFD", str(x)).encode("ascii", "ignore").decode().lower().strip()
chave = lambda x: re.sub(r"[^a-z0-9]+", " ", fold(x)).strip()

def carregar_js(nome):
    txt = (RAIZ / nome).read_text(encoding="utf-8")
    return json.loads(txt[txt.index("=") + 1:].strip().rstrip(";"))

def ler(xlsx):
    topo = pd.read_excel(xlsx, header=None, nrows=12)
    lin = next(i for i in range(len(topo)) if str(topo.iloc[i, 0]).strip() == "Contrato")
    periodo_txt = next((str(v) for v in topo.iloc[:lin, 0] if "Período" in str(v)), "")
    df = pd.read_excel(xlsx, header=lin, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df, periodo_txt

def main(arquivos):
    cob = carregar_js("cobertura.js")
    hubs = {chave(h["nome"]): h for h in cob["hubs"]}
    cid_hub = {}                                   # (hub_id, cidade) -> UF
    cid_uf = {}                                    # cidade -> {UFs}
    for r in cob["rows"]:
        cid_hub[(r["h"], fold(r["c"]))] = r["uf"]
        cid_uf.setdefault(fold(r["c"]), set()).add(r["uf"])

    def uf_de(row):
        proj = chave(re.sub(r"^OP[- ]", "", str(row["Projeto"]), flags=re.I))
        proj = re.sub(r" (pb|rn|pe|ba|sp|pr|mg|ce|pa)$", "", proj)
        h = hubs.get(APELIDO_PROJETO.get(proj, proj))
        c = fold(row["Município"])
        if h:
            if (h["id"], c) in cid_hub: return cid_hub[(h["id"], c)]
            ufs = cid_uf.get(c, set())
            return next(iter(ufs)) if len(ufs) == 1 else h["uf"]
        if proj in UF_PROJETO:
            ufs = cid_uf.get(c, set())
            return next(iter(ufs)) if len(ufs) == 1 else UF_PROJETO[proj]
        return UF_CONTRATO.get(str(row["Contrato"])[-1])

    partes, periodos = [], []
    for xlsx in arquivos:
        df, ptxt = ler(xlsx)
        periodos.append(ptxt)
        df["UF"] = df.apply(uf_de, axis=1)
        df["Natureza"] = df["Natureza"].fillna("SEM NATUREZA")
        for c in ["Qtde. Aplic.", "Qtde. Remov."]:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        df["k"] = df["Município"].map(fold) + "|" + df["UF"].fillna("??")
        partes.append(df)
        n_ser = (df["N° Serie"].astype(str).str.strip() != "-").sum()
        print(f"   {Path(xlsx).name}: {len(df)} linhas, {n_ser} serializadas")
    df = pd.concat(partes, ignore_index=True)

    datas = re.findall(r"(\d{2})/(\d{2})/(\d{4})", periodos[0])
    periodo = f"{MESES[int(datas[0][1]) - 1]}/{datas[0][2]}" if datas else "período"

    eq = df[~df["N° Serie"].astype(str).str.strip().isin(["-", "", "nan"])].copy()
    eq["T"] = eq["Material"].astype(str).str.extract(r"^(REAP|RET)-")[0].fillna("NOVO")
    eq["q"] = eq["Qtde. Aplic."].where(eq["T"] != "RET", eq["Qtde. Remov."])

    cid = {k: {"n": 0, "r": 0, "t": 0, "os": int(g["OS"].nunique()), "nat": {}} for k, g in df.groupby("k")}
    idx = {"NOVO": 0, "REAP": 1, "RET": 2}
    for (k, nat, t), q in eq.groupby(["k", "Natureza", "T"])["q"].sum().items():
        c = cid[k]; i = idx[t]
        c["nrt"[i]] += int(q)
        c["nat"].setdefault(nat, [0, 0, 0])[i] += int(q)

    mapa = {f"{fold(r['c'])}|{r['uf']}" for r in cob["rows"]}
    no_mapa = {k: v for k, v in cid.items() if k in mapa}
    fora = [v for k, v in cid.items() if k not in mapa and (v["n"] or v["r"] or v["t"])]
    reg = lambda k: "ss" if k.split("|")[-1] in ("SP", "PR", "MG") else "ne"
    soma = lambda vs: {"cid": len(vs), "n": sum(v["n"] for v in vs), "r": sum(v["r"] for v in vs), "t": sum(v["t"] for v in vs)}
    fora_reg = {g: soma([v for k, v in cid.items() if k not in mapa and (v["n"] or v["r"] or v["t"]) and reg(k) == g]) for g in ("ne", "ss")}
    tot = lambda tipo: int(eq.loc[eq["T"] == tipo, "q"].sum())
    saida = {
        "meta": {
            "periodo": periodo,
            "fonte": f"SINAPSE · Materiais Aplicados/Retirados ({periodos[0].split(':', 1)[-1].strip()}) · {len(arquivos)} relatório(s)",
            "tot": {"n": tot("NOVO"), "r": tot("REAP"), "t": tot("RET")},
            "fora": {"cid": len(fora), "n": sum(v["n"] for v in fora), "r": sum(v["r"] for v in fora), "t": sum(v["t"] for v in fora)},
            "foraReg": fora_reg,
        },
        "c": no_mapa,
    }
    destino = RAIZ / "equipamentos.js"
    destino.write_text("/* gerado automaticamente — não editar à mão */\nwindow.DADOS_EQUIPAMENTOS = "
                       + json.dumps(saida, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
    m = saida["meta"]
    print(f"OK · {periodo}: {m['tot']['n']} novos, {m['tot']['r']} reaproveitados, {m['tot']['t']} retirados")
    print(f"   {len(no_mapa)} cidades no mapa · {m['fora']['cid']} cidades sem vínculo com almoxarifado (fora do mapa)")
    print(f"   gravado em {destino.name}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("Uso: python atualizar_equipamentos.py planilha1.xlsx [planilha2.xlsx ...]")
    main(sys.argv[1:])
