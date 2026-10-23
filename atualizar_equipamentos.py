"""
Atualiza equipamentos.js a partir de um ou mais relatórios SINAPSE
(Relatório de Materiais Aplicados/Retirados), do mesmo mês.

Uso:
    python atualizar_equipamentos.py planilha1.xlsx [planilha2.xlsx ...] [--categorias CATEGORIA_TECNOLOGICA.xlsx]

Exemplo (Nordeste + Sul/Sudeste):
    python atualizar_equipamentos.py aplicados_nordeste.xlsx aplicados_sul.xlsx

Regras:
  - Só projetos de operação (coluna Projeto começando com "OP-"); ESTRUT, MANUT REDES etc. ficam fora.
  - Só linhas com N° de Série (equipamentos serializados); conectores, cabos etc. ficam fora.
  - Prefixo do material: "REAP-" = Reaproveitado, "RET-" = Retirado, sem prefixo = Novo.
  - Estado (UF) da cidade:
      1. pelo almoxarifado da coluna "Projeto" (ex.: OP-OSASCO -> Osasco/SP), procurando a cidade
         entre as atendidas por ele no cobertura.js;
      2. senão, pelo contrato: final 1 = PB, 2 = RN, 3 = PE, 4 = BA (operação Nordeste).
  - O cabeçalho da planilha é localizado automaticamente (linha que começa com "Contrato").
  - Só entram dados agregados por cidade (nenhum dado de cliente vai para o site).
  - Aparelho de categoria OBSOLETO aplicado não conta como novo nem reaproveitado (campo "o").
  - Categoria tecnológica (opcional, --categorias): pelo nome do aparelho na planilha de referência
    (coluna "Equipamentos" -> "Categoria Tecnologia"); modelos que não estão nela são classificados
    por palavras-chave do nome (ver categoria_por_nome).
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

PREFIXO = r"^((REAP|RET|REA|DEF|DET|SUC|PO)\s*-\s*)+"
base_nome = lambda x: re.sub(PREFIXO, "", re.sub(r"\s+", " ", str(x).strip().upper())).strip()

ANTIGAS = {"ONT 5G": "ONT WIFI 5", "ROTEADOR 5G": "ROTEADOR WIFI 5", "ROTEADOR 5G MESH": "ROTEADOR WIFI 5", "MISCELÂNEA": "MISCELANEA"}

def ler_categorias(xlsx):
    """Planilha DE PARA: codigo_material | descricao_material | categoria_tecnologia [| base_empresa].
    Procura primeiro pelo código do material e depois pelo nome (sem prefixo REAP-/RET-/DEF-...)."""
    ref = pd.read_excel(xlsx, dtype=str)
    ref = ref.iloc[:, :3]; ref.columns = ["cod", "nome", "cat"]
    ref = ref[~ref["cat"].astype(str).str.strip().isin(["", "nan"])]
    norm = lambda c: ANTIGAS.get(str(c).strip().upper(), str(c).strip().upper())
    out = {"cod:" + str(k).strip().upper(): norm(c) for k, c in zip(ref["cod"], ref["cat"])}
    out.update({base_nome(n): norm(c) for n, c in zip(ref["nome"], ref["cat"])})
    return out

WIFI6 = r"WIFI ?6|WI-FI ?6|\bW6\b|\bAX ?\d{3,4}|AX3000|AX1800|X6\b|X6-|F6600|G1426|G-2426|VANTIVA|HG6145F|HG6145 - AX|XC220|EX 141"
AC = r"\bAC\b|\bAC ?\d{3,4}|AC\d{3,4}|1200|DUAL BAND|8145|8245|HS8546|G ?-? ?140 ?W|G-240W-G|G-?1425|G-0425|F67\d|HG6245|HG6143|HG6145D|121 ?AC|GW24|V2801SGA|DM985-424|DM986-416|WS5200"
PON = r"\bONU\b|\bONT\b|ONU-|GPON|XPON|EPON|ECHOLIFE|FIBERHOME|NOKIA|HG[68]\d|EG8\d|HS85|AN5506|G-140|G-240|G-1425|G-0425|VANTIVA|XZ000|F6\d\d"
RADIO = r"^RADIO|^ANTENA|\bLHG\b|LITEBEAM|\bSXT\b|AIRGRID|POWER ?BEAM|NANO ?STATION|WOM AC|5GHZ.*DBI"

def categoria_por_nome(nome, ref, cod=None):
    if cod is not None and "cod:" + str(cod).strip().upper() in ref: return ref["cod:" + str(cod).strip().upper()]
    b = base_nome(nome)
    if b in ref: return ref[b]
    if re.search(RADIO, b) and not re.search(r"\bONU\b", b): return "RADIO"
    if re.search(PON, b):
        if re.search(WIFI6, b): return "ONT WIFI 6"
        if re.search(AC, b): return "ONT WIFI 5"
        return "ONU"
    if re.search(r"ROTEADOR|^ROT |WS7\d|WS5\d|DIR-|ARCHER|MERCUSYS|GREATEK|AC12G|MW301|EX 141", b):
        if re.search(r"WS7\d", b): return "ROTEADOR WIFI 6"    # mesma regra da planilha DE PARA (WS7001-AX2)
        if re.search(WIFI6, b): return "ROTEADOR WIFI 6"
        if re.search(AC + r"|MESH|\bC[2-8]0?\b|H19\dA|ZT199|EMG3524|HUAWEI$", b): return "ROTEADOR WIFI 5"
    return "SEM INFORMACAO"

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

def main(arquivos, cat_xlsx=None):
    ref_cat = ler_categorias(cat_xlsx) if cat_xlsx else None
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
        df = df[df["Projeto"].astype(str).str.upper().str.startswith("OP-")].copy()   # só operações
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
    if ref_cat is not None:
        eq["Cat"] = [categoria_por_nome(m, ref_cat, c) for m, c in zip(eq["Material"], eq["Cód. Material"])]
        # obsoleto aplicado não conta como novo nem reaproveitado: vira status próprio "OBS"
        eq.loc[(eq["Cat"] == "OBSOLETO") & eq["T"].isin(["NOVO", "REAP"]), "T"] = "OBS"

    cid = {k: {"n": 0, "r": 0, "t": 0, "o": 0, "os": int(g["OS"].nunique()), "nat": {}} for k, g in df.groupby("k")}
    idx = {"NOVO": 0, "REAP": 1, "RET": 2}
    for (k, nat, t), q in eq.groupby(["k", "Natureza", "T"])["q"].sum().items():
        c = cid[k]
        if t == "OBS": c["o"] += int(q); continue
        i = idx[t]
        c["nrt"[i]] += int(q)
        c["nat"].setdefault(nat, [0, 0, 0])[i] += int(q)
    if ref_cat is not None:
        for (k, ct, t), q in eq[eq["T"] != "OBS"].groupby(["k", "Cat", "T"])["q"].sum().items():
            cid[k].setdefault("cat", {}).setdefault(ct, [0, 0, 0])[idx[t]] += int(q)

    mapa = {f"{fold(r['c'])}|{r['uf']}" for r in cob["rows"]}
    for v in cid.values():
        if not v["o"]: v.pop("o")      # só grava "o" (obsoletos aplicados) onde houver
    no_mapa = {k: v for k, v in cid.items() if k in mapa}
    fora = [v for k, v in cid.items() if k not in mapa and (v["n"] or v["r"] or v["t"] or v.get("o"))]
    N_UF = ("AC", "AM", "AP", "PA", "RO", "RR", "TO")
    reg = lambda k: "ss" if k.split("|")[-1] in ("SP", "PR", "MG") else "no" if k.split("|")[-1] in N_UF else "ne"
    soma = lambda vs: {"cid": len(vs), "n": sum(v["n"] for v in vs), "r": sum(v["r"] for v in vs), "t": sum(v["t"] for v in vs), "o": sum(v.get("o", 0) for v in vs)}
    fora_reg = {g: soma([v for k, v in cid.items() if k not in mapa and (v["n"] or v["r"] or v["t"] or v.get("o")) and reg(k) == g]) for g in ("no", "ne", "ss")}
    def soma_cat(vs):
        out = {}
        for v in vs:
            for ct, a in v.get("cat", {}).items():
                o = out.setdefault(ct, [0, 0, 0]); o[0] += a[0]; o[1] += a[1]; o[2] += a[2]
        return out
    tot = lambda tipo: int(eq.loc[eq["T"] == tipo, "q"].sum())
    saida = {
        "meta": {
            "periodo": periodo,
            "fonte": f"SINAPSE · Materiais Aplicados/Retirados ({periodos[0].split(':', 1)[-1].strip()}) · {len(arquivos)} relatório(s)",
            "tot": {"n": tot("NOVO"), "r": tot("REAP"), "t": tot("RET"), "o": tot("OBS")},
            "fora": soma(fora),
            "foraReg": fora_reg,
            **({"foraCat": {g: soma_cat([v for k, v in cid.items() if k not in mapa and reg(k) == g]) for g in ("no", "ne", "ss")},
                "categorias": sorted(eq["Cat"].unique().tolist(), key=lambda c: (c in ("SEM INFORMACAO", "OBSOLETO"), c))} if ref_cat is not None else {}),
        },
        "c": no_mapa,
    }
    destino = RAIZ / "equipamentos.js"
    destino.write_text("/* gerado automaticamente — não editar à mão */\nwindow.DADOS_EQUIPAMENTOS = "
                       + json.dumps(saida, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
    m = saida["meta"]
    print(f"OK · {periodo}: {m['tot']['n']} novos, {m['tot']['r']} reaproveitados, {m['tot']['t']} retirados, {m['tot']['o']} obsoletos aplicados")
    print(f"   {len(no_mapa)} cidades no mapa · {m['fora']['cid']} cidades sem vínculo com almoxarifado (fora do mapa)")
    if ref_cat is not None:
        print("   por categoria (novos, reaprov., retirados):")
        tab = eq.groupby(["Cat", "T"])["q"].sum().unstack(fill_value=0)
        for ct, r in tab.iterrows(): print(f"     {ct:<18} {int(r.get('NOVO', 0)):>6} {int(r.get('REAP', 0)):>6} {int(r.get('RET', 0)):>6}")
        outros = eq.loc[eq["Cat"] == "SEM INFORMACAO", "Material"].map(base_nome).value_counts()
        if len(outros): print("   SEM INFORMACAO (assim na planilha, ou sem regra):", ", ".join(f"{n} ({q})" for n, q in outros.items()))
    print(f"   gravado em {destino.name}")

if __name__ == "__main__":
    args, cat = sys.argv[1:], None
    if "--categorias" in args:
        i = args.index("--categorias"); cat = args[i + 1]; args = args[:i] + args[i + 2:]
    if not args:
        sys.exit("Uso: python atualizar_equipamentos.py planilha1.xlsx [planilha2.xlsx ...] [--categorias CATEGORIAS.xlsx]")
    main(args, cat)
