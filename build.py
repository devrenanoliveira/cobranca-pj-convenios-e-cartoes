"""Lê os 2 Excel da pasta de dados mais recente e gera dashboard.html (autônomo).

Uso:  python build.py [pasta_com_os_xlsx] [--publico]
  sem flag -> dashboard.html (nominal, só local); --publico -> index.html (sem identificação, vai pro Pages)
Sem argumento, usa a pasta do RELATORIO INADIMPLENCIA mais recente em Dados/.
"""
import glob, json, os, re, sys
from datetime import date, timedelta
import pandas as pd

AQUI = os.path.dirname(os.path.abspath(__file__))
PRODUTOS = [  # (código, nome) — ordem de exibição
    ("567", "Cartão Corporativo"), ("341", "Cartão Mercado"), ("936", "Empréstimo PJ"),
    ("76", "CDC Funcionário"), ("239", "Venda faturada Condor"), ("109", "Cartão presente"),
    ("162", "Venda faturada Postos"), ("163", "Cartão Frota"),
    ("161", "Agenda 161 (sem descrição)"), ("CHQ", "Cheques pré-datados"),
]
COD = {c for c, _ in PRODUTOS}


def acha_pasta():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if args:
        return args[0]
    rel = glob.glob(os.path.join(AQUI, "Dados", "**", "RELATORIO*.xlsx"), recursive=True)
    if not rel:
        sys.exit("Nenhum RELATORIO*.xlsx em Dados/")
    return os.path.dirname(max(rel, key=os.path.getmtime))


class Pool:  # dicionário de strings repetidas (loja, devedor) -> índice
    def __init__(self): self.l, self.m = [], {}
    def __call__(self, s):
        s = "" if pd.isna(s) else str(s).strip()
        if s not in self.m: self.m[s] = len(self.l); self.l.append(s)
        return self.m[s]


def iso(x): return "" if pd.isna(x) else pd.Timestamp(x).strftime("%Y-%m-%d")


def ler_agendas(pasta, lojas, devs):
    arq = glob.glob(os.path.join(pasta, "RELATORIO*.xlsx"))[0]
    x = pd.ExcelFile(arq)
    rows, ref_votos = [], pd.Series(dtype=int)
    for s in x.sheet_names:
        d = x.parse(s)
        if "VALOR NOTA" not in d:  # ex.: "AGENDA 76- A VENCER" só tem observação
            print(f"  aba sem dados: {s.strip()}"); continue
        d.columns = [re.sub(r"[^A-Z0-9. ]", "?", c) for c in d.columns]
        em = next(c for c in d.columns if c.startswith("DATA EMISS"))
        st = "V" if "VENCIDOS" in s else "A"
        ag = str(d["AGENDA"].iloc[0])
        assert ag in COD, f"agenda nova sem tradução: {ag}"
        if st == "V":
            ref_votos = pd.concat([ref_votos, (d["DATA VENCIMENTO"] + pd.to_timedelta(d["ATRASO"], "D")).value_counts()])
        for r in d.itertuples(index=False):
            g = dict(zip(d.columns, r))
            nome = re.sub(r"^\d+-", "", str(g["FORNECEDOR"]))
            rows.append([ag, st, int(g["ATRASO"]), iso(g["DATA VENCIMENTO"]), round(float(g["VALOR NOTA"]), 2),
                         devs(nome), str(g["CNPJ"]), lojas(g["LOJA"]), str(g["NOTA FISCAL"]), int(g["PARCELA"]),
                         iso(g[em]), ""])
    ref = ref_votos.groupby(level=0).sum().idxmax()
    return rows, ref.date()


def ler_cheques(pasta, ref, lojas, devs):
    x = pd.ExcelFile(glob.glob(os.path.join(pasta, "CARTEIRA*.xlsx"))[0])
    resumo = x.parse("RESUMO PORTADOR").set_index("PORTADOR")
    rows = []
    for s in x.sheet_names:
        m = re.match(r"Portador (\d+)", s)
        if not m: continue
        port = m.group(1)
        d = x.parse(s)
        d.columns = d.columns.str.strip()
        d["V"] = pd.to_datetime(d["VENCTO"], errors="coerce")
        d["VAL"] = pd.to_numeric(d["VLR.CHEQUE"], errors="coerce")
        d = d[d.V.notna() & d.VAL.notna() & d.CHEQUE.notna()]  # tira total, traços e lixo de impressão
        ok = (len(d), round(d.VAL.sum(), 2)) == (resumo.loc[int(port), "QUANTIDADEW"], round(resumo.loc[int(port), "TOTAL"], 2))
        print(f"  portador {port}: {len(d)} cheques, R$ {d.VAL.sum():,.2f} — {'confere com RESUMO' if ok else 'DIVERGE do RESUMO!'}")
        for r in d.itertuples():
            loja = "" if pd.isna(r.LOJA) else f"Loja {int(float(r.LOJA))}"
            rows.append(["CHQ", "V", (ref - r.V.date()).days, iso(r.V), round(float(r.VAL), 2), devs(r.NOME),
                         str(r.CODIGO).replace(".0", ""), lojas(loja), str(r.CHEQUE).replace(".0", "") if isinstance(r.CHEQUE, float) else str(r.CHEQUE), 0, iso(r.EMISSAO) if isinstance(r.EMISSAO, pd.Timestamp) else "", port])
    return rows


if __name__ == "__main__":
    pasta = acha_pasta()
    print("Pasta:", pasta)
    lojas, devs = Pool(), Pool()
    rows, ref = ler_agendas(pasta, lojas, devs)
    print("Data de referência (vencido = atraso desde):", ref)
    rows += ler_cheques(pasta, ref, lojas, devs)
    # vencimento <= ref é vencido; qualquer linha "A vencer" com atraso>0 seria inconsistência
    assert all(r[2] <= 0 for r in rows if r[1] == "A"), "linha a vencer com atraso"
    pub = "--publico" in sys.argv
    if pub:  # tira tudo que identifica devedor: nome, CNPJ/CPF, NF/cheque, emissão
        rows = [r[:5] + [0, "", r[7], "", r[9], "", r[11]] for r in rows]
        devs.l = ["(oculto)"]
    data = {"pub": pub, "ref": ref.isoformat(), "pasta": os.path.basename(pasta), "prods": PRODUTOS,
            "lojas": lojas.l, "devs": devs.l, "rows": rows}
    tpl = open(os.path.join(AQUI, "template.html"), encoding="utf-8").read()
    out = tpl.replace("/*DATA*/null", json.dumps(data, ensure_ascii=False, separators=(",", ":")))
    nome = "index.html" if pub else "dashboard.html"
    open(os.path.join(AQUI, nome), "w", encoding="utf-8").write(out)
    print(f"{nome}: {len(rows)} linhas, {len(out)/1e6:.1f} MB")
