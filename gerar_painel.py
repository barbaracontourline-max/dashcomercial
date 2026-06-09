# -*- coding: utf-8 -*-
"""Gerador do Painel Comercial Contourline.
Baixa a planilha (Google Sheets) e reconstroi o painel HTML com os numeros atuais.
Rodar:  python -X utf8 gerar_painel.py
"""
import urllib.request, io, os, datetime, openpyxl

FILE_ID = os.environ.get("SHEET_ID", "").strip()  # vem do segredo do GitHub (nao fica no codigo)
if not FILE_ID:
    raise SystemExit("Defina o segredo SHEET_ID em Settings > Secrets and variables > Actions.")
URL = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"
MONTH = os.environ.get("MONTH", "MAIO").upper()
MONTH_LABEL = {"MAIO":"Maio","JUNHO":"Junho"}.get(MONTH, MONTH.title())
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "painel-comercial-prototipo.html")

def br(x):
    try: return "R$ " + f"{float(x):,.2f}".replace(",", "#").replace(".", ",").replace("#", ".")
    except Exception: return "R$ 0,00"
def pctf(x, d=0):
    return f"{x:.{d}f}".replace(".", ",") + "%"
def num(v): return isinstance(v, (int, float))
def gv(v): return float(v) if num(v) else 0.0
def _isdeal(ws, r):
    """Linha real de negocio: VALOR DE VENDA numerico + CLIENTE preenchido.
    Robusto a erros de formula na coluna # (ex: #VALUE! que antes derrubava a contagem)."""
    q = ws.cell(r, 17).value
    b = ws.cell(r, 2).value
    return isinstance(q, (int, float)) and isinstance(b, str) and b.strip() != ""

# ---------- baixar dados ----------
import time
data = None
for _try in range(3):
    try:
        data = urllib.request.urlopen(URL, timeout=90).read(); break
    except Exception as _e:
        print(f"  tentativa {_try+1} falhou ({_e}); repetindo...")
        time.sleep(3)
if data is None:
    raise SystemExit("ERRO: nao consegui baixar a planilha (rede). Tente de novo.")
wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
R = wb["RESUMÃO"]
def c(col, row):
    return R[f"{col}{row}"].value

# ---------- EQUIP / LINHAS ----------
consum_qtd = int(gv(c("D", 4)))            # consumiveis vendidos (so existem em BH Nao Med)
consum_val_vend = gv(c("B", 4)); consum_val_ab = gv(c("B", 6)); consum_ab_qtd = int(gv(c("D", 6)))
equip_nm_val_vend = gv(c("B", 5)); equip_nm_qtd = int(gv(c("D", 5)))
equip_nm_val_ab = gv(c("B", 7)); equip_nm_ab_qtd = int(gv(c("D", 7)))

lines = [
    {"nome": "Body Health — Não Médico", "meta": gv(c("B",3)), "vend": gv(c("B",10)),
     "vqtd": equip_nm_qtd + consum_qtd, "ab": consum_val_ab + equip_nm_val_ab,
     "abqtd": consum_ab_qtd + equip_nm_ab_qtd, "full": True,
     "bd": (equip_nm_qtd, consum_qtd),
     "chips": [("Equipamentos (não consumíveis)", equip_nm_val_vend, equip_nm_qtd, equip_nm_val_ab, equip_nm_ab_qtd),
               ("Consumíveis", consum_val_vend, consum_qtd, consum_val_ab, consum_ab_qtd)]},
    {"nome": "Body Health — Médico", "meta": gv(c("B",13)), "vend": gv(c("B",14)),
     "vqtd": int(gv(c("D",14))), "ab": gv(c("B",15)), "abqtd": int(gv(c("D",15)))},
    {"nome": "Lumenis", "meta": gv(c("B",21)), "vend": gv(c("B",22)),
     "vqtd": int(gv(c("D",22))), "ab": gv(c("B",23)), "abqtd": int(gv(c("D",23)))},
    {"nome": "Cynosure Lutronic", "meta": gv(c("B",29)), "vend": gv(c("B",30)),
     "vqtd": int(gv(c("D",30))), "ab": gv(c("B",31)), "abqtd": int(gv(c("D",31)))},
    {"nome": "Visbody", "meta": gv(c("B",37)), "vend": gv(c("B",38)),
     "vqtd": int(gv(c("D",38))), "ab": gv(c("B",39)), "abqtd": int(gv(c("D",39)))},
]

# ---------- GESTORES ----------
def gest(base, nome, img, bd=None):
    return {"nome": nome, "img": img, "meta": gv(c("G",base)), "vend": gv(c("G",base+1)),
            "vqtd": int(gv(c("I",base+1))), "ab": gv(c("G",base+2)), "abqtd": int(gv(c("I",base+2))), "bd": bd}
jess_qtd = int(gv(c("I", 4)))
gestores = [
    gest(19, "Xavier Ramos", "xavier_av.jpg"),
    gest(11, "Nikole Cestaro", "nikole_av.jpg"),
    gest(3,  "Jéssica Oliveira", "jessica_av.jpg", bd=(jess_qtd - consum_qtd, consum_qtd)),
]
gestores.sort(key=lambda g: -(g["vend"]/g["meta"] if g["meta"] else 0))

# ---------- VENDEDORES (metas de VENDAS 2026 + vendido/pendente de NEGOCIAÇÕES) ----------
import unicodedata
def _norm(s):
    s = str(s or "").strip().upper()
    s = "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")
    return " ".join(s.split())
def ger_key(s):
    s = _norm(s)
    if "JESSICA" in s: return "JÉSSICA"
    if "NIKOLE" in s: return "NIKOLE"
    if "XAVIER" in s: return "XAVIER"
    return None
# Metas dos vendedores embutidas (fallback p/ quando rodar na nuvem sem o VENDAS 2026.xlsx local).
# Atualizar quando as metas mudarem de mês/valor.
METAS_EMBED_ALL = {
"MAIO": {
    'BRANDINA VIDAL': ('JÉSSICA', 975000.0, 'Brandina Vidal'),
    'CLAUDIA JARDIM': ('JÉSSICA', 975000.0, 'Cláudia Jardim'),
    'DIANA CARVALHO': ('JÉSSICA', 0.0, 'Diana Carvalho'),
    'FLAVIANA CECILIA': ('JÉSSICA', 975000.0, 'Flaviana Cecília'),
    'GABRIEL LISBOA': ('JÉSSICA', 545000.0, 'Gabriel Lisboa'),
    'GRAZIANE ALVES': ('JÉSSICA', 350000.0, 'Graziane Alves'),
    'JESSICA CARVALHO': ('JÉSSICA', 450000.0, 'Jéssica Carvalho'),
    'LUIS FERNANDO': ('JÉSSICA', 545000.0, 'Luis Fernando'),
    'RAFAELA CALDEIRA': ('JÉSSICA', 975000.0, 'Rafaela Caldeira'),
    'THIAGO MENDES': ('JÉSSICA', 975000.0, 'Thiago Mendes'),
    'REVEND. | REPRESENTANTE': ('JÉSSICA', 1300000.0, 'Revend. | Representante'),
    'FABIANO ROSESTOLATO': ('XAVIER', 900000.0, 'Fabiano Rosestolato'),
    'FELIPE GONTIJO': ('XAVIER', 900000.0, 'Felipe Gontijo'),
    'JULIANA ALMEIDA': ('XAVIER', 415000.0, 'Juliana Almeida'),
    'MAYA CAMPOS': ('XAVIER', 700000.0, 'Maya Campos'),
    'PAULA RAPOSO': ('XAVIER', 700000.0, 'Paula Raposo'),
    'RAQUEL BURNS': ('XAVIER', 600000.0, 'Raquel Burns'),
    'AIANI MARTINS': ('NIKOLE', 0.0, 'Aiani Martins'),
    'AMANDA RIBEIRO': ('NIKOLE', 800000.0, 'Amanda Ribeiro'),
    'FILIPE ANJOS': ('NIKOLE', 1050000.0, 'Filipe Anjos'),
    'JAIRO SALES': ('NIKOLE', 805000.0, 'Jairo Sales'),
    'KARINA GRANDINI': ('NIKOLE', 1350000.0, 'Karina Grandini'),
    'MIRELLA CAMARGO': ('NIKOLE', 805000.0, 'Mirella Camargo'),
},
"JUNHO": {
    'BRANDINA VIDAL': ('JÉSSICA', 966500.0, 'Brandina Vidal'),
    'CLAUDIA JARDIM': ('JÉSSICA', 483000.0, 'Cláudia Jardim'),
    'DIANA CARVALHO': ('JÉSSICA', 0.0, 'Diana Carvalho'),
    'FLAVIANA CECILIA': ('JÉSSICA', 966500.0, 'Flaviana Cecília'),
    'GABRIEL LISBOA': ('JÉSSICA', 676750.0, 'Gabriel Lisboa'),
    'GRAZIANE ALVES': ('JÉSSICA', 400000.0, 'Graziane Alves'),
    'JESSICA CARVALHO': ('JÉSSICA', 966500.0, 'Jéssica Carvalho'),
    'LUIS FERNANDO': ('JÉSSICA', 676750.0, 'Luis Fernando'),
    'RAFAELA CALDEIRA': ('JÉSSICA', 966500.0, 'Rafaela Caldeira'),
    'THIAGO MENDES': ('JÉSSICA', 966500.0, 'Thiago Mendes'),
    'REVEND. | REPRESENTANTE': ('JÉSSICA', 966500.0, 'Revend. | Representante'),
    'FABIANO ROSESTOLATO': ('XAVIER', 950000.0, 'Fabiano Rosestolato'),
    'FELIPE GONTIJO': ('XAVIER', 900000.0, 'Felipe Gontijo'),
    'JULIANA ALMEIDA': ('XAVIER', 700000.0, 'Juliana Almeida'),
    'PAULA RAPOSO': ('XAVIER', 950000.0, 'Paula Raposo'),
    'RAQUEL BURNS': ('XAVIER', 700000.0, 'Raquel Burns'),
    'VAGO': ('XAVIER', 449300.0, 'Vago'),
    'AMANDA RIBEIRO': ('NIKOLE', 750000.0, 'Amanda Ribeiro'),
    'FILIPE ANJOS': ('NIKOLE', 830000.0, 'Filipe Anjos'),
    'JAIRO SALES': ('NIKOLE', 830000.0, 'Jairo Sales'),
    'KARINA GRANDINI': ('NIKOLE', 1200000.0, 'Karina Grandini'),
    'MARIANA LOPES': ('NIKOLE', 400000.0, 'Mariana Lopes'),
    'MIRELLA CAMARGO': ('NIKOLE', 900000.0, 'Mirella Camargo'),
},
}
METAS_EMBED = METAS_EMBED_ALL.get(MONTH, {})
import shutil, tempfile
VENDAS_PATH = os.path.join(os.path.dirname(HERE), "VENDAS 2026.xlsx")
metas_v = {}
try:
    _vtmp = os.path.join(tempfile.gettempdir(), "vendas2026_copy.xlsx")
    shutil.copy2(VENDAS_PATH, _vtmp)  # copia 1o p/ contornar arquivo aberto no Excel
    wv = openpyxl.load_workbook(_vtmp, data_only=True)
    for gs in ("JÉSSICA", "XAVIER", "NIKOLE"):
        if gs not in wv.sheetnames: continue
        ws = wv[gs]; atual = None
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value; cc = ws.cell(r, 3).value
            if str(b).strip() == "✅" and cc: atual = str(cc).strip()
            elif str(b).strip().upper() == MONTH and atual:
                d = ws.cell(r, 4).value
                metas_v[_norm(atual)] = (gs, gv(d), atual); atual = None
except Exception as _e:
    print("AVISO: VENDAS 2026 local nao lido (", _e, ") -> usando metas embutidas")
    metas_v = {}
if not metas_v:
    metas_v = {k: v for k, v in METAS_EMBED.items()}   # fallback nuvem
N = wb["NEGOCIAÇÕES"]
def _hasop(g): return isinstance(g, datetime.datetime) or (isinstance(g,str) and g.strip() not in ("","N/A","N/T"))
REPRES = {_norm(x) for x in ["BAZAR LASER","BCMED","DIMITRI","DIOGO MELO","ESSENZA",
    "GILMAR DAMASCENO","LAGUI EQUIPAMENTOS FITNESS","ISP","LARI VENANCIO","OLASERTECH",
    "PAULO NEOTOX","PROW","ZELOTECH"]}
SOCIOS = {_norm(x) for x in ["ANDREIA ALMEIDA","CAROLINE PEREIRA","ÉGIO ROBERTO",
    "GLÁUCIA MOURA","JÉSSICA OLIVEIRA"]}
EXCLUIR = {_norm(x) for x in ["DIANA CARVALHO"]}  # vendedores desligados
REVEND_KEY = _norm("REVEND. | REPRESENTANTE")
revend_ger = metas_v.get(REVEND_KEY, ("JÉSSICA", 0.0, None))[0]
revend_meta = metas_v.get(REVEND_KEY, (None, 0.0, None))[1]

# vendido/pendente agrupado pelo GERENTE DO NEGÓCIO (reconcilia com o total do gestor)
det = {"JÉSSICA":{}, "NIKOLE":{}, "XAVIER":{}}
for r in range(3, N.max_row + 1):
    a = N.cell(r,1).value; q = N.cell(r,17).value
    if not _isdeal(N, r): continue
    vd = N.cell(r,13).value; gk = ger_key(N.cell(r,14).value)
    if not vd or not gk: continue
    vn = _norm(vd)
    d = det[gk].setdefault(vn, {"op":0.0,"pe":0.0,"oq":0,"pq":0,"raw":str(vd).strip()})
    if _hasop(N.cell(r,7).value): d["op"]+=q; d["oq"]+=1
    else: d["pe"]+=q; d["pq"]+=1

def _row(nome, meta, d): return {"nome":nome,"meta":meta,"op":d["op"],"pe":d["pe"],"oq":d["oq"],"pq":d["pq"]}
def _zero(): return {"op":0.0,"pe":0.0,"oq":0,"pq":0}
vend_by = {"JÉSSICA":[], "NIKOLE":[], "XAVIER":[]}
seen = {gk:set() for gk in det}
for gk in det:
    rev=_zero(); soc=_zero()
    for vn,d in det[gk].items():
        if vn in EXCLUIR: continue
        if vn in REPRES:
            for k in ("op","pe","oq","pq"): rev[k]+=d[k]
        elif vn in SOCIOS:
            for k in ("op","pe","oq","pq"): soc[k]+=d[k]
        else:
            vend_by[gk].append(_row(d["raw"].title(), metas_v.get(vn,(None,0.0,None))[1], d)); seen[gk].add(vn)
    if rev["op"] or rev["pe"] or gk==revend_ger:
        vend_by[gk].append(_row("Revend. | Representante", revend_meta if gk==revend_ger else 0.0, rev))
    if soc["op"] or soc["pe"] or gk=="JÉSSICA":
        vend_by[gk].append(_row("Sócios | Outros", 0.0, soc))
# vendedores com meta mas sem vendas no mês (aparecem com 0)
for vn,(gsheet,meta,disp) in metas_v.items():
    if vn==REVEND_KEY or vn in EXCLUIR or vn in REPRES or vn in SOCIOS: continue
    if gsheet in seen and vn in seen[gsheet]: continue
    vend_by.setdefault(gsheet,[]).append(_row(disp.title(), meta, _zero()))
for gk in vend_by: vend_by[gk].sort(key=lambda v:-v["op"])
for G in gestores:
    G["vends"] = vend_by.get(ger_key(G["nome"]), [])
def br0(x):
    try: return "R$ " + f"{float(x):,.0f}".replace(",", ".")
    except Exception: return "R$ 0"

# simulador por equipamento (modelo)
eqsim = {}
for r in range(3, N.max_row + 1):
    a = N.cell(r,1).value; q = N.cell(r,17).value
    if not _isdeal(N, r): continue
    if not _hasop(N.cell(r,7).value): continue   # só fechado com OP
    eq = N.cell(r,5).value; sim = N.cell(r,16).value
    if not eq: continue
    e = str(eq).strip()
    d = eqsim.setdefault(e, [0.0, 0.0, 0])
    d[0] += sim if num(sim) else 0; d[1] += q; d[2] += 1
eqlist = sorted(([e, a[0], a[1], a[2]] for e,a in eqsim.items()), key=lambda x:-x[2])

# Global (hero) = TODOS os negócios da aba NEGOCIAÇÕES (nenhuma venda fica de fora)
g_meta = gv(c("G",27))
_NG = wb["NEGOCIAÇÕES"]
g_vend = g_ab = 0.0; g_vqtd = g_abqtd = 0
for _r in range(3, _NG.max_row + 1):
    _a = _NG.cell(_r,1).value; _q = _NG.cell(_r,17).value
    if not _isdeal(_NG, _r): continue
    _op = _NG.cell(_r,7).value
    if isinstance(_op, datetime.datetime) or (isinstance(_op, str) and _op.strip() not in ("","N/A","N/T")):
        g_vend += _q; g_vqtd += 1
    else:
        g_ab += _q; g_abqtd += 1
falta_real = g_meta - g_vend            # só o já liberado (OP)
falta_proj = g_meta - g_vend - g_ab     # liberado + em aberto
hero_equip = g_vqtd - consum_qtd

# ---------- FICHAS ----------
fichas = []
for r in range(35, 43):
    lab = (c("F", r) or "").strip(); val = gv(c("G", r)); qt = int(gv(c("I", r)))
    if lab.upper() == "FECHADO": lab = "Finalizado"
    else: lab = lab.capitalize() if lab.isupper() else lab.title()
    fichas.append((lab, val, qt))
fichas_tot_val = gv(c("G", 43)); fichas_tot_qtd = int(gv(c("I", 43)))
fmax = max([v for _, v, _ in fichas] + [1])
def fcolor(lab):
    l = lab.lower()
    if l in ("contrato", "finalizado"): return "c-pos"
    if l in ("análise", "analise", "negociação", "negociacao", "decidindo"): return "c-mid"
    if l in ("suspeita fraude",): return "c-warn"
    return "c-neg"

# ---------- FINANCEIRO (NEGOCIAÇÕES) ----------
N = wb["NEGOCIAÇÕES"]
fcols = {"SIM":16,"VENDA":17,"RD":18,"ENT":19,"SANT":20,"AREC":21}
def nb(): return {k:0.0 for k in fcols}
G_, E_, P_ = nb(), nb(), nb()
ne = npd = 0
for r in range(3, N.max_row+1):
    a = N.cell(r,1).value; q = N.cell(r,17).value; op = N.cell(r,7).value
    if not _isdeal(N, r): continue
    has = isinstance(op, datetime.datetime) or (isinstance(op,str) and op.strip() not in ("","N/A","N/T"))
    t = E_ if has else P_
    if has: ne += 1
    else: npd += 1
    for k,col in fcols.items():
        v = N.cell(r,col).value
        if num(v): G_[k]+=v; t[k]+=v
def fpc(a,b): return pctf(a/b*100,2) if b else "0%"
def frow(label, key, base_key, sub):
    cells = ""
    for grp, isto in ((E_,False),(P_,False),(G_,True)):
        v = grp[key]; b = grp[base_key]
        pcell = f'<span class="pc">{fpc(v,b)}</span>' if base_key else ''
        cls = ' class="tot"' if isto else ''
        cells += f'<td{cls}>{br(v)}{pcell}</td>'
    bsub = f'<span class="bs">{sub}</span>' if sub else ''
    return f'<tr><td>{label}{bsub}</td>{cells}</tr>'
total_cash = G_["ENT"] + G_["SANT"]
recebido = total_cash - G_["AREC"]
pct_receb = recebido/total_cash*100 if total_cash else 0
pct_arec = G_["AREC"]/total_cash*100 if total_cash else 0

# =================== RENDER ===================
def barcls(p):
    if p >= 80: return "fill-green"
    if p >= 40: return "fill-amber"
    return "fill-red"

def line_card(L, idx):
    meta, vend, ab = L["meta"], L["vend"], L["ab"]
    p = vend/meta*100 if meta else 0
    batida = p >= 100
    falta = meta - vend
    third = (f'<div class="m"><div class="label">Superou em</div><div class="v money sm" style="color:var(--green)">{br(vend-meta)}</div></div>'
             if batida else
             f'<div class="m"><div class="label">Falta p/ meta</div><div class="v money sm">{br(falta)}</div></div>')
    if "bd" in L and L["bd"]:
        eq, co = L["bd"]
        vend_block = (f'<div class="v money green">{br(vend)}</div>'
                      f'<div class="qbreak"><div><span>Equipamentos</span><b>{eq}</b></div>'
                      f'<div><span>Consumíveis</span><b>{co}</b></div>'
                      f'<div class="tt"><span>Total</span><b>{eq+co}</b></div></div>')
    else:
        vend_block = f'<div class="v money green">{br(vend)}</div><div class="qty"><b>{L["vqtd"]}</b> equipamentos</div>'
    badge = '<span class="badge ok">✓ Meta batida</span>' if batida else ''
    chips = ''
    if L.get("chips"):
        cs = ''
        for nm, vv, vq, av, aq in L["chips"]:
            cs += (f'<div class="chip"><div class="t">{nm}</div><div class="a money">{br(vv)}</div>'
                   f'<div class="b"><b>{vq}</b> vendidos · aberto {br(av)} (<b>{aq}</b>)</div></div>')
        chips = f'<div class="split">{cs}</div>'
    style = ' style="grid-column:1 / -1"' if L.get("full") else ''
    return f'''
    <div class="card" id="lin{idx}"{style}>
      <div class="head"><h3>{L["nome"]}</h3>{badge}</div>
      <div class="metrics">
        <div class="m"><div class="label">Meta</div><div class="v money">{br(meta)}</div></div>
        <div class="m"><div class="label">Vendido (OP)</div>{vend_block}</div>
        {third}
      </div>
      <div class="barline"><div class="bar {barcls(p)}" style="flex:1"><span style="width:{min(p,100):.0f}%"></span></div><div class="pct"{' style="color:var(--green)"' if batida else ''}>{p:.0f}%</div></div>
      {chips}
      <div class="foot"><span>Em aberto <b class="money">{br(ab)}</b> · {L["abqtd"]} negociações</span></div>
    </div>'''

def gest_card(G, idx):
    meta, vend, ab = G["meta"], G["vend"], G["ab"]
    p = vend/meta*100 if meta else 0
    batida = p >= 100
    if G.get("bd"):
        eq, co = G["bd"]
        vend_block = (f'<div class="v money green">{br(vend)}</div>'
                      f'<div class="qbreak"><div><span>Equipamentos</span><b>{eq}</b></div>'
                      f'<div><span>Consumíveis</span><b>{co}</b></div>'
                      f'<div class="tt"><span>Total</span><b>{eq+co}</b></div></div>')
    else:
        vend_block = f'<div class="v money green">{br(vend)}</div><div class="qty"><b>{G["vqtd"]}</b> vendas</div>'
    third = (f'<div class="m"><div class="label">Superou em</div><div class="v money sm" style="color:var(--green)">{br(vend-meta)}</div></div>'
             if batida else
             f'<div class="m"><div class="label">Falta p/ meta</div><div class="v money sm">{br(meta-vend)}</div></div>')
    vrows = ('<div class="vrow vhdr"><span class="nm">Vendedor</span>'
             '<span class="r col-meta">Meta</span><span class="r">Vendido (OP)</span>'
             '<span class="r col-pend">Pendente</span><span class="r">%</span></div>')
    for v in G.get("vends", []):
        if v["meta"] > 0:
            vp = v["op"]/v["meta"]*100
            pc = "pct-g" if vp >= 80 else "pct-a" if vp >= 40 else "pct-r"
            meta_txt = br(v["meta"]); pct_txt = f"{vp:.0f}%"
        else:
            pc = "pct-n"; meta_txt = "—"; pct_txt = "—"
        vrows += (f'<div class="vrow"><span class="nm">{v["nome"]}</span>'
                  f'<span class="r col-meta">{meta_txt}</span>'
                  f'<span class="r">{br(v["op"])}<span class="q">({v["oq"]})</span></span>'
                  f'<span class="r col-pend">{br(v["pe"])}<span class="q">({v["pq"]})</span></span>'
                  f'<span class="r pct {pc}">{pct_txt}</span></div>')
    vlist = f'<div class="vlist"><div class="vhead">Vendedores · do maior pro menor</div>{vrows}</div>' if G.get("vends") else ''
    return f'''
    <div class="card" id="ger{idx}">
      <div class="ghead"><div class="avatar" style="background-image:url('Imagens/{G["img"]}')"></div>
        <div class="gname"><h3>{G["nome"]}</h3><span>{p:.0f}% da meta</span></div></div>
      <div class="metrics">
        <div class="m"><div class="label">Meta</div><div class="v money">{br(meta)}</div></div>
        <div class="m"><div class="label">Vendido (OP)</div>{vend_block}</div>
        {third}
      </div>
      <div class="barline"><div class="bar {barcls(p)}" style="flex:1"><span style="width:{min(p,100):.0f}%"></span></div><div class="pct">{p:.0f}%</div></div>
      <div class="foot"><span>Em aberto <b class="money">{br(ab)}</b> · {G["abqtd"]} negociações</span></div>
      {vlist}
    </div>'''

fichas_rows = ""
for lab, val, qt in fichas:
    w = val/fichas_tot_val*100 if fichas_tot_val else 0
    fichas_rows += (f'<div class="row"><div class="nm">{lab}</div>'
                    f'<div class="fbar {fcolor(lab)}"><span style="width:{w:.0f}%"></span></div>'
                    f'<div class="val"><b class="money">{br(val)}</b> · {qt}</div>'
                    f'<div class="fpct">{w:.0f}%</div></div>\n      ')

fin_rows = (
    f'<tr><td>Simulador</td><td>{br(E_["SIM"])}</td><td>{br(P_["SIM"])}</td><td class="tot">{br(G_["SIM"])}</td></tr>'
    f'<tr class="hl"><td>Valor de Venda</td><td>{br(E_["VENDA"])}</td><td>{br(P_["VENDA"])}</td><td class="tot">{br(G_["VENDA"])}</td></tr>'
    + frow("Valor RD", "RD", "VENDA", "(do valor de venda)")
    + frow("Entrada / Cartão", "ENT", "VENDA", "(do valor de venda)")
    + frow("Santander/Glória", "SANT", "VENDA", "(do valor de venda)")
    + frow("A Receber", "AREC", "ENT", "(pendente entradas)")
)

def ind_card(nome, grp):
    e = grp["ENT"] + grp["SANT"]
    p = e/grp["VENDA"]*100 if grp["VENDA"] else 0
    ok = p >= 30
    cls = "ok" if ok else "alert"
    status = "✓ Acima de 30%" if ok else "⚠ Abaixo de 30%"
    return (f'<div class="ind {cls}"><div class="label">{nome}</div>'
            f'<div class="ind-pct">{pctf(p,1)}</div>'
            f'<div class="ind-status">{status}</div></div>')
ind_block = (f'<div class="label" style="margin:22px 0 10px">Entrada recebida · meta mínima 30% '
             f'<span style="color:var(--muted);font-weight:600;text-transform:none;letter-spacing:0">(Entrada/Cartão + Santander/Glória ÷ Valor de Venda)</span></div>'
             f'<div class="indgrid">{ind_card("Já entrou", E_)}{ind_card("Pendente", P_)}{ind_card("Total", G_)}</div>')

eq_rows = ('<div class="eqrow eqhdr"><span class="nm">Equipamento</span>'
           '<span class="r">Nº</span><span class="r">Simulador</span><span class="r">Valor de Venda</span><span class="r">% meta</span></div>')
_tsim = _tven = 0.0; _tn = 0
for e, sim, ven, n in eqlist:
    pm = ven/g_meta*100 if g_meta else 0
    _tsim += sim; _tven += ven; _tn += n
    eq_rows += (f'<div class="eqrow"><span class="nm">{e}</span>'
                f'<span class="r">{n}</span>'
                f'<span class="r">{br(sim)}</span>'
                f'<span class="r sim">{br(ven)}</span>'
                f'<span class="r pm">{pm:.1f}%</span></div>')
eq_rows += (f'<div class="eqrow eqtot"><span class="nm">TOTAL</span>'
            f'<span class="r">{_tn}</span>'
            f'<span class="r">{br(_tsim)}</span>'
            f'<span class="r">{br(_tven)}</span>'
            f'<span class="r">{_tven/g_meta*100 if g_meta else 0:.1f}%</span></div>')

hg_p = g_vend/g_meta*100 if g_meta else 0
agora = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-3))).strftime("%d/%m/%Y %H:%M")
# chips de mes (Junho fica em /dashcomercial/, Maio fica em /dashcomercial/maio/)
if MONTH == "JUNHO":
    href_maio, href_junho = "maio/", "./"
else:  # MAIO
    href_maio, href_junho = "./", "../"
cls_maio  = "m-chip active" if MONTH == "MAIO"  else "m-chip"
cls_junho = "m-chip active" if MONTH == "JUNHO" else "m-chip"
if falta_proj > 0:
    falta_top = f'<div class="v money">{br(falta_proj)}</div><div class="qty">já liberado + em aberto</div>'
else:
    falta_top = f'<div class="v money green">R$ 0,00</div><div class="qty" style="color:var(--green)">+ {br(-falta_proj)} acima (c/ em aberto)</div>'

_fab_links = "".join(f'<a href="#lin{i}">{L["nome"]}</a>' for i,L in enumerate(lines))
_ger_links = "".join(f'<a href="#ger{i}">{G["nome"]}</a>' for i,G in enumerate(gestores))
navbar = (f'<nav class="navbar">'
          f'<details class="nav-dd"><summary>Por fabricante</summary><div class="dd-menu">{_fab_links}</div></details>'
          f'<details class="nav-dd"><summary>Por gestor</summary><div class="dd-menu">{_ger_links}</div></details>'
          f'<a class="nav-chip" href="#equipamentos">Por equipamento</a>'
          f'<a class="nav-chip" href="#fichas">Fichas de {MONTH_LABEL}</a>'
          f'<a class="nav-chip" href="#financeiro">Financeiro</a>'
          f'</nav>')
SCRIPT = r"""<script>
document.querySelectorAll('.dd-menu a').forEach(function(a){a.addEventListener('click',function(){document.querySelectorAll('.nav-dd[open]').forEach(function(d){d.removeAttribute('open');});});});
document.addEventListener('click',function(e){document.querySelectorAll('.nav-dd[open]').forEach(function(d){if(!d.contains(e.target)){d.removeAttribute('open');}});});
</script>"""

CSS = r"""<style>
  :root{--bg:#f4f5f7;--card:#fff;--ink:#1b1f2a;--muted:#8b909c;--line:#eceef1;--track:#eef0f3;--green:#16a34a;--amber:#f59e0b;--red:#ef4444;--blue:#3b82f6;--ink2:#3d4350;--brand:#103868;--brand2:#1f5f9e}
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:var(--bg);color:var(--ink);line-height:1.45;-webkit-font-smoothing:antialiased}
  .wrap{max-width:1080px;margin:0 auto;padding:32px 20px 56px}
  .topbar{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:14px;background:linear-gradient(120deg,#0d2f57,#103868);border-radius:16px;padding:18px 26px;margin-bottom:14px;box-shadow:0 2px 10px rgba(16,56,104,.20)}
  .topbar .brand{display:flex;align-items:center;gap:18px}
  .topbar .logo{height:38px;width:auto;display:block}
  .topbar h1{font-size:18px;font-weight:700;color:#fff;letter-spacing:-.2px}
  .topbar .tw{border-left:1px solid rgba(255,255,255,.28);padding-left:18px}
  .topbar .sub{font-size:12px;color:rgba(255,255,255,.72);margin-top:2px}
  .live{display:inline-flex;align-items:center;gap:6px;font-size:12px;color:#fff;font-weight:600}
  .live .dot{width:7px;height:7px;border-radius:50%;background:#6BE8A0;box-shadow:0 0 0 3px rgba(107,232,160,.25)}
  .month-pick{display:flex;gap:6px;margin-left:4px;padding-left:14px;border-left:1px solid rgba(255,255,255,.28)}
  .m-chip{font-size:12px;font-weight:700;padding:6px 12px;border-radius:8px;color:rgba(255,255,255,.85);text-decoration:none;background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.18);transition:.15s}
  .m-chip:hover{background:rgba(255,255,255,.18);color:#fff}
  .m-chip.active{background:#fff;color:var(--brand);cursor:default;border-color:#fff}
  .label{font-size:11px;font-weight:600;letter-spacing:.7px;text-transform:uppercase;color:var(--muted)}
  .money{font-variant-numeric:tabular-nums}
  .qty{font-size:11.5px;color:var(--muted);font-weight:600;margin-top:3px;font-variant-numeric:tabular-nums}
  .qty b{color:var(--ink2);font-weight:700}
  .qbreak{margin-top:8px;font-size:11.5px;color:var(--muted);font-variant-numeric:tabular-nums;max-width:190px}
  .qbreak>div{display:flex;justify-content:space-between;gap:14px;padding:1.5px 0}
  .qbreak>div b{color:var(--ink2);font-weight:700}
  .qbreak .tt{border-top:1px solid var(--line);margin-top:3px;padding-top:4px}
  .qbreak .tt span,.qbreak .tt b{color:var(--ink);font-weight:700}
  .hero{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:24px 26px;margin-bottom:14px;box-shadow:0 1px 2px rgba(20,25,40,.04)}
  .hero .topgrid{display:grid;grid-template-columns:repeat(4,1fr);gap:18px;margin-bottom:18px}
  .hero .topgrid .v{font-size:23px;font-weight:700;margin-top:3px;letter-spacing:-.5px}
  .hero .topgrid .v.green{color:var(--green)}
  .bar{height:12px;background:var(--track);border-radius:99px;overflow:hidden}
  .bar>span{display:block;height:100%;border-radius:99px}
  .bar.lg{height:16px}
  .fill-red>span{background:linear-gradient(90deg,#f87171,#ef4444)}
  .fill-amber>span{background:linear-gradient(90deg,#fbbf24,#f59e0b)}
  .fill-green>span{background:linear-gradient(90deg,#34d399,#16a34a)}
  .barline{display:flex;align-items:center;gap:12px;margin-top:10px}
  .barline .pct{font-size:14px;font-weight:700;min-width:48px;text-align:right;font-variant-numeric:tabular-nums}
  .section-title{font-size:13px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;color:var(--brand);margin:26px 4px 12px}
  .grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
  .grid.g3{grid-template-columns:repeat(3,1fr)}
  .ghead{display:flex;align-items:center;gap:12px;margin-bottom:14px}
  .avatar{width:46px;height:46px;flex:0 0 46px;border-radius:50%;background:linear-gradient(135deg,#103868,#1f5f9e);overflow:hidden;background-size:cover;background-position:center}
  .gname h3{font-size:15px;font-weight:700;line-height:1.2}
  .gname span{font-size:12px;color:var(--muted);font-weight:600}
  .card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px 20px;box-shadow:0 1px 2px rgba(20,25,40,.04)}
  .card .head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:14px}
  .card .head h3{font-size:15px;font-weight:700}
  .badge{font-size:10.5px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;padding:4px 9px;border-radius:99px}
  .badge.ok{background:rgba(22,163,74,.12);color:var(--green)}
  .metrics{display:flex;gap:22px;margin-bottom:6px;flex-wrap:wrap}
  .metrics .m .v{font-size:18px;font-weight:700;margin-top:2px;letter-spacing:-.3px}
  .metrics .m .v.sm{font-size:15px;color:var(--ink2)}
  .foot{display:flex;gap:18px;margin-top:14px;padding-top:13px;border-top:1px dashed var(--line);font-size:12.5px;color:var(--muted);flex-wrap:wrap}
  .foot b{color:var(--ink2);font-weight:700}
  .split{display:flex;gap:10px;margin-top:14px}
  .split .chip{flex:1;background:#fafbfc;border:1px solid var(--line);border-radius:10px;padding:10px 12px}
  .split .chip .t{font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.4px}
  .split .chip .a{font-size:14px;font-weight:700;margin-top:3px}
  .split .chip .b{font-size:11.5px;color:var(--muted);margin-top:1px}
  .split .chip .b b{color:var(--ink2);font-weight:700}
  .funnel .row{display:grid;grid-template-columns:120px 1fr 165px 52px;align-items:center;gap:10px;margin:11px 0}
  .funnel .row .nm{font-size:12.5px;font-weight:600;color:var(--ink2)}
  .funnel .row .val{font-size:12px;text-align:right;color:var(--muted);font-variant-numeric:tabular-nums}
  .funnel .row .val b{color:var(--ink);font-weight:700}
  .funnel .row .fpct{text-align:right;font-size:13px;font-weight:700;color:var(--ink2);font-variant-numeric:tabular-nums}
  .fbar{height:9px;background:var(--track);border-radius:99px;overflow:hidden}
  .fbar>span{display:block;height:100%;border-radius:99px}
  .c-pos>span{background:#16a34a}.c-mid>span{background:#1f5f9e}.c-warn>span{background:#f59e0b}.c-neg>span{background:#cbd0d8}
  .fintable{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums}
  .fintable th{font-size:10px;font-weight:600;letter-spacing:.4px;text-transform:uppercase;color:var(--muted);text-align:right;padding:0 0 10px 6px}
  .fintable th:first-child{text-align:left}
  .fintable td{font-size:13px;padding:8px 6px;text-align:right;border-top:1px solid var(--line);color:var(--ink2)}
  .fintable td:first-child{text-align:left;font-weight:600}
  .fintable td.tot{font-weight:700;color:var(--ink)}
  .fintable tr.hl td{background:#fafbfc}
  .fintable .pc{display:block;font-size:11px;font-weight:700;color:var(--brand2);margin-top:2px}
  .fintable td .bs{font-weight:500;font-size:10.5px;color:var(--muted);margin-left:6px}
  .indgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:8px}
  .ind{border-radius:12px;padding:14px 16px;border:1px solid}
  .ind.ok{background:rgba(22,163,74,.08);border-color:rgba(22,163,74,.28)}
  .ind.alert{background:rgba(239,68,68,.07);border-color:rgba(239,68,68,.28)}
  .ind-pct{font-size:25px;font-weight:800;margin-top:4px;letter-spacing:-.5px}
  .ind.ok .ind-pct{color:var(--green)}.ind.alert .ind-pct{color:var(--red)}
  .ind-status{font-size:11.5px;font-weight:700;margin-top:2px}
  .ind.ok .ind-status{color:var(--green)}.ind.alert .ind-status{color:var(--red)}
  @media(max-width:760px){.indgrid{grid-template-columns:1fr}}
  .gstack{display:flex;flex-direction:column;gap:14px}
  .vlist{margin-top:14px;padding-top:12px;border-top:1px solid var(--line)}
  .vhead{font-size:10px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;color:var(--muted);margin-bottom:6px}
  .vrow{display:grid;grid-template-columns:1fr 130px 170px 170px 60px;gap:12px;align-items:center;font-size:12px;padding:7px 2px;border-top:1px solid #f3f4f6}
  .vrow.vhdr{font-size:10px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;color:var(--muted);border-top:none;padding-bottom:5px}
  .vrow .nm{font-weight:700;color:var(--ink2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .vrow .r{text-align:right;font-variant-numeric:tabular-nums;color:var(--ink2)}
  .vrow .q{color:var(--muted);font-weight:500;font-size:10.5px;margin-left:3px}
  .vrow .pct{text-align:right;font-weight:700;font-variant-numeric:tabular-nums}
  .pct-g{color:var(--green)}.pct-a{color:var(--amber)}.pct-r{color:var(--red)}.pct-n{color:var(--muted)}
  @media(max-width:760px){.vrow{grid-template-columns:1fr 80px 60px;font-size:10.5px;gap:6px}.vrow .col-meta,.vrow .col-pend{display:none}}
  html{scroll-behavior:smooth}
  [id]{scroll-margin-top:16px}
  .navbar{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px}
  .nav-chip,.nav-dd>summary{display:inline-flex;align-items:center;gap:6px;background:#fff;border:1px solid var(--line);border-radius:10px;padding:9px 14px;font-size:12.5px;font-weight:700;color:var(--brand);cursor:pointer;list-style:none;box-shadow:0 1px 2px rgba(20,25,40,.04);white-space:nowrap;text-decoration:none}
  .nav-dd>summary::-webkit-details-marker{display:none}
  .nav-dd>summary::marker{content:""}
  .nav-dd>summary::after{content:"▾";font-size:10px;opacity:.65}
  .nav-chip:hover,.nav-dd>summary:hover,.nav-dd[open]>summary{background:var(--brand);color:#fff;border-color:var(--brand)}
  .nav-dd{position:relative}
  .dd-menu{position:absolute;top:calc(100% + 6px);left:0;z-index:30;background:#fff;border:1px solid var(--line);border-radius:10px;box-shadow:0 10px 28px rgba(16,56,104,.16);padding:6px;min-width:235px;display:flex;flex-direction:column;gap:2px}
  .dd-menu a{text-decoration:none;color:var(--ink2);font-size:12.5px;font-weight:600;padding:8px 10px;border-radius:7px;white-space:nowrap}
  .dd-menu a:hover{background:#eef2f7;color:var(--brand)}
  .eqtab .eqrow{display:grid;grid-template-columns:1fr 50px 150px 160px 70px;gap:10px;align-items:center;padding:8px 4px;border-top:1px solid #f3f4f6;font-size:12.5px}
  .eqtab .eqrow.eqhdr{border-top:none;font-size:10px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;color:var(--muted);padding-bottom:6px}
  .eqtab .eqrow.eqtot{border-top:2px solid var(--line);font-weight:700;color:var(--ink);margin-top:2px}
  .eqtab .eqrow .nm{font-weight:700;color:var(--ink2)}
  .eqtab .eqrow .r{text-align:right;font-variant-numeric:tabular-nums;color:var(--ink2)}
  .eqtab .eqrow .sim{font-weight:700;color:var(--brand)}
  .eqtab .eqrow .pm{font-weight:700}
  @media(max-width:760px){.eqtab .eqrow{grid-template-columns:1fr 38px 100px 108px 52px;font-size:10.5px;gap:5px}}
  .note{font-size:12px;color:var(--muted);margin-top:24px;text-align:center;line-height:1.6}
  .note2{display:flex;justify-content:space-between;align-items:baseline;flex-wrap:wrap;gap:8px;font-size:12px;color:var(--muted);margin-top:24px}
  .note2 b{color:var(--ink);font-size:14px;font-weight:700}
  @media(max-width:760px){.hero .topgrid{grid-template-columns:repeat(2,1fr)}.grid,.grid.g3{grid-template-columns:1fr}.fintable td .bs{display:block;margin-left:0}}
</style>"""

html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Painel Comercial — Contourline</title>
{CSS}
</head><body><div class="wrap">
  <div class="topbar">
    <div class="brand"><img class="logo" src="Imagens/contourline-branca.png" alt="Contourline"><div class="tw"><h1>Painel Comercial</h1><div class="sub">atualizado em {agora}</div></div><div class="month-pick"><a href="{href_maio}" class="{cls_maio}">Maio</a><a href="{href_junho}" class="{cls_junho}">Junho</a></div></div>
    <span class="live"><span class="dot"></span> ao vivo</span>
  </div>

  {navbar}

  <div class="hero">
    <div class="topgrid">
      <div><div class="label">Meta Global</div><div class="v money">{br(g_meta)}</div></div>
      <div><div class="label">Vendido (OP liberada)</div><div class="v green money">{br(g_vend)}</div>
        <div class="qbreak"><div><span>Equipamentos</span><b>{hero_equip}</b></div><div><span>Consumíveis</span><b>{consum_qtd}</b></div><div class="tt"><span>Total</span><b>{g_vqtd}</b></div></div></div>
      <div><div class="label">Em aberto</div><div class="v money" style="color:var(--ink2)">{br(g_ab)}</div><div class="qty"><b>{g_abqtd}</b> negociações</div></div>
      <div><div class="label">Falta p/ meta</div>{falta_top}</div>
    </div>
    <div class="label">% da meta já batido</div>
    <div class="barline"><div class="bar lg {barcls(hg_p)}" style="flex:1"><span style="width:{min(hg_p,100):.0f}%"></span></div><div class="pct">{hg_p:.0f}%</div></div>
  </div>

  <div class="section-title" id="fabricantes">Por linha de equipamento · cada uma com sua meta</div>
  <div class="grid">{''.join(line_card(L, i) for i,L in enumerate(lines))}
  </div>

  <div class="section-title" id="equipamentos">Por equipamento · simulador (fechado com OP) <span style="color:var(--muted);font-weight:600;text-transform:none;letter-spacing:0">(% sobre a meta total)</span></div>
  <div class="card eqtab">{eq_rows}</div>

  <div class="section-title" id="gestores">Por gestor · ranking de meta (com vendedores)</div>
  <div class="gstack">{''.join(gest_card(G, i) for i,G in enumerate(gestores))}
  </div>

  <div class="section-title" id="fichas">Fichas · {MONTH_LABEL} <span style="color:var(--muted);font-weight:600;text-transform:none;letter-spacing:0">({fichas_tot_qtd} fichas · {br(fichas_tot_val)})</span></div>
  <div class="card funnel">
      {fichas_rows}</div>

  <div class="section-title" id="financeiro">Financeiro · Negociações <span style="color:var(--muted);font-weight:600;text-transform:none;letter-spacing:0">({ne+npd} negócios)</span></div>
  <div class="card">
    <div class="metrics" style="margin-bottom:8px">
      <div class="m"><div class="label">Já recebido</div><div class="v money green">{br(recebido)}</div><div style="color:var(--green);font-size:13px;font-weight:700;margin-top:2px">{pctf(pct_receb,1)}</div><div class="qty">Entrada + Santander/Glória − A Receber</div></div>
      <div class="m"><div class="label">A receber</div><div class="v money" style="color:#d97706">{br(G_["AREC"])}</div><div style="color:#d97706;font-size:13px;font-weight:700;margin-top:2px">{pctf(pct_arec,1)}</div><div class="qty">cliente ainda não pagou</div></div>
      <div class="m"><div class="label">Total</div><div class="v money">{br(total_cash)}</div><div style="color:var(--ink2);font-size:13px;font-weight:700;margin-top:2px">100%</div><div class="qty">Entrada/Cartão + Santander/Glória</div></div>
    </div>
    {ind_block}
    <div class="label" style="margin:22px 0 10px">Detalhamento por negócio</div>
    <table class="fintable">
      <thead><tr><th>Valor</th><th>Já entrou · {ne}</th><th>Pendente · {npd}</th><th>Total · {ne+npd}</th></tr></thead>
      <tbody>{fin_rows}</tbody>
    </table>
  </div>

  <div class="note2"><span>Gerado automaticamente da planilha em {agora}.</span><span>Falta p/ meta REAL: <b>{br(falta_real)}</b></span></div>
</div>{SCRIPT}</body></html>"""

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)
print("OK ->", OUT)
print(f"Linhas:{len(lines)} Gestores:{len(gestores)} Fichas:{len(fichas)} Negocios:{ne+npd} (entrou {ne}/pend {npd})")
print("Recebido:", br(recebido), "| A receber:", br(G_['AREC']))
for G in gestores:
    _s = sum(v["op"] for v in G["vends"])
    print(f"  conferencia {G['nome']}: card {br(G['vend'])} = vendedores {br(_s)}  (dif {G['vend']-_s:.2f})")
