#!/usr/bin/env python3
"""
gerar_dashboard.py
------------------
Lê o arquivo Excel de inadimplência e gera o index.html do dashboard.

Uso local:
    pip install pandas openpyxl
    python gerar_dashboard.py

No GitHub Actions esse script roda automaticamente ao fazer push
de um novo arquivo .xlsx na pasta /dados/.
"""

import json
import sys
import os
import pandas as pd
from pathlib import Path

# ── Configuração ────────────────────────────────────────────────────────────
DADOS_DIR = Path("dados")          # pasta onde ficam os arquivos Excel
OUTPUT    = Path("index.html")     # arquivo gerado

PRODUTOS_ORDER = [
    "Venda faturada", "Z-ON corporativo", "Convênio vale mercado",
    "Cartão presente", "Empréstimo PJ", "CDC 76", "Cartão frota"
]

# ── Leitura do Excel mais recente ───────────────────────────────────────────
excels = sorted(DADOS_DIR.glob("*.xlsx"))
if not excels:
    sys.exit("Nenhum arquivo .xlsx encontrado em /dados/")

excel_path = excels[-1]   # pega o mais recente (ordem alfabética)
print(f"Lendo: {excel_path}")

xl = pd.ExcelFile(excel_path)
abas = xl.sheet_names
print(f"Abas encontradas: {abas}")

partes = []
for aba in abas:
    parte = pd.read_excel(xl, sheet_name=aba)
    parte.columns = parte.columns.str.strip()

    # Se a aba não tiver coluna PRODUTO, usa o nome da aba como produto
    if "PRODUTO" not in parte.columns:
        print(f"  Aba '{aba}': coluna PRODUTO ausente — usando nome da aba como produto")
        parte["PRODUTO"] = aba

    partes.append(parte)

df = pd.concat(partes, ignore_index=True)
df.columns = df.columns.str.strip()

# ── Validação de colunas ────────────────────────────────────────────────────
required = {"CÓDIGO", "D. ATRASO", "VALOR", "PRODUTO"}
missing = required - set(df.columns)
if missing:
    sys.exit(f"Colunas ausentes no Excel: {missing}")

# ── Codificação dos produtos ────────────────────────────────────────────────
produtos_na_base = df["PRODUTO"].unique().tolist()
# Mantém ordem padrão + adiciona novos produtos ao final
produtos = [p for p in PRODUTOS_ORDER if p in produtos_na_base]
produtos += [p for p in produtos_na_base if p not in produtos]
produto_map = {p: i for i, p in enumerate(produtos)}

# ── Codificação das empresas ────────────────────────────────────────────────
empresa_map = {c: i for i, c in enumerate(df["CÓDIGO"].unique())}

# ── Montagem do array compacto ──────────────────────────────────────────────
rows = []
for _, row in df.iterrows():
    rows.append([
        int(row["D. ATRASO"]),
        round(float(row["VALOR"]), 2),
        produto_map.get(row["PRODUTO"], 0),
        empresa_map[row["CÓDIGO"]],
    ])

raw_json = json.dumps(rows, separators=(",", ":"))

# Data da base para exibir no dashboard
from datetime import date
data_base = date.today().strftime("%d/%m/%Y")

# ── Geração do HTML ──────────────────────────────────────────────────────────
TMPL_BEFORE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Carteira de Inadimplência</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:system-ui,-apple-system,sans-serif;background:#0F172A;color:#E2E8F0;min-height:100vh;padding:14px}
h1{font-size:20px;font-weight:700;color:#F8FAFC;letter-spacing:-0.5px}
.sub{font-size:11px;color:#64748B;margin-top:3px}
.header{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:8px}
.btn-group{display:flex;gap:6px}
.btn{padding:5px 14px;border-radius:6px;font-size:12px;cursor:pointer;background:#1E293B;border:1px solid #334155;color:#94A3B8;font-weight:500;transition:all .15s}
.btn.active{background:#3B82F6;border-color:#3B82F6;color:#fff}
.card{background:#1E293B;border:1px solid #334155;border-radius:10px;padding:13px}
.kpi-grid{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:13px}
.kpi{background:#1E293B;border:1px solid #334155;border-radius:10px;padding:13px 8px;text-align:center}
.kpi-val{font-size:18px;font-weight:700}
.kpi-lbl{font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:.5px;margin-top:2px}
.kpi-sub{font-size:10px;color:#475569;margin-top:2px}
.filter-box{background:#1E293B;border:1px solid #334155;border-radius:10px;padding:13px;margin-bottom:13px}
.filter-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}
.filter-title{font-size:10px;color:#94A3B8;text-transform:uppercase;letter-spacing:.8px;margin-bottom:7px;font-weight:600;display:flex;justify-content:space-between}
.filter-title a{color:#3B82F6;cursor:pointer;font-weight:400;text-transform:none;letter-spacing:0}
.chips{display:flex;flex-wrap:wrap;gap:5px}
.chip{display:inline-flex;align-items:center;padding:3px 10px;border-radius:4px;font-size:11px;cursor:pointer;border:1px solid #334155;color:#475569;background:#0F172A;transition:all .15s;user-select:none;opacity:0.45;text-decoration:none}
.chip.active{background:#1E3A5F;border-color:#3B82F6;color:#CBD5E1;font-weight:600;opacity:1}
.val-inputs{display:flex;gap:6px;align-items:center;margin-top:8px}
.val-inputs input{background:#0F172A;border:1px solid #334155;border-radius:6px;padding:5px 8px;color:#E2E8F0;font-size:12px;width:80px}
.val-inputs span{color:#475569;font-size:12px}
.val-inputs .clear{color:#EF4444;cursor:pointer;font-size:12px}
.tabs{display:flex;gap:6px;margin-bottom:12px}
.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.chart-card{background:#1E293B;border:1px solid #334155;border-radius:10px;padding:14px}
.chart-title{font-size:12px;color:#94A3B8;font-weight:600;margin-bottom:12px}
.chart-full{background:#1E293B;border:1px solid #334155;border-radius:10px;padding:14px;margin-bottom:0}
.insight-box{background:#0F2A1A;border:1px solid #16A34A;border-radius:10px;padding:14px;margin-top:13px}
.insight-title{font-size:11px;font-weight:700;color:#4ADE80;text-transform:uppercase;letter-spacing:.6px;margin-bottom:9px}
.insight-item{font-size:13px;color:#D1FAE5;line-height:1.65;margin-bottom:3px}
.insight-footer{margin-top:9px;font-size:11px;color:#065F46;border-top:1px solid #166534;padding-top:7px}
canvas{display:block}
.hidden{display:none}
</style>
</head>
<body>
<div class="header">
  <div>
    <h1>📋 Carteira de Inadimplência</h1>
    <div class="sub" id="sub-info">Base: 21/07/2026</div>
  </div>
  <div class="btn-group">
    <button class="btn active" id="btn-qty" onclick="setMetrica('qty')">Por Títulos</button>
    <button class="btn" id="btn-unique" onclick="setMetrica('unique')">Empresas Únicas</button>
    <button class="btn" id="btn-val" onclick="setMetrica('val')">Valor (R$)</button>
    <button class="btn active" id="btn-filters" onclick="toggleFilters()">▲ Filtros</button>
  </div>
</div>

<div class="filter-box" id="filter-box">
  <div class="filter-grid">
    <div>
      <div class="filter-title">Produto <a onclick="toggleAllProdutos()">limpar</a></div>
      <div class="chips" id="chips-produto"></div>
    </div>
    <div>
      <div class="filter-title">Faixa de Atraso (dias) <a onclick="toggleAllBands()">limpar</a></div>
      <div class="chips" id="chips-band"></div>
    </div>
    <div>
      <div class="filter-title">Faixa de Valor (R$)</div>
      <div class="chips" id="chips-valor"></div>
      <div class="val-inputs">
        <input type="number" id="val-min" placeholder="Mín" oninput="applyFilters()"/>
        <span>—</span>
        <input type="number" id="val-max" placeholder="Máx" oninput="applyFilters()"/>
        <span class="clear" onclick="clearValor()">✕</span>
      </div>
    </div>
  </div>
</div>

<div class="kpi-grid">
  <div class="kpi"><div class="kpi-val" id="k-qty" style="color:#60A5FA">—</div><div class="kpi-lbl">Títulos</div><div class="kpi-sub" id="k-qty-sub">—</div></div>
  <div class="kpi"><div class="kpi-val" id="k-val" style="color:#6EE7B7">—</div><div class="kpi-lbl">Valor Total</div><div class="kpi-sub" id="k-val-sub">—</div></div>
  <div class="kpi"><div class="kpi-val" id="k-avg" style="color:#94A3B8">—</div><div class="kpi-lbl">Atraso Médio</div><div class="kpi-sub" id="k-avg-sub">—</div></div>
  <div class="kpi"><div class="kpi-val" id="k-med" style="color:#94A3B8">—</div><div class="kpi-lbl">Atraso Mediano</div><div class="kpi-sub" id="k-med-sub">—</div></div>
  <div class="kpi"><div class="kpi-val" id="k-emp" style="color:#94A3B8">—</div><div class="kpi-lbl">Empresas</div><div class="kpi-sub">clientes únicos</div></div>
  <div class="kpi"><div class="kpi-val" id="k-tkt" style="color:#94A3B8">—</div><div class="kpi-lbl">Ticket Médio</div><div class="kpi-sub">por título</div></div>
</div>

<div class="tabs">
  <button class="btn active" id="tab-atraso" onclick="setTab('atraso')">📅 Por Faixa de Atraso</button>
  <button class="btn" id="tab-produto" onclick="setTab('produto')">📦 Por Produto</button>
  <button class="btn" id="tab-valor" onclick="setTab('valor')">💵 Por Faixa de Valor</button>
</div>

<div id="view-atraso">
  <div class="chart-full"><div class="chart-title" id="title-atraso">Distribuição por Faixa de Atraso — Quantidade de Títulos</div><div style="position:relative;height:300px"><canvas id="chart-atraso"></canvas></div></div>
</div>
<div id="view-produto" class="hidden">
  <div class="charts-row">
    <div class="chart-card"><div class="chart-title" id="title-produto">Títulos por Produto</div><div style="position:relative;height:280px"><canvas id="chart-prod-qty"></canvas></div></div>
    <div class="chart-card"><div class="chart-title">Valor Total por Produto (R$)</div><div style="position:relative;height:280px"><canvas id="chart-prod-val"></canvas></div></div>
  </div>
</div>
<div id="view-valor" class="hidden">
  <div class="charts-row">
    <div class="chart-card"><div class="chart-title" id="title-val-qty">Títulos por Faixa de Valor</div><div style="position:relative;height:280px"><canvas id="chart-val-qty"></canvas></div></div>
    <div class="chart-card"><div class="chart-title">Valor Total por Faixa (R$)</div><div style="position:relative;height:280px"><canvas id="chart-val-val"></canvas></div></div>
  </div>
</div>

<div class="insight-box">
  <div class="insight-title">📌 Resumo Analítico da Carteira</div>
  <div id="insight-content"></div>
  <div class="insight-footer" id="insight-footer"></div>
</div>

<script>
const PRODUTOS = ["Venda faturada","Z-ON corporativo","Convênio vale mercado","Cartão presente","Empréstimo PJ","CDC 76","Cartão frota"];
const PRODUTO_COLORS = ["#3B82F6","#2563EB","#1D4ED8","#1E40AF","#1E3A8A","#60A5FA","#93C5FD"];
const DELAY_BANDS = [
  {label:"01-30",    min:1,    max:30,       color:"#BFDBFE"},
  {label:"31-60",    min:31,   max:60,       color:"#93C5FD"},
  {label:"61-90",    min:61,   max:90,       color:"#60A5FA"},
  {label:"91-120",   min:91,   max:120,      color:"#3B82F6"},
  {label:"121-150",  min:121,  max:150,      color:"#2563EB"},
  {label:"151-180",  min:151,  max:180,      color:"#1D4ED8"},
  {label:"181-360",  min:181,  max:360,      color:"#1E40AF"},
  {label:"361-720",  min:361,  max:720,      color:"#1E3A8A"},
  {label:"721-1080", min:721,  max:1080,     color:"#172554"},
  {label:"1081-1440",min:1081, max:1440,     color:"#0F172A"},
  {label:"1441-1800",min:1441, max:1800,     color:"#0A1020"},
  {label:">1800",    min:1801, max:Infinity, color:"#060810"},
];
const VALUE_BANDS = [
  {label:"Até R$100",  min:0,   max:100},
  {label:"R$100-500",  min:100, max:500},
  {label:"R$500-1k",   min:500, max:1000},
  {label:"R$1k-5k",    min:1000,max:5000},
  {label:"R$5k-10k",   min:5000,max:10000},
  {label:"Acima R$10k",min:10000,max:Infinity},
];

const RAW = """

TMPL_AFTER = """

const fmt = n => new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(n);
const fmtK = n => n>=1e6?`R$ ${(n/1e6).toFixed(2)}M`:n>=1e3?`R$ ${(n/1e3).toFixed(1)}K`:fmt(n);
const fmtN = n => new Intl.NumberFormat("pt-BR").format(Math.round(n));
const getBand = d => DELAY_BANDS.findIndex(b => d >= b.min && d <= b.max);

// State
let selProdutos = new Set(PRODUTOS.map((_,i)=>i));
let selBands    = new Set(DELAY_BANDS.map((_,i)=>i));
let selValBand  = null;
let metrica     = "qty";
let activeTab   = "atraso";
let filtersOpen = true;

// Charts
let cAtraso, cProdQty, cProdVal, cValQty, cValVal;

const TOTAL_QTY = RAW.length;
const TOTAL_VAL = RAW.reduce((s,r)=>s+r[1],0);

// --- INIT CHIPS ---
function initChips(){
  const cp = document.getElementById("chips-produto");
  PRODUTOS.forEach((p,i)=>{
    const el = document.createElement("span");
    el.className="chip active";
    el.textContent=p;
    el.dataset.idx=i;
    el.onclick=()=>{
      selProdutos.has(i)?selProdutos.delete(i):selProdutos.add(i);
      el.classList.toggle("active",selProdutos.has(i));
      applyFilters();
    };
    cp.appendChild(el);
  });

  const cb = document.getElementById("chips-band");
  DELAY_BANDS.forEach((b,i)=>{
    const el = document.createElement("span");
    el.className="chip active";
    el.textContent=b.label;
    el.dataset.idx=i;
    el.onclick=()=>{
      selBands.has(i)?selBands.delete(i):selBands.add(i);
      el.classList.toggle("active",selBands.has(i));
      applyFilters();
    };
    cb.appendChild(el);
  });

  const cv = document.getElementById("chips-valor");
  VALUE_BANDS.forEach((vb,i)=>{
    const el = document.createElement("span");
    el.className="chip";
    el.textContent=vb.label;
    el.onclick=()=>{
      if(selValBand===i){
        selValBand=null; el.classList.remove("active");
        clearValor(false);
      } else {
        document.querySelectorAll("#chips-valor .chip").forEach(c=>{
          c.classList.remove("active");
        });
        selValBand=i; el.classList.add("active");
        document.getElementById("val-min").value = vb.min===0?"":vb.min;
        document.getElementById("val-max").value = vb.max===Infinity?"":vb.max;
      }
      applyFilters();
    };
    cv.appendChild(el);
  });
}

function clearValor(doApply=true){
  selValBand=null;
  document.getElementById("val-min").value="";
  document.getElementById("val-max").value="";
  document.querySelectorAll("#chips-valor .chip").forEach(c=>{
    c.classList.remove("active");
  });
  if(doApply) applyFilters();
}

function toggleAllProdutos(){
  const all = selProdutos.size===PRODUTOS.length;
  if(all){ selProdutos.clear(); } else { PRODUTOS.forEach((_,i)=>selProdutos.add(i)); }
  document.querySelectorAll("#chips-produto .chip").forEach((el,i)=>{
    el.classList.toggle("active",!all);
    el.style.background=!all?PRODUTO_COLORS[i]+"22":"#0F172A";
  });
  document.querySelector("#filter-box .filter-title a").textContent=all?"todos":"limpar";
  applyFilters();
}

function toggleAllBands(){
  const all = selBands.size===DELAY_BANDS.length;
  if(all){ selBands.clear(); } else { DELAY_BANDS.forEach((_,i)=>selBands.add(i)); }
  document.querySelectorAll("#chips-band .chip").forEach((el,i)=>{
    el.classList.toggle("active",!all);
  });
  applyFilters();
}

// --- FILTER & COMPUTE ---
let filtered = [];
function applyFilters(){
  const vMin = parseFloat(document.getElementById("val-min").value)||0;
  const vMax = parseFloat(document.getElementById("val-max").value)||Infinity;
  filtered = RAW.filter(([delay,value,pidx])=>{
    if(!selProdutos.has(pidx)) return false;
    const b=getBand(delay);
    if(b===-1||!selBands.has(b)) return false;
    if(value<vMin||value>vMax) return false;
    return true;
  });
  updateKPIs();
  updateCharts();
  updateInsights();
}

function updateKPIs(){
  const n=filtered.length;
  const tv=filtered.reduce((s,r)=>s+r[1],0);
  const delays=filtered.map(r=>r[0]);
  const avg=n?delays.reduce((a,b)=>a+b,0)/n:0;
  const sorted=[...delays].sort((a,b)=>a-b);
  const med=n?sorted[Math.floor(n/2)]:0;
  const emps=new Set(filtered.map(r=>r[3])).size;
  const tkt=n?tv/n:0;
  set("k-qty",fmtN(n),"k-qty-sub",`de ${fmtN(TOTAL_QTY)} total`,"#60A5FA");
  set("k-val",fmtK(tv),"k-val-sub",n?`${((tv/TOTAL_VAL)*100).toFixed(1)}% da carteira`:"—","#6EE7B7");
  set("k-avg",fmtN(avg)+"d","k-avg-sub",`~${(avg/365).toFixed(1)} anos`,"#94A3B8");
  set("k-med",fmtN(med)+"d","k-med-sub",`~${(med/365).toFixed(1)} anos`,"#94A3B8");
  document.getElementById("k-emp").textContent=fmtN(emps);
  set("k-tkt",fmtK(tkt),"","","#EC4899");
}
function set(id,v,subId,subV,color){
  const el=document.getElementById(id); if(el){el.textContent=v; if(color)el.style.color=color;}
  if(subId){const s=document.getElementById(subId); if(s)s.textContent=subV;}
}

function getDelayAgg(){
  return DELAY_BANDS.map((b,i)=>{
    const rows=filtered.filter(r=>getBand(r[0])===i);
    return {
      label:b.label,color:b.color,
      qty:rows.length,
      val:rows.reduce((s,r)=>s+r[1],0),
      unique:new Set(rows.map(r=>r[3])).size,
    };
  });
}
function getProdAgg(){
  return PRODUTOS.map((p,i)=>{
    const rows=filtered.filter(r=>r[2]===i);
    return {
      name:p,color:PRODUTO_COLORS[i],
      qty:rows.length,
      val:rows.reduce((s,r)=>s+r[1],0),
      unique:new Set(rows.map(r=>r[3])).size,
    };
  });
}
function getValAgg(){
  return VALUE_BANDS.map(vb=>{
    const rows=filtered.filter(r=>r[1]>=vb.min&&r[1]<vb.max);
    return {
      label:vb.label,
      qty:rows.length,
      val:rows.reduce((s,r)=>s+r[1],0),
      unique:new Set(rows.map(r=>r[3])).size,
    };
  });
}

const gridColor="#1E3A5F";
const baseOpts={
  responsive:true,maintainAspectRatio:false,
  plugins:{legend:{display:false},tooltip:{
    backgroundColor:"#0F172A",borderColor:"#334155",borderWidth:1,
    titleColor:"#94A3B8",bodyColor:"#F8FAFC",callbacks:{
      label:ctx=>ctx.dataset.isVal?fmtK(ctx.raw):metricLabel(ctx.raw,metrica)
    }
  }},
  scales:{
    x:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:10}}},
    y:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:10},callback:(v,i,t)=>t[i]?.label?.length>14?t[i].label.slice(0,13)+"…":t[i]?.label}},
  }
};

function mkChart(id,type,data,opts){
  const ctx=document.getElementById(id).getContext("2d");
  return new Chart(ctx,{type,data,options:{...baseOpts,...opts}});
}

function initCharts(){
  const da=getDelayAgg();
  cAtraso=mkChart("chart-atraso","bar",
    {labels:da.map(d=>d.label),datasets:[{data:da.map(d=>d.qty),backgroundColor:da.map(d=>d.color),borderRadius:4}]},
    {plugins:{...baseOpts.plugins},scales:{x:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:11}}},y:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:11},callback:v=>fmtN(v)}}}}
  );

  const prodByQty=[...getProdAgg()].sort((a,b)=>b.qty-a.qty);
  cProdQty=mkChart("chart-prod-qty","bar",
    {labels:prodByQty.map(p=>p.name),datasets:[{data:prodByQty.map(p=>p.qty),backgroundColor:prodByQty.map(p=>p.color),borderRadius:4}]},
    {indexAxis:"y",plugins:{...baseOpts.plugins},scales:{x:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:10},callback:v=>fmtN(v)}},y:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:11}}}}}
  );

  const prodByVal=[...getProdAgg()].sort((a,b)=>b.val-a.val);
  cProdVal=mkChart("chart-prod-val","bar",
    {labels:prodByVal.map(p=>p.name),datasets:[{data:prodByVal.map(p=>p.val),backgroundColor:prodByVal.map(p=>p.color),borderRadius:4,isVal:true}]},
    {indexAxis:"y",plugins:{...baseOpts.plugins},scales:{x:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:10},callback:v=>fmtK(v)}},y:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:11}}}}}
  );

  const va=getValAgg();
  cValQty=mkChart("chart-val-qty","bar",
    {labels:va.map(v=>v.label),datasets:[{data:va.map(v=>v.qty),backgroundColor:"#3B82F6",borderRadius:4}]},
    {plugins:{...baseOpts.plugins},scales:{x:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:10}}},y:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:11},callback:v=>fmtN(v)}}}}
  );
  cValVal=mkChart("chart-val-val","bar",
    {labels:va.map(v=>v.label),datasets:[{data:va.map(v=>v.val),backgroundColor:"#3B82F6",borderRadius:4,isVal:true}]},
    {plugins:{...baseOpts.plugins},scales:{x:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:10}}},y:{grid:{color:gridColor},ticks:{color:"#64748B",font:{size:11},callback:v=>fmtK(v)}}}}
  );
}

function metricLabel(v,key){
  if(key==="val") return fmtK(v);
  return fmtN(v)+(key==="unique"?" empresas":" títulos");
}
function metricTitle(key){
  if(key==="qty") return "Títulos";
  if(key==="unique") return "Empresas Únicas";
  return "Valor (R$)";
}
function updateCharts(){
  const da=getDelayAgg();
  const key=metrica;
  cAtraso.data.datasets[0].data=da.map(d=>d[key]);
  cAtraso.data.datasets[0].backgroundColor=da.map(d=>d.color);
  cAtraso.options.scales.y.ticks.callback=v=>key==="val"?fmtK(v):fmtN(v);
  cAtraso.options.plugins.tooltip.callbacks.label=ctx=>metricLabel(ctx.raw,key);
  document.getElementById("title-atraso").textContent=`Distribuição por Faixa de Atraso — ${metricTitle(key)}`;
  cAtraso.update();

  const prod=getProdAgg();
  const prodSorted=[...prod].sort((a,b)=>b[key]-a[key]);
  const prodTitle=document.getElementById("title-produto");
  if(prodTitle) prodTitle.textContent=`${metricTitle(key)} por Produto`;
  cProdQty.data.labels=prodSorted.map(p=>p.name);
  cProdQty.data.datasets[0].data=prodSorted.map(p=>p[key]);
  cProdQty.data.datasets[0].backgroundColor=prodSorted.map(p=>p.color);
  cProdQty.options.scales.x.ticks.callback=v=>key==="val"?fmtK(v):fmtN(v);
  cProdQty.options.plugins.tooltip.callbacks.label=ctx=>metricLabel(ctx.raw,key);
  cProdQty.update();

  const prodByVal=[...prod].sort((a,b)=>b.val-a.val);
  cProdVal.data.labels=prodByVal.map(p=>p.name);
  cProdVal.data.datasets[0].data=prodByVal.map(p=>p.val);
  cProdVal.data.datasets[0].backgroundColor=prodByVal.map(p=>p.color);
  cProdVal.update();

  const va=getValAgg();
  cValQty.data.datasets[0].data=va.map(v=>v[key]);
  cValQty.options.scales.y.ticks.callback=v=>key==="val"?fmtK(v):fmtN(v);
  cValQty.options.plugins.tooltip.callbacks.label=ctx=>metricLabel(ctx.raw,key);
  const valQtyTitle=document.getElementById("title-val-qty");
  if(valQtyTitle) valQtyTitle.textContent=`${metricTitle(key)} por Faixa de Valor`;
  cValQty.update();
  cValVal.data.datasets[0].data=va.map(v=>v.val); cValVal.update();
}

function updateInsights(){
  const n=filtered.length, tv=filtered.reduce((s,r)=>s+r[1],0);
  const lines=[];
  if(n===0){lines.push("Nenhum registro corresponde aos filtros.");}
  else{
    const a360=filtered.filter(r=>r[0]>360).length;
    const a1800=filtered.filter(r=>r[0]>1800).length;
    const v360=filtered.filter(r=>r[0]>360).reduce((s,r)=>s+r[1],0);
    const v1800=filtered.filter(r=>r[0]>1800).reduce((s,r)=>s+r[1],0);
    const sorted=[...filtered.map(r=>r[0])].sort((a,b)=>a-b);
    const med=sorted[Math.floor(n/2)];
    const prodStats=PRODUTOS.map((p,i)=>({p,qty:filtered.filter(r=>r[2]===i).length,val:filtered.filter(r=>r[2]===i).reduce((s,r)=>s+r[1],0)}));
    const tpq=[...prodStats].sort((a,b)=>b.qty-a.qty)[0];
    const tpv=[...prodStats].sort((a,b)=>b.val-a.val)[0];
    if(a360>0) lines.push(`🔴 ${((a360/n)*100).toFixed(1)}% dos títulos (${((v360/tv)*100).toFixed(1)}% do valor) estão acima de 360 dias de atraso.`);
    if(a1800>0) lines.push(`⚠️ ${((a1800/n)*100).toFixed(1)}% dos títulos (${((v1800/tv)*100).toFixed(1)}% do valor) ultrapassam 1.800 dias — alta improbabilidade de recuperação.`);
    lines.push(`📊 Mediana de atraso: ${fmtN(med)} dias (~${(med/365).toFixed(1)} anos). Carteira com perfil predominantemente envelhecido.`);
    if(tpq&&tpq.qty>0) lines.push(`📦 Por quantidade: "${tpq.p}" domina com ${((tpq.qty/n)*100).toFixed(1)}% dos títulos (${new Set(filtered.filter(r=>r[2]===PRODUTOS.indexOf(tpq.p)).map(r=>r[3])).size} empresas únicas).`);
    if(tpv&&tpv.val>0) lines.push(`💰 Por valor: "${tpv.p}" representa ${((tpv.val/tv)*100).toFixed(1)}% do saldo — ticket médio de ${fmt(tpv.val/tpv.qty)}.`);
    const totalUnique=new Set(filtered.map(r=>r[3])).size;
    lines.push(`🏢 ${fmtN(totalUnique)} empresas únicas com saldo em aberto nesta seleção.`);
  }
  const cont=document.getElementById("insight-content");
  cont.innerHTML=lines.map(l=>`<div class="insight-item">${l}</div>`).join("");
  document.getElementById("insight-footer").textContent=n>0?`Exibindo ${fmtN(n)} de ${fmtN(TOTAL_QTY)} títulos · Valor filtrado: ${fmt(tv)}`:"";
}

function setMetrica(m){
  metrica=m;
  document.getElementById("btn-qty").classList.toggle("active",m==="qty");
  document.getElementById("btn-unique").classList.toggle("active",m==="unique");
  document.getElementById("btn-val").classList.toggle("active",m==="val");
  updateCharts();
}

function setTab(t){
  activeTab=t;
  ["atraso","produto","valor"].forEach(v=>{
    document.getElementById("view-"+v).classList.toggle("hidden",v!==t);
    document.getElementById("tab-"+v).classList.toggle("active",v===t);
  });
}

function toggleFilters(){
  filtersOpen=!filtersOpen;
  document.getElementById("filter-box").classList.toggle("hidden",!filtersOpen);
  document.getElementById("btn-filters").textContent=(filtersOpen?"▲":"▼")+" Filtros";
  document.getElementById("btn-filters").classList.toggle("active",filtersOpen);
}

// Boot
window.addEventListener("DOMContentLoaded",()=>{
  document.getElementById("sub-info").textContent=`Base: 21/07/2026 · ${fmtN(TOTAL_QTY)} títulos · ${fmt(TOTAL_VAL)} em saldo`;
  filtered=[...RAW];
  initChips();
  initCharts();
  updateKPIs();
  updateInsights();
});
</script>
</body>
</html>"""

html = TMPL_BEFORE + raw_json + TMPL_AFTER

# Atualiza a data de base dinamicamente
html = html.replace("Base: 21/07/2026", f"Base: {data_base}")

OUTPUT.write_text(html, encoding="utf-8")
print(f"✅  index.html gerado com {len(rows):,} títulos — {len(html)//1024} KB")
