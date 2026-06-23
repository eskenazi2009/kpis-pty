# -*- coding: utf-8 -*-
"""
Build a standalone, offline KPI dashboard (dashboard.html) from the
Sport Factory Panama yearly KPI workbooks.

Reads every "SPORT FACTORY_KPI <year>.xlsx" file in this folder, parses the
month/store/KPI grid, and writes a single self-contained HTML file with the
data embedded as JSON and charts drawn with hand-rolled inline SVG (no CDN,
fully offline).

Re-run this script whenever the spreadsheets change:
    python build_dashboard.py
"""

import json
import glob
import os
import re
import base64
import getpass
import secrets
import openpyxl
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

HERE = os.path.dirname(os.path.abspath(__file__))
PBKDF2_ITERATIONS = 200000  # must match the value used in the browser


def encrypt_payload(plaintext, passphrase):
    """Encrypt with AES-256-GCM using a PBKDF2-SHA256 derived key.
    Layout is byte-for-byte compatible with the browser Web Crypto API."""
    salt = secrets.token_bytes(16)
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32,
                     salt=salt, iterations=PBKDF2_ITERATIONS)
    key = kdf.derive(passphrase.encode("utf-8"))
    iv = secrets.token_bytes(12)
    ct = AESGCM(key).encrypt(iv, plaintext.encode("utf-8"), None)  # ciphertext||tag
    b64 = lambda b: base64.b64encode(b).decode("ascii")
    return {"salt": b64(salt), "iv": b64(iv), "ct": b64(ct), "iter": PBKDF2_ITERATIONS}

MONTHS = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
          "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
MONTH_SET = set(MONTHS)
KPI_ORDER = ["ventas", "uds", "fact", "upt", "vpt"]  # column order inside a month block
BRAND_HEADERS = {"TNB": "NB", "TRB": "RB"}


def num(v):
    """Return a float for a real numeric cell, else None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def parse_workbook(path):
    """Parse one workbook -> { store_name: { MONTH: {ventas,uds,fact,upt,vpt} } }."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    # 1) Find month start columns from any row that holds month names (row 3).
    month_cols = {}
    for r in range(1, min(ws.max_row, 6) + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if isinstance(val, str) and val.strip().upper() in MONTH_SET:
                month_cols[val.strip().upper()] = c
        if month_cols:
            break
    # keep only the 12 calendar months (ignore ACUM / anything else)
    month_cols = {m: c for m, c in month_cols.items() if m in MONTH_SET}

    # 2) Walk rows; track current brand from TNB/TRB header rows.
    stores = {}
    brand = None
    for r in range(1, ws.max_row + 1):
        raw = ws.cell(r, 1).value
        label = (str(raw).strip() if raw is not None else "")
        if not label:
            continue
        up = label.upper()
        if up in BRAND_HEADERS:          # brand header / KPI sub-header row
            brand = BRAND_HEADERS[up]
            continue
        if up == "SPORT FACTORY" or up == "TNB" or up == "TRB":
            continue
        if up == "TOTAL":
            store = (brand or "?") + " TOTAL"
        else:
            store = label

        # only treat as a data row if it actually has KPI numbers
        months = {}
        for m, base in month_cols.items():
            ventas = num(ws.cell(r, base).value)
            if ventas is None or ventas == 0:
                continue
            rec = {
                "ventas": ventas,
                "uds":    num(ws.cell(r, base + 1).value),
                "fact":   num(ws.cell(r, base + 2).value),
                "upt":    num(ws.cell(r, base + 3).value),
                "vpt":    num(ws.cell(r, base + 4).value),
            }
            months[m] = rec
        if months:
            stores[store] = months
    return stores


def year_from_name(path):
    m = re.search(r"(20\d{2})", os.path.basename(path))
    return m.group(1) if m else os.path.splitext(os.path.basename(path))[0]


def build_data():
    files = sorted(glob.glob(os.path.join(HERE, "SPORT FACTORY_KPI *.xlsx")))
    data = {}
    for f in files:
        if os.path.basename(f).startswith("~$"):
            continue
        data[year_from_name(f)] = parse_workbook(f)
    return data


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sport Factory Panamá — KPI Dashboard</title>
<style>
  :root{
    --bg:#0f1419; --panel:#1a2230; --panel2:#222c3d; --line:#2e3a4d;
    --text:#e8edf4; --muted:#93a1b5; --accent:#4da3ff; --accent2:#ff8c42;
    --good:#36d399; --bad:#f87272;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--text);
       font-family:Segoe UI,Roboto,Helvetica,Arial,sans-serif;font-size:14px}
  header{padding:18px 24px;border-bottom:1px solid var(--line);
         display:flex;align-items:baseline;gap:14px;flex-wrap:wrap}
  header h1{font-size:20px;margin:0;font-weight:600}
  header .sub{color:var(--muted);font-size:13px}
  .wrap{padding:18px 24px;max-width:1280px;margin:0 auto}
  .controls{display:flex;gap:18px;flex-wrap:wrap;align-items:flex-end;margin-bottom:16px}
  .ctl{display:flex;flex-direction:column;gap:5px}
  .ctl label{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}
  select{background:var(--panel2);color:var(--text);border:1px solid var(--line);
         border-radius:8px;padding:8px 12px;font-size:14px;min-width:170px}
  .stores-wrap{margin-bottom:20px}
  .stores-wrap .lab{font-size:12px;color:var(--muted);text-transform:uppercase;
                    letter-spacing:.04em;margin-bottom:7px}
  .stores{display:flex;flex-wrap:wrap;gap:8px}
  .sb{background:var(--panel2);border:1px solid var(--line);color:var(--text);
      border-radius:8px;padding:8px 14px;cursor:pointer;font-size:13px;transition:.12s}
  .sb:hover{border-color:var(--accent)}
  .sb.active{background:var(--accent);border-color:var(--accent);color:#04263f;font-weight:700}
  .sb.total{border-style:dashed}
  .cards{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:22px}
  @media(max-width:880px){.cards{grid-template-columns:repeat(2,1fr)}}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px}
  .card .k{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}
  .card .v{font-size:22px;font-weight:700;margin-top:6px}
  .card .prev{font-size:12px;color:var(--muted);margin-top:2px}
  .delta{font-size:13px;font-weight:600;margin-top:6px}
  .delta.up{color:var(--good)} .delta.down{color:var(--bad)} .delta.flat{color:var(--muted)}
  .card.sel{outline:2px solid var(--accent)}
  .card.click{cursor:pointer}
  .panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;
         padding:16px 18px;margin-bottom:22px}
  .panel h2{font-size:15px;margin:0 0 14px;font-weight:600}
  .panel h2 .hint{color:var(--muted);font-weight:400;font-size:12px;margin-left:8px}
  .chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:22px}
  @media(max-width:980px){.chart-grid{grid-template-columns:1fr}}
  .chartbox{position:relative}
  svg{width:100%;height:auto;display:block}
  .hz{cursor:crosshair}
  .legend{display:flex;gap:18px;margin-top:8px;font-size:12px;color:var(--muted);flex-wrap:wrap}
  .legend i{display:inline-block;width:11px;height:11px;border-radius:3px;margin-right:6px;vertical-align:-1px}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{padding:9px 12px;text-align:right;border-bottom:1px solid var(--line)}
  th:first-child,td:first-child{text-align:left}
  th{color:var(--muted);font-weight:600;cursor:pointer;user-select:none;text-transform:uppercase;
     font-size:11px;letter-spacing:.04em;white-space:nowrap}
  th.sortdir::after{content:" \25BE";opacity:.7}
  th.sortdir.asc::after{content:" \25B4"}
  tbody tr:hover{background:var(--panel2)}
  td.pos{color:var(--good)} td.neg{color:var(--bad)}
  .foot{color:var(--muted);font-size:12px;margin-top:8px}
  .tag{display:inline-block;background:var(--panel2);border:1px solid var(--line);
       border-radius:999px;padding:2px 10px;font-size:12px;color:var(--muted)}
  #tip{position:fixed;display:none;pointer-events:none;z-index:50;background:#0b0f15;
       border:1px solid var(--line);border-radius:8px;padding:8px 10px;font-size:12px;
       color:var(--text);box-shadow:0 6px 20px rgba(0,0,0,.45);min-width:130px}
  #tip .tt{color:var(--muted);font-size:11px;margin-bottom:5px}
  #tip .row{display:flex;justify-content:space-between;gap:16px;line-height:1.55}
  #tip .row i{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:6px}
  #lock{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;
        background:var(--bg);z-index:100;padding:20px}
  .lockbox{background:var(--panel);border:1px solid var(--line);border-radius:16px;
           padding:32px 28px;max-width:380px;width:100%;text-align:center}
  .lockbox .ico{font-size:34px;line-height:1}
  .lockbox h2{font-size:17px;margin:14px 0 4px;font-weight:600}
  .lockbox p{color:var(--muted);font-size:13px;margin:0 0 18px}
  .lockbox input{width:100%;background:var(--panel2);border:1px solid var(--line);color:var(--text);
                 border-radius:9px;padding:11px 13px;font-size:15px;margin-bottom:12px}
  .lockbox input:focus{outline:none;border-color:var(--accent)}
  .lockbox button{width:100%;background:var(--accent);color:#04263f;border:none;border-radius:9px;
                  padding:11px;font-size:15px;font-weight:700;cursor:pointer}
  .lockbox button:disabled{opacity:.6;cursor:default}
  .lockerr{color:var(--bad);font-size:13px;margin-top:12px;display:none}
</style>
</head>
<body>
<header>
  <h1>Sport Factory Panamá</h1>
  <span class="sub">KPI Dashboard</span>
  <span class="tag" id="yearsTag"></span>
</header>
<div id="lock">
  <div class="lockbox">
    <div class="ico">🔒</div>
    <h2>Sport Factory Panamá — KPI Dashboard</h2>
    <p>Ingresa la contraseña para ver el tablero.</p>
    <input id="pw" type="password" placeholder="Contraseña" autocomplete="off" autocapitalize="off" spellcheck="false">
    <button id="unlockBtn">Entrar</button>
    <div id="lockErr" class="lockerr"></div>
  </div>
</div>

<div class="wrap" id="app" style="display:none">
  <div class="controls">
    <div class="ctl">
      <label for="kpi">Indicador (KPI)</label>
      <select id="kpi"></select>
    </div>
    <div class="ctl">
      <label for="cmp">Comparar años</label>
      <select id="cmp"></select>
    </div>
  </div>

  <div class="stores-wrap">
    <div class="lab">Tienda</div>
    <div class="stores" id="stores"></div>
  </div>

  <div class="cards" id="cards"></div>

  <div class="chart-grid">
    <div class="panel">
      <h2>Tendencia mensual <span class="hint" id="trendHint"></span></h2>
      <div class="chartbox" id="trend"></div>
      <div class="legend" id="trendLegend"></div>
    </div>
    <div class="panel">
      <h2>Comparación año contra año <span class="hint" id="yoyHint"></span></h2>
      <div class="chartbox" id="yoy"></div>
      <div class="legend" id="yoyLegend"></div>
    </div>
  </div>

  <div class="panel">
    <h2>Todos los KPIs juntos <span class="hint" id="allHint"></span></h2>
    <div class="chartbox" id="allk"></div>
    <div class="legend" id="allLegend"></div>
    <div class="foot">Cada KPI se escala a su propio máximo del año (0–100%) para poder verse en el mismo gráfico. Pasa el mouse para ver los valores reales.</div>
  </div>

  <div class="panel">
    <h2>Comparación por tienda <span class="hint" id="tableHint"></span></h2>
    <table id="tbl">
      <thead><tr id="thr"></tr></thead>
      <tbody id="tb"></tbody>
    </table>
    <div class="foot" id="tblFoot"></div>
  </div>

  <div class="foot">Generado desde los archivos .xlsx · funciona sin conexión.</div>
</div>
<div id="tip"></div>

<script>
const ENCRYPTED = /*__ENC__*/;
let DATA = {};
let YEARS = [];
const MONTHS = /*__MONTHS__*/;
const MONTH_LABEL = {ENE:"Ene",FEB:"Feb",MAR:"Mar",ABR:"Abr",MAY:"May",JUN:"Jun",
                     JUL:"Jul",AGO:"Ago",SEP:"Sep",OCT:"Oct",NOV:"Nov",DIC:"Dic"};
const KPIS = [
  {id:"ventas", label:"VENTAS", desc:"Ingreso total ($)",     fmt:"money", color:"#4da3ff"},
  {id:"uds",    label:"UDS",    desc:"Unidades vendidas",     fmt:"int",   color:"#ff8c42"},
  {id:"fact",   label:"FACT",   desc:"Cantidad de facturas",  fmt:"int",   color:"#36d399"},
  {id:"upt",    label:"UPT",    desc:"Unidades por factura",  fmt:"dec2",  color:"#c084fc"},
  {id:"vpt",    label:"VPT",    desc:"Valor ($) por factura", fmt:"money", color:"#fbbf24"},
];
const KMAP = Object.fromEntries(KPIS.map(k=>[k.id,k]));
const COL = {a:"#4da3ff", b:"#ff8c42"}; // older year, newer year

// ---- formatting -----------------------------------------------------------
function fmtVal(v, kind){
  if(v===null||v===undefined||isNaN(v)) return "—";
  if(kind==="money") return "$"+v.toLocaleString("en-US",{maximumFractionDigits:0});
  if(kind==="int")   return Math.round(v).toLocaleString("en-US");
  if(kind==="dec2")  return v.toFixed(2);
  return String(v);
}
function pct(a,b){ // change from a(old) -> b(new)
  if(a===null||b===null||a===0||a===undefined||b===undefined) return null;
  return (b-a)/a*100;
}
function deltaHtml(p){
  if(p===null) return '<div class="delta flat">—</div>';
  const cls = p>0.05?"up":(p<-0.05?"down":"flat");
  const arrow = p>0.05?"▲":(p<-0.05?"▼":"▬");
  return `<div class="delta ${cls}">${arrow} ${p>=0?"+":""}${p.toFixed(1)}%</div>`;
}

// ---- aggregation ----------------------------------------------------------
function monthsWithData(year, store){
  const s = (DATA[year]||{})[store]||{};
  return MONTHS.filter(m=>s[m]);
}
function commonMonths(store, yA, yB){
  const a=new Set(monthsWithData(yA,store)), b=new Set(monthsWithData(yB,store));
  return MONTHS.filter(m=>a.has(m)&&b.has(m));
}
function ytd(year, store, kpi, months){
  const s=(DATA[year]||{})[store]||{};
  let ventas=0,uds=0,fact=0,n=0;
  for(const m of months){ if(!s[m]) continue;
    ventas+=s[m].ventas||0; uds+=s[m].uds||0; fact+=s[m].fact||0; n++; }
  if(!n) return null;
  switch(kpi){
    case "ventas": return ventas;
    case "uds":    return uds;
    case "fact":   return fact;
    case "upt":    return fact? uds/fact : null;
    case "vpt":    return fact? ventas/fact : null;
  }
}

// ---- state ----------------------------------------------------------------
const state = {kpi:"ventas", store:null, yA:null, yB:null,
               sortKey:"db", sortDir:-1};

function allStores(){
  const set=[];
  for(const y of YEARS) for(const s of Object.keys(DATA[y])) if(!set.includes(s)) set.push(s);
  return set;
}

// ---- tooltip --------------------------------------------------------------
function tipHtml(title, rows){
  return `<div class="tt">${title}</div>`+rows.map(r=>
    `<div class="row"><span>${r.color?`<i style="background:${r.color}"></i>`:""}${r.label}</span><b>${r.val}</b></div>`).join("");
}
function showTip(html, ev){
  const t=document.getElementById("tip");
  t.innerHTML=html; t.style.display="block";
  const pad=14; let x=ev.clientX+pad, y=ev.clientY+pad;
  const r=t.getBoundingClientRect();
  if(x+r.width>window.innerWidth)  x=ev.clientX-r.width-pad;
  if(y+r.height>window.innerHeight) y=ev.clientY-r.height-pad;
  t.style.left=x+"px"; t.style.top=y+"px";
}
function hideTip(){ document.getElementById("tip").style.display="none"; }

function attachHover(container, model){
  const svg=container.querySelector("svg");
  const ns="http://www.w3.org/2000/svg";
  const guide=document.createElementNS(ns,"line");
  guide.setAttribute("stroke","#93a1b5"); guide.setAttribute("stroke-width","1");
  guide.setAttribute("stroke-dasharray","3 3"); guide.style.display="none";
  guide.setAttribute("y1",model.guideTop); guide.setAttribute("y2",model.guideBottom);
  svg.appendChild(guide);
  container.querySelectorAll(".hz").forEach(z=>{
    const i=+z.dataset.i;
    z.addEventListener("mouseenter",()=>{ guide.style.display="";
      guide.setAttribute("x1",model.xs[i]); guide.setAttribute("x2",model.xs[i]); });
    z.addEventListener("mousemove",e=>showTip(model.months[i].html,e));
    z.addEventListener("mouseleave",()=>{ guide.style.display="none"; hideTip(); });
  });
}

// ---- SVG helpers ----------------------------------------------------------
function svgEl(w,h){ return {w,h,parts:[],
  add(s){this.parts.push(s);},
  out(){return `<svg viewBox="0 0 ${this.w} ${this.h}" preserveAspectRatio="xMidYMid meet">${this.parts.join("")}</svg>`;}};}
function txt(x,y,s,opt={}){const a=opt.anchor||"middle";const fs=opt.size||11;
  const col=opt.color||"#93a1b5";const w=opt.weight||"400";
  return `<text x="${x}" y="${y}" text-anchor="${a}" font-size="${fs}" fill="${col}" font-weight="${w}">${s}</text>`;}
function niceMax(v){
  if(v<=0) return 1;
  const pow=Math.pow(10,Math.floor(Math.log10(v)));
  const n=v/pow; const step = n<=1?1:n<=2?2:n<=5?5:10;
  return step*pow;
}

// ---- KPI cards ------------------------------------------------------------
function renderCards(){
  const wrap=document.getElementById("cards"); wrap.innerHTML="";
  const cm = commonMonths(state.store, state.yA, state.yB);
  const periodLbl = cm.length ? `${MONTH_LABEL[cm[0]]}–${MONTH_LABEL[cm[cm.length-1]]}` : "—";
  KPIS.forEach(k=>{
    const a=ytd(state.yA, state.store, k.id, cm);
    const b=ytd(state.yB, state.store, k.id, cm);
    const p=pct(a,b);
    const div=document.createElement("div");
    div.className="card click"+(k.id===state.kpi?" sel":"");
    div.innerHTML=`<div class="k">${k.label}</div>
      <div class="v">${fmtVal(b,k.fmt)}</div>
      <div class="prev">${state.yB} · ${periodLbl}</div>
      ${deltaHtml(p)}
      <div class="prev">vs ${state.yA}: ${fmtVal(a,k.fmt)}</div>`;
    div.onclick=()=>{state.kpi=k.id; render();};
    wrap.appendChild(div);
  });
}

// ---- trend (line, single KPI, both years) ---------------------------------
function renderTrend(){
  const k=KMAP[state.kpi];
  document.getElementById("trendHint").textContent = `${k.label} · ${state.store}`;
  const W=560,H=300,PL=64,PR=16,PT=16,PB=34;
  const sA=(DATA[state.yA]||{})[state.store]||{};
  const sB=(DATA[state.yB]||{})[state.store]||{};
  const vals=[];
  MONTHS.forEach(m=>{ if(sA[m]&&sA[m][k.id]!=null) vals.push(sA[m][k.id]);
                      if(sB[m]&&sB[m][k.id]!=null) vals.push(sB[m][k.id]); });
  const max = niceMax(vals.length?Math.max(...vals):1);
  const plotW=W-PL-PR, step=plotW/(MONTHS.length-1);
  const x=i=> PL + i*step;
  const y=v=> H-PB - (v/max)*(H-PT-PB);
  const g=svgEl(W,H);
  for(let t=0;t<=4;t++){const yy=PT+(H-PT-PB)*t/4; const val=max*(1-t/4);
    g.add(`<line x1="${PL}" y1="${yy}" x2="${W-PR}" y2="${yy}" stroke="#2e3a4d" stroke-width="1"/>`);
    g.add(txt(PL-8,yy+4,fmtVal(val,k.fmt),{anchor:"end",size:10}));}
  MONTHS.forEach((m,i)=>g.add(txt(x(i),H-PB+18,MONTH_LABEL[m],{size:10})));
  function line(store,color){
    let d="",pts="",started=false;
    MONTHS.forEach((m,i)=>{const rec=store[m]; if(rec&&rec[k.id]!=null){
      const X=x(i),Y=y(rec[k.id]); d+=(started?"L":"M")+X+" "+Y+" "; started=true;
      pts+=`<circle cx="${X}" cy="${Y}" r="3.2" fill="${color}"/>`;}});
    if(d) g.add(`<path d="${d}" fill="none" stroke="${color}" stroke-width="2.4"/>`);
    g.add(pts);
  }
  line(sA,COL.a); line(sB,COL.b);
  // hover model
  const model={xs:MONTHS.map((m,i)=>x(i)), guideTop:PT, guideBottom:H-PB, months:[]};
  MONTHS.forEach((m,i)=>{
    const rows=[];
    if(sA[m]&&sA[m][k.id]!=null) rows.push({color:COL.a,label:state.yA,val:fmtVal(sA[m][k.id],k.fmt)});
    if(sB[m]&&sB[m][k.id]!=null) rows.push({color:COL.b,label:state.yB,val:fmtVal(sB[m][k.id],k.fmt)});
    if(rows.length){
      model.months[i]={html:tipHtml(`${MONTH_LABEL[m]} · ${k.label}`,rows)};
      const zx=Math.max(PL,x(i)-step/2), zw=Math.min(step,W-PR-zx);
      g.add(`<rect class="hz" data-i="${i}" x="${zx}" y="${PT}" width="${zw}" height="${H-PT-PB}" fill="#000" fill-opacity="0" pointer-events="all"/>`);
    } else model.months[i]=null;
  });
  const c=document.getElementById("trend"); c.innerHTML=g.out(); attachHover(c,model);
  document.getElementById("trendLegend").innerHTML=
    `<span><i style="background:${COL.a}"></i>${state.yA}</span>
     <span><i style="background:${COL.b}"></i>${state.yB}</span>`;
}

// ---- YoY (grouped bars) ---------------------------------------------------
function renderYoY(){
  const k=KMAP[state.kpi];
  document.getElementById("yoyHint").textContent = `${k.label} · ${state.store}`;
  const cm = commonMonths(state.store, state.yA, state.yB);
  const W=560,H=300,PL=64,PR=16,PT=16,PB=34;
  const g=svgEl(W,H);
  if(!cm.length){ g.add(txt(W/2,H/2,"Sin meses comparables",{size:13}));
    document.getElementById("yoy").innerHTML=g.out();
    document.getElementById("yoyLegend").innerHTML=""; return; }
  const sA=DATA[state.yA][state.store], sB=DATA[state.yB][state.store];
  const vals=[]; cm.forEach(m=>{vals.push(sA[m][k.id],sB[m][k.id]);});
  const max=niceMax(Math.max(...vals));
  const y=v=> H-PB - (v/max)*(H-PT-PB);
  const band=(W-PL-PR)/cm.length, bw=Math.min(26,band/3);
  for(let t=0;t<=4;t++){const yy=PT+(H-PT-PB)*t/4; const val=max*(1-t/4);
    g.add(`<line x1="${PL}" y1="${yy}" x2="${W-PR}" y2="${yy}" stroke="#2e3a4d" stroke-width="1"/>`);
    g.add(txt(PL-8,yy+4,fmtVal(val,k.fmt),{anchor:"end",size:10}));}
  const model={xs:[], guideTop:PT, guideBottom:H-PB, months:[]};
  cm.forEach((m,i)=>{
    const cx=PL+band*i+band/2, va=sA[m][k.id], vb=sB[m][k.id];
    g.add(`<rect x="${cx-bw-2}" y="${y(va)}" width="${bw}" height="${H-PB-y(va)}" fill="${COL.a}" rx="2"/>`);
    g.add(`<rect x="${cx+2}" y="${y(vb)}" width="${bw}" height="${H-PB-y(vb)}" fill="${COL.b}" rx="2"/>`);
    g.add(txt(cx,H-PB+18,MONTH_LABEL[m],{size:10}));
    model.xs.push(cx);
    const p=pct(va,vb);
    model.months.push({html:tipHtml(`${MONTH_LABEL[m]} · ${k.label}`,[
      {color:COL.a,label:state.yA,val:fmtVal(va,k.fmt)},
      {color:COL.b,label:state.yB,val:fmtVal(vb,k.fmt)},
      {color:"",label:"Δ",val:(p===null?"—":(p>=0?"+":"")+p.toFixed(1)+"%")},
    ])});
    g.add(`<rect class="hz" data-i="${i}" x="${PL+band*i}" y="${PT}" width="${band}" height="${H-PT-PB}" fill="#000" fill-opacity="0" pointer-events="all"/>`);
  });
  const c=document.getElementById("yoy"); c.innerHTML=g.out(); attachHover(c,model);
  document.getElementById("yoyLegend").innerHTML=
    `<span><i style="background:${COL.a}"></i>${state.yA}</span>
     <span><i style="background:${COL.b}"></i>${state.yB}</span>`;
}

// ---- all KPIs overlay (normalized) ----------------------------------------
function renderAllKpi(){
  const year=state.yB, s=(DATA[year]||{})[state.store]||{};
  document.getElementById("allHint").textContent = `${state.store} · ${year} (normalizado)`;
  const W=1140,H=320,PL=48,PR=16,PT=16,PB=34;
  const plotW=W-PL-PR, step=plotW/(MONTHS.length-1);
  const x=i=> PL + i*step;
  const y=pc=> H-PB - (pc/100)*(H-PT-PB);
  const g=svgEl(W,H);
  for(let t=0;t<=4;t++){const yy=PT+(H-PT-PB)*t/4; const val=100*(1-t/4);
    g.add(`<line x1="${PL}" y1="${yy}" x2="${W-PR}" y2="${yy}" stroke="#2e3a4d" stroke-width="1"/>`);
    g.add(txt(PL-8,yy+4,val+"%",{anchor:"end",size:10}));}
  MONTHS.forEach((m,i)=>g.add(txt(x(i),H-PB+18,MONTH_LABEL[m],{size:10})));
  // per-KPI max for normalization
  const maxes={};
  KPIS.forEach(k=>{ let mx=0; MONTHS.forEach(m=>{ if(s[m]&&s[m][k.id]!=null) mx=Math.max(mx,s[m][k.id]); }); maxes[k.id]=mx||1; });
  KPIS.forEach(k=>{
    let d="",pts="",started=false;
    MONTHS.forEach((m,i)=>{ if(s[m]&&s[m][k.id]!=null){
      const X=x(i),Y=y(s[m][k.id]/maxes[k.id]*100);
      d+=(started?"L":"M")+X+" "+Y+" "; started=true;
      pts+=`<circle cx="${X}" cy="${Y}" r="2.8" fill="${k.color}"/>`;}});
    if(d) g.add(`<path d="${d}" fill="none" stroke="${k.color}" stroke-width="2.2"/>`);
    g.add(pts);
  });
  const model={xs:MONTHS.map((m,i)=>x(i)), guideTop:PT, guideBottom:H-PB, months:[]};
  MONTHS.forEach((m,i)=>{
    if(s[m]){
      const rows=KPIS.filter(k=>s[m][k.id]!=null).map(k=>({color:k.color,label:k.label,val:fmtVal(s[m][k.id],k.fmt)}));
      model.months[i]={html:tipHtml(`${MONTH_LABEL[m]} · ${year}`,rows)};
      const zx=Math.max(PL,x(i)-step/2), zw=Math.min(step,W-PR-zx);
      g.add(`<rect class="hz" data-i="${i}" x="${zx}" y="${PT}" width="${zw}" height="${H-PT-PB}" fill="#000" fill-opacity="0" pointer-events="all"/>`);
    } else model.months[i]=null;
  });
  const c=document.getElementById("allk"); c.innerHTML=g.out(); attachHover(c,model);
  document.getElementById("allLegend").innerHTML=
    KPIS.map(k=>`<span><i style="background:${k.color}"></i>${k.label}</span>`).join("");
}

// ---- store buttons --------------------------------------------------------
function renderStoreButtons(){
  const wrap=document.getElementById("stores"); wrap.innerHTML="";
  allStores().forEach(s=>{
    const b=document.createElement("button");
    b.className="sb"+(s===state.store?" active":"")+(s.includes("TOTAL")?" total":"");
    b.textContent=s;
    b.onclick=()=>{state.store=s; render();};
    wrap.appendChild(b);
  });
}

// ---- store comparison table ----------------------------------------------
function renderTable(){
  const k=KMAP[state.kpi];
  const thr=document.getElementById("thr"); thr.innerHTML="";
  const cols=[
    {key:"store", label:"Tienda"},
    {key:"da",    label:state.yA},
    {key:"db",    label:state.yB},
    {key:"delta", label:"Δ%"},
  ];
  cols.forEach(c=>{
    const th=document.createElement("th"); th.textContent=c.label; th.dataset.key=c.key;
    if(c.key===state.sortKey){th.classList.add("sortdir"); if(state.sortDir>0)th.classList.add("asc");}
    th.onclick=()=>{ if(state.sortKey===c.key) state.sortDir*=-1;
      else {state.sortKey=c.key; state.sortDir=(c.key==="store")?1:-1;} render(); };
    thr.appendChild(th);
  });
  const rows=allStores().map(s=>{
    const cm=commonMonths(s,state.yA,state.yB);
    const da=ytd(state.yA,s,k.id,cm), db=ytd(state.yB,s,k.id,cm);
    return {store:s, da, db, delta:pct(da,db)};
  });
  const sk=state.sortKey, dir=state.sortDir;
  rows.sort((a,b)=>{ let x=a[sk], y=b[sk];
    if(sk==="store") return dir*String(x).localeCompare(String(y));
    if(x===null) return 1; if(y===null) return -1; return dir*(x-y); });
  const tb=document.getElementById("tb"); tb.innerHTML="";
  rows.forEach(r=>{
    const tr=document.createElement("tr");
    const dcls = r.delta===null?"":(r.delta>0?"pos":(r.delta<0?"neg":""));
    const dtxt = r.delta===null?"—":`${r.delta>=0?"+":""}${r.delta.toFixed(1)}%`;
    tr.innerHTML=`<td>${r.store}${r.store===state.store?' &#9679;':''}</td>
      <td>${fmtVal(r.da,k.fmt)}</td><td>${fmtVal(r.db,k.fmt)}</td>
      <td class="${dcls}">${dtxt}</td>`;
    tr.style.cursor="pointer";
    tr.onclick=()=>{state.store=r.store; render();};
    tb.appendChild(tr);
  });
  const cm=commonMonths(state.store,state.yA,state.yB);
  const per = cm.length?`${MONTH_LABEL[cm[0]]}–${MONTH_LABEL[cm[cm.length-1]]}`:"—";
  document.getElementById("tableHint").textContent=`${k.label} · acumulado comparable (${per})`;
  document.getElementById("tblFoot").textContent=
    `Los acumulados (YTD) comparan solo los meses presentes en ambos años. UPT y VPT se recalculan sobre los acumulados, no se promedian.`;
}

// ---- controls -------------------------------------------------------------
function initControls(){
  const kpiSel=document.getElementById("kpi");
  KPIS.forEach(k=>kpiSel.add(new Option(`${k.label} — ${k.desc}`, k.id)));
  kpiSel.value=state.kpi;
  kpiSel.onchange=e=>{state.kpi=e.target.value; render();};

  const cmp=document.getElementById("cmp");
  if(YEARS.length>=2){
    for(let i=0;i<YEARS.length;i++) for(let j=i+1;j<YEARS.length;j++)
      cmp.add(new Option(`${YEARS[i]} vs ${YEARS[j]}`, YEARS[i]+"|"+YEARS[j]));
    cmp.value = state.yA+"|"+state.yB;
    cmp.onchange=e=>{const [a,b]=e.target.value.split("|"); state.yA=a; state.yB=b; render();};
  } else { cmp.add(new Option(YEARS[0], YEARS[0])); cmp.disabled=true; }

  state.store = allStores()[0];
  document.getElementById("yearsTag").textContent = YEARS.join(" · ");
}

function render(){
  renderStoreButtons();
  renderCards();
  renderTrend();
  renderYoY();
  renderAllKpi();
  renderTable();
}

// ---- decryption / unlock --------------------------------------------------
async function decryptData(passphrase){
  const dec = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
  const salt = dec(ENCRYPTED.salt), iv = dec(ENCRYPTED.iv), ct = dec(ENCRYPTED.ct);
  const baseKey = await crypto.subtle.importKey(
    "raw", new TextEncoder().encode(passphrase), "PBKDF2", false, ["deriveKey"]);
  const key = await crypto.subtle.deriveKey(
    {name:"PBKDF2", salt, iterations:ENCRYPTED.iter, hash:"SHA-256"},
    baseKey, {name:"AES-GCM", length:256}, false, ["decrypt"]);
  const plain = await crypto.subtle.decrypt({name:"AES-GCM", iv}, key, ct);
  return JSON.parse(new TextDecoder().decode(plain));
}

function boot(data){
  DATA = data;
  YEARS = Object.keys(DATA).sort();
  state.yA = YEARS[0];
  state.yB = YEARS[YEARS.length-1];
  initControls();
  render();
  document.getElementById("lock").style.display = "none";
  document.getElementById("app").style.display = "";
}

async function unlock(){
  const inp = document.getElementById("pw");
  const err = document.getElementById("lockErr");
  const btn = document.getElementById("unlockBtn");
  if(!window.crypto || !crypto.subtle){
    err.textContent = "Este navegador no permite descifrar desde un archivo local. Abre el enlace web (https).";
    err.style.display = "block"; return;
  }
  btn.disabled = true; err.style.display = "none";
  try{
    const data = await decryptData(inp.value);
    boot(data);
  }catch(e){
    err.textContent = "Contraseña incorrecta.";
    err.style.display = "block"; inp.select();
  }finally{ btn.disabled = false; }
}

document.getElementById("unlockBtn").onclick = unlock;
document.getElementById("pw").addEventListener("keydown", e=>{ if(e.key==="Enter") unlock(); });
document.getElementById("pw").focus();
</script>
</body>
</html>
"""


def main():
    data = build_data()
    if not data:
        raise SystemExit("No 'SPORT FACTORY_KPI *.xlsx' files found in this folder.")

    # Passphrase: from $DASH_PASSPHRASE, else prompt (hidden input).
    passphrase = os.environ.get("DASH_PASSPHRASE")
    if not passphrase:
        passphrase = getpass.getpass("Contraseña para el tablero: ")
    if not passphrase:
        raise SystemExit("A passphrase is required to encrypt the dashboard.")

    enc = encrypt_payload(json.dumps(data, ensure_ascii=False), passphrase)
    html = (HTML_TEMPLATE
            .replace("/*__ENC__*/", json.dumps(enc))
            .replace("/*__MONTHS__*/", json.dumps(MONTHS)))
    # Write both dashboard.html (local use) and index.html (served by GitHub Pages).
    for name in ("dashboard.html", "index.html"):
        with open(os.path.join(HERE, name), "w", encoding="utf-8") as f:
            f.write(html)

    print("Wrote dashboard.html and index.html")
    for y in sorted(data):
        stores = data[y]
        months = sorted({m for s in stores.values() for m in s}, key=MONTHS.index)
        print(f"  {y}: {len(stores)} filas, meses con datos: {', '.join(months) or '—'}")


if __name__ == "__main__":
    main()
