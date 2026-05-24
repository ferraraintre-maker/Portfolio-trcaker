import { useState, useCallback } from "react";
import { PieChart, Pie, Cell, ResponsiveContainer, LineChart, Line, XAxis, YAxis, Tooltip } from "recharts";
import * as XLSX from "xlsx";

const TYPE_LIST = ["Azioni","Obbligazioni","Materie Prime","Liquidità"];
const GEO_LIST  = ["Nord America","Europa","Asia Pac.","Mercati Em.","Globale","Altro"];
const C  = { Azioni:"#00D4FF", Obbligazioni:"#FF6B35", "Materie Prime":"#7B61FF", Liquidità:"#00FF9F" };
const GC = { "Nord America":"#00D4FF", Europa:"#FF6B35", "Asia Pac.":"#7B61FF", "Mercati Em.":"#AAFFCC", Globale:"#FFD700", Altro:"#FF8080" };

const INIT_HOLDINGS = [
  { id:1, name:"MSCI World ETF", ticker:"IWDA", type:"Azioni",        geo:"Nord America", value:12840, buyPrice:11200, quantity:45 },
  { id:2, name:"S&P 500 ETF",    ticker:"VUSA", type:"Azioni",        geo:"Nord America", value:9320,  buyPrice:8500,  quantity:62 },
  { id:3, name:"BTP Italia",     ticker:"BTP27",type:"Obbligazioni",  geo:"Europa",       value:7200,  buyPrice:7000,  quantity:70 },
  { id:4, name:"iShares Corp",   ticker:"LQDE", type:"Obbligazioni",  geo:"Europa",       value:5846,  buyPrice:5600,  quantity:48 },
  { id:5, name:"Gold ETC",       ticker:"PHAU", type:"Materie Prime", geo:"Globale",      value:4890,  buyPrice:4200,  quantity:30 },
  { id:6, name:"Emerging ETF",   ticker:"VFEM", type:"Azioni",        geo:"Mercati Em.",  value:3920,  buyPrice:3700,  quantity:55 },
  { id:7, name:"Conto Deposito", ticker:"CASH", type:"Liquidità",     geo:"Europa",       value:4324,  buyPrice:4324,  quantity:1  },
];
const INIT_TARGETS = {
  type: { Azioni:55, Obbligazioni:25, "Materie Prime":10, Liquidità:10 },
  geo:  { "Nord America":40, Europa:30, "Asia Pac.":15, "Mercati Em.":10, Globale:5, Altro:0 },
};

// ─── Excel export ────────────────────────────────────────────────
function doExport(holdings, targets) {
  const wb = XLSX.utils.book_new();
  const wsH = XLSX.utils.aoa_to_sheet([
    ["ID","Nome","Ticker","Tipo","Area Geo","Valore €","Prezzo Acquisto €","Quantità","G/P %"],
    ...holdings.map(h=>[
      h.id, h.name, h.ticker, h.type, h.geo, h.value, h.buyPrice, h.quantity,
      h.buyPrice ? +((h.value-h.buyPrice)/h.buyPrice*100).toFixed(2) : 0,
    ]),
  ]);
  wsH["!cols"]=[6,22,8,14,14,12,18,10,8].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, wsH, "Posizioni");

  const wsT = XLSX.utils.aoa_to_sheet([
    ["Asset Class","Target %"],
    ...Object.entries(targets.type).map(([k,v])=>[k,v]),
    ["TOTALE",{f:`SUM(B2:B${Object.keys(targets.type).length+1})`}],
  ]);
  wsT["!cols"]=[{wch:18},{wch:12}];
  XLSX.utils.book_append_sheet(wb, wsT, "Target Asset");

  const wsG = XLSX.utils.aoa_to_sheet([
    ["Area Geografica","Target %"],
    ...Object.entries(targets.geo).map(([k,v])=>[k,v]),
    ["TOTALE",{f:`SUM(B2:B${Object.keys(targets.geo).length+1})`}],
  ]);
  wsG["!cols"]=[{wch:18},{wch:12}];
  XLSX.utils.book_append_sheet(wb, wsG, "Target Geo");
  XLSX.writeFile(wb, "portafoglio.xlsx");
}

// ─── Excel import ────────────────────────────────────────────────
function doImport(file, cb) {
  const r = new FileReader();
  r.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result),{type:"array"});
      const rows = XLSX.utils.sheet_to_json(wb.Sheets["Posizioni"]||{});
      const holdings = rows.map((r,i)=>({
        id: r["ID"]||Date.now()+i, name:r["Nome"]||"", ticker:r["Ticker"]||"",
        type:r["Tipo"]||"Azioni", geo:r["Area Geo"]||"Europa",
        value:+r["Valore €"]||0, buyPrice:+r["Prezzo Acquisto €"]||0, quantity:+r["Quantità"]||0,
      }));
      const tRows = XLSX.utils.sheet_to_json(wb.Sheets["Target Asset"]||{});
      const gRows = XLSX.utils.sheet_to_json(wb.Sheets["Target Geo"]||{});
      const tType={}, tGeo={};
      tRows.forEach(r=>{ if(r["Asset Class"]&&r["Asset Class"]!=="TOTALE") tType[r["Asset Class"]]=+r["Target %"]||0; });
      gRows.forEach(r=>{ if(r["Area Geografica"]&&r["Area Geografica"]!=="TOTALE") tGeo[r["Area Geografica"]]=+r["Target %"]||0; });
      cb({ holdings, targets:{ type:Object.keys(tType).length?tType:INIT_TARGETS.type, geo:Object.keys(tGeo).length?tGeo:INIT_TARGETS.geo } });
    } catch(err){ alert("Errore: "+err.message); }
  };
  r.readAsArrayBuffer(file);
}

// ─── Stats ───────────────────────────────────────────────────────
function calcStats(h) {
  const total=h.reduce((s,x)=>s+x.value,0);
  const byType=TYPE_LIST.map(t=>{const v=h.filter(x=>x.type===t).reduce((s,x)=>s+x.value,0);return{name:t,value:total?Math.round(v/total*100):0,amount:v,color:C[t]};});
  const gm={}; h.forEach(x=>{gm[x.geo]=(gm[x.geo]||0)+x.value;});
  const byGeo=GEO_LIST.filter(g=>gm[g]).map(g=>({name:g,value:total?Math.round(gm[g]/total*100):0,color:GC[g]}));
  const cost=h.reduce((s,x)=>s+(x.buyPrice||0),0);
  return{total,byType,byGeo,gainPct:cost?((total-cost)/cost*100).toFixed(1):0};
}
function buildPerf(total) {
  const months=["Giu","Lug","Ago","Set","Ott","Nov","Dic","Gen","Feb","Mar","Apr","Mag"];
  let p=total*0.85,t=total*0.85;
  return months.map(m=>{p+=(Math.random()-0.42)*total*0.03;t+=total*0.85*0.008;return{month:m,portfolio:Math.round(p),target:Math.round(t)};});
}

// ─── UI primitives ───────────────────────────────────────────────
const bg="#070B11", card="linear-gradient(135deg,#0D1825,#0A1420)", border="1px solid #0F1E30";
function Card({ch,style={},onClick}){return <div onClick={onClick} style={{background:card,border,borderRadius:16,padding:16,marginBottom:12,cursor:onClick?"pointer":"default",...style}}>{ch}</div>;}
function Sec({t}){return <div style={{fontSize:10,color:"#2A4060",letterSpacing:3,textTransform:"uppercase",fontFamily:"'DM Mono',monospace",marginBottom:12,marginTop:4}}>{t}</div>;}
const IS={width:"100%",background:"#070B11",border:"1px solid #1A2A3A",borderRadius:10,color:"#E8F0FF",padding:"10px 12px",fontSize:13,fontFamily:"'DM Sans',sans-serif",outline:"none",boxSizing:"border-box"};
function Btn({label,onClick,color="#00D4FF",disabled=false,style={}}){
  return <button disabled={disabled} onClick={onClick} style={{border:"none",borderRadius:12,padding:"11px 0",fontSize:13,fontWeight:700,cursor:disabled?"not-allowed":"pointer",fontFamily:"'DM Sans',sans-serif",width:"100%",background:disabled?"#1A2A3A":color,color:disabled?"#3A5070":color==="#00D4FF"||color.startsWith("linear")?"#070B11":"#E8F0FF",...style}}>{label}</button>;
}
function TTip({active,payload}){if(!active||!payload?.length)return null;return <div style={{background:"#0D1117",border:"1px solid #1E2A3A",borderRadius:8,padding:"8px 12px"}}>{payload.map((p,i)=><div key={i} style={{color:p.color,fontSize:11,fontFamily:"monospace"}}>{p.name}: €{(+p.value).toLocaleString("it-IT")}</div>)}</div>;}

// ─── HoldingForm ─────────────────────────────────────────────────
function HoldingForm({initial,onSave,onClose,onDelete}){
  const blank={name:"",ticker:"",type:"Azioni",geo:"Europa",value:"",buyPrice:"",quantity:""};
  const [f,setF]=useState(initial||blank);
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const ok=f.name&&f.ticker&&+f.value>0&&+f.quantity>0;
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.88)",zIndex:300,display:"flex",alignItems:"flex-end",maxWidth:430,margin:"0 auto",left:"50%",transform:"translateX(-50%)"}}>
      <div style={{width:"100%",background:"#0D1825",borderRadius:"20px 20px 0 0",border:"1px solid #1A2A3A",padding:"20px 20px 40px",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
          <span style={{fontSize:16,fontWeight:700}}>{initial?"Modifica Posizione":"➕ Nuova Posizione"}</span>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#4A6080",fontSize:24,cursor:"pointer",lineHeight:1,padding:0}}>✕</button>
        </div>
        {[["Nome strumento","name","es. iShares MSCI World"],["Ticker / Codice","ticker","es. IWDA"]].map(([lb,k,ph])=>(
          <div key={k} style={{marginBottom:12}}>
            <div style={{fontSize:9,color:"#3A5070",letterSpacing:2,textTransform:"uppercase",fontFamily:"'DM Mono',monospace",marginBottom:5}}>{lb}</div>
            <input style={IS} value={f[k]} placeholder={ph} onChange={e=>set(k,e.target.value)}/>
          </div>
        ))}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
          {[["Tipo","type",TYPE_LIST],["Area Geo","geo",GEO_LIST]].map(([lb,k,opts])=>(
            <div key={k}>
              <div style={{fontSize:9,color:"#3A5070",letterSpacing:2,textTransform:"uppercase",fontFamily:"'DM Mono',monospace",marginBottom:5}}>{lb}</div>
              <select style={{...IS,appearance:"none"}} value={f[k]} onChange={e=>set(k,e.target.value)}>{opts.map(o=><option key={o}>{o}</option>)}</select>
            </div>
          ))}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:22}}>
          {[["Valore €","value","12000"],["Prezzo acq.","buyPrice","10500"],["Quantità","quantity","45"]].map(([lb,k,ph])=>(
            <div key={k}>
              <div style={{fontSize:8,color:"#3A5070",letterSpacing:1,textTransform:"uppercase",fontFamily:"'DM Mono',monospace",marginBottom:5}}>{lb}</div>
              <input style={IS} type="number" value={f[k]} placeholder={ph} onChange={e=>set(k,e.target.value)}/>
            </div>
          ))}
        </div>
        <Btn label={initial?"💾 Salva Modifiche":"➕ Aggiungi al Portafoglio"} disabled={!ok}
          color="linear-gradient(135deg,#00D4FF,#0096BB)"
          onClick={()=>ok&&onSave({...f,id:initial?.id||Date.now(),value:+f.value,buyPrice:+f.buyPrice,quantity:+f.quantity})}/>
        {initial&&<Btn label="🗑 Rimuovi posizione" onClick={()=>onDelete(initial.id)} color="rgba(255,80,80,0.15)" style={{marginTop:10,color:"#FF5050",border:"1px solid #FF505030"}}/>}
      </div>
    </div>
  );
}

// ─── TargetForm ──────────────────────────────────────────────────
function TargetForm({targets,onSave,onClose}){
  const [tT,setTT]=useState({...INIT_TARGETS.type,...targets.type});
  const [tG,setTG]=useState({...INIT_TARGETS.geo, ...targets.geo});
  const [tab,setTab]=useState("type");
  const sum=obj=>Object.values(obj).reduce((s,v)=>s+(+v||0),0);
  const sT=sum(tT), sG=sum(tG), ok=sT===100&&sG===100;

  function slide(setMap,key,val,map){
    const keys=Object.keys(map), others=keys.filter(k=>k!==key);
    const cur=others.reduce((s,k)=>s+(+map[k]||0),0);
    const nm={...map,[key]:val};
    if(cur>0){others.forEach(k=>{nm[k]=Math.max(0,Math.round((map[k]/cur)*(100-val)));});
      const tot=Object.values(nm).reduce((s,v)=>s+v,0);
      if(tot!==100){const last=others[others.length-1];nm[last]=Math.max(0,(nm[last]||0)+(100-tot));}}
    setMap(nm);
  }

  function SliderSec({map,setMap,colors,s}){
    return(<div>
      {Object.entries(map).map(([k,v])=>(
        <div key={k} style={{marginBottom:16}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
            <span style={{fontSize:12,color:"#8AA0B8",display:"flex",alignItems:"center",gap:6}}>
              <span style={{width:8,height:8,borderRadius:"50%",background:colors[k]||"#888",display:"inline-block"}}/>{k}
            </span>
            <div style={{display:"flex",alignItems:"center",gap:6}}>
              <input type="number" min={0} max={100} value={v}
                onChange={e=>slide(setMap,k,Math.min(100,Math.max(0,+e.target.value||0)),map)}
                style={{...IS,width:60,padding:"4px 8px",fontSize:12,textAlign:"right",fontFamily:"'DM Mono',monospace"}}/>
              <span style={{fontSize:11,color:colors[k]||"#888",fontFamily:"'DM Mono',monospace",minWidth:14}}>%</span>
            </div>
          </div>
          <input type="range" min={0} max={100} value={v} onChange={e=>slide(setMap,k,+e.target.value,map)} style={{width:"100%",accentColor:colors[k]||"#00D4FF"}}/>
        </div>
      ))}
      <div style={{display:"flex",justifyContent:"space-between",padding:"10px 14px",background:s===100?"rgba(0,255,159,0.06)":"rgba(255,80,80,0.06)",border:`1px solid ${s===100?"#00FF9F30":"#FF505030"}`,borderRadius:12,marginTop:4}}>
        <span style={{fontSize:12,color:"#8AA0B8"}}>Totale allocato</span>
        <span style={{fontSize:15,fontWeight:700,fontFamily:"'DM Mono',monospace",color:s===100?"#00FF9F":"#FF5050"}}>{s}% {s===100?"✓":s<100?`(mancano ${100-s}%)`:`(eccede di ${s-100}%)`}</span>
      </div>
    </div>);
  }

  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.88)",zIndex:300,display:"flex",alignItems:"flex-end",maxWidth:430,margin:"0 auto",left:"50%",transform:"translateX(-50%)"}}>
      <div style={{width:"100%",background:"#0D1825",borderRadius:"20px 20px 0 0",border:"1px solid #1A2A3A",padding:"20px 20px 40px",maxHeight:"92vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <span style={{fontSize:16,fontWeight:700}}>🎯 Imposta Target</span>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#4A6080",fontSize:24,cursor:"pointer",padding:0}}>✕</button>
        </div>
        <div style={{display:"flex",background:"#070B11",borderRadius:12,padding:4,marginBottom:20,gap:4}}>
          {[["type","Asset Class",sT===100],["geo","Area Geo",sG===100]].map(([id,lb,chk])=>(
            <button key={id} onClick={()=>setTab(id)} style={{flex:1,background:tab===id?"#0D1825":"none",border:tab===id?"1px solid #1A2A3A":"1px solid transparent",color:tab===id?"#E8F0FF":"#3A5070",borderRadius:9,padding:"8px 0",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>
              {chk&&<span style={{color:"#00FF9F",fontSize:9}}>✓</span>}{lb}
            </button>
          ))}
        </div>
        {tab==="type"&&<SliderSec map={tT} setMap={setTT} colors={C}  s={sT}/>}
        {tab==="geo" &&<SliderSec map={tG} setMap={setTG} colors={GC} s={sG}/>}
        <Btn label={ok?"💾 Salva Target":"Completa entrambe le sezioni (100% ciascuna)"}
          disabled={!ok} color="linear-gradient(135deg,#00D4FF,#0096BB)" style={{marginTop:20}}
          onClick={()=>ok&&onSave({type:tT,geo:tG})}/>
      </div>
    </div>
  );
}

// ─── BackupModal ─────────────────────────────────────────────────
function BackupModal({holdings,targets,onRestore,onClose}){
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.88)",zIndex:300,display:"flex",alignItems:"flex-end",maxWidth:430,margin:"0 auto",left:"50%",transform:"translateX(-50%)"}}>
      <div style={{width:"100%",background:"#0D1825",borderRadius:"20px 20px 0 0",border:"1px solid #1A2A3A",padding:"20px 20px 40px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <span style={{fontSize:16,fontWeight:700}}>📊 Excel Backup</span>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#4A6080",fontSize:24,cursor:"pointer",padding:0}}>✕</button>
        </div>
        <div style={{background:"#070B11",border,borderRadius:14,padding:16,marginBottom:12}}>
          <div style={{fontSize:12,color:"#8AA0B8",marginBottom:4,fontWeight:600}}>📤 Esporta in Excel</div>
          <div style={{fontSize:11,color:"#4A6080",lineHeight:1.6,marginBottom:12}}>
            Genera <b style={{color:"#00D4FF"}}>portafoglio.xlsx</b> con 3 fogli:<br/>
            Posizioni · Target Asset · Target Geo
          </div>
          <Btn label="⬇ Scarica portafoglio.xlsx" color="linear-gradient(135deg,#00D4FF,#0096BB)" onClick={()=>doExport(holdings,targets)}/>
        </div>
        <div style={{background:"#070B11",border,borderRadius:14,padding:16,marginBottom:12}}>
          <div style={{fontSize:12,color:"#8AA0B8",marginBottom:4,fontWeight:600}}>📥 Importa da Excel</div>
          <div style={{fontSize:11,color:"#4A6080",lineHeight:1.6,marginBottom:12}}>
            Carica un file esportato da questa app. Sostituisce i dati attuali.
          </div>
          <label style={{display:"block",width:"100%",background:"rgba(123,97,255,0.12)",color:"#7B61FF",border:"1px solid #7B61FF30",borderRadius:12,padding:"11px 0",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",textAlign:"center",boxSizing:"border-box"}}>
            📂 Carica file .xlsx
            <input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>e.target.files[0]&&doImport(e.target.files[0],d=>{onRestore(d);onClose();})}/>
          </label>
        </div>
        <div style={{background:"rgba(0,212,255,0.04)",border:"1px solid #00D4FF18",borderRadius:12,padding:12}}>
          <div style={{fontSize:10,color:"#4A6080",lineHeight:1.7}}>💡 Salva su <span style={{color:"#00D4FF"}}>Google Drive / iCloud / Dropbox</span> per avere un backup sempre accessibile.</div>
        </div>
      </div>
    </div>
  );
}

// ════════════════════════════════════════════════════════════════
export default function App(){
  const [holdings,setHoldings]=useState(INIT_HOLDINGS);
  const [targets,setTargets]  =useState(INIT_TARGETS);
  const [tab,setTab]          =useState("overview");
  const [form,setForm]        =useState(null);   // null | "new" | holding obj
  const [showTgt,setShowTgt]  =useState(false);
  const [showBack,setShowBack]=useState(false);

  const {total,byType,byGeo,gainPct}=calcStats(holdings);
  const perf=buildPerf(total);

  const saveHolding=useCallback(h=>{
    setHoldings(hs=>{const ex=hs.find(x=>x.id===h.id);return ex?hs.map(x=>x.id===h.id?h:x):[...hs,h];});
    setForm(null);
  },[]);
  const delHolding=useCallback(id=>{setHoldings(hs=>hs.filter(x=>x.id!==id));setForm(null);},[]);
  const restore=useCallback(d=>{if(d?.holdings&&d?.targets){setHoldings(d.holdings);setTargets(d.targets);}},[]);

  const TABS=[
    {id:"overview",icon:"◈",label:"Home"},
    {id:"allocation",icon:"◉",label:"Asset"},
    {id:"geo",icon:"◎",label:"Geo"},
    {id:"performance",icon:"▲",label:"Trend"},
    {id:"holdings",icon:"☰",label:"Posizioni"},
  ];

  return(
    <div style={{background:bg,minHeight:"100vh",maxWidth:430,margin:"0 auto",fontFamily:"'DM Sans',sans-serif",color:"#E8F0FF",position:"relative",overflowX:"hidden"}}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
      {/* Glows */}
      <div style={{position:"fixed",top:-100,left:-80,width:300,height:300,background:"radial-gradient(circle,rgba(0,212,255,0.07),transparent 70%)",pointerEvents:"none",zIndex:0}}/>
      <div style={{position:"fixed",bottom:80,right:-80,width:260,height:260,background:"radial-gradient(circle,rgba(123,97,255,0.06),transparent 70%)",pointerEvents:"none",zIndex:0}}/>

      {/* ── Header ── */}
      <div style={{padding:"48px 20px 12px",position:"relative",zIndex:1}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div>
            <div style={{fontSize:10,color:"#4A6080",letterSpacing:3,textTransform:"uppercase",fontFamily:"'DM Mono',monospace"}}>IL MIO PORTAFOGLIO</div>
            <div style={{fontSize:32,fontWeight:700,letterSpacing:-1,marginTop:2,lineHeight:1}}>€{total.toLocaleString("it-IT",{minimumFractionDigits:2})}</div>
            <div style={{display:"inline-flex",alignItems:"center",gap:4,marginTop:6,background:+gainPct>=0?"rgba(0,255,159,0.1)":"rgba(255,80,80,0.1)",color:+gainPct>=0?"#00FF9F":"#FF5050",padding:"3px 10px",borderRadius:20,fontSize:11,fontFamily:"'DM Mono',monospace"}}>
              {+gainPct>=0?"▲":"▼"} {Math.abs(gainPct)}% rendimento totale
            </div>
          </div>
          {/* Action buttons */}
          <div style={{display:"flex",gap:6,alignItems:"center",marginTop:4}}>
            <IconBtn icon="📊" title="Excel backup" onClick={()=>setShowBack(true)} color="#00D4FF"/>
            <IconBtn icon="🎯" title="Target"       onClick={()=>setShowTgt(true)}  color="#7B61FF"/>
            <IconBtn icon="+"  title="Aggiungi"     onClick={()=>setForm("new")}    color="#00D4FF" big/>
          </div>
        </div>
      </div>

      {/* ── Page content ── */}
      <div style={{padding:"0 16px 100px",position:"relative",zIndex:1}}>
        {tab==="overview"    && <PageOverview    holdings={holdings} byType={byType} onEdit={setForm}/>}
        {tab==="allocation"  && <PageAllocation  byType={byType} targets={targets} onEditTarget={()=>setShowTgt(true)}/>}
        {tab==="geo"         && <PageGeo         byGeo={byGeo} total={total} targets={targets} onEditTarget={()=>setShowTgt(true)}/>}
        {tab==="performance" && <PagePerformance perf={perf}/>}
        {tab==="holdings"    && <PageHoldings    holdings={holdings} onEdit={setForm} onNew={()=>setForm("new")}/>}
      </div>

      {/* ── Bottom nav ── */}
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:430,background:"rgba(7,11,17,0.95)",backdropFilter:"blur(20px)",borderTop:"1px solid #0F1A28",padding:"8px 0 20px",display:"flex",justifyContent:"space-around",zIndex:100}}>
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{background:"none",border:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:3,padding:"5px 10px"}}>
            <span style={{fontSize:16,color:tab===t.id?"#00D4FF":"#2A3A50",filter:tab===t.id?"drop-shadow(0 0 6px #00D4FF80)":"none",transition:"color .2s"}}>{t.icon}</span>
            <span style={{fontSize:8,letterSpacing:.5,textTransform:"uppercase",color:tab===t.id?"#00D4FF":"#2A3A50",fontFamily:"'DM Mono',monospace"}}>{t.label}</span>
          </button>
        ))}
      </div>

      {/* ── Modals ── */}
      {form      && <HoldingForm initial={form==="new"?null:form} onSave={saveHolding} onClose={()=>setForm(null)} onDelete={delHolding}/>}
      {showTgt   && <TargetForm  targets={targets} onSave={t=>{setTargets(t);setShowTgt(false);}} onClose={()=>setShowTgt(false)}/>}
      {showBack  && <BackupModal holdings={holdings} targets={targets} onRestore={restore} onClose={()=>setShowBack(false)}/>}
    </div>
  );
}

function IconBtn({icon,title,onClick,color,big=false}){
  return(
    <button onClick={onClick} title={title} style={{width:big?42:38,height:big?42:38,borderRadius:"50%",background:`${color}18`,border:`1px solid ${color}40`,color,fontSize:big?22:16,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",lineHeight:1,flexShrink:0}}>
      {icon}
    </button>
  );
}

// ════════════════ PAGE COMPONENTS ══════════════════════════════

function PageOverview({holdings,byType,onEdit}){
  const top=[...holdings].sort((a,b)=>b.value-a.value).slice(0,6);
  return(<div>
    <Sec t="Composizione"/>
    <Card ch={
      <div style={{display:"flex",alignItems:"center",gap:8}}>
        <div style={{width:128,height:128,flexShrink:0}}>
          <ResponsiveContainer width="100%" height="100%">
            <PieChart><Pie data={byType.filter(x=>x.value>0)} cx="50%" cy="50%" innerRadius={36} outerRadius={56} paddingAngle={3} dataKey="value" strokeWidth={0}>{byType.map((e,i)=><Cell key={i} fill={e.color}/>)}</Pie></PieChart>
          </ResponsiveContainer>
        </div>
        <div style={{flex:1}}>
          {byType.map((a,i)=>(
            <div key={i} style={{marginBottom:8}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                <span style={{fontSize:11,color:"#8AA0B8",display:"flex",alignItems:"center",gap:5}}><span style={{width:7,height:7,borderRadius:"50%",background:a.color,display:"inline-block"}}/>{a.name}</span>
                <span style={{fontSize:12,fontWeight:600,fontFamily:"'DM Mono',monospace",color:a.color}}>{a.value}%</span>
              </div>
              <div style={{height:3,background:"#0F1E30",borderRadius:3,overflow:"hidden"}}><div style={{height:"100%",width:`${a.value}%`,background:a.color,borderRadius:3}}/></div>
            </div>
          ))}
        </div>
      </div>
    }/>
    <Sec t="Posizioni principali · tocca per modificare"/>
    {top.map(h=>{
      const gp=h.buyPrice?(((h.value-h.buyPrice)/h.buyPrice)*100).toFixed(1):null;
      return(<Card key={h.id} style={{padding:"12px 14px"}} onClick={()=>onEdit(h)} ch={
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div style={{display:"flex",alignItems:"center",gap:9}}>
            <div style={{width:38,height:38,borderRadius:10,background:`${C[h.type]}18`,border:`1px solid ${C[h.type]}30`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,color:C[h.type],fontFamily:"'DM Mono',monospace",fontWeight:700}}>{h.ticker.slice(0,2)}</div>
            <div><div style={{fontSize:13,fontWeight:600}}>{h.name}</div><div style={{fontSize:9,color:"#3A5070",fontFamily:"'DM Mono',monospace"}}>{h.ticker} · {h.quantity} pz</div></div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:13,fontWeight:600,fontFamily:"'DM Mono',monospace"}}>€{h.value.toLocaleString("it-IT")}</div>
            {gp!==null&&<div style={{fontSize:10,color:+gp>=0?"#00FF9F":"#FF5050",fontFamily:"'DM Mono',monospace"}}>{+gp>=0?"+":""}{gp}%</div>}
          </div>
        </div>
      }/>);
    })}
    {holdings.length===0&&<Card ch={<div style={{textAlign:"center",color:"#3A5070",padding:"20px 0"}}>Nessuna posizione.<br/>Premi <b>+</b> in alto per iniziare.</div>}/>}
  </div>);
}

function PageAllocation({byType,targets,onEditTarget}){
  return(<div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
      <Sec t="Asset Class vs Target"/>
      <button onClick={onEditTarget} style={{background:"#00D4FF18",border:"1px solid #00D4FF30",color:"#00D4FF",borderRadius:20,padding:"4px 12px",fontSize:10,cursor:"pointer",fontFamily:"'DM Mono',monospace",marginBottom:12}}>🎯 Modifica Target</button>
    </div>
    {byType.map(a=>{
      const tgt=targets.type[a.name]||0, diff=a.value-tgt;
      return(<Card key={a.name} ch={<div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div style={{display:"flex",alignItems:"center",gap:7}}><div style={{width:9,height:9,borderRadius:"50%",background:a.color}}/><span style={{fontSize:14,fontWeight:600}}>{a.name}</span></div>
          <span style={{fontSize:10,padding:"2px 8px",borderRadius:20,background:Math.abs(diff)<=3?"rgba(0,255,159,0.1)":"rgba(255,107,53,0.1)",color:Math.abs(diff)<=3?"#00FF9F":"#FF6B35",fontFamily:"'DM Mono',monospace"}}>{diff>0?"+":""}{diff}% vs target</span>
        </div>
        <div style={{display:"flex",gap:8,marginBottom:10}}>
          {[["Attuale",`${a.value}%`,a.color],["Target",`${tgt}%`,"#2A4060"],["Valore",`€${(a.amount/1000).toFixed(1)}k`,"#E8F0FF"]].map(([l,v,c],j)=>(
            <div key={j} style={{flex:1,textAlign:"center",background:"#0A1420",borderRadius:10,padding:"7px 0"}}>
              <div style={{fontSize:15,fontWeight:700,color:c,fontFamily:"'DM Mono',monospace"}}>{v}</div>
              <div style={{fontSize:8,color:"#3A5070",textTransform:"uppercase",letterSpacing:.8}}>{l}</div>
            </div>
          ))}
        </div>
        <div style={{height:6,background:"#0A1420",borderRadius:3,overflow:"hidden",marginBottom:4}}><div style={{height:"100%",width:`${a.value}%`,background:a.color,borderRadius:3}}/></div>
        <div style={{height:3,background:"#0A1420",borderRadius:3,overflow:"hidden"}}><div style={{height:"100%",width:`${tgt}%`,background:"#2A4060",borderRadius:3}}/></div>
      </div>}/>);
    })}
  </div>);
}

function PageGeo({byGeo,total,targets,onEditTarget}){
  return(<div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
      <Sec t="Geografia vs Target"/>
      <button onClick={onEditTarget} style={{background:"#00D4FF18",border:"1px solid #00D4FF30",color:"#00D4FF",borderRadius:20,padding:"4px 12px",fontSize:10,cursor:"pointer",fontFamily:"'DM Mono',monospace",marginBottom:12}}>🎯 Modifica Target</button>
    </div>
    <Card ch={<div>
      <div style={{height:180}}>
        <ResponsiveContainer width="100%" height="100%">
          <PieChart><Pie data={byGeo} cx="50%" cy="50%" innerRadius={50} outerRadius={76} paddingAngle={2} dataKey="value" strokeWidth={0}>{byGeo.map((e,i)=><Cell key={i} fill={e.color}/>)}</Pie></PieChart>
        </ResponsiveContainer>
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:8,justifyContent:"center"}}>
        {byGeo.map((g,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:4}}><div style={{width:7,height:7,borderRadius:"50%",background:g.color}}/><span style={{fontSize:10,color:"#8AA0B8"}}>{g.name}</span><span style={{fontSize:10,color:g.color,fontFamily:"'DM Mono',monospace"}}>{g.value}%</span></div>)}
      </div>
    </div>}/>
    {byGeo.map((g,i)=>{
      const tgt=targets.geo[g.name]||0, diff=g.value-tgt;
      return(<Card key={i} style={{padding:"12px 14px"}} ch={<div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
          <span style={{fontSize:13,fontWeight:600}}>{g.name}</span>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:10,padding:"2px 8px",borderRadius:20,background:Math.abs(diff)<=3?"rgba(0,255,159,0.1)":"rgba(255,107,53,0.1)",color:Math.abs(diff)<=3?"#00FF9F":"#FF6B35",fontFamily:"'DM Mono',monospace"}}>{diff>0?"+":""}{diff}%</span>
            <span style={{fontFamily:"'DM Mono',monospace",fontSize:14,color:g.color,fontWeight:700}}>{g.value}%</span>
          </div>
        </div>
        <div style={{height:6,background:"#0A1420",borderRadius:3,overflow:"hidden",marginBottom:3}}><div style={{height:"100%",width:`${g.value}%`,background:`linear-gradient(90deg,${g.color}70,${g.color})`,borderRadius:3}}/></div>
        <div style={{height:3,background:"#0A1420",borderRadius:3,overflow:"hidden",marginBottom:4}}><div style={{height:"100%",width:`${tgt}%`,background:"#2A4060",borderRadius:3}}/></div>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:9,color:"#3A5070",fontFamily:"'DM Mono',monospace"}}><span>€{(total*g.value/100).toLocaleString("it-IT",{maximumFractionDigits:0})}</span><span>target: {tgt}%</span></div>
      </div>}/>);
    })}
    {byGeo.length===0&&<Card ch={<div style={{textAlign:"center",color:"#3A5070",padding:"20px 0"}}>Aggiungi posizioni con area geografica.</div>}/>}
  </div>);
}

function PagePerformance({perf}){
  const s=perf[0].portfolio,e=perf[perf.length-1].portfolio;
  const r=((e-s)/s*100).toFixed(1), tr=((perf[perf.length-1].target-perf[0].target)/perf[0].target*100).toFixed(1);
  return(<div>
    <Sec t="Andamento 12 Mesi"/>
    <div style={{display:"flex",gap:8,marginBottom:12}}>
      {[[`+${r}%`,"Portafoglio","#00D4FF"],[`+${tr}%`,"Target","#FFD700"]].map(([v,l,c],i)=>(
        <div key={i} style={{flex:1,background:"#0D1825",border,borderRadius:12,padding:"10px 6px",textAlign:"center"}}>
          <div style={{fontSize:18,fontWeight:700,color:c,fontFamily:"'DM Mono',monospace"}}>{v}</div>
          <div style={{fontSize:8,color:"#3A5070",textTransform:"uppercase",letterSpacing:.5,marginTop:2}}>{l}</div>
        </div>
      ))}
    </div>
    <Card style={{padding:"14px 8px 8px"}} ch={<div>
      <div style={{height:200}}>
        <ResponsiveContainer width="100%" height="100%">
          <LineChart data={perf} margin={{top:4,right:8,left:0,bottom:0}}>
            <XAxis dataKey="month" tick={{fontSize:8,fill:"#2A4060",fontFamily:"DM Mono"}} tickLine={false} axisLine={false}/>
            <YAxis hide domain={["auto","auto"]}/>
            <Tooltip content={<TTip/>}/>
            <Line type="monotone" dataKey="portfolio" name="Portafoglio" stroke="#00D4FF" strokeWidth={2} dot={false}/>
            <Line type="monotone" dataKey="target"    name="Target"      stroke="#FFD700" strokeWidth={1.5} strokeDasharray="4 4" dot={false}/>
          </LineChart>
        </ResponsiveContainer>
      </div>
      <div style={{display:"flex",justifyContent:"center",gap:16,marginTop:8}}>
        {[["#00D4FF","Portafoglio"],["#FFD700","Target"]].map(([c,l],i)=><div key={i} style={{display:"flex",alignItems:"center",gap:4}}><div style={{width:14,height:2,background:c}}/><span style={{fontSize:9,color:"#4A6080",fontFamily:"'DM Mono',monospace"}}>{l}</span></div>)}
      </div>
    </div>}/>
    <Card style={{border:"1px solid #1E2A3A"}} ch={<div style={{fontSize:11,color:"#4A6080",lineHeight:1.7}}>ℹ️ Curva indicativa basata sui valori inseriti. I rendimenti passati non garantiscono quelli futuri.</div>}/>
  </div>);
}

function PageHoldings({holdings,onEdit,onNew}){
  const [fil,setFil]=useState("Tutti");
  const list=fil==="Tutti"?holdings:holdings.filter(h=>h.type===fil);
  return(<div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
      <Sec t={`Posizioni (${holdings.length})`}/>
      <button onClick={onNew} style={{background:"#00D4FF",color:"#070B11",border:"none",borderRadius:20,padding:"5px 14px",fontSize:11,fontWeight:700,cursor:"pointer",marginBottom:12}}>+ Aggiungi</button>
    </div>
    <div style={{display:"flex",gap:6,marginBottom:12,overflowX:"auto",paddingBottom:4}}>
      {["Tutti",...TYPE_LIST].map(f=><button key={f} onClick={()=>setFil(f)} style={{background:fil===f?"#00D4FF":"#0D1825",border:fil===f?"none":"1px solid #0F1E30",color:fil===f?"#070B11":"#4A6080",borderRadius:20,padding:"5px 13px",fontSize:10,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap",fontFamily:"'DM Sans',sans-serif"}}>{f}</button>)}
    </div>
    {list.length===0&&<Card ch={<div style={{textAlign:"center",color:"#3A5070",padding:"20px 0"}}>Nessuna posizione. Premi <b>+ Aggiungi</b>.</div>}/>}
    {list.map(h=>{
      const gp=h.buyPrice?(((h.value-h.buyPrice)/h.buyPrice)*100).toFixed(1):null;
      return(<Card key={h.id} style={{padding:"13px 14px"}} onClick={()=>onEdit(h)} ch={
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div style={{display:"flex",gap:9,alignItems:"center"}}>
            <div style={{width:40,height:40,borderRadius:11,background:`${C[h.type]}14`,border:`1px solid ${C[h.type]}30`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:C[h.type],fontFamily:"'DM Mono',monospace"}}>{h.ticker.slice(0,2)}</div>
            <div><div style={{fontSize:13,fontWeight:600}}>{h.name}</div><div style={{fontSize:9,color:"#3A5070",fontFamily:"'DM Mono',monospace",marginTop:1}}>{h.ticker} · {h.geo} · {h.quantity} pz</div></div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:14,fontWeight:700,fontFamily:"'DM Mono',monospace"}}>€{h.value.toLocaleString("it-IT")}</div>
            {gp!==null&&<div style={{fontSize:10,color:+gp>=0?"#00FF9F":"#FF5050",fontFamily:"'DM Mono',monospace"}}>{+gp>=0?"+":""}{gp}%</div>}
          </div>
        </div>
      }/>);
    })}
  </div>);
}
