import { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

// ── SAMPLE DATA ───────────────────────────────────────────────────────────────
const SAMPLE = [
  {
    id:"1", tag:"宮崎-0042", name:"黒姫", sex:"雌", breed:"黒毛和種",
    birthDate:"2024-03-15", introDate:"2025-01-10",
    farm:"宮崎中央市場", pen:"1号棟A", shippingPlan:"2026-08-01",
    expectedPrice:1500000, memo:"おとなしい。発育良好。",
    pedigree:{
      sire:{ name:"安福久",
        sire:{ name:"安平",    sire:{name:"第1藤良"}, dam:{name:"菊平"} },
        dam: { name:"福姫",    sire:{name:"糸福"},    dam:{name:"花月"} } },
      dam: { name:"第6福桜",
        sire:{ name:"忠富士",  sire:{name:"紋次郎"},  dam:{name:"福美"} },
        dam: { name:"桜姫",    sire:{name:"昭和"},    dam:{name:"梅花"} } }
    },
    costs:{ purchasePrice:820000, roughageDaily:400, compoundKgPerDay:8, compoundKgPrice:80, otherDaily:200, fixedOther:30000, vetCosts:0 },
    weights:[{date:"2025-01-10",weight:280},{date:"2025-03-01",weight:362},{date:"2025-05-01",weight:435}],
    vaccines:[
      {date:"2025-01-15", name:"口蹄疫ワクチン", nextDate:"2026-05-12"},
    ],
    treatments:[], status:"肥育中", result:null,
  },
  {
    id:"2", tag:"鹿児島-0051", name:"武蔵", sex:"去勢", breed:"黒毛和種",
    birthDate:"2024-01-20", introDate:"2025-02-05",
    farm:"鹿児島中央市場", pen:"1号棟B", shippingPlan:"2026-05-18",
    expectedPrice:1400000, memo:"3月に軽い下痢。現在回復済み。",
    pedigree:{
      sire:{ name:"糸福",
        sire:{ name:"第12茂重波", sire:{name:"気高波"},  dam:{name:"岩梅"} },
        dam: { name:"第9茂福",    sire:{name:"茂重波"},  dam:{name:"光福"} } },
      dam: { name:"福乃国",
        sire:{ name:"北国7の8",   sire:{name:"北国"},    dam:{name:"菊恵"} },
        dam: { name:"雪乃",       sire:{name:"百合茂"},  dam:{name:"松乃"} } }
    },
    costs:{ purchasePrice:760000, roughageDaily:380, compoundKgPerDay:8, compoundKgPrice:80, otherDaily:180, fixedOther:30000, vetCosts:15000 },
    weights:[{date:"2025-02-05",weight:310},{date:"2025-04-01",weight:395},{date:"2025-11-01",weight:530},{date:"2026-03-01",weight:642}],
    vaccines:[
      {date:"2025-02-10", name:"BRDワクチン", nextDate:"2026-05-08"},
    ],
    treatments:[{date:"2025-03-10",name:"下痢",drug:"抗生剤",vet:"田中獣医師",cost:15000}],
    status:"肥育中", result:null,
  },
  {
    id:"4", tag:"岩手-0077", name:"錦波", sex:"去勢", breed:"黒毛和種",
    birthDate:"2023-08-22", introDate:"2024-06-10",
    farm:"岩手中央市場", pen:"2号棟A", shippingPlan:"2026-05-25",
    expectedPrice:1480000, memo:"発育良好。出荷準備中。",
    pedigree:{
      sire:{ name:"安福久",
        sire:{ name:"安平",      sire:{name:"第1藤良"}, dam:{name:"菊平"} },
        dam: { name:"福恵",      sire:{name:"百合茂"},  dam:{name:"花菊"} } },
      dam: { name:"第3波乃",
        sire:{ name:"北国7の8",  sire:{name:"北国"},    dam:{name:"雪姫"} },
        dam: { name:"波乃花",    sire:{name:"茂重波"},  dam:{name:"桜波"} } }
    },
    costs:{ purchasePrice:800000, roughageDaily:410, compoundKgPerDay:9, compoundKgPrice:78, otherDaily:190, fixedOther:30000, vetCosts:5000 },
    weights:[{date:"2024-06-10",weight:298},{date:"2024-09-01",weight:405},{date:"2025-01-15",weight:530},{date:"2025-06-01",weight:648},{date:"2026-03-10",weight:720}],
    vaccines:[
      {date:"2024-06-15", name:"口蹄疫ワクチン", nextDate:"2026-06-15"},
      {date:"2025-01-20", name:"BRDワクチン",    nextDate:"2026-05-06"},
    ],
    treatments:[], status:"肥育中", result:null,
  },
  {
    id:"3", tag:"宮崎-0011", name:"大黒", sex:"去勢", breed:"黒毛和種",
    birthDate:"2022-08-10", introDate:"2023-06-01",
    farm:"宮崎中央市場", pen:"出荷済", shippingPlan:"2025-03-10",
    expectedPrice:1600000, memo:"出荷済。",
    pedigree:{
      sire:{ name:"安平",
        sire:{ name:"第1藤良",   sire:{name:"藤良"},   dam:{name:"光月"} },
        dam: { name:"菊平",      sire:{name:"平茂勝"}, dam:{name:"菊月"} } },
      dam: { name:"第5大和",
        sire:{ name:"忠富士",    sire:{name:"富士"},   dam:{name:"忠子"} },
        dam: { name:"大和姫",    sire:{name:"昭和"},   dam:{name:"千代"} } }
    },
    costs:{ purchasePrice:850000, roughageDaily:420, compoundKgPerDay:9, compoundKgPrice:78, otherDaily:200, fixedOther:30000, vetCosts:20000 },
    weights:[{date:"2023-06-01",weight:310},{date:"2023-09-01",weight:430},{date:"2024-01-01",weight:560},{date:"2024-06-01",weight:690}],
    vaccines:[], treatments:[], status:"出荷済",
    result:{sellPrice:1620000,bms:9,loinArea:62,ribThickness:8.2,yieldGrade:"A",grade:"A5",dg:0.97},
  },
];

// ── UTILS ──────────────────────────────────────────────────────────────────────
const fmtDate = (d) => d ? new Date(d).toLocaleDateString("ja-JP",{year:"numeric",month:"2-digit",day:"2-digit"}) : "―";
const fmtMoney = (n) => n != null ? `¥${Number(n).toLocaleString()}` : "―";
const fmtM = (n) => n != null ? `¥${(Number(n)/10000).toFixed(0)}万` : "―";
const daysSince = (d) => d ? Math.floor((Date.now()-new Date(d))/86400000) : 0;
const daysUntil = (d) => d ? Math.ceil((new Date(d)-Date.now())/86400000) : null;
const calcAge = (b) => {
  if(!b) return "―";
  const m = Math.floor((Date.now()-new Date(b))/(30.44*86400000));
  return m>=12 ? `${Math.floor(m/12)}歳${m%12}ヶ月` : `${m}ヶ月`;
};
const latestWeight = (ws) => ws?.length ? ws[ws.length-1].weight : null;
const calcDG = (ws) => {
  if(!ws||ws.length<2) return null;
  const f=ws[0], l=ws[ws.length-1];
  const days=Math.floor((new Date(l.date)-new Date(f.date))/86400000);
  return days>0 ? (l.weight-f.weight)/days : null;
};
const predictWeight = (cow) => {
  const dg=calcDG(cow.weights), lw=latestWeight(cow.weights), du=daysUntil(cow.shippingPlan);
  if(!dg||!lw||!du||du<0) return null;
  return Math.round(lw+dg*du);
};
const calcCosts = (cow) => {
  const c = cow.costs || {};
  const days = daysSince(cow.introDate);
  const compoundDaily = (c.compoundKgPerDay||0)*(c.compoundKgPrice||0);
  const dailyTotal = (c.roughageDaily||0)+compoundDaily+(c.otherDaily||0);
  const runningCost = dailyTotal*days;
  const totalCost = (c.purchasePrice||0)+runningCost+(c.fixedOther||0)+(c.vetCosts||0);
  const sellPrice = cow.result?.sellPrice||cow.expectedPrice||null;
  const profit = sellPrice ? sellPrice-totalCost : null;
  return { compoundDaily, dailyTotal, runningCost, totalCost, profit };
};
const monthKey = (d) => d ? d.slice(0,7) : null;
const avg = (arr) => arr.length ? arr.reduce((s,v)=>s+v,0)/arr.length : null;

// ── EMPTY PEDIGREE ──────────────────────────────────────────────────────────
const emptyPedigree = () => ({
  sire:{ name:"", sire:{ name:"", sire:{name:""}, dam:{name:""} }, dam:{ name:"", sire:{name:""}, dam:{name:""} } },
  dam: { name:"", sire:{ name:"", sire:{name:""}, dam:{name:""} }, dam:{ name:"", sire:{name:""}, dam:{name:""} } },
});
const emptyCosts = () => ({ purchasePrice:0, roughageDaily:400, compoundKgPerDay:8, compoundKgPrice:80, otherDaily:200, fixedOther:30000, vetCosts:0 });

// ── THEME ──────────────────────────────────────────────────────────────────────
const C = {
  bg:"#f4f9fc", surface:"#ffffff", card:"#ffffff", cardSub:"#f0f7fb",
  border:"#cfe8f4", borderLight:"#e4f2f9",
  accent:"#4ab8e8", accentMid:"#7ecef2", accentLight:"#e0f4fd", accentDark:"#2a8ec4",
  green:"#4abf8a", greenLight:"#e6f8f0",
  amber:"#f5a623", amberLight:"#fff4e0",
  red:"#e86060", redLight:"#fdeaea",
  purple:"#8b7cf8", purpleLight:"#f0eeff",
  teal:"#2bc0b4", tealLight:"#e0f8f6",
  text:"#1e3a4a", textMid:"#4a7a92", textDim:"#8ab4c8",
  shadow:"0 2px 12px rgba(74,184,232,0.10)",
  shadowMd:"0 4px 20px rgba(74,184,232,0.15)",
};

// ── UI COMPONENTS ──────────────────────────────────────────────────────────────
const Tag = ({label,color=C.accent,bg}) => (
  <span style={{background:bg||color+"18",color,border:`1px solid ${color}44`,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{label}</span>
);
const Btn = ({children,onClick,variant="primary",sm,full,disabled,icon}) => {
  const s = {
    primary:{bg:`linear-gradient(135deg,${C.accent},${C.accentDark})`,color:"#fff",border:"none"},
    outline:{bg:"#fff",color:C.accent,border:`1.5px solid ${C.accent}`},
    soft:   {bg:C.accentLight,color:C.accentDark,border:`1px solid ${C.border}`},
    teal:   {bg:C.tealLight,color:C.teal,border:`1px solid ${C.teal}44`},
    danger: {bg:C.redLight,color:C.red,border:`1px solid ${C.red}44`},
  }[variant]||{bg:C.accentLight,color:C.accentDark,border:"none"};
  return (
    <button onClick={disabled?undefined:onClick} style={{background:s.bg,color:s.color,border:s.border,borderRadius:12,padding:sm?"7px 14px":"11px 20px",fontSize:sm?12:14,fontWeight:700,cursor:disabled?"not-allowed":"pointer",width:full?"100%":"auto",opacity:disabled?0.5:1,display:"flex",alignItems:"center",gap:5,justifyContent:"center",boxShadow:variant==="primary"?C.shadow:"none"}}>
      {icon&&<span>{icon}</span>}{children}
    </button>
  );
};
const Card = ({children,style,onClick}) => (
  <div onClick={onClick} style={{background:C.card,borderRadius:16,border:`1px solid ${C.border}`,padding:"16px",boxShadow:C.shadow,cursor:onClick?"pointer":"default",...style}}>{children}</div>
);
const InfoRow = ({label,value,accent,big,last}) => (
  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"11px 0",borderBottom:last?"none":`1px solid ${C.borderLight}`}}>
    <span style={{color:C.textDim,fontSize:12,flexShrink:0}}>{label}</span>
    <span style={{color:accent?C.accentDark:C.text,fontSize:big?16:13,fontWeight:accent||big?700:400,textAlign:"right",marginLeft:8}}>{value}</span>
  </div>
);
const SectionLabel = ({children,color}) => (
  <div style={{color:color||C.textDim,fontSize:11,fontWeight:700,letterSpacing:2,marginBottom:10,marginTop:4}}>{children}</div>
);
const Modal = ({title,onClose,children,wide}) => (
  <div style={{position:"fixed",inset:0,background:"rgba(30,58,74,0.4)",zIndex:300,display:"flex",alignItems:"flex-end",justifyContent:"center",backdropFilter:"blur(3px)"}}>
    <div style={{background:"#fff",borderRadius:"24px 24px 0 0",width:"100%",maxWidth:wide?600:520,maxHeight:"92vh",overflowY:"auto",padding:"8px 20px 40px",boxShadow:"0 -8px 40px rgba(74,184,232,0.18)"}}>
      <div style={{width:40,height:4,background:C.border,borderRadius:2,margin:"12px auto 20px"}}/>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <span style={{color:C.text,fontWeight:800,fontSize:16}}>{title}</span>
        <button onClick={onClose} style={{background:C.accentLight,border:"none",color:C.textMid,width:32,height:32,borderRadius:"50%",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
      </div>
      {children}
    </div>
  </div>
);
const FInput = ({label,children,hint}) => (
  <div style={{marginBottom:16}}>
    <div style={{color:C.textMid,fontSize:12,fontWeight:600,marginBottom:5}}>{label}</div>
    {children}
    {hint&&<div style={{color:C.textDim,fontSize:10,marginTop:3}}>{hint}</div>}
  </div>
);
const inp = {width:"100%",background:C.bg,border:`1.5px solid ${C.border}`,color:C.text,borderRadius:10,padding:"10px 14px",fontSize:14,boxSizing:"border-box",outline:"none"};

// ── かしこい黒牛SVG（再利用コンポーネント） ────────────────────────────────────
const SmartCowSvg = ({size=20}) => (
  <svg width={size} height={size} viewBox="0 0 80 80" fill="none">
    <path d="M24 22 Q18 10 22 5 Q28 10 26 20" fill="#1a2a1e" opacity="0.92"/>
    <path d="M56 22 Q62 10 58 5 Q52 10 54 20" fill="#1a2a1e" opacity="0.92"/>
    <ellipse cx="14" cy="32" rx="8" ry="11" fill="#1e3a2a"/>
    <ellipse cx="14" cy="32" rx="5" ry="7" fill="#2a5a3a" opacity="0.7"/>
    <ellipse cx="66" cy="32" rx="8" ry="11" fill="#1e3a2a"/>
    <ellipse cx="66" cy="32" rx="5" ry="7" fill="#2a5a3a" opacity="0.7"/>
    <ellipse cx="40" cy="44" rx="26" ry="28" fill="#1a2a1e"/>
    <ellipse cx="40" cy="28" rx="10" ry="6" fill="white" opacity="0.07"/>
    <rect x="18" y="36" width="16" height="11" rx="5.5" stroke="white" strokeWidth="2" fill="none" opacity="0.85"/>
    <rect x="46" y="36" width="16" height="11" rx="5.5" stroke="white" strokeWidth="2" fill="none" opacity="0.85"/>
    <line x1="34" y1="41" x2="46" y2="41" stroke="white" strokeWidth="1.8" opacity="0.85"/>
    <line x1="18" y1="41" x2="12" y2="38" stroke="white" strokeWidth="1.5" opacity="0.7"/>
    <line x1="62" y1="41" x2="68" y2="38" stroke="white" strokeWidth="1.5" opacity="0.7"/>
    <circle cx="26" cy="41" r="3.2" fill="white"/>
    <circle cx="26.8" cy="40.2" r="1.3" fill="#1a2a1e"/>
    <circle cx="54" cy="41" r="3.2" fill="white"/>
    <circle cx="54.8" cy="40.2" r="1.3" fill="#1a2a1e"/>
    <ellipse cx="40" cy="60" rx="12" ry="8.5" fill="#2a3d2e"/>
    <ellipse cx="35.5" cy="61" rx="2.8" ry="2.2" fill="#111e14"/>
    <ellipse cx="44.5" cy="61" rx="2.8" ry="2.2" fill="#111e14"/>
    <path d="M34 67 Q40 71 46 67" stroke="white" strokeWidth="1.4" strokeLinecap="round" fill="none" opacity="0.5"/>
    <circle cx="52" cy="20" r="3" fill="white" opacity="0.6"/>
    <circle cx="57" cy="15" r="1.8" fill="white" opacity="0.4"/>
    <circle cx="55" cy="24" r="1.2" fill="white" opacity="0.3"/>
  </svg>
);

const KpiCard = ({label,value,unit="",color=C.accent,bg=C.accentLight,icon}) => (
  <div style={{background:bg,borderRadius:14,padding:"12px 14px",border:`1px solid ${color}22`}}>
    <div style={{display:"flex",alignItems:"center",gap:4,marginBottom:5}}>
      {icon&&(
        typeof icon === "string"
          ? <span style={{fontSize:14}}>{icon}</span>
          : <span style={{display:"flex",alignItems:"center"}}>{icon}</span>
      )}
      <span style={{color:C.textMid,fontSize:10,fontWeight:600}}>{label}</span>
    </div>
    <div style={{color,fontSize:18,fontWeight:900,lineHeight:1}}>
      {value}<span style={{fontSize:10,fontWeight:500,marginLeft:2,color:C.textDim}}>{unit}</span>
    </div>
  </div>
);

// ── PEDIGREE TREE COMPONENT ────────────────────────────────────────────────────
function PedigreeTree({pedigree}) {
  if(!pedigree) return null;
  const p = pedigree;
  const cell = (name, color=C.text, bg=C.cardSub, size=11) => (
    <div style={{background:bg,border:`1px solid ${C.border}`,borderRadius:8,padding:"5px 8px",fontSize:size,fontWeight:700,color,textAlign:"center",minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
      {name||<span style={{color:C.textDim,fontWeight:400}}>未登録</span>}
    </div>
  );
  return (
    <div style={{overflowX:"auto",paddingBottom:4}}>
      <div style={{display:"grid",gridTemplateColumns:"2fr 1.6fr 1.3fr",gap:6,minWidth:280}}>
        {/* Col 1: 父・母 */}
        <div style={{display:"grid",gridTemplateRows:"1fr 1fr",gap:6,alignContent:"center"}}>
          {cell(p.sire?.name,"#fff",C.accentDark,12)}
          {cell(p.dam?.name,"#fff","#e06090cc",12)}
        </div>
        {/* Col 2: 祖父母 (4) */}
        <div style={{display:"grid",gridTemplateRows:"1fr 1fr 1fr 1fr",gap:4}}>
          {cell(p.sire?.sire?.name,C.accentDark,C.accentLight)}
          {cell(p.sire?.dam?.name,"#e06090",C.purpleLight)}
          {cell(p.dam?.sire?.name,C.accentDark,C.accentLight)}
          {cell(p.dam?.dam?.name,"#e06090",C.purpleLight)}
        </div>
        {/* Col 3: 曾祖父母 (8) */}
        <div style={{display:"grid",gridTemplateRows:"repeat(8,1fr)",gap:3}}>
          {[
            p.sire?.sire?.sire?.name, p.sire?.sire?.dam?.name,
            p.sire?.dam?.sire?.name,  p.sire?.dam?.dam?.name,
            p.dam?.sire?.sire?.name,  p.dam?.sire?.dam?.name,
            p.dam?.dam?.sire?.name,   p.dam?.dam?.dam?.name,
          ].map((n,i)=>
            <div key={i} style={{background:C.bg,border:`1px solid ${C.borderLight}`,borderRadius:6,padding:"3px 6px",fontSize:9,fontWeight:600,color:n?C.textMid:C.textDim,textAlign:"center",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
              {n||"―"}
            </div>
          )}
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"2fr 1.6fr 1.3fr",gap:6,marginTop:4}}>
        <div style={{textAlign:"center",fontSize:9,color:C.textDim}}>親（1代）</div>
        <div style={{textAlign:"center",fontSize:9,color:C.textDim}}>祖父母（2代）</div>
        <div style={{textAlign:"center",fontSize:9,color:C.textDim}}>曾祖父母（3代）</div>
      </div>
    </div>
  );
}

// ── 耳標番号表示（6〜9文字目を大きく太字） ────────────────────────────────
const TagDisplay = ({tag, size=14, highlightSize=22, color="#4ab8e8"}) => {
  if(!tag) return null;
  const digits = tag.replace(/[^0-9]/g,"");
  if(digits.length < 9) return <span style={{color,fontWeight:900,fontSize:size,fontFamily:"monospace"}}>{tag}</span>;
  const d = digits;
  const pre  = d.slice(0,5);
  const hi   = d.slice(5,9);
  const post = d.slice(9);
  return (
    <span style={{fontFamily:"monospace",display:"inline-flex",alignItems:"baseline",gap:1}}>
      <span style={{color,fontWeight:700,fontSize:size}}>{pre}</span>
      <span style={{color,fontWeight:900,fontSize:highlightSize,letterSpacing:1}}>{hi}</span>
      <span style={{color,fontWeight:700,fontSize:size}}>{post}</span>
    </span>
  );
};
function PedigreeForm({pedigree, onChange}) {
  const set = (path, val) => {
    const keys = path.split(".");
    const next = JSON.parse(JSON.stringify(pedigree));
    let obj = next;
    for(let i=0;i<keys.length-1;i++) obj=obj[keys[i]];
    obj[keys[keys.length-1]] = val;
    onChange(next);
  };

  // helper to read nested
  const get = (obj, path) => {
    return path.split(".").reduce((o,k)=>o?.[k], obj)||"";
  };
  const Field = ({label,path}) => (
    <div style={{marginBottom:8}}>
      <div style={{color:C.textDim,fontSize:10,marginBottom:3}}>{label}</div>
      <input value={get(pedigree,path)} onChange={e=>set(path,e.target.value)}
        placeholder={label} style={{...inp,fontSize:12,padding:"7px 10px"}}/>
    </div>
  );

  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        {/* 父系 */}
        <div style={{background:C.accentLight,borderRadius:12,padding:"12px 10px"}}>
          <div style={{color:C.accentDark,fontWeight:800,fontSize:12,marginBottom:8}}>🐂 父系</div>
          <Field label="父" path="sire.name"/>
          <div style={{paddingLeft:8,borderLeft:`2px solid ${C.accentMid}`}}>
            <Field label="父の父（祖父）" path="sire.sire.name"/>
            <div style={{paddingLeft:8,borderLeft:`2px solid ${C.border}`}}>
              <Field label="父の父の父" path="sire.sire.sire.name"/>
              <Field label="父の父の母" path="sire.sire.dam.name"/>
            </div>
            <Field label="父の母（祖母）" path="sire.dam.name"/>
            <div style={{paddingLeft:8,borderLeft:`2px solid ${C.border}`}}>
              <Field label="父の母の父" path="sire.dam.sire.name"/>
              <Field label="父の母の母" path="sire.dam.dam.name"/>
            </div>
          </div>
        </div>
        {/* 母系 */}
        <div style={{background:"#fce8f0",borderRadius:12,padding:"12px 10px"}}>
          <div style={{color:"#c0407a",fontWeight:800,fontSize:12,marginBottom:8}}>🐄 母系</div>
          <Field label="母" path="dam.name"/>
          <div style={{paddingLeft:8,borderLeft:"2px solid #f0a0c0"}}>
            <Field label="母の父（祖父）" path="dam.sire.name"/>
            <div style={{paddingLeft:8,borderLeft:`2px solid ${C.border}`}}>
              <Field label="母の父の父" path="dam.sire.sire.name"/>
              <Field label="母の父の母" path="dam.sire.dam.name"/>
            </div>
            <Field label="母の母（祖母）" path="dam.dam.name"/>
            <div style={{paddingLeft:8,borderLeft:`2px solid ${C.border}`}}>
              <Field label="母の母の父" path="dam.dam.sire.name"/>
              <Field label="母の母の母" path="dam.dam.dam.name"/>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── COST FORM ──────────────────────────────────────────────────────────────────
function CostForm({costs, onChange}) {
  const set = (k,v) => onChange({...costs,[k]:Number(v)||0});
  const compoundDaily = (costs.compoundKgPerDay||0)*(costs.compoundKgPrice||0);
  const totalDaily = (costs.roughageDaily||0)+compoundDaily+(costs.otherDaily||0);
  const NI = ({label,k,unit,hint}) => (
    <FInput label={label} hint={hint}>
      <div style={{display:"flex",alignItems:"center",gap:6}}>
        <input type="number" value={costs[k]||""} onChange={e=>set(k,e.target.value)} style={{...inp,flex:1}} placeholder="0"/>
        {unit&&<span style={{color:C.textDim,fontSize:12,whiteSpace:"nowrap"}}>{unit}</span>}
      </div>
    </FInput>
  );
  return (
    <div>
      <div style={{background:C.amberLight,borderRadius:12,padding:"12px 14px",marginBottom:14}}>
        <div style={{color:C.amber,fontWeight:800,fontSize:12,marginBottom:10}}>💴 導入コスト</div>
        <NI label="素牛購入価格" k="purchasePrice" unit="円"/>
      </div>

      <div style={{background:C.greenLight,borderRadius:12,padding:"12px 14px",marginBottom:14}}>
        <div style={{color:C.green,fontWeight:800,fontSize:12,marginBottom:10}}>🌾 飼養コスト（1日あたり）</div>
        <NI label="粗飼料費" k="roughageDaily" unit="円/日" hint="牧草・稲わらなど"/>
        <div style={{borderTop:`1px solid ${C.border}`,paddingTop:12,marginBottom:8}}>
          <div style={{color:C.textMid,fontSize:11,fontWeight:700,marginBottom:8}}>配合飼料</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <NI label="給与量" k="compoundKgPerDay" unit="kg/日"/>
            <NI label="単価" k="compoundKgPrice" unit="円/kg"/>
          </div>
          <div style={{background:"#fff",borderRadius:8,padding:"8px 12px",fontSize:12,color:C.accentDark,fontWeight:700}}>
            配合コスト: ¥{compoundDaily.toLocaleString()}/日
          </div>
        </div>
        <NI label="その他日常経費" k="otherDaily" unit="円/日" hint="光熱費・消耗品など"/>
        <div style={{background:"#fff",borderRadius:8,padding:"8px 12px",fontSize:12,color:C.green,fontWeight:700,marginTop:4}}>
          飼養コスト合計: ¥{totalDaily.toLocaleString()}/日
        </div>
      </div>

      <div style={{background:C.redLight,borderRadius:12,padding:"12px 14px"}}>
        <div style={{color:C.red,fontWeight:800,fontSize:12,marginBottom:10}}>🏥 固定・医療コスト</div>
        <NI label="固定経費（通期）" k="fixedOther" unit="円" hint="施設費・減価償却など"/>
        <NI label="獣医・医薬品費（累計）" k="vetCosts" unit="円"/>
      </div>
    </div>
  );
}

// ── COST BREAKDOWN DISPLAY ─────────────────────────────────────────────────────
function CostBreakdown({cow}) {
  const costs = cow.costs||{};
  const days = daysSince(cow.introDate);
  const cv = calcCosts(cow);
  const pred = predictWeight(cow);
  const dg = calcDG(cow.weights);

  const bars = [
    {label:"素牛代",val:costs.purchasePrice||0,color:C.amber},
    {label:"粗飼料",val:(costs.roughageDaily||0)*days,color:C.green},
    {label:"配合飼料",val:cv.compoundDaily*days,color:C.teal},
    {label:"その他日常",val:(costs.otherDaily||0)*days,color:C.accent},
    {label:"固定経費",val:costs.fixedOther||0,color:C.purple},
    {label:"獣医費",val:costs.vetCosts||0,color:C.red},
  ].filter(b=>b.val>0);
  const total = bars.reduce((s,b)=>s+b.val,0)||1;

  return (
    <div>
      {/* KPIs */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
        <KpiCard label="1日あたりコスト" value={`¥${cv.dailyTotal.toLocaleString()}`} icon="📅" color={C.amber} bg={C.amberLight}/>
        <KpiCard label="累計飼養コスト" value={fmtM(cv.runningCost)} icon="📈" color={C.red} bg={C.redLight}/>
        <KpiCard label="総コスト" value={fmtM(cv.totalCost)} icon="💴" color={C.red} bg={C.redLight}/>
        <KpiCard label="予想損益" value={cv.profit!=null?(cv.profit>=0?"+":"")+fmtM(cv.profit):"―"} icon="💰" color={cv.profit!=null?(cv.profit>=0?C.green:C.red):C.textDim} bg={cv.profit!=null?(cv.profit>=0?C.greenLight:C.redLight):C.cardSub}/>
      </div>

      {/* Cost breakdown bar */}
      <Card style={{marginBottom:14}}>
        <SectionLabel>コスト構成</SectionLabel>
        <div style={{marginBottom:10}}>
          <div style={{display:"flex",height:18,borderRadius:8,overflow:"hidden",gap:1}}>
            {bars.map((b,i)=>(
              <div key={i} style={{flex:b.val/total,background:b.color,minWidth:2}}/>
            ))}
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:"4px 12px",marginTop:8}}>
            {bars.map((b,i)=>(
              <span key={i} style={{fontSize:10,color:C.textMid,display:"flex",alignItems:"center",gap:3}}>
                <span style={{width:8,height:8,borderRadius:2,background:b.color,display:"inline-block"}}/>
                {b.label} {fmtM(b.val)}
              </span>
            ))}
          </div>
        </div>
        <InfoRow label="配合飼料" value={`${costs.compoundKgPerDay||0}kg × ¥${costs.compoundKgPrice||0}/kg = ¥${cv.compoundDaily.toLocaleString()}/日`}/>
        <InfoRow label="飼養日数" value={`${days}日`}/>
        <InfoRow label="1日コスト" value={`¥${cv.dailyTotal.toLocaleString()}/日`} accent/>
        <InfoRow label="累計コスト" value={fmtMoney(cv.runningCost)}/>
        <InfoRow label="固定・医療" value={fmtMoney((costs.fixedOther||0)+(costs.vetCosts||0))}/>
        <InfoRow label="総コスト合計" value={fmtMoney(cv.totalCost)} accent big last/>
      </Card>

      {/* P&L */}
      <Card>
        <SectionLabel>損益</SectionLabel>
        <InfoRow label="予想販売価格" value={fmtMoney(cow.expectedPrice)}/>
        {dg&&pred&&<InfoRow label={`出荷予測体重（DG ${dg.toFixed(2)}kg/日）`} value={`${pred}kg`}/>}
        <div style={{background:cv.profit!=null?(cv.profit>=0?C.greenLight:C.redLight):C.cardSub,borderRadius:14,padding:"16px",marginTop:12,textAlign:"center"}}>
          <div style={{color:C.textMid,fontSize:12,marginBottom:4}}>予想損益</div>
          <div style={{color:cv.profit!=null?(cv.profit>=0?C.green:C.red):C.textDim,fontWeight:900,fontSize:26}}>
            {cv.profit!=null?(cv.profit>=0?"+":"")+fmtMoney(cv.profit):"―"}
          </div>
          {cv.profit!=null&&<div style={{color:C.textDim,fontSize:11,marginTop:4}}>{fmtM(cv.profit)} / 頭</div>}
        </div>
      </Card>
    </div>
  );
}

// ── WEIGHT CHART ───────────────────────────────────────────────────────────────
function WeightChart({weights,shippingPlan,dg}) {
  if(!weights?.length) return <div style={{color:C.textDim,textAlign:"center",padding:"24px 0",fontSize:13}}>体重データがありません</div>;
  const points=[...weights];
  const du=daysUntil(shippingPlan),lw=latestWeight(weights);
  if(dg&&lw&&du>0&&shippingPlan) points.push({date:shippingPlan,weight:Math.round(lw+dg*du),predicted:true});
  const W=320,H=140,PAD={t:20,r:14,b:30,l:42};
  const iW=W-PAD.l-PAD.r,iH=H-PAD.t-PAD.b;
  const ws=points.map(p=>p.weight),ds=points.map(p=>new Date(p.date).getTime());
  const minW=Math.min(...ws)-25,maxW=Math.max(...ws)+25,minD=Math.min(...ds),maxD=Math.max(...ds);
  const px=(d)=>PAD.l+((new Date(d).getTime()-minD)/(maxD-minD||1))*iW;
  const py=(w)=>PAD.t+(1-(w-minW)/(maxW-minW))*iH;
  const real=points.filter(p=>!p.predicted);
  const path=real.map((p,i)=>`${i===0?"M":"L"}${px(p.date)},${py(p.weight)}`).join(" ");
  const pred=points.find(p=>p.predicted),lastR=real[real.length-1];
  return (
    <svg width="100%" viewBox={`0 0 ${W} ${H}`} style={{overflow:"visible"}}>
      <defs><linearGradient id="wg3" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={C.accent} stopOpacity="0.18"/><stop offset="100%" stopColor={C.accent} stopOpacity="0"/></linearGradient></defs>
      {[0,0.5,1].map(t=>(<g key={t}><line x1={PAD.l} y1={PAD.t+t*iH} x2={PAD.l+iW} y2={PAD.t+t*iH} stroke={C.borderLight} strokeWidth={1}/><text x={PAD.l-5} y={PAD.t+t*iH+4} textAnchor="end" fill={C.textDim} fontSize={9}>{Math.round(maxW-t*(maxW-minW))}</text></g>))}
      {real.length>1&&<path d={`${path} L${px(lastR.date)},${PAD.t+iH} L${px(real[0].date)},${PAD.t+iH} Z`} fill="url(#wg3)"/>}
      {real.length>1&&<path d={path} fill="none" stroke={C.accent} strokeWidth={2.5} strokeLinejoin="round" strokeLinecap="round"/>}
      {pred&&lastR&&<path d={`M${px(lastR.date)},${py(lastR.weight)} L${px(pred.date)},${py(pred.weight)}`} fill="none" stroke={C.amber} strokeWidth={2} strokeDasharray="5,4"/>}
      {points.map((p,i)=>(
        <g key={i}>
          <circle cx={px(p.date)} cy={py(p.weight)} r={5} fill={p.predicted?C.amber:C.accent} stroke="#fff" strokeWidth={2}/>
          <text x={px(p.date)} y={py(p.weight)-10} textAnchor="middle" fill={p.predicted?C.amber:C.textMid} fontSize={9} fontWeight={700}>{p.weight}kg{p.predicted?"★":""}</text>
          <text x={px(p.date)} y={H-8} textAnchor="middle" fill={C.textDim} fontSize={8}>{p.date.slice(5)}</text>
        </g>
      ))}
    </svg>
  );
}

function BmsMeter({value}) {
  if(!value) return <span style={{color:C.textDim}}>―</span>;
  const c=value<=3?"#aac4d4":value<=5?"#f5d76e":value<=7?"#f5a623":value<=9?"#e87040":"#d04020";
  return (
    <div style={{display:"flex",alignItems:"center",gap:8}}>
      <div style={{display:"flex",gap:3}}>{Array.from({length:12},(_,i)=><div key={i} style={{width:14,height:18,borderRadius:4,background:i<value?c:C.borderLight}}/>)}</div>
      <span style={{color:c,fontWeight:900,fontSize:15}}>{value}</span>
    </div>
  );
}

// ── OCR MODAL ──────────────────────────────────────────────────────────────────
function OcrModal({onClose, onApply}) {
  const fileRef = useRef(null);
  const [status, setStatus] = useState("idle"); // idle | loading | done | error
  const [preview, setPreview] = useState(null);
  const [result, setResult] = useState(null);
  const [errMsg, setErrMsg] = useState("");

  const handleFile = async (file) => {
    if(!file) return;
    setStatus("loading");
    setResult(null);
    setErrMsg("");

    // Preview
    const reader = new FileReader();
    reader.onload = (e) => setPreview(e.target.result);
    reader.readAsDataURL(file);

    // Base64
    const b64 = await new Promise((res,rej)=>{
      const r2=new FileReader();
      r2.onload=()=>res(r2.result.split(",")[1]);
      r2.onerror=()=>rej();
      r2.readAsDataURL(file);
    });
    const mediaType = file.type||"image/jpeg";

    const prompt = `この画像は日本の和牛子牛登記証明書または個体識別票です。
以下のJSON形式で情報を抽出してください。読み取れない項目は""にしてください。

{
  "tag": "耳標番号（10桁数字またはハイフン付き）",
  "name": "牛名",
  "sex": "去勢 または 雌 または 雄",
  "breed": "品種（例: 黒毛和種）",
  "birthDate": "生年月日（YYYY-MM-DD形式）",
  "pedigree": {
    "sire": {
      "name": "父名",
      "sire": { "name": "父の父（祖父）", "sire": {"name":"父の父の父"}, "dam": {"name":"父の父の母"} },
      "dam":  { "name": "父の母（祖母）", "sire": {"name":"父の母の父"}, "dam": {"name":"父の母の母"} }
    },
    "dam": {
      "name": "母名",
      "sire": { "name": "母の父", "sire": {"name":"母の父の父"}, "dam": {"name":"母の父の母"} },
      "dam":  { "name": "母の母", "sire": {"name":"母の母の父"}, "dam": {"name":"母の母の母"} }
    }
  }
}

JSONのみ返してください。前置き・説明・バッククォートは不要です。`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":window.ANTHROPIC_KEY||"","anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
        body:JSON.stringify({
          model:"claude-sonnet-4-20250514",
          max_tokens:1000,
          messages:[{role:"user",content:[
            {type:"image",source:{type:"base64",media_type:mediaType,data:b64}},
            {type:"text",text:prompt}
          ]}]
        })
      });
      const data = await res.json();
      const text = data.content?.map(c=>c.text||"").join("") || "";
      const clean = text.replace(/```json|```/g,"").trim();
      const parsed = JSON.parse(clean);
      setResult(parsed);
      setStatus("done");
    } catch(e) {
      console.error(e);
      setErrMsg("読み取りに失敗しました。鮮明な写真で再試行してください。");
      setStatus("error");
    }
  };

  return (
    <Modal title="📷 証明書から自動入力" onClose={onClose} wide>
      {status==="idle"&&(
        <div>
          <div style={{background:C.accentLight,borderRadius:14,padding:"16px",marginBottom:16,textAlign:"center"}}>
            <div style={{fontSize:36,marginBottom:8}}>📄</div>
            <div style={{color:C.accentDark,fontWeight:700,fontSize:14,marginBottom:6}}>子牛登記証明書を撮影・選択</div>
            <div style={{color:C.textMid,fontSize:12,lineHeight:1.6}}>証明書をカメラで撮影するか、保存済みの写真を選択してください。AIが血統情報を自動で読み取ります。</div>
          </div>
          <input ref={fileRef} type="file" accept="image/*" capture="environment" style={{display:"none"}} onChange={e=>handleFile(e.target.files?.[0])}/>
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            <Btn full variant="primary" icon="📷" onClick={()=>{if(fileRef.current){fileRef.current.setAttribute("capture","environment");fileRef.current.click();}}}>カメラで撮影</Btn>
            <Btn full variant="soft" icon="🖼️" onClick={()=>{if(fileRef.current){fileRef.current.removeAttribute("capture");fileRef.current.click();}}}>写真を選択</Btn>
          </div>
        </div>
      )}
      {status==="loading"&&(
        <div style={{textAlign:"center",padding:"32px 0"}}>
          {preview&&<img src={preview} style={{width:"100%",maxHeight:200,objectFit:"contain",borderRadius:12,marginBottom:16}} alt="preview"/>}
          <div style={{fontSize:32,marginBottom:12,animation:"spin 1.5s linear infinite"}}>🔍</div>
          <div style={{color:C.accentDark,fontWeight:700,fontSize:15,marginBottom:6}}>AIが読み取り中...</div>
          <div style={{color:C.textMid,fontSize:12}}>血統情報を解析しています。少々お待ちください。</div>
          <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
        </div>
      )}
      {status==="error"&&(
        <div style={{textAlign:"center",padding:"24px 0"}}>
          {preview&&<img src={preview} style={{width:"100%",maxHeight:160,objectFit:"contain",borderRadius:12,marginBottom:12}} alt="preview"/>}
          <div style={{fontSize:32,marginBottom:8}}>❌</div>
          <div style={{color:C.red,fontWeight:700,marginBottom:6}}>{errMsg}</div>
          <Btn variant="soft" onClick={()=>{setStatus("idle");setPreview(null);}}>再試行</Btn>
        </div>
      )}
      {status==="done"&&result&&(
        <div>
          {preview&&<img src={preview} style={{width:"100%",maxHeight:160,objectFit:"contain",borderRadius:12,marginBottom:14}} alt="preview"/>}
          <div style={{background:C.greenLight,border:`1px solid ${C.green}44`,borderRadius:12,padding:"12px 14px",marginBottom:14}}>
            <div style={{color:C.green,fontWeight:700,fontSize:13,marginBottom:8}}>✅ 読み取り完了</div>
            {[
              ["耳標",result.tag],["牛名",result.name],["性別",result.sex],["品種",result.breed],["生年月日",result.birthDate],
              ["父",result.pedigree?.sire?.name],["母",result.pedigree?.dam?.name],
              ["父の父",result.pedigree?.sire?.sire?.name],["母の父",result.pedigree?.dam?.sire?.name],
            ].filter(([,v])=>v).map(([k,v])=>(
              <div key={k} style={{display:"flex",gap:8,fontSize:12,marginBottom:3}}>
                <span style={{color:C.textDim,minWidth:70}}>{k}</span>
                <span style={{color:C.text,fontWeight:600}}>{v}</span>
              </div>
            ))}
          </div>
          <div style={{color:C.textDim,fontSize:11,marginBottom:12}}>※登録後も各項目を手動で修正できます</div>
          <Btn full onClick={()=>onApply(result)}>この内容で入力する</Btn>
        </div>
      )}
    </Modal>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// MAIN APP
// ══════════════════════════════════════════════════════════════════════════════
export default function App() {
  const [cattle,       setCattle]       = useState(SAMPLE);
  const [page,         setPage]         = useState("home");
  const [selectedId,   setSelectedId]   = useState(null);
  const [detailTab,    setDetailTab]    = useState("info");
  const [search,       setSearch]       = useState("");
  const [modal,        setModal]        = useState(null);
  const [showOcr,      setShowOcr]      = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [dbReady,      setDbReady]      = useState(false);
  const [sortKey,      setSortKey]      = useState("ship");
  const [showShipped,  setShowShipped]  = useState(false);
  const [pendingOcr,   setPendingOcr]   = useState(null);
  const [addFormKey,   setAddFormKey]   = useState(0);
  const [editCosts,    setEditCosts]    = useState(false);
  const [tmpCosts,     setTmpCosts]     = useState(null);
  const [settings,     setSettings]     = useState(()=>{
    try {
      const s = localStorage.getItem('wagyu_settings');
      if(s) return JSON.parse(s);
    } catch(e) {}
    return {
      farmName: "",
      farmId: "farm_" + Math.random().toString(36).slice(2,8),
      defaultCosts:{ roughageDaily:400, compoundKgPerDay:8, compoundKgPrice:80, otherDaily:200, fixedOther:30000 },
    };
  });
  const [tmpSettings, setTmpSettings] = useState(null);

  // ── 起動時にデータ読み込み ────────────────────────────────────────────────
  useEffect(()=>{
    const load = async () => {
      try {
        if(typeof window.loadAllCattle === "function") {
          const saved = await window.loadAllCattle();
          if(saved && Array.isArray(saved) && saved.length > 0) {
            setCattle(saved);
          }
        }
      } catch(e) { console.log("読み込みエラー:", e); }
      setDbReady(true);
    };
    load();
  },[]);

  // ── 1頭ずつ自動保存 ───────────────────────────────────────────────────────
  const saveCowDebounced = useCallback((cow)=>{
    if(typeof window.saveCow === "function") {
      window.saveCow(cow).catch(e=>console.log("保存エラー:",e));
    }
  },[]);

  // cattleが変わったら差分を保存
  const prevCattleRef = useRef([]);
  useEffect(()=>{
    if(!dbReady) return;
    const prev = prevCattleRef.current;
    const timer = setTimeout(()=>{
      cattle.forEach(cow => {
        const old = prev.find(c=>c.id===cow.id);
        if(!old || JSON.stringify(old) !== JSON.stringify(cow)) {
          saveCowDebounced(cow);
        }
      });
      prevCattleRef.current = cattle;
    }, 1000);
    return ()=>clearTimeout(timer);
  },[cattle, dbReady]);

  const makeEmptyNew = () => ({
    tag:"",name:"",sex:"去勢",breed:"黒毛和種",birthDate:"",
    introDate:new Date().toISOString().slice(0,10),
    farm:"",pen:"",shippingPlan:"",expectedPrice:"",memo:"",status:"肥育中",
    pedigree:emptyPedigree(),
    costs:{ purchasePrice:0, ...settings.defaultCosts, vetCosts:0 },
  });

  const emptyNew = makeEmptyNew();

  const cow = cattle.find(c=>c.id===selectedId);
  const dg = cow ? calcDG(cow.weights) : null;
  const pred = cow ? predictWeight(cow) : null;

  const filtered = useMemo(()=>{
    let list = cattle.filter(c=>{
      if(!showShipped && c.status==="出荷済") return false;
      if(!search) return true;
      return c.tag.includes(search)||c.name.includes(search)||(c.pen||"").includes(search)||(c.pedigree?.sire?.name||"").includes(search);
    });
    list = [...list].sort((a,b)=>{
      switch(sortKey){
        case "ship":
          return (daysUntil(a.shippingPlan)??9999)-(daysUntil(b.shippingPlan)??9999);
        case "age": {
          const ma = new Date(a.birthDate||"2099").getTime();
          const mb = new Date(b.birthDate||"2099").getTime();
          return ma - mb; // 古い（月齢大）順
        }
        case "sire":
          return (a.pedigree?.sire?.name||"zzz").localeCompare(b.pedigree?.sire?.name||"zzz","ja");
        case "profit": {
          const pa = calcCosts(a).profit ?? -Infinity;
          const pb = calcCosts(b).profit ?? -Infinity;
          return pb - pa; // 利益高い順
        }
        case "introWeight": {
          const wa = a.weights?.[0]?.weight ?? 0;
          const wb = b.weights?.[0]?.weight ?? 0;
          return wb - wa; // 導入体重重い順
        }
        case "dg": {
          const da = calcDG(a.weights) ?? -Infinity;
          const db = calcDG(b.weights) ?? -Infinity;
          return db - da;
        }
        case "bms": {
          const ba = a.result?.bms ?? -Infinity;
          const bb = b.result?.bms ?? -Infinity;
          return bb - ba;
        }
        case "grade": {
          const go = {"A5":0,"B5":1,"A4":2,"B4":3,"A3":4,"B3":5,"A2":6,"B2":7,"A1":8,"B1":9};
          return (go[a.result?.grade]??99)-(go[b.result?.grade]??99);
        }
        case "sellPrice": {
          const sa = a.result?.sellPrice ?? a.expectedPrice ?? 0;
          const sb = b.result?.sellPrice ?? b.expectedPrice ?? 0;
          return sb - sa;
        }
        default: return 0;
      }
    });
    return list;
  },[cattle,search,sortKey,showShipped]);

  const monthlyData = useMemo(()=>{
    const map={};
    cattle.filter(c=>c.shippingPlan).forEach(c=>{
      const mk=monthKey(c.shippingPlan);
      if(!map[mk]) map[mk]={month:mk,label:mk.slice(2).replace("-","/"),head:0,revenue:0,profit:0,cows:[]};
      map[mk].head++;
      map[mk].revenue+=c.result?.sellPrice||c.expectedPrice||0;
      map[mk].profit+=calcCosts(c).profit||0;
      map[mk].cows.push(c);
    });
    return Object.values(map).sort((a,b)=>a.month.localeCompare(b.month)).slice(0,12);
  },[cattle]);

  // 血統レベル別集計（1=父, 2=父の父, 3=父の父の父）
  const buildLineageData = (level) => {
    const getKey = (c) => {
      if(level===1) return c.pedigree?.sire?.name;
      if(level===2) return c.pedigree?.dam?.sire?.name;   // 母の父
      if(level===3) return c.pedigree?.dam?.dam?.sire?.name; // 母の母の父
      return null;
    };
    const map={};
    cattle.forEach(c=>{
      const key=getKey(c)||"不明";
      if(!map[key]) map[key]={sire:key,head:0,dgs:[],bmsList:[],loinList:[],profits:[],cows:[]};
      map[key].head++;
      const d=calcDG(c.weights); if(d) map[key].dgs.push(d);
      if(c.result?.bms)      map[key].bmsList.push(c.result.bms);
      if(c.result?.loinArea) map[key].loinList.push(c.result.loinArea);
      const p=calcCosts(c).profit; if(p!=null) map[key].profits.push(p);
      map[key].cows.push(c);
    });
    return Object.values(map)
      .map(s=>({...s,avgDG:avg(s.dgs),avgBMS:avg(s.bmsList),avgLoin:avg(s.loinList),avgProfit:avg(s.profits)}))
      .sort((a,b)=>b.head-a.head);
  };
  const sireData = useMemo(()=>buildLineageData(1),[cattle]);

  const update = (id,fn) => setCattle(p=>p.map(c=>c.id===id?fn(c):c));
  const goDetail = (id) => { setSelectedId(id); setDetailTab("info"); setPage("detail"); };

  // Apply OCR result to new form
  const applyOcr = (result) => {
    setPendingOcr(result);
    setShowOcr(false);
    setPage("add");
  };
  const mergeNode = (base, src) => {
    if(!src) return base;
    return {
      name: src.name||base.name,
      sire: mergeNode(base.sire||{name:"",sire:{name:""},dam:{name:""}}, src.sire),
      dam:  mergeNode(base.dam ||{name:"",sire:{name:""},dam:{name:""}}, src.dam),
    };
  };

  // ── APP HEADER ─────────────────────────────────────────────────────────────
  const AppHeader = ({subtitle, showGear=false}) => (
    <div style={{background:"#fff",borderBottom:`1px solid ${C.border}`,padding:"12px 16px 10px",boxShadow:"0 1px 8px rgba(74,184,232,0.08)",position:"sticky",top:0,zIndex:80}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div style={{display:"flex",alignItems:"baseline",gap:7,flexWrap:"wrap"}}>
          <span style={{background:`linear-gradient(135deg,${C.accent},${C.accentDark})`,WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",fontSize:22,fontWeight:900,letterSpacing:2,fontFamily:"Georgia,serif"}}>WAGYU AI</span>
          {settings.farmName&&<span style={{color:C.text,fontSize:14,fontWeight:700}}>{settings.farmName}</span>}
          {subtitle&&<span style={{color:C.textDim,fontSize:11}}>{subtitle}</span>}
        </div>
        {(showGear||!subtitle)&&(
          <button onClick={()=>{setTmpSettings(JSON.parse(JSON.stringify(settings)));setShowSettings(true);}} style={{background:C.cardSub,border:`1px solid ${C.border}`,color:C.textMid,borderRadius:10,width:34,height:34,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",fontSize:16}}>⚙️</button>
        )}
      </div>
    </div>
  );

  // ── 設定モーダル ────────────────────────────────────────────────────────────
  const SettingsModal = () => {
    if(!tmpSettings) return null;
    const [loc, setLoc] = React.useState(tmpSettings);
    const sf = (k,v) => setLoc(p=>({...p,[k]:v}));
    const sc = (k,v) => setLoc(p=>({...p,defaultCosts:{...p.defaultCosts,[k]:Number(v)||0}}));
    const cd = (loc.defaultCosts.compoundKgPerDay||0)*(loc.defaultCosts.compoundKgPrice||0);
    const td = (loc.defaultCosts.roughageDaily||0)+cd+(loc.defaultCosts.otherDaily||0);
    const save = () => {
      setSettings(loc); setTmpSettings(loc);
      try { localStorage.setItem('wagyu_settings', JSON.stringify(loc)); } catch(e) {}
      setShowSettings(false);
    };
    const numRow = (label, val, onChange, unit, hint) => (
      <div style={{marginBottom:14}}>
        <div style={{color:C.textMid,fontSize:12,fontWeight:600,marginBottom:5}}>{label}</div>
        <div style={{display:"flex",alignItems:"center",gap:6}}>
          <input type="number" value={val||""} onChange={onChange} style={{...inp,flex:1}} placeholder="0"/>
          {unit&&<span style={{color:C.textDim,fontSize:12,whiteSpace:"nowrap"}}>{unit}</span>}
        </div>
        {hint&&<div style={{color:C.textDim,fontSize:10,marginTop:3}}>{hint}</div>}
      </div>
    );
    return (
      <div style={{position:"fixed",inset:0,background:"rgba(30,58,74,0.4)",zIndex:400,display:"flex",alignItems:"flex-end",justifyContent:"center",backdropFilter:"blur(3px)"}}>
        <div style={{background:"#fff",borderRadius:"24px 24px 0 0",width:"100%",maxWidth:520,maxHeight:"92vh",overflowY:"auto",padding:"8px 20px 40px",boxShadow:"0 -8px 40px rgba(74,184,232,0.18)"}}>
          <div style={{width:40,height:4,background:C.border,borderRadius:2,margin:"12px auto 18px"}}/>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:22}}>
            <div>
              <div style={{color:C.text,fontWeight:900,fontSize:17}}>⚙️ 農場設定</div>
              <div style={{color:C.textDim,fontSize:11,marginTop:2}}>設定したコストは新規導入時のデフォルトに反映されます</div>
            </div>
            <button onClick={()=>setShowSettings(false)} style={{background:C.accentLight,border:"none",color:C.textMid,width:32,height:32,borderRadius:"50%",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
          </div>

          <div style={{background:C.accentLight,borderRadius:14,padding:"14px 16px",marginBottom:16,border:`1px solid ${C.border}`}}>
            <div style={{color:C.accentDark,fontWeight:800,fontSize:13,marginBottom:10}}>🏡 農場情報</div>
            <div style={{marginBottom:14}}>
              <div style={{color:C.textMid,fontSize:12,fontWeight:600,marginBottom:5}}>農場名</div>
              <input value={loc.farmName} onChange={e=>sf("farmName",e.target.value)} placeholder="例: 田中和牛農場" style={inp}/>
            </div>
            <div style={{marginBottom:8}}>
              <div style={{color:C.textMid,fontSize:12,fontWeight:600,marginBottom:5}}>農場ID（複数スマホで同じIDにすると同じデータが見れます）</div>
              <input value={loc.farmId||""} onChange={e=>sf("farmId",e.target.value)} placeholder="例: tanaka_farm_001" style={inp}/>
            </div>
            <div style={{color:C.textDim,fontSize:11}}>⚠️ 全スマホで同じIDにしてください</div>
          </div>

          <div style={{background:C.greenLight,borderRadius:14,padding:"14px 16px",marginBottom:16,border:`1px solid ${C.green}22`}}>
            <div style={{color:C.green,fontWeight:800,fontSize:13,marginBottom:12}}>🌾 基本飼養コスト</div>
            {numRow("粗飼料費（円/日）", loc.defaultCosts.roughageDaily, e=>sc("roughageDaily",e.target.value), "円/日", "牧草・稲わらなど")}
            <div style={{borderTop:`1px solid ${C.border}`,paddingTop:12,marginBottom:10}}>
              <div style={{color:C.textMid,fontSize:12,fontWeight:700,marginBottom:8}}>配合飼料</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                {numRow("給与量（kg/日）", loc.defaultCosts.compoundKgPerDay, e=>sc("compoundKgPerDay",e.target.value), "kg")}
                {numRow("単価（円/kg）",   loc.defaultCosts.compoundKgPrice,   e=>sc("compoundKgPrice",e.target.value),   "円")}
              </div>
              <div style={{background:"#fff",borderRadius:8,padding:"8px 12px",fontSize:12,color:C.green,fontWeight:700}}>
                配合コスト: ¥{cd.toLocaleString()}/日
              </div>
            </div>
            {numRow("その他日常経費（円/日）", loc.defaultCosts.otherDaily, e=>sc("otherDaily",e.target.value), "円/日", "光熱費など")}
            <div style={{background:"#fff",border:`1.5px solid ${C.green}44`,borderRadius:10,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{color:C.textMid,fontSize:13}}>飼養コスト合計</span>
              <span style={{color:C.green,fontWeight:900,fontSize:18}}>¥{td.toLocaleString()}<span style={{fontSize:11,fontWeight:400}}>/日</span></span>
            </div>
          </div>

          <div style={{background:C.amberLight,borderRadius:14,padding:"14px 16px",marginBottom:20,border:`1px solid ${C.amber}22`}}>
            <div style={{color:C.amber,fontWeight:800,fontSize:13,marginBottom:10}}>🏗️ 固定経費（通期）</div>
            {numRow("固定経費（円）", loc.defaultCosts.fixedOther, e=>sc("fixedOther",e.target.value), "円", "施設費・減価償却など")}
          </div>

          <Btn full onClick={save}>設定を保存する</Btn>
        </div>
      </div>
    );
  };

  // ── HOME ───────────────────────────────────────────────────────────────────
  const HomeScreen = () => {
    const now       = new Date();
    const thisYM    = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`; // "2026-05"
    // 今期: 農業年度（4月始まり）
    const fiscalStart = now.getMonth() >= 3
      ? `${now.getFullYear()}-04`
      : `${now.getFullYear()-1}-04`;
    const fiscalEnd   = now.getMonth() >= 3
      ? `${now.getFullYear()+1}-03`
      : `${now.getFullYear()}-03`;

    const active      = cattle.filter(c => c.status === "肥育中");
    const soon30      = active.filter(c => { const d=daysUntil(c.shippingPlan); return d!==null&&d<=30; });
    const vaxSoon     = cattle.flatMap(c => c.vaccines.filter(v=>v.nextDate).map(v=>({cow:c,v,days:daysUntil(v.nextDate)}))).filter(x=>x.days!==null&&x.days<=30);
    const has30       = soon30.length > 0;
    const hasVax      = vaxSoon.length > 0;

    // 導入総額（肥育中のみ）
    const activePurchaseTotal = active.reduce((s,c) => s+(c.costs?.purchasePrice||0), 0);

    // 今月の利益：shippingPlan が今月の個体の損益合計
    const thisMonthCows   = cattle.filter(c => c.shippingPlan?.startsWith(thisYM));
    const thisMonthProfit = thisMonthCows.reduce((s,c) => {
      const p = calcCosts(c).profit;
      return p != null ? s + p : s;
    }, 0);
    const hasThisMonth = thisMonthCows.length > 0;

    // 今期の利益：shippingPlan が今期内の個体の損益合計
    const thisFiscalCows   = cattle.filter(c => c.shippingPlan && c.shippingPlan >= fiscalStart+"-01" && c.shippingPlan <= fiscalEnd+"-31");
    const thisFiscalProfit = thisFiscalCows.reduce((s,c) => {
      const p = calcCosts(c).profit;
      return p != null ? s + p : s;
    }, 0);
    const hasFiscal = thisFiscalCows.length > 0;

    return (
      <div style={{paddingBottom:90}}>
        <AppHeader subtitle="和牛AI管理システム" showGear={true}/>
        <div style={{padding:"16px 16px"}}>

          {/* KPI 3行 */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
            {/* 行1: 肥育中 ／ 導入総額（肥育中） */}
            <KpiCard label="肥育中" value={active.length} unit="頭"
              icon="🐂"
              color={C.accent} bg={C.accentLight}/>
            <KpiCard label="導入総額（肥育中）" value={fmtM(activePurchaseTotal)} icon="💴"
              color={C.accentDark} bg={C.accentLight}/>

            {/* 行2: 出荷30日以内 ／ 出荷90日以内 */}
            <KpiCard label="出荷30日以内" value={soon30.length} unit="頭" icon="🚚"
              color={soon30.length>0?C.red:C.textDim}
              bg   ={soon30.length>0?C.redLight:C.cardSub}/>
            <KpiCard label="出荷90日以内"
              value={active.filter(c=>{const d=daysUntil(c.shippingPlan);return d!==null&&d<=90;}).length}
              unit="頭" icon="📅"
              color={C.amber} bg={C.amberLight}/>

            {/* 行3: 今月の利益 ／ 今期の利益 */}
            <div style={{
              background: hasThisMonth ? (thisMonthProfit>=0?C.greenLight:C.redLight) : C.cardSub,
              borderRadius:14, padding:"12px 14px",
              border:`1px solid ${hasThisMonth?(thisMonthProfit>=0?C.green:C.red)+"22":C.border}`,
            }}>
              <div style={{display:"flex",alignItems:"center",gap:4,marginBottom:5}}>
                <span style={{fontSize:14}}>📆</span>
                <span style={{color:C.textMid,fontSize:10,fontWeight:600}}>今月の利益</span>
              </div>
              <div style={{color:hasThisMonth?(thisMonthProfit>=0?C.green:C.red):C.textDim,fontSize:16,fontWeight:900,lineHeight:1}}>
                {hasThisMonth ? (thisMonthProfit>=0?"+":"")+fmtM(thisMonthProfit) : "―"}
              </div>
              {hasThisMonth&&<div style={{color:C.textDim,fontSize:9,marginTop:3}}>{thisMonthCows.length}頭対象</div>}
            </div>
            <div style={{
              background: hasFiscal ? (thisFiscalProfit>=0?C.greenLight:C.redLight) : C.cardSub,
              borderRadius:14, padding:"12px 14px",
              border:`1px solid ${hasFiscal?(thisFiscalProfit>=0?C.green:C.red)+"22":C.border}`,
            }}>
              <div style={{display:"flex",alignItems:"center",gap:4,marginBottom:5}}>
                <span style={{fontSize:14}}>📊</span>
                <span style={{color:C.textMid,fontSize:10,fontWeight:600}}>今期の利益</span>
              </div>
              <div style={{color:hasFiscal?(thisFiscalProfit>=0?C.green:C.red):C.textDim,fontSize:16,fontWeight:900,lineHeight:1}}>
                {hasFiscal ? (thisFiscalProfit>=0?"+":"")+fmtM(thisFiscalProfit) : "―"}
              </div>
              {hasFiscal&&<div style={{color:C.textDim,fontSize:9,marginTop:3}}>{thisFiscalCows.length}頭対象</div>}
            </div>
          </div>

          {/* アラートバナー（出荷30日 or ワクチン期限あり） */}
          {(has30||hasVax)&&(
            <div onClick={()=>setPage("alerts")} style={{
              background:`linear-gradient(135deg,${C.red}15,${C.amber}08)`,
              border:`1.5px solid ${C.red}44`, borderRadius:14,
              padding:"12px 16px", marginBottom:14, cursor:"pointer",
              display:"flex", alignItems:"center", justifyContent:"space-between",
            }}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:26}}>🔔</span>
                <div>
                  <div style={{color:C.red,fontWeight:800,fontSize:14}}>要対応アラートあり</div>
                  <div style={{color:C.textMid,fontSize:11,marginTop:2}}>
                    {has30&&`🚚 出荷30日以内 ${soon30.length}頭　`}
                    {hasVax&&`💉 ワクチン期限 ${vaxSoon.length}件`}
                  </div>
                </div>
              </div>
              <div style={{background:C.red,color:"#fff",borderRadius:20,padding:"5px 14px",fontSize:13,fontWeight:800}}>確認 →</div>
            </div>
          )}
          {/* ━━ 新規導入スタートボタン ━━ */}
          <div onClick={()=>setPage("intake")} style={{
            background:`linear-gradient(135deg, ${C.accent}, ${C.accentDark})`,
            borderRadius:18, padding:"18px 20px", marginBottom:10, cursor:"pointer",
            boxShadow:`0 6px 24px ${C.accent}44`,
            display:"flex", alignItems:"center", justifyContent:"space-between",
          }}>
            <div style={{display:"flex",alignItems:"center",gap:14}}>
              <div style={{
                width:52, height:52, borderRadius:16,
                background:"rgba(255,255,255,0.22)",
                display:"flex", alignItems:"center", justifyContent:"center", fontSize:30,
              }}>🐂</div>
              <div>
                <div style={{color:"#fff",fontWeight:900,fontSize:18,letterSpacing:0.5}}>新規導入スタート</div>
                <div style={{color:"rgba(255,255,255,0.8)",fontSize:12,marginTop:3}}>
                  伝票撮影 → 登記書撮影 → 一括登録
                </div>
              </div>
            </div>
            <div style={{
              background:"rgba(255,255,255,0.25)", borderRadius:12,
              padding:"8px 14px", color:"#fff", fontWeight:800, fontSize:14,
            }}>START →</div>
          </div>

          {/* ━━ 出荷成績入力ボタン ━━ */}
          <div onClick={()=>setPage("shipResult")} style={{
            background:`linear-gradient(135deg, ${C.amber}, #e08a10)`,
            borderRadius:18, padding:"16px 20px", marginBottom:14, cursor:"pointer",
            boxShadow:`0 5px 20px ${C.amber}44`,
            display:"flex", alignItems:"center", justifyContent:"space-between",
          }}>
            <div style={{display:"flex",alignItems:"center",gap:14}}>
              <div style={{
                width:48, height:48, borderRadius:14,
                background:"rgba(255,255,255,0.22)",
                display:"flex", alignItems:"center", justifyContent:"center", fontSize:26,
              }}>🏆</div>
              <div>
                <div style={{color:"#fff",fontWeight:900,fontSize:17,letterSpacing:0.5}}>出荷成績入力</div>
                <div style={{color:"rgba(255,255,255,0.8)",fontSize:12,marginTop:3}}>
                  出荷伝票（PDF・写真）から複数頭一括読み込み
                </div>
              </div>
            </div>
            <div style={{
              background:"rgba(255,255,255,0.25)", borderRadius:12,
              padding:"7px 14px", color:"#fff", fontWeight:800, fontSize:13,
            }}>入力 →</div>
          </div>

          <div style={{display:"flex",gap:8,marginBottom:10}}>
            <div style={{flex:1,position:"relative"}}>
              <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:C.textDim}}>🔍</span>
              <input placeholder="耳標・名前・牛舎・血統で検索" value={search} onChange={e=>setSearch(e.target.value)} style={{...inp,paddingLeft:34,borderRadius:24}}/>
            </div>
            <Btn variant="outline" onClick={()=>{setAddFormKey(k=>k+1); setPage("add");}}>1頭追加</Btn>
          </div>

          {/* ── 並び替え ── */}
          <div style={{marginBottom:10}}>
            <div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:4,scrollbarWidth:"none"}}>
              {[
                {key:"ship",       label:"📅 出荷日"},
                {key:"age",        label:"🐂 月齢"},
                {key:"sire",       label:"🧬 血統"},
                {key:"profit",     label:"💴 利益"},
                {key:"introWeight",label:"⚖️ 導入体重"},
                {key:"dg",         label:"📈 DG"},
                {key:"bms",        label:"🥩 BMS"},
                {key:"grade",      label:"🏆 等級"},
                {key:"sellPrice",  label:"💰 販売額"},
              ].map(({key,label})=>(
                <button key={key} onClick={()=>setSortKey(key)} style={{
                  background:sortKey===key?`linear-gradient(135deg,${C.accent},${C.accentDark})`:"#fff",
                  color:sortKey===key?"#fff":C.textMid,
                  border:`1.5px solid ${sortKey===key?C.accent:C.border}`,
                  borderRadius:20, padding:"6px 14px",
                  fontSize:11, fontWeight:700, cursor:"pointer",
                  whiteSpace:"nowrap", flexShrink:0,
                  boxShadow:sortKey===key?C.shadow:"none",
                }}>{label}</button>
              ))}
            </div>
          </div>

          {/* ── 出荷済み表示トグル ── */}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <SectionLabel>個体一覧（{filtered.length}頭）</SectionLabel>
            <button onClick={()=>setShowShipped(v=>!v)} style={{
              background:showShipped?C.accentLight:"#fff",
              color:showShipped?C.accentDark:C.textDim,
              border:`1.5px solid ${showShipped?C.accent:C.border}`,
              borderRadius:20, padding:"5px 12px",
              fontSize:11, fontWeight:700, cursor:"pointer",
            }}>
              {showShipped?"✓ 出荷済み表示中":"出荷済みを表示"}
            </button>
          </div>

          {filtered.map(c=>{
            const dg_c=calcDG(c.weights),du=daysUntil(c.shippingPlan),lw=latestWeight(c.weights),cv=calcCosts(c);
            const urgent=du!==null&&du<=60&&c.status!=="出荷済";
            const shipped=c.status==="出荷済";
            const introW=c.weights?.[0]?.weight;
            const ageStr=calcAge(c.birthDate);
            return (
              <div key={c.id} onClick={()=>goDetail(c.id)} style={{
                background: shipped?"#f8f8f8":"#fff",
                border:`1.5px solid ${urgent?C.amber+"66":shipped?C.border:C.border}`,
                borderRadius:18,padding:"14px 16px",marginBottom:10,cursor:"pointer",
                boxShadow:urgent?`0 2px 12px ${C.amber}22`:C.shadow,
                opacity:shipped?0.85:1,
              }}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                  <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                    <TagDisplay tag={c.tag} size={13} highlightSize={17} color={C.accent}/>
                    <span style={{color:C.text,fontWeight:700,fontSize:15}}>{c.name}</span>
                  </div>
                  <div style={{display:"flex",gap:4,flexWrap:"wrap",justifyContent:"flex-end"}}>
                    <Tag label={c.sex} color={c.sex==="雌"?"#e06090":C.purple}/>
                    {shipped&&<Tag label="出荷済" color={C.textDim} bg="#efefef"/>}
                    {urgent&&<Tag label={`🚚 ${du}日前`} color={C.amber} bg={C.amberLight}/>}
                  </div>
                </div>

                {/* 父血統 */}
                {c.pedigree?.sire?.name&&(
                  <div style={{fontSize:11,color:C.textDim,marginBottom:8,display:"flex",alignItems:"center",gap:4}}>
                    <span>🐂</span>
                    <span style={{fontWeight:600,color:C.purple}}>{c.pedigree.sire.name}</span>
                    {c.pedigree.sire.sire?.name&&<span style={{color:C.textDim}}> ／ {c.pedigree.sire.sire.name}</span>}
                    {ageStr!=="―"&&<span style={{marginLeft:"auto",color:C.textDim}}>{ageStr}</span>}
                  </div>
                )}

                {/* 肥育中：体重・DG・牛舎 */}
                {!shipped&&(
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:8}}>
                    {[
                      ["⚖️ 最新体重", lw?`${lw}kg`:"未計測"],
                      ["📈 DG",       dg_c?`+${dg_c.toFixed(2)}`:"―"],
                      ["📦 導入体重", introW?`${introW}kg`:"―"],
                    ].map(([k,v])=>(
                      <div key={k} style={{background:C.cardSub,borderRadius:10,padding:"6px 10px"}}>
                        <div style={{color:C.textDim,fontSize:9,marginBottom:2}}>{k}</div>
                        <div style={{color:C.text,fontSize:12,fontWeight:700}}>{v}</div>
                      </div>
                    ))}
                  </div>
                )}

                {/* 出荷済み：成績サマリー */}
                {shipped&&c.result&&(
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:5,marginBottom:8}}>
                    {[
                      {label:"等級",     val:c.result.grade,                                    color: c.result.grade?.startsWith("A5")?"#d04020":c.result.grade?.startsWith("A4")?C.amber:C.green},
                      {label:"BMS",      val:c.result.bms!=null?`${c.result.bms}`:null,          color:C.red},
                      {label:"販売額",   val:fmtM(c.result.sellPrice),                          color:C.amber},
                      {label:"枝肉重量", val:c.result.coldWeight?`${c.result.coldWeight}kg`:null, color:C.accentDark},
                      {label:"枝肉DG",  val:c.result.dg?`${c.result.dg}kg/日`:null,             color:C.green},
                      {label:"歩留",     val:c.result.yieldGrade||null,                          color:C.purple},
                    ].map(({label,val,color})=>(
                      <div key={label} style={{background:C.cardSub,borderRadius:8,padding:"5px 8px"}}>
                        <div style={{color:C.textDim,fontSize:9,marginBottom:1}}>{label}</div>
                        <div style={{color:val!=null?color:C.textDim,fontSize:12,fontWeight:700}}>{val??"―"}</div>
                      </div>
                    ))}
                  </div>
                )}

                {/* 損益 */}
                {cv.profit!=null&&(
                  <div style={{display:"flex",justifyContent:"flex-end"}}>
                    <span style={{
                      background:cv.profit>=0?C.greenLight:C.redLight,
                      color:cv.profit>=0?C.green:C.red,
                      borderRadius:20,padding:"3px 12px",fontSize:12,fontWeight:700,
                    }}>
                      {shipped?"確定":"予想"}損益 {cv.profit>=0?"+":""}{fmtMoney(cv.profit)}
                    </span>
                  </div>
                )}
              </div>
            );
          })}
          {filtered.length===0&&(
            <div style={{textAlign:"center",padding:"32px 0",color:C.textDim,fontSize:13}}>
              {showShipped?"該当する個体がいません":"肥育中の個体がいません"}
            </div>
          )}
        </div>
      </div>
    );
  };

  // ── ALERTS ────────────────────────────────────────────────────────────────
  const AlertsScreen = () => {
    const today = Date.now();

    // 出荷30日以内（肥育中のみ）
    const shipSoon = cattle
      .filter(c => c.status !== "出荷済" && c.shippingPlan)
      .map(c => ({ cow:c, days: daysUntil(c.shippingPlan) }))
      .filter(x => x.days !== null && x.days <= 30)
      .sort((a,b) => a.days - b.days);

    // ワクチン期限（過去〜60日以内）
    const vaxAlerts = cattle
      .flatMap(c => c.vaccines
        .filter(v => v.nextDate)
        .map(v => ({ cow:c, v, days: daysUntil(v.nextDate) }))
      )
      .filter(x => x.days !== null && x.days <= 60)
      .sort((a,b) => a.days - b.days);

    const urgentVax  = vaxAlerts.filter(x => x.days <= 0);
    const soonVax    = vaxAlerts.filter(x => x.days > 0 && x.days <= 14);
    const nearVax    = vaxAlerts.filter(x => x.days > 14 && x.days <= 60);

    const totalAlerts = shipSoon.length + urgentVax.length + soonVax.length;

    const AlertCard = ({icon, title, subtitle, days, color, bg, onClick}) => (
      <div onClick={onClick} style={{
        display:"flex", alignItems:"center", gap:14,
        background:"#fff", border:`1.5px solid ${color}44`,
        borderLeft:`4px solid ${color}`,
        borderRadius:14, padding:"14px 16px", marginBottom:8,
        cursor:"pointer", boxShadow:`0 2px 10px ${color}15`,
      }}>
        <div style={{fontSize:28,flexShrink:0}}>{icon}</div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{color:C.text,fontWeight:700,fontSize:14,marginBottom:2}}>{title}</div>
          <div style={{color:C.textMid,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{subtitle}</div>
        </div>
        <div style={{textAlign:"right",flexShrink:0}}>
          <div style={{
            background: days < 0 ? C.red : bg,
            color: days < 0 ? "#fff" : color,
            borderRadius:20, padding:"4px 10px",
            fontSize:12, fontWeight:800, whiteSpace:"nowrap",
          }}>
            {days < 0 ? `${Math.abs(days)}日超過` : days === 0 ? "今日" : `あと${days}日`}
          </div>
        </div>
      </div>
    );

    return (
      <div style={{paddingBottom:90}}>
        <AppHeader subtitle="アラート"/>

        {/* 全体サマリー */}
        <div style={{
          background: totalAlerts > 0
            ? `linear-gradient(135deg, ${C.red}15, ${C.amber}10)`
            : `linear-gradient(135deg, ${C.greenLight}, ${C.accentLight})`,
          borderBottom:`1px solid ${C.border}`,
          padding:"16px 20px",
        }}>
          {totalAlerts > 0 ? (
            <div style={{display:"flex",alignItems:"center",gap:12}}>
              <div style={{fontSize:36}}>🔔</div>
              <div>
                <div style={{color:C.red,fontWeight:900,fontSize:18}}>要対応 {totalAlerts}件</div>
                <div style={{color:C.textMid,fontSize:12,marginTop:2}}>
                  {shipSoon.length > 0 && `出荷30日以内 ${shipSoon.length}頭　`}
                  {(urgentVax.length+soonVax.length) > 0 && `ワクチン要対応 ${urgentVax.length+soonVax.length}件`}
                </div>
              </div>
            </div>
          ) : (
            <div style={{display:"flex",alignItems:"center",gap:12}}>
              <div style={{fontSize:36}}>✅</div>
              <div>
                <div style={{color:C.green,fontWeight:900,fontSize:18}}>問題ありません</div>
                <div style={{color:C.textMid,fontSize:12,marginTop:2}}>現在、緊急のアラートはありません</div>
              </div>
            </div>
          )}
        </div>

        <div style={{padding:"16px 16px"}}>

          {/* ── 出荷30日以内 ── */}
          <div style={{
            display:"flex", alignItems:"center", gap:8,
            marginBottom:12,
          }}>
            <div style={{
              background: shipSoon.length > 0 ? C.red : C.green,
              color:"#fff", borderRadius:20, padding:"4px 12px",
              fontSize:12, fontWeight:800,
            }}>🚚 出荷30日以内</div>
            <div style={{
              background: shipSoon.length > 0 ? C.redLight : C.greenLight,
              color: shipSoon.length > 0 ? C.red : C.green,
              borderRadius:20, padding:"3px 10px", fontSize:12, fontWeight:700,
            }}>{shipSoon.length}頭</div>
          </div>

          {shipSoon.length === 0 ? (
            <div style={{background:C.greenLight,borderRadius:12,padding:"14px 16px",marginBottom:20,display:"flex",gap:10,alignItems:"center"}}>
              <span style={{fontSize:22}}>✅</span>
              <span style={{color:C.green,fontSize:13,fontWeight:600}}>30日以内に出荷予定の個体はいません</span>
            </div>
          ) : (
            <div style={{marginBottom:20}}>
              {shipSoon.map(({cow:c, days}) => {
                const lw = latestWeight(c.weights);
                const dg_c = calcDG(c.weights);
                const cv = calcCosts(c);
                return (
                  <div key={c.id}>
                    <AlertCard
                      icon={days <= 7 ? "🚨" : "🚚"}
                      title={`${c.name}（${c.tag}）`}
                      subtitle={`${c.pen}　体重${lw?lw+"kg":"未計測"}　${dg_c?`DG+${dg_c.toFixed(2)}`:""}`}
                      days={days}
                      color={days <= 7 ? C.red : C.amber}
                      bg={days <= 7 ? C.redLight : C.amberLight}
                      onClick={()=>goDetail(c.id)}
                    />
                    {/* 出荷チェックリスト */}
                    <div style={{
                      background:C.cardSub, borderRadius:12, padding:"12px 14px",
                      marginBottom:10, marginTop:-4,
                      border:`1px solid ${C.borderLight}`,
                    }}>
                      <div style={{color:C.textDim,fontSize:10,fontWeight:700,letterSpacing:1,marginBottom:8}}>出荷チェックリスト</div>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
                        {[
                          {label:"出荷予定日",val:fmtDate(c.shippingPlan),ok:!!c.shippingPlan},
                          {label:"予想販売価格",val:fmtM(c.expectedPrice),ok:!!c.expectedPrice},
                          {label:"最新体重",val:lw?`${lw}kg`:"未計測",ok:!!lw},
                          {label:"予想損益",val:cv.profit!=null?(cv.profit>=0?"+":"")+fmtM(cv.profit):"―",ok:cv.profit!=null},
                        ].map(({label,val,ok})=>(
                          <div key={label} style={{background:"#fff",borderRadius:8,padding:"7px 10px",border:`1px solid ${ok?C.green+"44":C.red+"33"}`}}>
                            <div style={{display:"flex",alignItems:"center",gap:4,marginBottom:2}}>
                              <span style={{fontSize:10}}>{ok?"✅":"⚠️"}</span>
                              <span style={{color:C.textDim,fontSize:9}}>{label}</span>
                            </div>
                            <div style={{color:ok?C.text:C.red,fontSize:12,fontWeight:700}}>{val}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* ── ワクチン期限 ── */}
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
            <div style={{
              background: urgentVax.length > 0 ? C.red : soonVax.length > 0 ? C.amber : C.green,
              color:"#fff", borderRadius:20, padding:"4px 12px", fontSize:12, fontWeight:800,
            }}>💉 ワクチン期限</div>
            <div style={{
              background: urgentVax.length > 0 ? C.redLight : soonVax.length > 0 ? C.amberLight : C.greenLight,
              color: urgentVax.length > 0 ? C.red : soonVax.length > 0 ? C.amber : C.green,
              borderRadius:20, padding:"3px 10px", fontSize:12, fontWeight:700,
            }}>{vaxAlerts.length}件</div>
          </div>

          {/* 期限超過 */}
          {urgentVax.length > 0 && (
            <div style={{marginBottom:6}}>
              <div style={{color:C.red,fontSize:11,fontWeight:700,marginBottom:6,display:"flex",alignItems:"center",gap:6}}>
                <span style={{background:C.red,color:"#fff",borderRadius:10,padding:"2px 8px",fontSize:10}}>期限超過</span>
              </div>
              {urgentVax.map((x,i) => (
                <AlertCard key={i}
                  icon="🚨"
                  title={`${x.cow.name}（${x.cow.tag}）`}
                  subtitle={x.v.name}
                  days={x.days}
                  color={C.red} bg={C.redLight}
                  onClick={()=>goDetail(x.cow.id)}
                />
              ))}
            </div>
          )}

          {/* 14日以内 */}
          {soonVax.length > 0 && (
            <div style={{marginBottom:6}}>
              <div style={{color:C.amber,fontSize:11,fontWeight:700,marginBottom:6}}>
                <span style={{background:C.amberLight,color:C.amber,border:`1px solid ${C.amber}44`,borderRadius:10,padding:"2px 8px",fontSize:10}}>14日以内</span>
              </div>
              {soonVax.map((x,i) => (
                <AlertCard key={i}
                  icon="⚠️"
                  title={`${x.cow.name}（${x.cow.tag}）`}
                  subtitle={x.v.name}
                  days={x.days}
                  color={C.amber} bg={C.amberLight}
                  onClick={()=>goDetail(x.cow.id)}
                />
              ))}
            </div>
          )}

          {/* 60日以内 */}
          {nearVax.length > 0 && (
            <div style={{marginBottom:6}}>
              <div style={{color:C.textDim,fontSize:11,fontWeight:700,marginBottom:6}}>
                <span style={{background:C.cardSub,color:C.textMid,border:`1px solid ${C.border}`,borderRadius:10,padding:"2px 8px",fontSize:10}}>60日以内</span>
              </div>
              {nearVax.map((x,i) => (
                <AlertCard key={i}
                  icon="💉"
                  title={`${x.cow.name}（${x.cow.tag}）`}
                  subtitle={x.v.name}
                  days={x.days}
                  color={C.accentDark} bg={C.accentLight}
                  onClick={()=>goDetail(x.cow.id)}
                />
              ))}
            </div>
          )}

          {vaxAlerts.length === 0 && (
            <div style={{background:C.greenLight,borderRadius:12,padding:"14px 16px",display:"flex",gap:10,alignItems:"center"}}>
              <span style={{fontSize:22}}>✅</span>
              <span style={{color:C.green,fontSize:13,fontWeight:600}}>期限が近いワクチンはありません</span>
            </div>
          )}
        </div>
      </div>
    );
  };

  // ── SCHEDULE ───────────────────────────────────────────────────────────────
  const ScheduleScreen = () => {
    const [selMonth,setSelMonth]=useState(null);
    const totalRev=monthlyData.reduce((s,m)=>s+m.revenue,0);
    const totalHead=monthlyData.reduce((s,m)=>s+m.head,0);
    const totalProfit=monthlyData.reduce((s,m)=>s+m.profit,0);
    return (
      <div style={{paddingBottom:90}}>
        <AppHeader subtitle="月別出荷・売上予定"/>
        <div style={{padding:"16px 16px"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:16}}>
            <KpiCard label="出荷合計" value={`${totalHead}頭`} icon="🐄" color={C.accent} bg={C.accentLight}/>
            <KpiCard label="予定売上" value={fmtM(totalRev)} icon="💴" color={C.amber} bg={C.amberLight}/>
            <KpiCard label="予想利益" value={fmtM(totalProfit)} icon="📊" color={totalProfit>=0?C.green:C.red} bg={totalProfit>=0?C.greenLight:C.redLight}/>
          </div>
          <SectionLabel>月別詳細（タップで個体一覧）</SectionLabel>
          {monthlyData.map(m=>{
            const open=selMonth===m.month;
            return (
              <div key={m.month} style={{marginBottom:10}}>
                <div onClick={()=>setSelMonth(open?null:m.month)} style={{background:open?C.accentLight:"#fff",border:`1.5px solid ${open?C.accent:C.border}`,borderRadius:16,padding:"14px 16px",cursor:"pointer",boxShadow:open?`0 4px 20px ${C.accent}22`:C.shadow}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                    <span style={{color:C.text,fontWeight:800,fontSize:16}}>{m.month.replace("-","年")}月</span>
                    <div style={{display:"flex",gap:6,alignItems:"center"}}><Tag label={`${m.head}頭`} color={C.accent}/><span style={{color:C.textDim}}>{open?"▲":"▼"}</span></div>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                    {[{label:"予定売上",val:fmtM(m.revenue),color:C.amber,bg:C.amberLight},{label:"予想利益",val:fmtM(m.profit),color:m.profit>=0?C.green:C.red,bg:m.profit>=0?C.greenLight:C.redLight}].map(({label,val,color,bg})=>(
                      <div key={label} style={{background:bg,borderRadius:10,padding:"8px 12px"}}><div style={{color:C.textDim,fontSize:10,marginBottom:2}}>{label}</div><div style={{color,fontSize:15,fontWeight:800}}>{val}</div></div>
                    ))}
                  </div>
                  {open&&(
                    <div style={{marginTop:14,borderTop:`1px solid ${C.border}`,paddingTop:12}}>
                      {m.cows.map(c=>{
                        const du=daysUntil(c.shippingPlan);
                        return (
                          <div key={c.id} onClick={e=>{e.stopPropagation();goDetail(c.id);}} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:"10px 14px",marginBottom:6,cursor:"pointer",boxShadow:C.shadow}}>
                            <div>
                              <span style={{color:C.accent,fontWeight:800,marginRight:8,fontFamily:"monospace",fontSize:12}}>{c.tag}</span>
                              <span style={{color:C.text,fontSize:13,fontWeight:600}}>{c.name}</span>
                            </div>
                            <div style={{textAlign:"right"}}>
                              <div style={{color:C.amber,fontSize:13,fontWeight:700}}>{fmtM(c.result?.sellPrice||c.expectedPrice)}</div>
                              {du!==null&&<div style={{color:du<=30?C.red:C.textDim,fontSize:10}}>あと{du}日</div>}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              </div>
            );
          })}
          {!monthlyData.length&&<Card><div style={{color:C.textDim,textAlign:"center",padding:24,fontSize:13}}>出荷予定日が設定されている個体がいません</div></Card>}
        </div>
      </div>
    );
  };

  // ── 繁殖農家分析ページ ────────────────────────────────────────────────────
  const BreederScreen = () => {
    const [selBreeder, setSelBreeder] = useState(null);
    const [sortMode,   setSortMode]   = useState("head"); // head|dg|bms|profit

    const breederData = useMemo(()=>{
      const map = {};
      cattle.forEach(c=>{
        const key = c.farm || "不明";
        if(!map[key]) map[key]={breeder:key, head:0, dgs:[], bmsList:[], loinList:[], profits:[], cows:[]};
        map[key].head++;
        const d = calcDG(c.weights); if(d) map[key].dgs.push(d);
        if(c.result?.bms)      map[key].bmsList.push(c.result.bms);
        if(c.result?.loinArea) map[key].loinList.push(c.result.loinArea);
        const p = calcCosts(c).profit; if(p!=null) map[key].profits.push(p);
        map[key].cows.push(c);
      });
      const list = Object.values(map).map(s=>({
        ...s,
        avgDG:     avg(s.dgs),
        avgBMS:    avg(s.bmsList),
        avgLoin:   avg(s.loinList),
        avgProfit: avg(s.profits),
      }));
      return list.sort((a,b)=>{
        if(sortMode==="head")   return b.head - a.head;
        if(sortMode==="dg")     return (b.avgDG??-Infinity) - (a.avgDG??-Infinity);
        if(sortMode==="bms")    return (b.avgBMS??-Infinity) - (a.avgBMS??-Infinity);
        if(sortMode==="profit") return (b.avgProfit??-Infinity) - (a.avgProfit??-Infinity);
        return 0;
      });
    },[cattle, sortMode]);

    const bc = "#e06040";

    return (
      <div style={{paddingBottom:90}}>
        <AppHeader subtitle="導入元・繁殖農家分析"/>
        <div style={{padding:"16px 16px"}}>

          {/* サマリー */}
          <div style={{background:`linear-gradient(135deg,${bc}18,${C.amberLight})`,borderRadius:14,padding:"12px 16px",marginBottom:12,border:`1px solid ${bc}22`}}>
            <div style={{color:bc,fontWeight:700,fontSize:13,marginBottom:3}}>🏡 導入元（家畜市場）別分析</div>
            <div style={{color:C.textMid,fontSize:12}}>{breederData.length}市場・農家 / {cattle.length}頭</div>
          </div>

          {/* 並び替え */}
          <div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:6,marginBottom:14,scrollbarWidth:"none"}}>
            {[
              {key:"head",   label:"🐂 頭数"},
              {key:"dg",     label:"📈 DG"},
              {key:"bms",    label:"🥩 BMS"},
              {key:"profit", label:"💴 損益"},
            ].map(({key,label})=>(
              <button key={key} onClick={()=>setSortMode(key)} style={{
                background: sortMode===key?`linear-gradient(135deg,${bc},#c04020)`:"#fff",
                color:       sortMode===key?"#fff":C.textMid,
                border:`1.5px solid ${sortMode===key?bc:C.border}`,
                borderRadius:20, padding:"6px 14px",
                fontSize:11, fontWeight:700, cursor:"pointer",
                whiteSpace:"nowrap", flexShrink:0,
              }}>{label}</button>
            ))}
          </div>

          {breederData.map((s,idx)=>{
            const open = selBreeder===s.breeder;
            return (
              <div key={s.breeder} style={{marginBottom:10}}>
                <div onClick={()=>setSelBreeder(open?null:s.breeder)} style={{
                  background: open?`${bc}14`:"#fff",
                  border:`1.5px solid ${open?bc:C.border}`,
                  borderRadius:16, padding:"14px 16px", cursor:"pointer",
                  boxShadow: open?`0 4px 20px ${bc}22`:C.shadow,
                }}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      {/* 順位バッジ */}
                      <div style={{
                        width:28,height:28,borderRadius:"50%",flexShrink:0,
                        background: idx===0?"#f5a623":idx===1?"#aaa":idx===2?"#cd7f32":C.cardSub,
                        display:"flex",alignItems:"center",justifyContent:"center",
                        color: idx<3?"#fff":C.textDim, fontWeight:900, fontSize:12,
                      }}>{idx+1}</div>
                      <div style={{width:32,height:32,borderRadius:10,background:open?bc:C.cardSub,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>🏡</div>
                      <div>
                        <div style={{color:open?bc:C.text,fontWeight:900,fontSize:14}}>{s.breeder}</div>
                        <div style={{color:C.textDim,fontSize:11}}>{s.head}頭導入</div>
                      </div>
                    </div>
                    <span style={{color:C.textDim,fontSize:16}}>{open?"▲":"▼"}</span>
                  </div>

                  {/* 成績グリッド */}
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                    {[
                      {label:"平均DG",   val:s.avgDG   ?`${s.avgDG.toFixed(2)} kg/日`:null, color:C.accent,  bg:C.accentLight},
                      {label:"平均BMS",  val:s.avgBMS  ?`${s.avgBMS.toFixed(1)}`:null,       color:C.red,     bg:C.redLight},
                      {label:"ロース芯", val:s.avgLoin ?`${s.avgLoin.toFixed(1)} cm²`:null,  color:C.accentDark,bg:C.accentLight},
                      {label:"平均損益", val:s.avgProfit!=null?(s.avgProfit>=0?"+":"")+fmtM(s.avgProfit):null, color:s.avgProfit>=0?C.green:C.red, bg:s.avgProfit>=0?C.greenLight:C.redLight},
                    ].map(({label,val,color,bg})=>(
                      <div key={label} style={{background:val?bg:C.cardSub,borderRadius:10,padding:"8px 12px"}}>
                        <div style={{color:C.textDim,fontSize:9,marginBottom:3}}>{label}</div>
                        <div style={{color:val?color:C.textDim,fontSize:13,fontWeight:700}}>{val||"データなし"}</div>
                      </div>
                    ))}
                  </div>

                  {/* 展開：個体一覧 */}
                  {open&&(
                    <div style={{marginTop:14,borderTop:`1px solid ${C.border}`,paddingTop:12}}>
                      {s.cows.map(c=>{
                        const dg_c = calcDG(c.weights);
                        return (
                          <div key={c.id} onClick={e=>{e.stopPropagation();goDetail(c.id);}} style={{
                            background:"#fff",border:`1px solid ${C.border}`,
                            borderRadius:14,padding:"12px 14px",marginBottom:8,
                            cursor:"pointer",boxShadow:C.shadow,
                          }}>
                            <div style={{display:"flex",justifyContent:"space-between",marginBottom:6,alignItems:"center"}}>
                              <TagDisplay tag={c.tag} size={11} highlightSize={15} color={C.accent}/>
                              <div style={{display:"flex",gap:4,alignItems:"center"}}>
                                {c.status==="出荷済"&&<Tag label="出荷済" color={C.textDim} bg="#efefef"/>}
                                <Tag label={c.sex} color={c.sex==="雌"?"#e06090":C.purple}/>
                              </div>
                            </div>
                            {/* 繁殖農家名 */}
                            {c.name&&<div style={{color:C.textDim,fontSize:11,marginBottom:6}}>🏡 繁殖農家: {c.name}</div>}
                            {/* 父血統 */}
                            {c.pedigree?.sire?.name&&(
                              <div style={{fontSize:11,color:C.textDim,marginBottom:8}}>
                                🐂 父: <b style={{color:C.purple}}>{c.pedigree.sire.name}</b>
                                {c.pedigree?.dam?.sire?.name&&<span>　母父: {c.pedigree.dam.sire.name}</span>}
                              </div>
                            )}
                            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
                              {[
                                {label:"DG",     val:dg_c?`${dg_c.toFixed(2)}`:null, color:C.accent},
                                {label:"BMS",    val:c.result?.bms!=null?`${c.result.bms}`:null, color:C.red},
                                {label:"枝肉DG", val:c.result?.dg?`${c.result.dg}`:null, color:C.green},
                              ].map(({label,val,color})=>(
                                <div key={label} style={{background:C.cardSub,borderRadius:8,padding:"5px 8px"}}>
                                  <div style={{color:C.textDim,fontSize:9}}>{label}</div>
                                  <div style={{color:val!=null?color:C.textDim,fontSize:12,fontWeight:700}}>{val??"―"}</div>
                                </div>
                              ))}
                            </div>
                            {c.result&&(
                              <div style={{marginTop:8,background:C.amberLight,borderRadius:8,padding:"6px 10px",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                                <Tag label={c.result.grade} color={C.amber}/>
                                <span style={{color:C.amber,fontSize:12,fontWeight:700}}>{fmtMoney(c.result.sellPrice)}</span>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              </div>
            );
          })}

          {breederData.length===0&&(
            <div style={{textAlign:"center",padding:32,color:C.textDim,fontSize:13}}>
              導入元のデータがありません
            </div>
          )}
        </div>
      </div>
    );
  };

  // ── GENETICS ───────────────────────────────────────────────────────────────
  const GeneticsScreen = () => {
    const [genLevel, setGenLevel] = useState(1);   // 1=一代祖, 2=二代祖, 3=三代祖
    const [selSire,  setSelSire]  = useState(null);

    const data = useMemo(()=>buildLineageData(genLevel),[genLevel]);

    const levelLabel = genLevel===1?"父（一代祖）":genLevel===2?"母の父（二代祖）":"母の母の父（三代祖）";
    const levelColor = genLevel===1?C.purple:genLevel===2?"#7b5ea7":"#5a3e8a";
    const levelDesc  = genLevel===1?"父牛別":genLevel===2?"母の父別":"母の母の父別";

    const AncestorCard = ({s, open, onToggle}) => (
      <div style={{marginBottom:10}}>
        <div onClick={onToggle} style={{
          background:open?`${levelColor}18`:"#fff",
          border:`1.5px solid ${open?levelColor:C.border}`,
          borderRadius:16, padding:"14px 16px", cursor:"pointer",
          boxShadow:open?`0 4px 20px ${levelColor}22`:C.shadow,
        }}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <div style={{
                width:36,height:36,borderRadius:10,
                background:open?levelColor:C.cardSub,
                display:"flex",alignItems:"center",justifyContent:"center",
                fontSize:18,flexShrink:0,
              }}>🐂</div>
              <div>
                <div style={{color:open?levelColor:C.text,fontWeight:900,fontSize:15}}>{s.sire}</div>
                <div style={{color:C.textDim,fontSize:11}}>{levelDesc} / {s.head}頭</div>
              </div>
            </div>
            <span style={{color:C.textDim,fontSize:16}}>{open?"▲":"▼"}</span>
          </div>

          {/* 成績グリッド */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:s.avgProfit!=null?10:0}}>
            {[
              {label:"平均 DG",   val:s.avgDG   ?`${s.avgDG.toFixed(2)} kg/日`:null, color:C.accent,  bg:C.accentLight},
              {label:"平均 BMS",  val:s.avgBMS  ?`${s.avgBMS.toFixed(1)}`:null,       color:C.red,     bg:C.redLight},
              {label:"ロース芯",  val:s.avgLoin ?`${s.avgLoin.toFixed(1)} cm²`:null,  color:C.accentDark,bg:C.accentLight},
              {label:"平均損益",  val:s.avgProfit!=null?(s.avgProfit>=0?"+":"")+fmtM(s.avgProfit):null, color:s.avgProfit>=0?C.green:C.red, bg:s.avgProfit>=0?C.greenLight:C.redLight},
            ].map(({label,val,color,bg})=>(
              <div key={label} style={{background:val?bg:C.cardSub,borderRadius:10,padding:"8px 12px"}}>
                <div style={{color:C.textDim,fontSize:9,marginBottom:3}}>{label}</div>
                <div style={{color:val?color:C.textDim,fontSize:13,fontWeight:700}}>{val||"データなし"}</div>
              </div>
            ))}
          </div>

          {/* 展開：個体一覧 */}
          {open&&(
            <div style={{marginTop:14,borderTop:`1px solid ${C.border}`,paddingTop:12}}>
              {s.cows.map(c=>{
                const dg_c=calcDG(c.weights);
                return (
                  <div key={c.id} onClick={e=>{e.stopPropagation();goDetail(c.id);}} style={{
                    background:"#fff",border:`1px solid ${C.border}`,
                    borderRadius:14,padding:"12px 14px",marginBottom:8,
                    cursor:"pointer",boxShadow:C.shadow,
                  }}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                      <div style={{display:"flex",gap:6,alignItems:"center"}}>
                        <span style={{color:C.accent,fontWeight:800,fontFamily:"monospace",fontSize:12}}>{c.tag}</span>
                        <span style={{color:C.text,fontSize:13,fontWeight:600}}>{c.name}</span>
                        {c.status==="出荷済"&&<Tag label="出荷済" color={C.textDim} bg="#efefef"/>}
                      </div>
                      <Tag label={c.sex} color={c.sex==="雌"?"#e06090":C.purple}/>
                    </div>
                    {/* 三世代血統ミニ表示 */}
                    <div style={{background:C.cardSub,borderRadius:8,padding:"7px 10px",marginBottom:8,fontSize:10,lineHeight:1.9}}>
                      <div><span style={{color:C.textDim,minWidth:70,display:"inline-block"}}>父</span><b style={{color:C.purple}}>{c.pedigree?.sire?.name||"―"}</b></div>
                      <div><span style={{color:C.textDim,minWidth:70,display:"inline-block"}}>母の父</span>{c.pedigree?.dam?.sire?.name||"―"}</div>
                      <div><span style={{color:C.textDim,minWidth:70,display:"inline-block"}}>母の母の父</span>{c.pedigree?.dam?.dam?.sire?.name||"―"}</div>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
                      {[{label:"DG",val:dg_c?`${dg_c.toFixed(2)}`:null,color:C.accent},{label:"BMS",val:c.result?.bms??null,color:C.red},{label:"ロース芯",val:c.result?.loinArea?`${c.result.loinArea}cm²`:null,color:C.accentDark}].map(({label,val,color})=>(
                        <div key={label} style={{background:C.cardSub,borderRadius:8,padding:"5px 8px"}}>
                          <div style={{color:C.textDim,fontSize:9}}>{label}</div>
                          <div style={{color:val!=null?color:C.textDim,fontSize:12,fontWeight:700}}>{val??"―"}</div>
                        </div>
                      ))}
                    </div>
                    {c.result&&(
                      <div style={{marginTop:8,padding:"6px 10px",background:C.amberLight,borderRadius:8,display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                        <Tag label={c.result.grade} color={C.amber}/>
                        <Tag label={`歩留${c.result.yieldGrade}`} color={C.accentDark}/>
                        <span style={{color:C.amber,fontSize:12,fontWeight:700}}>{fmtMoney(c.result.sellPrice)}</span>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    );

    return (
      <div style={{paddingBottom:90}}>
        <AppHeader subtitle="血統別成績分析"/>

        {/* 世代タブ */}
        <div style={{
          background:"#fff", borderBottom:`1px solid ${C.border}`,
          padding:"12px 16px",
          display:"flex", gap:8,
        }}>
          {[
            {level:1, label:"一代祖", sub:"父"},
            {level:2, label:"二代祖", sub:"母の父"},
            {level:3, label:"三代祖", sub:"母の母の父"},
          ].map(({level,label,sub})=>(
            <button key={level} onClick={()=>{setGenLevel(level);setSelSire(null);}} style={{
              flex:1,
              background: genLevel===level
                ? `linear-gradient(135deg,${level===1?C.purple:level===2?"#7b5ea7":"#5a3e8a"},${level===1?"#6b5cd8":level===2?"#5a3e8a":"#3d2a6a"})`
                : "#fff",
              color: genLevel===level?"#fff":C.textMid,
              border:`1.5px solid ${genLevel===level?(level===1?C.purple:level===2?"#7b5ea7":"#5a3e8a"):C.border}`,
              borderRadius:12, padding:"9px 6px",
              cursor:"pointer", transition:"all 0.15s",
              boxShadow:genLevel===level?`0 3px 12px ${C.purple}44`:"none",
            }}>
              <div style={{fontWeight:800,fontSize:13}}>{label}</div>
              <div style={{fontSize:9,opacity:genLevel===level?0.85:0.6,marginTop:2}}>{sub}</div>
            </button>
          ))}
        </div>

        <div style={{padding:"16px 16px"}}>
          {/* サマリーバー */}
          <div style={{
            background:`linear-gradient(135deg,${levelColor}18,${C.accentLight})`,
            borderRadius:14,padding:"12px 16px",marginBottom:16,
            border:`1px solid ${levelColor}22`,
          }}>
            <div style={{color:levelColor,fontWeight:700,fontSize:13,marginBottom:3}}>🧬 {levelLabel}別分析</div>
            <div style={{color:C.textMid,fontSize:12}}>{data.length}血統 / {cattle.length}頭　個体が増えるほど精度が上がります</div>
          </div>

          {data.map(s=>(
            <AncestorCard key={s.sire} s={s} open={selSire===s.sire} onToggle={()=>setSelSire(selSire===s.sire?null:s.sire)}/>
          ))}
          {data.length===0&&(
            <Card><div style={{color:C.textDim,textAlign:"center",padding:24,fontSize:13}}>血統データがありません</div></Card>
          )}
        </div>
      </div>
    );
  };

  // ── ADD SCREEN ─────────────────────────────────────────────────────────────
  const AddScreen = () => {
    // ── フォームstateはここで管理（Appの再描画を防ぐ）──
    const [newForm, setNewForm] = useState(makeEmptyNew());
    const [addTab, setAddTab] = useState("basic");

    // OCRデータが来たらフォームに反映
    useEffect(()=>{
      if(!pendingOcr) return;
      const merge = (base, src) => {
        if(!src) return base;
        return { name:src.name||base?.name||"", sire:merge(base?.sire,src.sire), dam:merge(base?.dam,src.dam) };
      };
      setNewForm(prev=>({
        ...prev,
        tag:      pendingOcr.tag||prev.tag,
        name:     pendingOcr.name||prev.name,
        sex:      pendingOcr.sex||prev.sex,
        breed:    pendingOcr.breed||prev.breed,
        birthDate:pendingOcr.birthDate||prev.birthDate,
        pedigree: pendingOcr.pedigree ? {
          sire: merge(prev.pedigree.sire, pendingOcr.pedigree.sire),
          dam:  merge(prev.pedigree.dam,  pendingOcr.pedigree.dam),
        } : prev.pedigree,
      }));
      setPendingOcr(null);
    },[pendingOcr]);

    const set = (k,v) => setNewForm(p=>({...p,[k]:v}));

    const submit=()=>{
      if(!newForm.tag) return;
      setCattle(p=>[...p,{...newForm,id:Date.now().toString(),expectedPrice:Number(newForm.expectedPrice)||0,weights:[],vaccines:[],treatments:[],result:null}]);
      setPage("home");
    };
    const F=({label,k,type="text",placeholder="",opts})=>(
      <FInput label={label}>
        {opts?<select value={newForm[k]} onChange={e=>set(k,e.target.value)} style={inp}>{opts.map(o=><option key={o}>{o}</option>)}</select>
          :<input type={type} placeholder={placeholder} value={newForm[k]} onChange={e=>set(k,e.target.value)} style={inp}/>}
      </FInput>
    );
    return (
      <div style={{paddingBottom:40}}>
        <div style={{background:"#fff",borderBottom:`1px solid ${C.border}`,padding:"12px 16px",position:"sticky",top:0,zIndex:80,display:"flex",alignItems:"center",gap:10}}>
          <button onClick={()=>setPage("home")} style={{background:C.accentLight,border:"none",color:C.accentDark,borderRadius:10,padding:"7px 14px",fontSize:13,fontWeight:700,cursor:"pointer"}}>← 戻る</button>
          <span style={{color:C.text,fontWeight:800,fontSize:15}}>新規個体登録</span>
        </div>

        {/* OCR banner */}
        <div style={{margin:"14px 16px 0",background:`linear-gradient(135deg,${C.teal}22,${C.accentLight})`,border:`1px solid ${C.teal}44`,borderRadius:14,padding:"12px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:10}}>
          <div>
            <div style={{color:C.teal,fontWeight:800,fontSize:13,marginBottom:3}}>📷 証明書から自動入力</div>
            <div style={{color:C.textMid,fontSize:11}}>子牛登記証明書を撮影するだけでAIが血統情報を読み取ります</div>
          </div>
          <Btn sm variant="teal" onClick={()=>setShowOcr(true)} icon="📷">撮影</Btn>
        </div>

        {/* Sub tabs */}
        <div style={{display:"flex",gap:6,padding:"12px 16px 0",overflowX:"auto",scrollbarWidth:"none"}}>
          {[["basic","📋 基本"],["pedigree","🧬 血統"],["costs","💴 コスト"]].map(([k,v])=>(
            <button key={k} onClick={()=>setAddTab(k)} style={{background:addTab===k?`linear-gradient(135deg,${C.accent},${C.accentDark})`:"transparent",color:addTab===k?"#fff":C.textMid,border:addTab===k?"none":`1px solid ${C.border}`,borderRadius:20,padding:"7px 16px",fontSize:12,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap",flexShrink:0}}>
              {v}
            </button>
          ))}
        </div>

        <div style={{padding:"16px 16px"}}>
          {addTab==="basic"&&(
            <>
              {F({label:"耳標番号 ＊", k:"tag", placeholder:"例: 宮崎-0099"})}
              {F({label:"名前", k:"name", placeholder:"例: 黒王"})}
              {F({label:"性別", k:"sex", opts:["去勢","雌","雄"]})}
              {F({label:"品種", k:"breed", opts:["黒毛和種","褐毛和種","日本短角種","無角和種","交雑種"]})}
              {F({label:"生年月日", k:"birthDate", type:"date"})}
              {F({label:"導入日 ＊", k:"introDate", type:"date"})}
              {F({label:"導入元市場", k:"farm", placeholder:"例: 宮崎中央市場"})}
              {F({label:"牛舎・ペン", k:"pen", placeholder:"例: 2号棟A"})}
              {F({label:"出荷予定日", k:"shippingPlan", type:"date"})}
              {F({label:"予想販売価格（円）", k:"expectedPrice", type:"number", placeholder:"例: 1500000"})}
              <FInput label="メモ"><textarea value={newForm.memo} onChange={e=>set("memo",e.target.value)} style={{...inp,height:70,resize:"vertical"}}/></FInput>
            </>
          )}
          {addTab==="pedigree"&&(
            <PedigreeForm pedigree={newForm.pedigree} onChange={p=>setNewForm(prev=>({...prev,pedigree:p}))}/>
          )}
          {addTab==="costs"&&(
            <CostForm costs={newForm.costs} onChange={c=>setNewForm(prev=>({...prev,costs:c}))}/>
          )}
          <div style={{marginTop:16}}>
            <Btn full onClick={submit}>登録する</Btn>
          </div>
        </div>
      </div>
    );
  };

  // ── DETAIL ─────────────────────────────────────────────────────────────────
  const DetailScreen = () => {
    if(!cow) return null;
    // ── モーダルフォームstateはここで管理（Appの再描画を防ぐ）──
    const [wForm, setWForm] = useState({date:new Date().toISOString().slice(0,10),weight:""});
    const [vForm, setVForm] = useState({date:new Date().toISOString().slice(0,10),name:"",nextDate:""});
    const [tForm, setTForm] = useState({date:new Date().toISOString().slice(0,10),name:"",drug:"",vet:"",cost:""});
    const [rForm, setRForm] = useState({sellPrice:"",bms:"",loinArea:"",ribThickness:"",yieldGrade:"A",grade:"A5",dg:""});
    // ── 編集・削除state ──────────────────────────────────────────────────────
    const [showEdit,      setShowEdit]      = useState(false);
    const [showDelConfirm,setShowDelConfirm] = useState(false);
    const [editForm,      setEditForm]      = useState(null);

    const openEdit = () => {
      setEditForm({
        tag:          cow.tag||"",
        name:         cow.name||"",
        sex:          cow.sex||"去勢",
        breed:        cow.breed||"黒毛和種",
        birthDate:    cow.birthDate||"",
        introDate:    cow.introDate||"",
        farm:         cow.farm||"",
        pen:          cow.pen||"",
        shippingPlan: cow.shippingPlan||"",
        expectedPrice:String(cow.expectedPrice||""),
        memo:         cow.memo||"",
        pedigree:     JSON.parse(JSON.stringify(cow.pedigree||emptyPedigree())),
        costs:        JSON.parse(JSON.stringify(cow.costs||emptyCosts())),
      });
      setShowEdit(true);
    };

    const saveEdit = () => {
      update(cow.id, c=>({
        ...c,
        tag:          editForm.tag,
        name:         editForm.name,
        sex:          editForm.sex,
        breed:        editForm.breed,
        birthDate:    editForm.birthDate,
        introDate:    editForm.introDate,
        farm:         editForm.farm,
        pen:          editForm.pen,
        shippingPlan: editForm.shippingPlan,
        expectedPrice:Number(editForm.expectedPrice)||0,
        memo:         editForm.memo,
        pedigree:     editForm.pedigree,
        costs:        editForm.costs,
      }));
      setShowEdit(false);
    };

    const deleteCow = () => {
      setCattle(p=>p.filter(c=>c.id!==cow.id));
      if(typeof window.deleteCowRemote === "function") window.deleteCowRemote(cow.id);
      setSelectedId(null);
      setPage("home");
    };

    const du=daysUntil(cow.shippingPlan);
    const cv=calcCosts(cow);
    return (
      <div style={{paddingBottom:90}}>
        <div style={{background:"#fff",borderBottom:`1px solid ${C.border}`,padding:"12px 16px",position:"sticky",top:0,zIndex:80,boxShadow:"0 1px 8px rgba(74,184,232,0.08)"}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
            <button onClick={()=>setPage("home")} style={{background:C.accentLight,border:"none",color:C.accentDark,borderRadius:10,padding:"6px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>← 戻る</button>
            <div style={{flex:1}}>
              <TagDisplay tag={cow.tag} size={14} highlightSize={19} color={C.accent}/>
              <span style={{color:C.text,fontWeight:800,fontSize:16,marginLeft:8}}>{cow.name}</span>
            </div>
            {/* 編集・削除ボタン */}
            <button onClick={openEdit} style={{background:C.accentLight,border:`1px solid ${C.border}`,color:C.accentDark,borderRadius:10,padding:"6px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✏️ 編集</button>
            <button onClick={()=>setShowDelConfirm(true)} style={{background:C.redLight,border:`1px solid ${C.red}44`,color:C.red,borderRadius:10,padding:"6px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🗑️</button>
          </div>
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            <Tag label={cow.breed}/>
            <Tag label={cow.sex} color={cow.sex==="雌"?"#e06090":C.purple}/>
            {cow.pedigree?.sire?.name&&<Tag label={`父:${cow.pedigree.sire.name}`} color={C.purple} bg={C.purpleLight}/>}
            {cow.pen&&<Tag label={cow.pen} color={C.textDim} bg={C.cardSub}/>}
            {du!==null&&<Tag label={`出荷${du}日前`} color={du<=60?C.amber:C.textDim} bg={du<=60?C.amberLight:C.cardSub}/>}
            {cow.status==="出荷済"&&<Tag label="出荷済" color={C.textDim} bg="#f0f0f0"/>}
          </div>
        </div>

        {/* Sub tabs */}
        <div style={{background:"#fff",borderBottom:`1px solid ${C.border}`,padding:"10px 14px",display:"flex",gap:6,overflowX:"auto",scrollbarWidth:"none"}}>
          {[["info","📋 基本"],["pedigree","🧬 血統"],["weight","⚖️ 体重"],["health","💉 衛生"],["costs","💴 損益"],["result","🏆 実績"]].map(([k,v])=>(
            <button key={k} onClick={()=>setDetailTab(k)} style={{background:detailTab===k?`linear-gradient(135deg,${C.accent},${C.accentDark})`:"transparent",color:detailTab===k?"#fff":C.textMid,border:detailTab===k?"none":`1px solid ${C.border}`,borderRadius:20,padding:"7px 16px",fontSize:12,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap",flexShrink:0,boxShadow:detailTab===k?C.shadow:"none"}}>
              {v}
            </button>
          ))}
        </div>

        <div style={{padding:"16px 16px"}}>
          {/* INFO */}
          {detailTab==="info"&&(
            <Card>
              {[["耳標",cow.tag],["名前",cow.name],["性別",cow.sex],["品種",cow.breed],
                ["生年月日",fmtDate(cow.birthDate)],["月齢",calcAge(cow.birthDate)],
                ["導入日",fmtDate(cow.introDate)],["飼養日数",`${daysSince(cow.introDate)}日`],
                ["導入元",cow.farm||"―"],["牛舎",cow.pen||"―"],
                ["購入価格",fmtMoney(cow.costs?.purchasePrice||0)],
                ["出荷予定",cow.shippingPlan?`${fmtDate(cow.shippingPlan)}（${du}日後）`:"―"],
              ].map(([k,v],i,arr)=><InfoRow key={k} label={k} value={v} last={i===arr.length-1}/>)}
              {cow.memo&&<div style={{marginTop:14,padding:"12px",background:C.cardSub,borderRadius:12}}><div style={{color:C.textDim,fontSize:11,marginBottom:4}}>📝 メモ</div><div style={{color:C.text,fontSize:13,lineHeight:1.7}}>{cow.memo}</div></div>}
            </Card>
          )}

          {/* PEDIGREE */}
          {detailTab==="pedigree"&&(
            <>
              <Card style={{marginBottom:14}}>
                <SectionLabel>三代祖　血統表</SectionLabel>
                <div style={{display:"flex",gap:6,marginBottom:10}}>
                  <div style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:C.textDim}}><div style={{width:12,height:12,borderRadius:3,background:C.accentDark}}/>父系</div>
                  <div style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:C.textDim}}><div style={{width:12,height:12,borderRadius:3,background:"#e06090cc"}}/>母系</div>
                  <div style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:C.textDim}}><div style={{width:12,height:12,borderRadius:3,background:C.accentLight,border:`1px solid ${C.border}`}}/>祖父母</div>
                  <div style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:C.textDim}}><div style={{width:12,height:12,borderRadius:3,background:C.bg,border:`1px solid ${C.borderLight}`}}/>曾祖父母</div>
                </div>
                <PedigreeTree pedigree={cow.pedigree}/>
              </Card>
              <Card>
                <SectionLabel>詳細血統</SectionLabel>
                {[
                  ["父",cow.pedigree?.sire?.name],
                  ["　父の父（祖父）",cow.pedigree?.sire?.sire?.name],
                  ["　　父の父の父",cow.pedigree?.sire?.sire?.sire?.name],
                  ["　　父の父の母",cow.pedigree?.sire?.sire?.dam?.name],
                  ["　父の母（祖母）",cow.pedigree?.sire?.dam?.name],
                  ["　　父の母の父",cow.pedigree?.sire?.dam?.sire?.name],
                  ["　　父の母の母",cow.pedigree?.sire?.dam?.dam?.name],
                  ["母",cow.pedigree?.dam?.name],
                  ["　母の父（祖父）",cow.pedigree?.dam?.sire?.name],
                  ["　　母の父の父",cow.pedigree?.dam?.sire?.sire?.name],
                  ["　　母の父の母",cow.pedigree?.dam?.sire?.dam?.name],
                  ["　母の母（祖母）",cow.pedigree?.dam?.dam?.name],
                  ["　　母の母の父",cow.pedigree?.dam?.dam?.sire?.name],
                  ["　　母の母の母",cow.pedigree?.dam?.dam?.dam?.name],
                ].filter(([,v])=>v).map(([k,v],i,arr)=>(
                  <InfoRow key={k} label={k} value={v} last={i===arr.length-1} accent={k==="父"||k==="母"}/>
                ))}
              </Card>
            </>
          )}

          {/* WEIGHT */}
          {detailTab==="weight"&&(
            <>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:14}}>
                <KpiCard label="最新体重" value={latestWeight(cow.weights)?`${latestWeight(cow.weights)}`:"―"} unit="kg" icon="⚖️" color={C.accent} bg={C.accentLight}/>
                <KpiCard label="DG実績" value={dg?`+${dg.toFixed(2)}`:"―"} unit="kg/日" icon="📈" color={C.amber} bg={C.amberLight}/>
                <KpiCard label="出荷予測" value={pred?`${pred}`:"―"} unit="kg" icon="🎯" color={C.amber} bg={C.amberLight}/>
              </div>
              <Card style={{marginBottom:14}}>
                <div style={{color:C.textDim,fontSize:11,marginBottom:10}}>発育グラフ{pred&&<span style={{color:C.amber}}>　★ = DG予測体重</span>}</div>
                <WeightChart weights={cow.weights} shippingPlan={cow.shippingPlan} dg={dg}/>
              </Card>
              <div style={{display:"flex",justifyContent:"flex-end",marginBottom:10}}>
                <Btn sm variant="soft" onClick={()=>{setWForm({date:new Date().toISOString().slice(0,10),weight:""});setModal("weight");}}>＋ 体重記録</Btn>
              </div>
              {cow.weights.slice().reverse().map((w,i)=>(
                <Card key={i} style={{marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 16px"}}>
                  <span style={{color:C.textMid,fontSize:13}}>{fmtDate(w.date)}</span>
                  <span style={{color:C.accent,fontWeight:800,fontSize:17}}>{w.weight} kg</span>
                </Card>
              ))}
              {!cow.weights.length&&<div style={{color:C.textDim,textAlign:"center",padding:"24px 0",fontSize:13}}>体重記録がありません</div>}
            </>
          )}

          {/* HEALTH */}
          {detailTab==="health"&&(
            <>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <span style={{color:C.text,fontWeight:700,fontSize:14}}>💉 ワクチン記録</span>
                <Btn sm variant="soft" onClick={()=>{setVForm({date:new Date().toISOString().slice(0,10),name:"",nextDate:""});setModal("vaccine");}}>＋ 追加</Btn>
              </div>
              {!cow.vaccines.length&&<div style={{color:C.textDim,fontSize:13,marginBottom:16,padding:"12px 0"}}>記録なし</div>}
              {cow.vaccines.map((v,i)=>(
                <Card key={i} style={{marginBottom:8}}>
                  <div style={{color:C.accentDark,fontWeight:700,marginBottom:5}}>{v.name}</div>
                  <div style={{color:C.textMid,fontSize:12}}>接種日: {fmtDate(v.date)}</div>
                  {v.nextDate&&<div style={{color:daysUntil(v.nextDate)<=30?C.red:C.textDim,fontSize:12,marginTop:2}}>次回: {fmtDate(v.nextDate)}（あと{daysUntil(v.nextDate)}日）</div>}
                </Card>
              ))}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"20px 0 10px"}}>
                <span style={{color:C.text,fontWeight:700,fontSize:14}}>🏥 治療記録</span>
                <Btn sm variant="soft" onClick={()=>{setTForm({date:new Date().toISOString().slice(0,10),name:"",drug:"",vet:"",cost:""});setModal("treatment");}}>＋ 追加</Btn>
              </div>
              {!cow.treatments.length&&<div style={{color:C.textDim,fontSize:13,padding:"12px 0"}}>記録なし</div>}
              {cow.treatments.map((t,i)=>(
                <Card key={i} style={{marginBottom:8}}>
                  <div style={{color:C.red,fontWeight:700,marginBottom:5}}>{t.name}</div>
                  <div style={{color:C.textMid,fontSize:12}}>{fmtDate(t.date)}　{t.drug||"薬剤未記録"}　{t.vet||"獣医未記録"}{t.cost?`　${fmtMoney(t.cost)}`:""}</div>
                </Card>
              ))}
            </>
          )}

          {/* COSTS */}
          {detailTab==="costs"&&(
            <>
              {!editCosts ? (
                <>
                  <div style={{display:"flex",justifyContent:"flex-end",marginBottom:10}}>
                    <Btn sm variant="soft" onClick={()=>{setTmpCosts({...cow.costs});setEditCosts(true);}}>✏️ コスト編集</Btn>
                  </div>
                  <CostBreakdown cow={cow}/>
                </>
              ):(
                <>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                    <span style={{color:C.text,fontWeight:700}}>コスト編集</span>
                    <div style={{display:"flex",gap:8}}>
                      <Btn sm variant="outline" onClick={()=>setEditCosts(false)}>キャンセル</Btn>
                      <Btn sm onClick={()=>{update(cow.id,c=>({...c,costs:tmpCosts}));setEditCosts(false);}}>保存</Btn>
                    </div>
                  </div>
                  <CostForm costs={tmpCosts} onChange={setTmpCosts}/>
                </>
              )}
            </>
          )}

          {/* RESULT */}
          {detailTab==="result"&&(
            !cow.result?(
              <Card style={{textAlign:"center",padding:"32px 16px"}}>
                <div style={{fontSize:40,marginBottom:12}}>🏆</div>
                <div style={{color:C.textMid,fontSize:14,marginBottom:20}}>出荷実績がまだ登録されていません</div>
                <Btn onClick={()=>{setRForm({sellPrice:"",bms:"",loinArea:"",ribThickness:"",yieldGrade:"A",grade:"A5",dg:""});setModal("result");}}>＋ 出荷実績を入力</Btn>
              </Card>
            ):(
              <>
                <Card style={{marginBottom:14}}>
                  <SectionLabel>枝肉成績</SectionLabel>
                  <div style={{marginBottom:14}}><div style={{color:C.textDim,fontSize:11,marginBottom:8}}>BMS（脂肪交雑）</div><BmsMeter value={cow.result.bms}/></div>
                  <InfoRow label="格付（等級）" value={cow.result.grade} accent/>
                  <InfoRow label="歩留等級" value={cow.result.yieldGrade}/>
                  <InfoRow label="ロース芯面積" value={cow.result.loinArea?`${cow.result.loinArea} cm²`:"―"}/>
                  <InfoRow label="バラ厚" value={cow.result.ribThickness?`${cow.result.ribThickness} cm`:"―"}/>
                  <InfoRow label="実績DG" value={cow.result.dg?`${cow.result.dg} kg/日`:"―"} accent/>
                  <InfoRow label="販売価格" value={fmtMoney(cow.result.sellPrice)} accent big last/>
                </Card>
                <Card>
                  <SectionLabel>確定損益</SectionLabel>
                  <InfoRow label="販売価格" value={fmtMoney(cow.result.sellPrice)}/>
                  <InfoRow label="総コスト" value={fmtMoney(calcCosts(cow).totalCost)}/>
                  <div style={{background:calcCosts(cow).profit>=0?C.greenLight:C.redLight,borderRadius:14,padding:"16px",marginTop:12,textAlign:"center"}}>
                    <div style={{color:C.textMid,fontSize:12,marginBottom:4}}>確定損益</div>
                    <div style={{color:calcCosts(cow).profit>=0?C.green:C.red,fontWeight:900,fontSize:26}}>
                      {calcCosts(cow).profit!=null?(calcCosts(cow).profit>=0?"+":"")+fmtMoney(calcCosts(cow).profit):"―"}
                    </div>
                  </div>
                </Card>
              </>
            )
          )}
        </div>

        {/* Modals */}
        {modal==="weight"&&(
          <Modal title="体重記録を追加" onClose={()=>setModal(null)}>
            <FInput label="計測日"><input type="date" value={wForm.date} onChange={e=>setWForm(p=>({...p,date:e.target.value}))} style={inp}/></FInput>
            <FInput label="体重（kg）"><input type="number" placeholder="例: 450" value={wForm.weight} onChange={e=>setWForm(p=>({...p,weight:e.target.value}))} style={inp}/></FInput>
            <Btn full onClick={()=>{if(!wForm.weight)return;update(cow.id,c=>({...c,weights:[...c.weights,{date:wForm.date,weight:Number(wForm.weight)}].sort((a,b)=>new Date(a.date)-new Date(b.date))}));setModal(null);}}>追加する</Btn>
          </Modal>
        )}
        {modal==="vaccine"&&(
          <Modal title="ワクチン記録を追加" onClose={()=>setModal(null)}>
            <FInput label="接種日"><input type="date" value={vForm.date} onChange={e=>setVForm(p=>({...p,date:e.target.value}))} style={inp}/></FInput>
            <FInput label="ワクチン名"><input type="text" placeholder="例: 口蹄疫ワクチン" value={vForm.name} onChange={e=>setVForm(p=>({...p,name:e.target.value}))} style={inp}/></FInput>
            <FInput label="次回接種予定日"><input type="date" value={vForm.nextDate} onChange={e=>setVForm(p=>({...p,nextDate:e.target.value}))} style={inp}/></FInput>
            <Btn full onClick={()=>{if(!vForm.name)return;update(cow.id,c=>({...c,vaccines:[...c.vaccines,{...vForm}]}));setModal(null);}}>追加する</Btn>
          </Modal>
        )}
        {modal==="treatment"&&(
          <Modal title="治療記録を追加" onClose={()=>setModal(null)}>
            <FInput label="治療日"><input type="date" value={tForm.date} onChange={e=>setTForm(p=>({...p,date:e.target.value}))} style={inp}/></FInput>
            <FInput label="病名・症状"><input type="text" placeholder="例: 下痢" value={tForm.name} onChange={e=>setTForm(p=>({...p,name:e.target.value}))} style={inp}/></FInput>
            <FInput label="使用薬剤"><input type="text" placeholder="例: 抗生剤" value={tForm.drug} onChange={e=>setTForm(p=>({...p,drug:e.target.value}))} style={inp}/></FInput>
            <FInput label="担当獣医"><input type="text" placeholder="例: 田中獣医師" value={tForm.vet} onChange={e=>setTForm(p=>({...p,vet:e.target.value}))} style={inp}/></FInput>
            <FInput label="治療費（円）"><input type="number" placeholder="例: 15000" value={tForm.cost} onChange={e=>setTForm(p=>({...p,cost:e.target.value}))} style={inp}/></FInput>
            <Btn full onClick={()=>{if(!tForm.name)return;update(cow.id,c=>({...c,treatments:[...c.treatments,{...tForm,cost:Number(tForm.cost)||0}]}));setModal(null);}}>追加する</Btn>
          </Modal>
        )}
        {modal==="result"&&(
          <Modal title="出荷実績を入力" onClose={()=>setModal(null)}>
            <FInput label="販売価格（円）"><input type="number" placeholder="例: 1620000" value={rForm.sellPrice} onChange={e=>setRForm(p=>({...p,sellPrice:e.target.value}))} style={inp}/></FInput>
            <FInput label="BMS（1〜12）"><input type="number" min="1" max="12" placeholder="例: 9" value={rForm.bms} onChange={e=>setRForm(p=>({...p,bms:e.target.value}))} style={inp}/></FInput>
            <FInput label="等級"><select value={rForm.grade} onChange={e=>setRForm(p=>({...p,grade:e.target.value}))} style={inp}>{["A5","A4","A3","A2","A1","B5","B4","B3","B2","B1"].map(g=><option key={g}>{g}</option>)}</select></FInput>
            <FInput label="歩留等級"><select value={rForm.yieldGrade} onChange={e=>setRForm(p=>({...p,yieldGrade:e.target.value}))} style={inp}>{["A","B","C"].map(g=><option key={g}>{g}</option>)}</select></FInput>
            <FInput label="ロース芯面積（cm²）"><input type="number" placeholder="例: 62" value={rForm.loinArea} onChange={e=>setRForm(p=>({...p,loinArea:e.target.value}))} style={inp}/></FInput>
            <FInput label="バラ厚（cm）"><input type="number" step="0.1" placeholder="例: 8.2" value={rForm.ribThickness} onChange={e=>setRForm(p=>({...p,ribThickness:e.target.value}))} style={inp}/></FInput>
            <FInput label="実績DG（kg/日）"><input type="number" step="0.01" placeholder="例: 0.97" value={rForm.dg} onChange={e=>setRForm(p=>({...p,dg:e.target.value}))} style={inp}/></FInput>
            <Btn full onClick={()=>{update(cow.id,c=>({...c,status:"出荷済",result:{sellPrice:Number(rForm.sellPrice)||0,bms:Number(rForm.bms)||null,loinArea:Number(rForm.loinArea)||null,ribThickness:Number(rForm.ribThickness)||null,yieldGrade:rForm.yieldGrade,grade:rForm.grade,dg:Number(rForm.dg)||null}}));setModal(null);}}>登録する</Btn>
          </Modal>
        )}

      {/* ── 編集モーダル ── */}
      {showEdit&&editForm&&(
        <Modal title="✏️ 個体情報を編集" onClose={()=>setShowEdit(false)}>
          {/* タブ */}
          <div style={{display:"flex",gap:6,marginBottom:16,borderBottom:`1px solid ${C.border}`,paddingBottom:10}}>
            {[["basic","📋 基本"],["pedigree","🧬 血統"]].map(([k,l])=>(
              <button key={k} onClick={()=>setDetailTab(k)} style={{
                background:detailTab===k?`linear-gradient(135deg,${C.accent},${C.accentDark})`:"transparent",
                color:detailTab===k?"#fff":C.textMid,
                border:`1px solid ${detailTab===k?C.accent:C.border}`,
                borderRadius:20,padding:"6px 16px",fontSize:12,fontWeight:700,cursor:"pointer",
              }}>{l}</button>
            ))}
          </div>

          {/* 基本情報タブ */}
          {detailTab!=="pedigree"&&(
            <>
              {[
                {label:"耳標番号",k:"tag",type:"text",ph:"例: 宮崎-0099"},
                {label:"繁殖農家",k:"name",type:"text",ph:"例: 黒王"},
                {label:"生年月日",k:"birthDate",type:"date"},
                {label:"導入日",k:"introDate",type:"date"},
                {label:"導入元市場",k:"farm",type:"text",ph:"例: 宮崎中央市場"},
                {label:"牛舎・ペン",k:"pen",type:"text",ph:"例: 2号棟A"},
                {label:"出荷予定日",k:"shippingPlan",type:"date"},
                {label:"予想販売価格（円）",k:"expectedPrice",type:"number"},
              ].map(({label,k,type,ph})=>(
                <FInput key={k} label={label}>
                  <input type={type} placeholder={ph||""} value={editForm[k]||""}
                    onChange={e=>setEditForm(p=>({...p,[k]:e.target.value}))} style={inp}/>
                </FInput>
              ))}
              <FInput label="性別">
                <select value={editForm.sex} onChange={e=>setEditForm(p=>({...p,sex:e.target.value}))} style={inp}>
                  {["去勢","雌","雄"].map(o=><option key={o}>{o}</option>)}
                </select>
              </FInput>
              <FInput label="品種">
                <select value={editForm.breed} onChange={e=>setEditForm(p=>({...p,breed:e.target.value}))} style={inp}>
                  {["黒毛和種","褐毛和種","日本短角種","無角和種","交雑種"].map(o=><option key={o}>{o}</option>)}
                </select>
              </FInput>
              <FInput label="メモ">
                <textarea value={editForm.memo} onChange={e=>setEditForm(p=>({...p,memo:e.target.value}))} style={{...inp,height:60,resize:"vertical"}}/>
              </FInput>
            </>
          )}

          {/* 血統タブ */}
          {detailTab==="pedigree"&&(
            <>
              {/* 父系 */}
              <div style={{background:`${C.purple}11`,borderRadius:12,padding:"12px 14px",marginBottom:12}}>
                <div style={{color:C.purple,fontWeight:700,fontSize:12,marginBottom:10}}>🐂 父系</div>
                {[
                  {label:"父",     path:"sire.name"},
                  {label:"父の父", path:"sire.sire.name"},
                  {label:"父の母", path:"sire.dam.name"},
                ].map(({label,path})=>{
                  const keys=path.split(".");
                  const getV=(obj,ks)=>ks.reduce((o,k)=>o?.[k],obj)||"";
                  const setV=(obj,ks,v)=>{
                    const n=JSON.parse(JSON.stringify(obj));
                    let o=n;
                    for(let i=0;i<ks.length-1;i++){
                      if(!o[ks[i]])o[ks[i]]={name:"",sire:{name:""},dam:{name:""}};
                      o=o[ks[i]];
                    }
                    o[ks[ks.length-1]]=v;
                    return n;
                  };
                  return (
                    <FInput key={label} label={label}>
                      <input value={getV(editForm.pedigree,keys)} onChange={e=>setEditForm(p=>({...p,pedigree:setV(p.pedigree,keys,e.target.value)}))} placeholder={label} style={inp}/>
                    </FInput>
                  );
                })}
              </div>
              {/* 母系 */}
              <div style={{background:"#ffebf488",borderRadius:12,padding:"12px 14px",marginBottom:12}}>
                <div style={{color:"#c0407a",fontWeight:700,fontSize:12,marginBottom:10}}>🐄 母系</div>
                {[
                  {label:"母",       path:"dam.name"},
                  {label:"母の父",   path:"dam.sire.name"},
                  {label:"母の母",   path:"dam.dam.name"},
                  {label:"母の母の父",path:"dam.dam.sire.name"},
                ].map(({label,path})=>{
                  const keys=path.split(".");
                  const getV=(obj,ks)=>ks.reduce((o,k)=>o?.[k],obj)||"";
                  const setV=(obj,ks,v)=>{
                    const n=JSON.parse(JSON.stringify(obj));
                    let o=n;
                    for(let i=0;i<ks.length-1;i++){
                      if(!o[ks[i]])o[ks[i]]={name:"",sire:{name:"",sire:{name:""},dam:{name:""}},dam:{name:"",sire:{name:""},dam:{name:""}}};
                      o=o[ks[i]];
                    }
                    o[ks[ks.length-1]]=v;
                    return n;
                  };
                  return (
                    <FInput key={label} label={label}>
                      <input value={getV(editForm.pedigree,keys)} onChange={e=>setEditForm(p=>({...p,pedigree:setV(p.pedigree,keys,e.target.value)}))} placeholder={label} style={inp}/>
                    </FInput>
                  );
                })}
              </div>
            </>
          )}

          <Btn full onClick={saveEdit}>保存する</Btn>
        </Modal>
      )}

      {/* ── 削除確認モーダル ── */}
      {showDelConfirm&&(
        <Modal title="削除の確認" onClose={()=>setShowDelConfirm(false)}>
          <div style={{textAlign:"center",padding:"8px 0"}}>
            <div style={{fontSize:40,marginBottom:12}}>⚠️</div>
            <div style={{color:C.text,fontWeight:700,fontSize:16,marginBottom:8}}>{cow.tag}　{cow.name}</div>
            <div style={{color:C.textMid,fontSize:13,marginBottom:24,lineHeight:1.7}}>
              この個体を削除します。<br/>
              <b style={{color:C.red}}>削除したデータは元に戻せません。</b>
            </div>
            <div style={{display:"flex",gap:10}}>
              <Btn full variant="outline" onClick={()=>setShowDelConfirm(false)}>キャンセル</Btn>
              <Btn full variant="danger" onClick={deleteCow}>削除する</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
    );
  };
  const ShipResultScreen = () => {
    const fileRef   = useRef(null);
    const [status,  setStatus]   = useState("idle"); // idle|loading|done|error
    const [preview, setPreview]  = useState(null);   // 画像プレビュー（PDF時はnull）
    const [isPdf,   setIsPdf]    = useState(false);
    const [errMsg,  setErrMsg]   = useState("");
    const [results, setResults]  = useState([]);     // AIが読み取った行
    const [matched, setMatched]  = useState([]);     // {cow, result, selected}

    /* ── ファイル受け取り ── */
    const handleFile = async (file) => {
      if(!file) return;
      const pdf = file.type === "application/pdf";
      setIsPdf(pdf);
      setStatus("loading");
      setResults([]); setMatched([]); setErrMsg("");

      // 画像プレビュー
      if(!pdf){
        const r = new FileReader();
        r.onload = e => setPreview(e.target.result);
        r.readAsDataURL(file);
      } else {
        setPreview(null);
      }

      // base64
      const b64 = await new Promise((res,rej)=>{
        const r2 = new FileReader();
        r2.onload = () => res(r2.result.split(",")[1]);
        r2.onerror = rej;
        r2.readAsDataURL(file);
      });

      const prompt = `この画像は和牛の出荷伝票・枝肉成績書・精算書です。
複数頭分の出荷成績をすべて読み取り、以下のJSON配列で返してください。
読み取れない項目は null にしてください。

[
  {
    "tag": "耳標番号（10桁数字またはハイフン付き）",
    "name": "牛名（あれば）",
    "sellPrice": 販売金額の数値,
    "grade": "格付等級（例: A5）",
    "yieldGrade": "歩留等級（A/B/C）",
    "bms": BMS数値（1〜12の整数）,
    "loinArea": ロース芯面積の数値（cm²）,
    "ribThickness": バラ厚の数値（cm）,
    "bft": 皮下脂肪厚の数値（cm）,
    "coldWeight": 枝肉重量の数値（kg）,
    "dg": 実績DGの数値（kg/日）,
    "shippingDate": "出荷日（YYYY-MM-DD）"
  }
]

全頭分を配列に含めてください。JSONのみ返してください。`;

      try {
        const contentItem = pdf
          ? { type:"document", source:{ type:"base64", media_type:"application/pdf", data:b64 } }
          : { type:"image",    source:{ type:"base64", media_type:file.type||"image/jpeg", data:b64 } };

        const resp = await fetch("https://api.anthropic.com/v1/messages",{
          method:"POST",
          headers:{"Content-Type":"application/json","x-api-key":window.ANTHROPIC_KEY||"","anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
          body:JSON.stringify({
            model:"claude-sonnet-4-20250514",
            max_tokens:1500,
            messages:[{ role:"user", content:[
              contentItem,
              { type:"text", text:prompt }
            ]}]
          })
        });
        const data = await resp.json();
        const text = data.content?.map(c=>c.text||"").join("")||"";
        const parsed = JSON.parse(text.replace(/```json|```/g,"").trim());
        const arr = Array.isArray(parsed) ? parsed : [parsed];
        setResults(arr);

        // 既存個体とマッチング（耳標で照合）
        const m = arr.map(r => {
          const cow = cattle.find(c =>
            c.tag && r.tag &&
            (c.tag.replace(/[^0-9]/g,"") === r.tag.replace(/[^0-9]/g,"") ||
             c.tag === r.tag ||
             (r.name && c.name === r.name))
          );
          return { ocr:r, cow: cow||null, selected: !!cow, overwrite: false };
        });
        setMatched(m);
        setStatus("done");
      } catch(e) {
        console.error(e);
        setErrMsg("読み取りに失敗しました。鮮明な写真・PDFで再試行してください。");
        setStatus("error");
      }
    };

    /* ── 一括保存 ── */
    const saveAll = () => {
      matched.filter(m=>m.selected&&m.cow).forEach(m=>{
        const r = m.ocr;
        update(m.cow.id, c=>({
          ...c,
          status:"出荷済",
          shippingPlan: r.shippingDate||c.shippingPlan,
          result:{
            sellPrice:    r.sellPrice    ?? c.result?.sellPrice    ?? 0,
            grade:        r.grade        ?? c.result?.grade        ?? "―",
            yieldGrade:   r.yieldGrade   ?? c.result?.yieldGrade   ?? "A",
            bms:          r.bms          ?? c.result?.bms          ?? null,
            loinArea:     r.loinArea     ?? c.result?.loinArea     ?? null,
            ribThickness: r.ribThickness ?? c.result?.ribThickness ?? null,
            bft:          r.bft          ?? c.result?.bft          ?? null,
            coldWeight:   r.coldWeight   ?? c.result?.coldWeight   ?? null,
            dg:           r.dg           ?? c.result?.dg           ?? null,
          }
        }));
      });
      setPage("home");
    };

    const toggleSelect = (i) => setMatched(p=>p.map((m,idx)=>idx===i?{...m,selected:!m.selected}:m));

    /* ── グレードカラー ── */
    const gradeColor = (g) =>
      g==="A5"||g==="B5" ? C.red :
      g==="A4"||g==="B4" ? C.amber :
      g==="A3"||g==="B3" ? C.green : C.textMid;

    return (
      <div style={{background:C.bg,minHeight:"100vh",fontFamily:"'Hiragino Kaku Gothic Pro','Noto Sans JP',sans-serif",maxWidth:520,margin:"0 auto",paddingBottom:40}}>
        {/* ヘッダー */}
        <div style={{
          background:`linear-gradient(135deg,${C.amber},#c87010)`,
          padding:"14px 20px 14px",
          position:"sticky", top:0, zIndex:80,
          display:"flex", alignItems:"center", justifyContent:"space-between",
        }}>
          <button onClick={()=>setPage("home")} style={{background:"rgba(255,255,255,0.22)",border:"none",color:"#fff",borderRadius:10,padding:"7px 14px",fontSize:13,fontWeight:700,cursor:"pointer"}}>← 戻る</button>
          <span style={{color:"#fff",fontWeight:900,fontSize:16}}>🏆 出荷成績入力</span>
          <div style={{width:60}}/>
        </div>

        <div style={{padding:"20px 18px"}}>

          {/* idle / error → アップロードUI */}
          {(status==="idle"||status==="error")&&(
            <div>
              <div style={{color:C.text,fontWeight:800,fontSize:17,marginBottom:6}}>出荷伝票を読み込む</div>
              <div style={{color:C.textMid,fontSize:13,marginBottom:20,lineHeight:1.7}}>
                出荷伝票・枝肉成績書・精算書の<br/>
                <b>PDF または 写真</b>を選択してください。<br/>
                複数頭分がまとまった伝票でも一括で読み取ります。
              </div>

              {status==="error"&&(
                <div style={{background:C.redLight,border:`1px solid ${C.red}33`,borderRadius:12,padding:"10px 14px",marginBottom:16}}>
                  <div style={{color:C.red,fontWeight:700,fontSize:13}}>{errMsg}</div>
                </div>
              )}

              <input ref={fileRef} type="file" accept="image/*,application/pdf" style={{display:"none"}}
                onChange={e=>handleFile(e.target.files?.[0])}/>
              <input ref={fileRef} type="file" accept="image/*,application/pdf" style={{display:"none"}}
                onChange={e=>handleFile(e.target.files?.[0])}/>

              {/* カメラ撮影 */}
              <div onClick={()=>{
                const inp2=document.createElement("input");
                inp2.type="file"; inp2.accept="image/*"; inp2.capture="environment";
                inp2.onchange=e=>handleFile(e.target.files?.[0]);
                inp2.click();
              }} style={{
                background:`linear-gradient(135deg,${C.amber}18,${C.amberLight})`,
                border:`2px dashed ${C.amber}`,
                borderRadius:18, padding:"28px 20px", textAlign:"center",
                cursor:"pointer", marginBottom:10,
              }}>
                <div style={{fontSize:44,marginBottom:8}}>📷</div>
                <div style={{color:C.amber,fontWeight:800,fontSize:16,marginBottom:4}}>カメラで撮影</div>
                <div style={{color:C.textMid,fontSize:12}}>出荷伝票をカメラで撮影</div>
              </div>

              {/* ファイル選択（画像・PDF両対応） */}
              <div onClick={()=>{
                const inp2=document.createElement("input");
                inp2.type="file"; inp2.accept="image/*,application/pdf";
                inp2.onchange=e=>handleFile(e.target.files?.[0]);
                inp2.click();
              }} style={{
                background:"#fff", border:`1.5px solid ${C.border}`,
                borderRadius:14, padding:"16px 20px", textAlign:"center",
                cursor:"pointer", display:"flex", alignItems:"center",
                justifyContent:"center", gap:10, boxShadow:C.shadow,
              }}>
                <span style={{fontSize:28}}>📄</span>
                <div style={{textAlign:"left"}}>
                  <div style={{color:C.text,fontWeight:700,fontSize:14}}>ファイルを選択</div>
                  <div style={{color:C.textDim,fontSize:11,marginTop:2}}>PDF・JPG・PNG 対応</div>
                </div>
              </div>
            </div>
          )}

          {/* loading */}
          {status==="loading"&&(
            <div style={{textAlign:"center",padding:"32px 0"}}>
              {preview&&<img src={preview} style={{width:"100%",maxHeight:220,objectFit:"contain",borderRadius:14,marginBottom:18}} alt="伝票"/>}
              {isPdf&&<div style={{fontSize:60,marginBottom:14}}>📄</div>}
              <div style={{
                width:60,height:60,borderRadius:"50%",
                background:`linear-gradient(135deg,${C.amber},#c87010)`,
                display:"flex",alignItems:"center",justifyContent:"center",
                fontSize:28,margin:"0 auto 14px",
                animation:"spin 1.2s linear infinite",
              }}>🔍</div>
              <div style={{color:"#c87010",fontWeight:800,fontSize:16,marginBottom:4}}>読み取り中...</div>
              <div style={{color:C.textMid,fontSize:12}}>AIが出荷成績を解析しています</div>
              <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
            </div>
          )}

          {/* done → 確認画面 */}
          {status==="done"&&(
            <div>
              {preview&&<img src={preview} style={{width:"100%",maxHeight:160,objectFit:"contain",borderRadius:12,marginBottom:14}} alt="伝票"/>}
              {isPdf&&<div style={{background:C.amberLight,borderRadius:12,padding:"10px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:22}}>📄</span><span style={{color:C.amber,fontWeight:700,fontSize:13}}>PDFを読み取りました</span></div>}

              {/* 読み取り結果サマリー */}
              <div style={{background:C.greenLight,border:`1px solid ${C.green}44`,borderRadius:12,padding:"12px 14px",marginBottom:16}}>
                <div style={{color:C.green,fontWeight:800,fontSize:13,marginBottom:4}}>✅ 読み取り完了</div>
                <div style={{color:C.text,fontSize:12}}>
                  {matched.length}頭分のデータを読み取りました。
                  うち <b style={{color:C.green}}>{matched.filter(m=>m.cow).length}頭</b> が既存個体と一致しました。
                </div>
              </div>

              <div style={{color:C.textDim,fontSize:11,fontWeight:700,letterSpacing:1,marginBottom:10}}>
                保存する個体を選択（タップで切替）
              </div>

              {matched.map((m,i)=>{
                const r = m.ocr;
                const gc = gradeColor(r.grade);
                return (
                  <div key={i} onClick={()=>m.cow&&toggleSelect(i)} style={{
                    background: m.selected&&m.cow ? "#fff" : C.cardSub,
                    border:`1.5px solid ${m.selected&&m.cow?C.green:m.cow?C.border:"#e0b070"}`,
                    borderRadius:16, padding:"14px 16px", marginBottom:10,
                    cursor:m.cow?"pointer":"default",
                    opacity:m.cow?1:0.7,
                    boxShadow: m.selected&&m.cow ? `0 3px 14px ${C.green}22` : C.shadow,
                  }}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        {m.cow&&(
                          <div style={{
                            width:22,height:22,borderRadius:"50%",
                            background:m.selected?C.green:C.cardSub,
                            border:`2px solid ${m.selected?C.green:C.border}`,
                            display:"flex",alignItems:"center",justifyContent:"center",
                            fontSize:12,color:"#fff",fontWeight:900,flexShrink:0,
                          }}>{m.selected?"✓":""}</div>
                        )}
                        <div>
                          <span style={{color:C.accent,fontWeight:800,fontFamily:"monospace",fontSize:13}}>{r.tag||"耳標不明"}</span>
                          {r.name&&<span style={{color:C.text,fontSize:12,marginLeft:6}}>{r.name}</span>}
                        </div>
                      </div>
                      {m.cow
                        ? <span style={{background:C.greenLight,color:C.green,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>✓ 照合済</span>
                        : <span style={{background:"#fff4e0",color:"#c87010",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>⚠ 未照合</span>
                      }
                    </div>

                    {m.cow&&(
                      <div style={{color:C.textDim,fontSize:11,marginBottom:8}}>
                        → {m.cow.name||m.cow.tag}（{m.cow.pen||"―"}）に保存
                      </div>
                    )}

                    {/* 成績グリッド */}
                    <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6}}>
                      {[
                        {label:"等級",     val:r.grade,              color:gc},
                        {label:"BMS",      val:r.bms,                color:C.red},
                        {label:"販売価格", val:r.sellPrice?fmtM(r.sellPrice):null, color:C.amber},
                        {label:"ロース芯", val:r.loinArea?`${r.loinArea}cm²`:null, color:C.accentDark},
                        {label:"枝肉重量", val:r.coldWeight?`${r.coldWeight}kg`:null, color:C.textMid},
                        {label:"歩留",     val:r.yieldGrade,         color:C.purple},
                      ].map(({label,val,color})=>(
                        <div key={label} style={{background:C.cardSub,borderRadius:8,padding:"6px 8px"}}>
                          <div style={{color:C.textDim,fontSize:9,marginBottom:2}}>{label}</div>
                          <div style={{color:val!=null?color:C.textDim,fontSize:12,fontWeight:700}}>{val??  "―"}</div>
                        </div>
                      ))}
                    </div>
                    {r.shippingDate&&<div style={{color:C.textDim,fontSize:11,marginTop:8}}>出荷日: {r.shippingDate}</div>}
                  </div>
                );
              })}

              {/* 未照合の説明 */}
              {matched.some(m=>!m.cow)&&(
                <div style={{background:"#fff4e0",border:"1px solid #e0b070",borderRadius:12,padding:"10px 14px",marginBottom:14}}>
                  <div style={{color:"#c87010",fontWeight:700,fontSize:12,marginBottom:4}}>⚠ 未照合の個体について</div>
                  <div style={{color:C.textMid,fontSize:11,lineHeight:1.6}}>
                    耳標番号が一致する個体が見つかりませんでした。<br/>
                    先に個体を登録してから再度お試しください。
                  </div>
                </div>
              )}

              <div style={{display:"flex",gap:10}}>
                <Btn variant="outline" full onClick={()=>{setStatus("idle");setPreview(null);}}>
                  再読み込み
                </Btn>
                <Btn full onClick={saveAll} disabled={!matched.some(m=>m.selected&&m.cow)}>
                  {matched.filter(m=>m.selected&&m.cow).length}頭を保存
                </Btn>
              </div>
            </div>
          )}

        </div>
      </div>
    );
  };

  // ── INTAKE FLOW（新規導入フロー） ─────────────────────────────────────────
  const IntakeScreen = () => {
    const fileRef = useRef(null);
    const [step, setStep]           = useState(0);   // 0=start 1=伝票OCR 2=伝票確認 3=登記書OCR 4=完了確認
    const [slipStatus, setSlipStatus] = useState("idle"); // idle|loading|done|error
    const [slipPreview, setSlipPreview] = useState(null);
    const [slipErr, setSlipErr]     = useState("");
    const [certIdx, setCertIdx]     = useState(0);   // 今何頭目の登記書を撮影中
    const [certStatus, setCertStatus] = useState("idle");
    const [certPreview, setCertPreview] = useState(null);
    // 導入共通情報
    const [introDate, setIntroDate] = useState(new Date().toISOString().slice(0,10));
    const [farm, setFarm]           = useState("");
    const [pen, setPen]             = useState("");
    const [feedCostPerDay, setFeedCostPerDay] = useState(String(settings.defaultCosts.compoundKgPerDay||8));
    // 頭リスト
    const [animals, setAnimals]     = useState([]); // [{tag,purchasePrice,name,sex,breed,birthDate,pedigree,certDone}]

    const totalSteps = 4;

    /* ── OCR helper ─────────────────────────────────── */
    const doOcr = async (file, prompt) => {
      const b64 = await new Promise((res,rej)=>{
        const r=new FileReader();
        r.onload=()=>res(r.result.split(",")[1]);
        r.onerror=rej;
        r.readAsDataURL(file);
      });
      const resp = await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":window.ANTHROPIC_KEY||"","anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
        body:JSON.stringify({
          model:"claude-sonnet-4-20250514",
          max_tokens:1200,
          messages:[{role:"user",content:[
            {type:"image",source:{type:"base64",media_type:file.type||"image/jpeg",data:b64}},
            {type:"text",text:prompt}
          ]}]
        })
      });
      const data = await resp.json();
      const text = data.content?.map(c=>c.text||"").join("")||"";
      return JSON.parse(text.replace(/```json|```/g,"").trim());
    };

    /* ── STEP 1: 伝票OCR ────────────────────────────── */
    const handleSlip = async (file) => {
      setSlipStatus("loading");
      setSlipErr("");
      const reader=new FileReader();
      reader.onload=e=>setSlipPreview(e.target.result);
      reader.readAsDataURL(file);
      try {
        const prompt = `この画像は和牛の購買伝票・市場の購買証明書です。
以下のJSON形式で読み取ってください。読めない項目は""にしてください。
{
  "introDate": "導入日（YYYY-MM-DD）",
  "farm": "市場名・農場名",
  "animals": [
    {"tag":"耳標番号","purchasePrice":購入価格の数値,"name":"牛名（あれば）","sex":"去勢か雌か雄","birthDate":"生年月日YYYY-MM-DD"}
  ]
}
複数頭ある場合はanimalsに全頭分入れてください。JSONのみ返してください。`;
        const result = await doOcr(file, prompt);
        if(result.introDate) setIntroDate(result.introDate);
        if(result.farm) setFarm(result.farm);
        const parsed = (result.animals||[]).map((a,i)=>({
          id: `new-${Date.now()}-${i}`,
          tag: a.tag||"",
          purchasePrice: a.purchasePrice||0,
          name: a.name||"",
          sex: a.sex||"去勢",
          breed: "黒毛和種",
          birthDate: a.birthDate||"",
          pedigree: emptyPedigree(),
          certDone: false,
        }));
        setAnimals(parsed.length>0 ? parsed : [newAnimal()]);
        setSlipStatus("done");
      } catch(e) {
        setSlipErr("読み取りできませんでした。手動で入力してください。");
        setSlipStatus("error");
        setAnimals([newAnimal()]);
      }
    };

    /* ── STEP 3: 登記書OCR ──────────────────────────── */
    const handleCert = async (file) => {
      setCertStatus("loading");
      const reader=new FileReader();
      reader.onload=e=>setCertPreview(e.target.result);
      reader.readAsDataURL(file);
      try {
        const prompt = `この画像は和牛子牛登記証明書です。以下のJSON形式で抽出してください。読み取れない項目は""にしてください。
{
  "tag":"耳標番号",
  "name":"牛名",
  "sex":"去勢か雌か雄",
  "breed":"品種",
  "birthDate":"生年月日YYYY-MM-DD",
  "pedigree":{
    "sire":{"name":"父","sire":{"name":"父の父","sire":{"name":"父の父の父"},"dam":{"name":"父の父の母"}},"dam":{"name":"父の母","sire":{"name":"父の母の父"},"dam":{"name":"父の母の母"}}},
    "dam":{"name":"母","sire":{"name":"母の父","sire":{"name":"母の父の父"},"dam":{"name":"母の父の母"}},"dam":{"name":"母の母","sire":{"name":"母の母の父"},"dam":{"name":"母の母の母"}}}
  }
}
JSONのみ返してください。`;
        const result = await doOcr(file, prompt);
        setAnimals(prev => prev.map((a,i) => {
          if(i !== certIdx) return a;
          // 耳標が一致するか、または現在のインデックスに適用
          return {
            ...a,
            tag:       result.tag||a.tag,
            name:      result.name||a.name,
            sex:       result.sex||a.sex,
            breed:     result.breed||a.breed,
            birthDate: result.birthDate||a.birthDate,
            pedigree:  result.pedigree ? mergeNode2(a.pedigree, result.pedigree) : a.pedigree,
            certDone:  true,
          };
        }));
        setCertStatus("done");
      } catch(e) {
        setCertStatus("error");
      }
    };

    const mergeNode2 = (base, src) => {
      if(!src) return base||{name:"",sire:{name:""},dam:{name:""}};
      return {
        name: src.name||base?.name||"",
        sire: mergeNode2(base?.sire, src.sire),
        dam:  mergeNode2(base?.dam,  src.dam),
      };
    };

    const newAnimal = () => ({
      id:`new-${Date.now()}-${Math.random()}`,
      tag:"", purchasePrice:0, name:"", sex:"去勢", breed:"黒毛和種",
      birthDate:"", pedigree:emptyPedigree(), certDone:false,
    });
    const updateAnimal = (i,fn) => setAnimals(prev=>prev.map((a,idx)=>idx===i?fn(a):a));

    /* ── Excel / CSV 読み込み ─────────────────────── */
    const [xlsxStatus, setXlsxStatus] = useState("idle"); // idle|loading|done|error
    const [xlsxMsg,    setXlsxMsg]    = useState("");

    const handleExcel = async (file) => {
      if(!file) return;
      setXlsxStatus("loading");
      setXlsxMsg("");
      try {
        let csvText = "";
        const name = file.name.toLowerCase();

        if(name.endsWith(".csv")) {
          // CSV: そのまま読む
          csvText = await file.text();
        } else {
          // Excel: SheetJSで変換
          const buf = await file.arrayBuffer();
          const wb  = XLSX.read(buf, {type:"array", cellDates:true});
          const ws  = wb.Sheets[wb.SheetNames[0]];
          csvText   = XLSX.utils.sheet_to_csv(ws);
        }

        // 先頭3000字をClaudeに渡して解析
        const prompt = `以下は和牛農場のExcel（CSV変換）データです。
各行を個体データとして解釈し、以下のJSON配列で返してください。
列名は日本語・英語・略称どれでも対応してください。
読み取れない項目はnullにしてください。

CSV:
${csvText.slice(0, 4000)}

返すJSON配列の型:
[{
  "tag": "耳標番号",
  "name": "牛名",
  "sex": "去勢|雌|雄",
  "breed": "品種",
  "birthDate": "YYYY-MM-DD",
  "introDate": "YYYY-MM-DD",
  "purchasePrice": 数値,
  "farm": "導入市場",
  "pen": "牛舎",
  "shippingPlan": "YYYY-MM-DD",
  "expectedPrice": 数値,
  "sire": "父名",
  "sireSire": "父の父名",
  "dam": "母名",
  "damSire": "母の父名",
  "memo": "メモ"
}]
ヘッダー行・空行・合計行は無視してください。JSONのみ返してください。`;

        const resp = await fetch("https://api.anthropic.com/v1/messages",{
          method:"POST",
          headers:{"Content-Type":"application/json","x-api-key":window.ANTHROPIC_KEY||"","anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
          body:JSON.stringify({
            model:"claude-sonnet-4-20250514",
            max_tokens:2000,
            messages:[{role:"user",content:prompt}]
          })
        });
        const data = await resp.json();
        const text = data.content?.map(c=>c.text||"").join("")||"";
        const parsed = JSON.parse(text.replace(/```json|```/g,"").trim());
        const arr = Array.isArray(parsed) ? parsed : [parsed];

        // animals配列に変換
        const mapped = arr.filter(r=>r.tag||r.name).map(r => ({
          id: `xl-${Date.now()}-${Math.random()}`,
          tag:           r.tag||"",
          name:          r.name||"",
          sex:           r.sex||"去勢",
          breed:         r.breed||"黒毛和種",
          birthDate:     r.birthDate||"",
          purchasePrice: Number(r.purchasePrice)||0,
          pen:           r.pen||"",
          pedigree: {
            sire:{ name:r.sire||"",
              sire:{ name:r.sireSire||"", sire:{name:""}, dam:{name:""}},
              dam: { name:"",             sire:{name:""}, dam:{name:""}} },
            dam: { name:r.dam||"",
              sire:{ name:r.damSire||"",  sire:{name:""}, dam:{name:""}},
              dam: { name:"",             sire:{name:""}, dam:{name:""}} },
          },
          certDone: false,
          _shippingPlan:  r.shippingPlan||"",
          _expectedPrice: Number(r.expectedPrice)||0,
          _memo:          r.memo||"",
        }));

        // 共通情報（最初の行から）
        if(arr[0]?.introDate)  setIntroDate(arr[0].introDate);
        if(arr[0]?.farm)       setFarm(arr[0].farm);

        setAnimals(mapped.length>0 ? mapped : [newAnimal()]);
        setXlsxStatus("done");
        setXlsxMsg(`${mapped.length}頭を読み込みました`);
        setStep(2); // 確認ステップへ
      } catch(e) {
        console.error(e);
        setXlsxStatus("error");
        setXlsxMsg("読み込みに失敗しました。ファイル形式を確認して再試行してください。");
      }
    };

    /* ── 一括登録 ───────────────────────────────────── */
    const registerAll = () => {
      const newCattle = animals.map(a=>({
        ...a,
        id: Date.now().toString()+Math.random(),
        introDate, farm, pen,
        shippingPlan:"", expectedPrice:0, memo:"", status:"肥育中", result:null,
        weights:[], vaccines:[], treatments:[],
        costs:{
          purchasePrice:    a.purchasePrice,
          roughageDaily:    settings.defaultCosts.roughageDaily,
          compoundKgPerDay: Number(feedCostPerDay)||settings.defaultCosts.compoundKgPerDay,
          compoundKgPrice:  settings.defaultCosts.compoundKgPrice,
          otherDaily:       settings.defaultCosts.otherDaily,
          fixedOther:       settings.defaultCosts.fixedOther,
          vetCosts: 0,
        },
      }));
      setCattle(prev=>[...prev,...newCattle]);
      setPage("home");
    };

    /* ── STEP INDICATOR ─────────────────────────────── */
    const StepDot = ({n,current}) => (
      <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
        <div style={{
          width:28, height:28, borderRadius:"50%",
          background: n<current?"#fff":n===current?`linear-gradient(135deg,${C.accent},${C.accentDark})`:"rgba(255,255,255,0.3)",
          color: n<current?C.accentDark:n===current?"#fff":"rgba(255,255,255,0.6)",
          display:"flex", alignItems:"center", justifyContent:"center",
          fontSize:12, fontWeight:900, border: n<current?`2px solid ${C.accentDark}`:"none",
        }}>
          {n<current?"✓":n}
        </div>
        <div style={{fontSize:8,color:n===current?"#fff":"rgba(255,255,255,0.6)",whiteSpace:"nowrap",fontWeight:n===current?700:400}}>
          {["","伝票","確認","登記書","登録"][n]}
        </div>
      </div>
    );

    const stepLabels = ["","📄 伝票撮影","🔍 内容確認","📋 登記書撮影","✅ 登録確認"];

    return (
      <div style={{background:C.bg,minHeight:"100vh",fontFamily:"'Hiragino Kaku Gothic Pro','Noto Sans JP',sans-serif",maxWidth:520,margin:"0 auto"}}>
        {/* ヘッダー */}
        <div style={{
          background:`linear-gradient(135deg,${C.accent},${C.accentDark})`,
          padding:"16px 20px 14px",
          position:"sticky", top:0, zIndex:80,
        }}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
            <button onClick={()=>setPage("home")} style={{background:"rgba(255,255,255,0.22)",border:"none",color:"#fff",borderRadius:10,padding:"7px 14px",fontSize:13,fontWeight:700,cursor:"pointer"}}>← 戻る</button>
            <span style={{color:"#fff",fontWeight:900,fontSize:16}}>🐄 新規導入フロー</span>
            <div style={{width:60}}/>
          </div>
          {/* ステップインジケーター */}
          <div style={{display:"flex",alignItems:"flex-start",justifyContent:"center",gap:0}}>
            {[1,2,3,4].map((n,i)=>(
              <div key={n} style={{display:"flex",alignItems:"center"}}>
                <StepDot n={n} current={step}/>
                {i<3&&<div style={{width:36,height:2,background:n<step?"rgba(255,255,255,0.8)":"rgba(255,255,255,0.25)",marginTop:-14,marginBottom:0}}/>}
              </div>
            ))}
          </div>
        </div>

        <div style={{padding:"20px 18px 100px"}}>

          {/* ━━━━ STEP 0: スタート ━━━━ */}
          {step===0&&(
            <div style={{padding:"12px 0"}}>
              <div style={{textAlign:"center",marginBottom:28}}>
                <div style={{fontSize:56,marginBottom:12}}>🐂</div>
                <div style={{color:C.text,fontWeight:900,fontSize:22,marginBottom:6}}>新規導入を開始します</div>
                <div style={{color:C.textMid,fontSize:13,lineHeight:1.8}}>
                  市場から複数頭を一括登録できます。
                </div>
              </div>

              {/* ━ Excel / CSV からインポート（おすすめ） ━ */}
              <div style={{
                background:`linear-gradient(135deg,${C.green}18,${C.teal}10)`,
                border:`2px solid ${C.green}66`,
                borderRadius:18, padding:"18px 18px", marginBottom:12,
              }}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                  <span style={{background:C.green,color:"#fff",borderRadius:8,padding:"3px 10px",fontSize:11,fontWeight:800}}>おすすめ</span>
                  <span style={{color:C.green,fontWeight:800,fontSize:15}}>📊 Excelから一括読み込み</span>
                </div>
                <div style={{color:C.textMid,fontSize:12,lineHeight:1.7,marginBottom:14}}>
                  今まで管理していたExcel・CSVファイルを選択するだけ。<br/>
                  AIが列名を自動判定して耳標・血統・価格などを一括取り込みます。
                </div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10,fontSize:11,color:C.textDim}}>
                  {["耳標番号","牛名","性別","品種","生年月日","導入日","購入価格","導入市場","牛舎","出荷予定","父名","母名"].map(t=>(
                    <span key={t} style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:10,padding:"2px 8px"}}>{t}</span>
                  ))}
                  <span style={{color:C.textDim}}>など自動対応</span>
                </div>
                {xlsxStatus==="error"&&(
                  <div style={{color:C.red,fontSize:12,marginBottom:8,background:C.redLight,borderRadius:8,padding:"6px 10px"}}>{xlsxMsg}</div>
                )}
                {xlsxStatus==="loading"&&(
                  <div style={{color:C.green,fontSize:13,fontWeight:700,textAlign:"center",padding:"8px 0"}}>
                    🔍 AIが解析中...
                  </div>
                )}
                <div style={{display:"flex",gap:8}}>
                  <button onClick={()=>{
                    const inp2=document.createElement("input");
                    inp2.type="file"; inp2.accept=".xlsx,.xls,.csv";
                    inp2.onchange=e=>handleExcel(e.target.files?.[0]);
                    inp2.click();
                  }} style={{
                    flex:1, background:C.green, color:"#fff", border:"none",
                    borderRadius:12, padding:"12px 0", fontSize:14, fontWeight:800,
                    cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", gap:6,
                  }}>📂 ファイルを選択（xlsx / csv）</button>
                </div>
              </div>

              {/* ━ 伝票撮影から開始 ━ */}
              <div onClick={()=>setStep(1)} style={{
                background:"#fff", border:`1.5px solid ${C.border}`,
                borderRadius:16, padding:"16px 18px", marginBottom:10,
                cursor:"pointer", display:"flex", alignItems:"center", gap:14,
                boxShadow:C.shadow,
              }}>
                <div style={{width:48,height:48,borderRadius:14,background:C.accentLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,flexShrink:0}}>📄</div>
                <div>
                  <div style={{color:C.text,fontWeight:800,fontSize:15}}>購買伝票から開始</div>
                  <div style={{color:C.textMid,fontSize:12,marginTop:2}}>伝票撮影 → 登記書撮影 → 一括登録</div>
                </div>
                <span style={{color:C.textDim,marginLeft:"auto",fontSize:18}}>→</span>
              </div>

              {/* ━ 手動入力 ━ */}
              <div onClick={()=>{setAnimals([newAnimal()]);setStep(2);}} style={{
                background:"#fff", border:`1.5px solid ${C.border}`,
                borderRadius:16, padding:"14px 18px",
                cursor:"pointer", display:"flex", alignItems:"center", gap:14,
                boxShadow:C.shadow,
              }}>
                <div style={{width:48,height:48,borderRadius:14,background:C.cardSub,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,flexShrink:0}}>✏️</div>
                <div>
                  <div style={{color:C.text,fontWeight:800,fontSize:15}}>手動で入力</div>
                  <div style={{color:C.textMid,fontSize:12,marginTop:2}}>1頭ずつ手入力して登録</div>
                </div>
                <span style={{color:C.textDim,marginLeft:"auto",fontSize:18}}>→</span>
              </div>
            </div>
          )}

          {/* ━━━━ STEP 1: 伝票撮影 ━━━━ */}
          {step===1&&(
            <div>
              <div style={{color:C.text,fontWeight:900,fontSize:18,marginBottom:4}}>📄 購買伝票を撮影</div>
              <div style={{color:C.textMid,fontSize:13,marginBottom:20,lineHeight:1.6}}>
                市場の購買伝票・購買証明書を撮影してください。<br/>
                AIが<b>導入日・市場名・各頭の耳標・価格</b>を自動で読み取ります。
              </div>

              <input ref={fileRef} type="file" accept="image/*" capture="environment" style={{display:"none"}}
                onChange={e=>handleSlip(e.target.files?.[0])}/>

              {slipStatus==="idle"&&(
                <div>
                  {/* カメラボタン */}
                  <div onClick={()=>{fileRef.current?.setAttribute("capture","environment");fileRef.current?.click();}}
                    style={{
                      background:`linear-gradient(135deg,${C.teal}22,${C.accentLight})`,
                      border:`2px dashed ${C.teal}`,
                      borderRadius:20, padding:"36px 20px", textAlign:"center",
                      cursor:"pointer", marginBottom:12,
                    }}>
                    <div style={{fontSize:48,marginBottom:10}}>📷</div>
                    <div style={{color:C.teal,fontWeight:800,fontSize:16,marginBottom:6}}>カメラで撮影</div>
                    <div style={{color:C.textMid,fontSize:12}}>購買伝票をカメラに向けてください</div>
                  </div>
                  <Btn full variant="soft" icon="🖼️"
                    onClick={()=>{fileRef.current?.removeAttribute("capture");fileRef.current?.click();}}>
                    ファイルから選択
                  </Btn>
                  <div style={{textAlign:"center",marginTop:14}}>
                    <button onClick={()=>{setAnimals([newAnimal()]);setStep(2);}} style={{background:"none",border:"none",color:C.textDim,fontSize:13,cursor:"pointer",textDecoration:"underline"}}>
                      伝票なしでスキップ →
                    </button>
                  </div>
                </div>
              )}

              {slipStatus==="loading"&&(
                <div style={{textAlign:"center",padding:"32px 0"}}>
                  {slipPreview&&<img src={slipPreview} style={{width:"100%",maxHeight:220,objectFit:"contain",borderRadius:14,marginBottom:16}} alt="伝票"/>}
                  <div style={{fontSize:36,marginBottom:10}}>🔍</div>
                  <div style={{color:C.accentDark,fontWeight:700,fontSize:16}}>AIが読み取り中...</div>
                  <div style={{color:C.textMid,fontSize:12,marginTop:6}}>耳標・価格を解析しています</div>
                </div>
              )}

              {(slipStatus==="done"||slipStatus==="error")&&(
                <div>
                  {slipPreview&&<img src={slipPreview} style={{width:"100%",maxHeight:180,objectFit:"contain",borderRadius:14,marginBottom:14}} alt="伝票"/>}
                  {slipStatus==="done"&&(
                    <div style={{background:C.greenLight,border:`1px solid ${C.green}44`,borderRadius:12,padding:"12px 14px",marginBottom:14}}>
                      <div style={{color:C.green,fontWeight:700,marginBottom:6}}>✅ 読み取り完了</div>
                      <div style={{fontSize:12,color:C.text}}>導入日: {introDate}　市場: {farm||"未読取"}</div>
                      <div style={{fontSize:12,color:C.text,marginTop:2}}>{animals.length}頭分を読み取りました</div>
                    </div>
                  )}
                  {slipStatus==="error"&&(
                    <div style={{background:C.redLight,border:`1px solid ${C.red}33`,borderRadius:12,padding:"12px 14px",marginBottom:14}}>
                      <div style={{color:C.red,fontWeight:700}}>{slipErr}</div>
                    </div>
                  )}
                  <Btn full onClick={()=>setStep(2)}>内容を確認する →</Btn>
                </div>
              )}
            </div>
          )}

          {/* ━━━━ STEP 2: 伝票内容確認 ━━━━ */}
          {step===2&&(
            <div>
              <div style={{color:C.text,fontWeight:900,fontSize:18,marginBottom:16}}>🔍 内容を確認</div>

              {/* 共通情報 */}
              <Card style={{marginBottom:16}}>
                <SectionLabel>導入共通情報</SectionLabel>
                <FInput label="導入日 ＊">
                  <input type="date" value={introDate} onChange={e=>setIntroDate(e.target.value)} style={inp}/>
                </FInput>
                <FInput label="導入市場・農場">
                  <input type="text" value={farm} onChange={e=>setFarm(e.target.value)} placeholder="例: 宮崎中央市場" style={inp}/>
                </FInput>
                <FInput label="牛舎・ペン">
                  <input type="text" value={pen} onChange={e=>setPen(e.target.value)} placeholder="例: 2号棟A" style={inp}/>
                </FInput>
                <FInput label="飼料費/日（円・全頭共通）">
                  <input type="number" value={feedCostPerDay} onChange={e=>setFeedCostPerDay(e.target.value)} placeholder="1200" style={inp}/>
                </FInput>
              </Card>

              {/* 個体リスト */}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <SectionLabel>{animals.length}頭分の個体情報</SectionLabel>
                <button onClick={()=>setAnimals(p=>[...p,newAnimal()])} style={{background:C.accentLight,border:`1px solid ${C.border}`,color:C.accentDark,borderRadius:10,padding:"5px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
                  ＋頭追加
                </button>
              </div>

              {animals.map((a,i)=>(
                <Card key={a.id} style={{marginBottom:10}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                    <span style={{color:C.accent,fontWeight:800,fontSize:14}}>{i+1}頭目</span>
                    {animals.length>1&&(
                      <button onClick={()=>setAnimals(p=>p.filter((_,idx)=>idx!==i))} style={{background:C.redLight,border:"none",color:C.red,borderRadius:8,padding:"4px 10px",fontSize:12,cursor:"pointer"}}>削除</button>
                    )}
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                    <FInput label="耳標番号">
                      <input value={a.tag} onChange={e=>updateAnimal(i,x=>({...x,tag:e.target.value}))} placeholder="例: 宮崎-0099" style={{...inp,fontSize:13,padding:"8px 10px"}}/>
                    </FInput>
                    <FInput label="購入価格（円）">
                      <input type="number" value={a.purchasePrice||""} onChange={e=>updateAnimal(i,x=>({...x,purchasePrice:Number(e.target.value)||0}))} placeholder="例: 820000" style={{...inp,fontSize:13,padding:"8px 10px"}}/>
                    </FInput>
                    <FInput label="牛名">
                      <input value={a.name} onChange={e=>updateAnimal(i,x=>({...x,name:e.target.value}))} placeholder="任意" style={{...inp,fontSize:13,padding:"8px 10px"}}/>
                    </FInput>
                    <FInput label="性別">
                      <select value={a.sex} onChange={e=>updateAnimal(i,x=>({...x,sex:e.target.value}))} style={{...inp,fontSize:13,padding:"8px 10px"}}>
                        {["去勢","雌","雄"].map(s=><option key={s}>{s}</option>)}
                      </select>
                    </FInput>
                  </div>
                </Card>
              ))}

              <div style={{marginTop:8}}>
                <Btn full onClick={()=>{setCertIdx(0);setCertStatus("idle");setCertPreview(null);setStep(3);}}>
                  登記書撮影へ進む →
                </Btn>
              </div>
            </div>
          )}

          {/* ━━━━ STEP 3: 登記書OCR ━━━━ */}
          {step===3&&(
            <div>
              {/* 進捗バー */}
              <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:16,overflowX:"auto",paddingBottom:4}}>
                {animals.map((a,i)=>(
                  <div key={a.id} onClick={()=>{setCertIdx(i);setCertStatus("idle");setCertPreview(null);}}
                    style={{
                      flexShrink:0, borderRadius:10, padding:"6px 12px", cursor:"pointer",
                      background: a.certDone ? C.greenLight : i===certIdx ? C.accentLight : C.cardSub,
                      border: `1.5px solid ${a.certDone?C.green:i===certIdx?C.accent:C.border}`,
                      fontSize:12, fontWeight:700,
                      color: a.certDone?C.green:i===certIdx?C.accentDark:C.textDim,
                    }}>
                    {a.certDone?"✅ ":"📋 "}{i+1}頭目{a.tag?` ${a.tag}`:""}
                  </div>
                ))}
              </div>

              <div style={{color:C.text,fontWeight:900,fontSize:18,marginBottom:4}}>
                📋 {certIdx+1}頭目の登記書
                <span style={{color:C.textDim,fontSize:13,fontWeight:400,marginLeft:8}}>
                  （{animals.length}頭中）
                </span>
              </div>
              {animals[certIdx]?.tag&&(
                <div style={{color:C.accent,fontFamily:"monospace",fontWeight:700,fontSize:14,marginBottom:10}}>
                  {animals[certIdx].tag}
                </div>
              )}
              <div style={{color:C.textMid,fontSize:13,marginBottom:20,lineHeight:1.6}}>
                子牛登記証明書を撮影してください。<br/>AIが<b>三代祖の血統</b>を自動で読み取ります。
              </div>

              <input ref={fileRef} type="file" accept="image/*" capture="environment" style={{display:"none"}}
                onChange={e=>handleCert(e.target.files?.[0])}/>

              {certStatus==="idle"&&(
                <div>
                  <div onClick={()=>{fileRef.current?.setAttribute("capture","environment");fileRef.current?.click();}}
                    style={{
                      background:`linear-gradient(135deg,${C.purple}18,${C.accentLight})`,
                      border:`2px dashed ${C.purple}`,
                      borderRadius:20, padding:"32px 20px", textAlign:"center",
                      cursor:"pointer", marginBottom:12,
                    }}>
                    <div style={{fontSize:44,marginBottom:10}}>📋</div>
                    <div style={{color:C.purple,fontWeight:800,fontSize:16,marginBottom:6}}>登記書を撮影</div>
                    <div style={{color:C.textMid,fontSize:12}}>証明書をカメラで撮影してください</div>
                  </div>
                  <Btn full variant="soft" icon="🖼️"
                    onClick={()=>{fileRef.current?.removeAttribute("capture");fileRef.current?.click();}}>
                    ファイルから選択
                  </Btn>
                  <div style={{textAlign:"center",marginTop:12}}>
                    <button onClick={()=>{updateAnimal(certIdx,a=>({...a,certDone:true}));if(certIdx<animals.length-1){setCertIdx(certIdx+1);setCertStatus("idle");setCertPreview(null);}else{setStep(4);}}} style={{background:"none",border:"none",color:C.textDim,fontSize:13,cursor:"pointer",textDecoration:"underline"}}>
                      この頭はスキップ →
                    </button>
                  </div>
                </div>
              )}

              {certStatus==="loading"&&(
                <div style={{textAlign:"center",padding:"28px 0"}}>
                  {certPreview&&<img src={certPreview} style={{width:"100%",maxHeight:200,objectFit:"contain",borderRadius:14,marginBottom:14}} alt="登記書"/>}
                  <div style={{fontSize:32,marginBottom:8}}>🔍</div>
                  <div style={{color:C.purple,fontWeight:700,fontSize:15}}>血統を読み取り中...</div>
                  <div style={{color:C.textMid,fontSize:12,marginTop:4}}>三代祖を解析しています</div>
                </div>
              )}

              {certStatus==="done"&&(
                <div>
                  {certPreview&&<img src={certPreview} style={{width:"100%",maxHeight:160,objectFit:"contain",borderRadius:14,marginBottom:12}} alt="登記書"/>}
                  <div style={{background:C.greenLight,border:`1px solid ${C.green}44`,borderRadius:12,padding:"12px 14px",marginBottom:16}}>
                    <div style={{color:C.green,fontWeight:700,marginBottom:6}}>✅ 血統読み取り完了</div>
                    {[
                      ["牛名",animals[certIdx]?.name],
                      ["父",animals[certIdx]?.pedigree?.sire?.name],
                      ["父の父",animals[certIdx]?.pedigree?.sire?.sire?.name],
                      ["母",animals[certIdx]?.pedigree?.dam?.name],
                      ["母の父",animals[certIdx]?.pedigree?.dam?.sire?.name],
                    ].filter(([,v])=>v).map(([k,v])=>(
                      <div key={k} style={{display:"flex",gap:8,fontSize:12,marginBottom:3}}>
                        <span style={{color:C.textDim,minWidth:56}}>{k}</span>
                        <span style={{color:C.text,fontWeight:600}}>{v}</span>
                      </div>
                    ))}
                  </div>
                  {certIdx < animals.length-1 ? (
                    <Btn full onClick={()=>{setCertIdx(certIdx+1);setCertStatus("idle");setCertPreview(null);}}>
                      次の頭（{certIdx+2}頭目）へ →
                    </Btn>
                  ):(
                    <Btn full onClick={()=>setStep(4)}>
                      全頭完了　→ 登録確認へ
                    </Btn>
                  )}
                </div>
              )}

              {certStatus==="error"&&(
                <div>
                  <div style={{background:C.redLight,borderRadius:12,padding:"12px 14px",marginBottom:12}}>
                    <div style={{color:C.red,fontWeight:700}}>読み取りに失敗しました</div>
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <Btn variant="outline" full onClick={()=>{setCertStatus("idle");setCertPreview(null);}}>再撮影</Btn>
                    <Btn variant="soft" full onClick={()=>{updateAnimal(certIdx,a=>({...a,certDone:true}));if(certIdx<animals.length-1){setCertIdx(certIdx+1);setCertStatus("idle");setCertPreview(null);}else{setStep(4);}}}>スキップ</Btn>
                  </div>
                </div>
              )}

              {/* 全頭撮影済みなら確認へ */}
              {animals.every(a=>a.certDone)&&certStatus!=="loading"&&(
                <div style={{marginTop:16}}>
                  <Btn full onClick={()=>setStep(4)}>全頭完了 → 登録確認へ</Btn>
                </div>
              )}
            </div>
          )}

          {/* ━━━━ STEP 4: 登録確認 ━━━━ */}
          {step===4&&(
            <div>
              <div style={{color:C.text,fontWeight:900,fontSize:18,marginBottom:16}}>✅ 登録内容の確認</div>

              <Card style={{marginBottom:14}}>
                <SectionLabel>導入共通情報</SectionLabel>
                <InfoRow label="導入日" value={introDate}/>
                <InfoRow label="市場・農場" value={farm||"―"}/>
                <InfoRow label="牛舎・ペン" value={pen||"―"}/>
                <InfoRow label="頭数" value={`${animals.length}頭`} accent/>
                <InfoRow label="飼料費/日" value={`¥${Number(feedCostPerDay||0).toLocaleString()}/日`} last/>
              </Card>

              <SectionLabel>{animals.length}頭の個体一覧</SectionLabel>
              {animals.map((a,i)=>(
                <Card key={a.id} style={{marginBottom:10}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                    <div style={{display:"flex",gap:6,alignItems:"center"}}>
                      <span style={{color:C.accent,fontWeight:800,fontFamily:"monospace",fontSize:13}}>{a.tag||"耳標未入力"}</span>
                      {a.name&&<span style={{color:C.text,fontSize:13}}>{a.name}</span>}
                    </div>
                    <div style={{display:"flex",gap:5}}>
                      <Tag label={a.sex} color={a.sex==="雌"?"#e06090":C.purple}/>
                      {a.certDone&&<Tag label="血統✓" color={C.green} bg={C.greenLight}/>}
                    </div>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
                    <div style={{background:C.cardSub,borderRadius:8,padding:"6px 10px"}}>
                      <div style={{color:C.textDim,fontSize:9}}>購入価格</div>
                      <div style={{color:C.amber,fontWeight:700,fontSize:13}}>{fmtMoney(a.purchasePrice)}</div>
                    </div>
                    <div style={{background:C.cardSub,borderRadius:8,padding:"6px 10px"}}>
                      <div style={{color:C.textDim,fontSize:9}}>父</div>
                      <div style={{color:C.purple,fontWeight:700,fontSize:12}}>{a.pedigree?.sire?.name||"―"}</div>
                    </div>
                  </div>
                </Card>
              ))}

              <div style={{
                background:C.amberLight,border:`1px solid ${C.amber}44`,borderRadius:12,
                padding:"12px 14px",marginBottom:20,
              }}>
                <div style={{color:C.amber,fontWeight:700,fontSize:12,marginBottom:4}}>💴 導入総額</div>
                <div style={{color:C.amber,fontWeight:900,fontSize:22}}>
                  {fmtMoney(animals.reduce((s,a)=>s+a.purchasePrice,0))}
                </div>
              </div>

              <Btn full onClick={registerAll} icon="✅">{animals.length}頭を一括登録する</Btn>
              <div style={{textAlign:"center",marginTop:12}}>
                <button onClick={()=>setStep(3)} style={{background:"none",border:"none",color:C.textDim,fontSize:13,cursor:"pointer",textDecoration:"underline"}}>
                  ← 登記書撮影に戻る
                </button>
              </div>
            </div>
          )}

        </div>
      </div>
    );
  };

  // ── AI 解析ページ ──────────────────────────────────────────────────────────
  const AiScreen = () => {
    const [loading, setLoading] = useState(false);
    const [result,  setResult]  = useState(null);
    const [question, setQuestion] = useState("");
    const [activePreset, setActivePreset] = useState(null);

    const farmSummary = () => {
      const active = cattle.filter(c=>c.status==="肥育中");
      return {
        totalHead: cattle.length,
        activeHead: active.length,
        avgDG: (()=>{
          const dgs=active.map(c=>calcDG(c.weights)).filter(Boolean);
          return dgs.length ? (dgs.reduce((s,v)=>s+v,0)/dgs.length).toFixed(2) : null;
        })(),
        totalCost: active.reduce((s,c)=>s+calcCosts(c).totalCost,0),
        totalExpectedRevenue: active.reduce((s,c)=>s+(c.expectedPrice||0),0),
        animals: cattle.map(c=>({
          tag:c.tag, name:c.name, sex:c.sex, breed:c.breed,
          sire:c.pedigree?.sire?.name,
          introDate:c.introDate, shippingPlan:c.shippingPlan,
          status:c.status,
          latestWeight:latestWeight(c.weights),
          dg:calcDG(c.weights)?.toFixed(2),
          totalCost:calcCosts(c).totalCost,
          expectedProfit:calcCosts(c).profit,
          result:c.result,
          pen:c.pen,
        })),
      };
    };

    const PRESETS = [
      { id:"overview", icon:"📊", label:"農場全体を分析",       color:C.accent,
        prompt:"この農場の現状を分析し、強み・課題・改善提案を教えてください。"},
      { id:"dg",       icon:"📈", label:"DGランキング＆予測",   color:C.green,
        prompt:"各個体のDG（日増体重）をランキングし、出荷時の予測体重と今後の発育傾向を評価してください。"},
      { id:"profit",   icon:"💴", label:"損益・コスト最適化",   color:C.amber,
        prompt:"各個体の損益状況を分析し、コスト削減や収益改善のアドバイスを具体的に提案してください。"},
      { id:"ship",     icon:"🚚", label:"出荷タイミング提案",   color:C.purple,
        prompt:"現在のDGや体重から、各個体の最適な出荷タイミングと予想販売価格について提案してください。"},
      { id:"blood",    icon:"🧬", label:"血統別パフォーマンス", color:"#e06090",
        prompt:"血統（父牛）ごとのDG・推定BMS・損益を比較し、今後どの血統の素牛を導入すべきか分析してください。"},
      { id:"feed",     icon:"🌾", label:"餌相談",               color:C.teal,
        prompt:`この農場の現在の飼養コスト（粗飼料費・配合飼料費・給与量）と各個体の発育状況をもとに、以下の点でアドバイスしてください。
1. 各ステージ（導入期・増体期・仕上げ期）に応じた配合飼料の給与量・切り替えタイミング
2. 日増体重（DG）が低い個体への対処法
3. コストを抑えつつ品質（BMS）を上げるための飼料設計の工夫
4. 粗飼料・配合飼料のバランスについての提案`},
      { id:"calf",     icon:"🐂", label:"素牛相場予想",         color:"#8b6914",
        prompt:`現在の農場データ（血統・導入価格・時期）と一般的な和牛市場の傾向をもとに、以下を分析してください。
1. 現在の素牛導入価格の水準評価（高い・適正・安い）
2. 今後3〜6ヶ月の素牛相場の見通しと価格変動要因
3. 次回導入に向けた最適な購入時期・市場・血統の提案
4. 導入コストを抑えるための戦略的アドバイス`},
      { id:"carcass",  icon:"🥩", label:"枝肉相場予想",         color:"#c0392b",
        prompt:`この農場の出荷予定個体（血統・DG・予測体重・出荷時期）と一般的な和牛枝肉市場の傾向をもとに、以下を分析してください。
1. 各個体の出荷時期における枝肉相場の見通し
2. BMS・等級別の価格動向と本農場個体への影響
3. 販売価格を最大化するための出荷時期・体重調整の提案
4. 季節需要や市場サイクルを踏まえた出荷戦略`},
      { id:"intro",    icon:"📋", label:"導入アドバイス",       color:C.accentDark,
        prompt:`この農場の現状（肥育中頭数・血統構成・損益・牛舎・飼養コスト）をもとに、次回の素牛導入に向けて以下をアドバイスしてください。
1. 現在の牛舎稼働状況から見た適正な導入頭数と時期
2. 過去の血統実績（DG・BMS・損益）から推奨する父牛・血統
3. 導入価格の目安と予算配分の提案
4. 現在の飼養コスト構造を踏まえた収益最大化のための導入戦略
5. リスク分散のための血統・月齢・性別の組み合わせ提案`},
    ];

    const ask = async (prompt) => {
      setLoading(true);
      setResult(null);
      try {
        const summary = farmSummary();
        const systemPrompt = `あなたは和牛肥育農家のAIアドバイザーです。以下の農場データをもとに、具体的で実践的なアドバイスを日本語で提供してください。
農場データ:
${JSON.stringify(summary, null, 2)}

回答はMarkdownを使わず、見出し・箇条書きは日本語の記号（■ ・ →）で読みやすく整形してください。`;

        const res = await fetch("https://api.anthropic.com/v1/messages",{
          method:"POST",
          headers:{"Content-Type":"application/json","x-api-key":window.ANTHROPIC_KEY||"","anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
          body:JSON.stringify({
            model:"claude-sonnet-4-20250514",
            max_tokens:1000,
            system: systemPrompt,
            messages:[{role:"user",content:prompt}]
          })
        });
        const data = await res.json();
        const text = data.content?.map(c=>c.text||"").join("")||"解析できませんでした。";
        setResult(text);
      } catch(e) {
        setResult("通信エラーが発生しました。再度お試しください。");
      }
      setLoading(false);
    };

    return (
      <div style={{paddingBottom:100}}>
        <AppHeader subtitle="AI解析"/>

        {/* Hero */}
        <div style={{
          background:`linear-gradient(135deg,${C.accent}22,${C.purple}15)`,
          borderBottom:`1px solid ${C.border}`,
          padding:"18px 20px 16px",
        }}>
          <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:10}}>
            <div style={{
              width:44,height:44,borderRadius:14,
              background:`linear-gradient(135deg,${C.accent},${C.accentDark})`,
              display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,
              boxShadow:`0 4px 14px ${C.accent}44`,
            }}>🤖</div>
            <div>
              <div style={{color:C.text,fontWeight:900,fontSize:18}}>WAGYU AI 解析</div>
              <div style={{color:C.textMid,fontSize:12}}>農場データをAIが分析・提案します</div>
            </div>
          </div>

          {/* 料金表示バナー */}
          <div style={{
            background:"#fff",
            border:`1.5px solid ${C.accent}66`,
            borderRadius:14, padding:"12px 16px",
            display:"flex", alignItems:"center", justifyContent:"space-between",
            boxShadow:`0 2px 10px ${C.accent}18`,
          }}>
            <div>
              <div style={{color:C.textDim,fontSize:10,fontWeight:600,letterSpacing:1,marginBottom:2}}>AI解析オプション</div>
              <div style={{display:"flex",alignItems:"baseline",gap:4}}>
                <span style={{color:C.accent,fontWeight:900,fontSize:22}}>¥3,980</span>
                <span style={{color:C.textDim,fontSize:12}}>/月から</span>
              </div>
              <div style={{color:C.textDim,fontSize:10,marginTop:2}}>14日間無料トライアル</div>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{
                background:`linear-gradient(135deg,${C.accent},${C.accentDark})`,
                color:"#fff", borderRadius:20, padding:"8px 18px",
                fontSize:13, fontWeight:800, cursor:"pointer",
                boxShadow:`0 3px 10px ${C.accent}44`,
              }}>
                申し込む →
              </div>
              <div style={{color:C.textDim,fontSize:10,marginTop:4}}>準備中</div>
            </div>
          </div>
        </div>

        <div style={{padding:"16px 16px"}}>

          {/* プリセットボタン */}
          <SectionLabel>解析メニュー</SectionLabel>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:20}}>
            {PRESETS.map(p=>(
              <div key={p.id} onClick={()=>{ setActivePreset(p.id); ask(p.prompt); setQuestion(""); }}
                style={{
                  background: activePreset===p.id ? p.color+"22" : "#fff",
                  border:`1.5px solid ${activePreset===p.id ? p.color : C.border}`,
                  borderRadius:16, padding:"14px 14px", cursor:"pointer",
                  boxShadow: activePreset===p.id ? `0 4px 16px ${p.color}28` : C.shadow,
                  transition:"all 0.15s",
                }}>
                <div style={{fontSize:26,marginBottom:6}}>{p.icon}</div>
                <div style={{color:p.color,fontWeight:800,fontSize:13,lineHeight:1.3}}>{p.label}</div>
              </div>
            ))}
          </div>

          {/* 自由質問 */}
          <SectionLabel>自由に質問する</SectionLabel>
          <div style={{display:"flex",gap:8,marginBottom:20}}>
            <input
              value={question}
              onChange={e=>setQuestion(e.target.value)}
              onKeyDown={e=>{if(e.key==="Enter"&&question.trim()){setActivePreset(null);ask(question);}}}
              placeholder="例：武蔵の出荷はいつが最適？"
              style={{...inp,flex:1,borderRadius:14}}
            />
            <Btn sm onClick={()=>{if(question.trim()){setActivePreset(null);ask(question);}}} disabled={!question.trim()||loading}>
              送信
            </Btn>
          </div>

          {/* 結果 */}
          {loading&&(
            <Card>
              <div style={{textAlign:"center",padding:"28px 0"}}>
                <div style={{
                  width:52,height:52,borderRadius:"50%",
                  background:`linear-gradient(135deg,${C.accent},${C.accentDark})`,
                  display:"flex",alignItems:"center",justifyContent:"center",
                  fontSize:24, margin:"0 auto 14px",
                  animation:"aiPulse 1.4s ease-in-out infinite",
                }}>🤖</div>
                <div style={{color:C.accentDark,fontWeight:700,fontSize:15,marginBottom:4}}>解析中...</div>
                <div style={{color:C.textMid,fontSize:12}}>農場データをAIが分析しています</div>
                <style>{`@keyframes aiPulse{0%,100%{transform:scale(1);box-shadow:0 4px 14px ${C.accent}44}50%{transform:scale(1.1);box-shadow:0 8px 28px ${C.accent}66}}`}</style>
              </div>
            </Card>
          )}

          {result&&!loading&&(
            <Card style={{position:"relative"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                <div style={{display:"flex",alignItems:"center",gap:6}}>
                  <div style={{width:28,height:28,borderRadius:8,background:`linear-gradient(135deg,${C.accent},${C.accentDark})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>🤖</div>
                  <span style={{color:C.accentDark,fontWeight:800,fontSize:13}}>AI解析結果</span>
                </div>
                <button onClick={()=>{setResult(null);setActivePreset(null);}} style={{background:C.cardSub,border:"none",color:C.textDim,borderRadius:8,padding:"4px 10px",fontSize:12,cursor:"pointer"}}>✕ 閉じる</button>
              </div>
              <div style={{
                color:C.text,fontSize:13,lineHeight:1.85,
                whiteSpace:"pre-wrap",wordBreak:"break-word",
                borderTop:`1px solid ${C.borderLight}`,paddingTop:12,
              }}>{result}</div>
            </Card>
          )}

          {!loading&&!result&&(
            <div style={{
              background:C.cardSub,borderRadius:16,padding:"24px 20px",
              textAlign:"center",border:`1px dashed ${C.border}`,
            }}>
              <div style={{fontSize:40,marginBottom:10}}>🤖</div>
              <div style={{color:C.textMid,fontSize:13,lineHeight:1.7}}>
                上のメニューを選ぶか、<br/>自由に質問してみてください
              </div>
            </div>
          )}
        </div>
      </div>
    );
  };

  // ── NAV ITEMS ──────────────────────────────────────────────────────────────
  const navItems=[
    {id:"home",     icon:"🏠", label:"ホーム"},
    {id:"schedule", icon:"📅", label:"出荷予定"},
    {id:"ai",       icon:"🤖", label:"AI解析", isCenter:true},
    {id:"breeder",  icon:"🏡", label:"繁殖農家"},
    {id:"genetics", icon:"🧬", label:"血統分析"},
  ];

  // ページ別早期return
  if(page==="shipResult") return <ShipResultScreen/>;
  if(page==="intake")     return <IntakeScreen/>;
  if(page==="add") return (
    <div key={addFormKey} style={{background:C.bg,minHeight:"100vh",color:C.text,fontFamily:"'Hiragino Kaku Gothic Pro','Noto Sans JP','YuGothic',sans-serif",maxWidth:520,margin:"0 auto"}}>
      <AddScreen/>
      {showOcr&&<OcrModal onClose={()=>setShowOcr(false)} onApply={applyOcr}/>}
    </div>
  );

  return (
    <div style={{background:C.bg,minHeight:"100vh",color:C.text,fontFamily:"'Hiragino Kaku Gothic Pro','Noto Sans JP','YuGothic',sans-serif",maxWidth:520,margin:"0 auto",position:"relative"}}>
      {page==="home"     && <HomeScreen/>}
      {page==="alerts"   && <AlertsScreen/>}
      {page==="schedule" && <ScheduleScreen/>}
      {page==="ai"       && <AiScreen/>}
      {page==="breeder"  && <BreederScreen/>}
      {page==="genetics" && <GeneticsScreen/>}
      {page==="detail"   && <DetailScreen/>}

      {/* 設定モーダル（全ページ共通） */}
      {showSettings && <SettingsModal/>}

      {/* ── Bottom Nav ── */}
      <div style={{
        position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",
        width:"100%",maxWidth:520,background:"#fff",
        borderTop:`1px solid ${C.border}`,
        display:"flex",justifyContent:"space-around",alignItems:"flex-end",
        padding:"6px 0 18px",zIndex:100,
        boxShadow:"0 -4px 20px rgba(74,184,232,0.10)",
      }}>
        {navItems.map(({id,icon,label,disabled,badge,isCenter})=>{
          if(isCenter) return (
            <button key={id} onClick={()=>setPage(id)} style={{
              background:"none",border:"none",cursor:"pointer",
              display:"flex",flexDirection:"column",alignItems:"center",gap:2,
              flex:1, padding:"0 0 2px", position:"relative", marginTop:-20,
            }}>
              <div style={{
                width:58, height:58, borderRadius:20,
                background: page===id ? "#1a2e22" : "#1e3a2a",
                display:"flex",alignItems:"center",justifyContent:"center",
                fontSize:28,
                boxShadow: page===id
                  ? `0 6px 22px rgba(30,58,42,0.55), 0 0 0 3px ${C.accent}66`
                  : `0 4px 18px rgba(30,58,42,0.45)`,
                border:"3px solid #fff",
                transition:"all 0.2s",
              }}>🤖</div>
              <span style={{fontSize:9,color:page===id?C.accentDark:C.textDim,fontWeight:page===id?900:600,marginTop:2}}>AI解析</span>
            </button>
          );
          return (
            <button key={id} onClick={()=>!disabled&&setPage(id)} style={{
              background:"none",border:"none",cursor:disabled?"default":"pointer",
              display:"flex",flexDirection:"column",alignItems:"center",gap:3,
              opacity:disabled?0.3:1,flex:1,padding:"4px 0",position:"relative",
            }}>
              <div style={{position:"relative",width:38,height:38}}>
                <div style={{
                  width:38,height:38,borderRadius:13,
                  background:page===id?`linear-gradient(135deg,${C.accent},${C.accentDark})`:C.cardSub,
                  display:"flex",alignItems:"center",justifyContent:"center",
                  fontSize:19,boxShadow:page===id?`0 3px 10px ${C.accent}44`:"none",
                  transition:"all 0.2s",
                }}>{icon}</div>
                {badge>0&&(
                  <div style={{position:"absolute",top:-4,right:-4,background:C.red,color:"#fff",borderRadius:10,minWidth:18,height:18,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:900,padding:"0 4px",border:"2px solid #fff"}}>
                    {badge}
                  </div>
                )}
              </div>
              <span style={{fontSize:9,color:page===id?C.accentDark:C.textDim,fontWeight:page===id?800:500}}>{label}</span>
            </button>
          );
        })}
      </div>
    </div>
  );
}
