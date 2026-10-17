/* ===== Lógica pura (fechas, cálculos, importación Zepp, Excel) ===== */
const DIAS = ["domingo","lunes","martes","miércoles","jueves","viernes","sábado"];
const MESES = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"];

function dayNum(iso){ if(!iso) return null; const [y,m,d]=iso.split("-").map(Number); return Date.UTC(y,m-1,d)/864e5; }
function isoFromDayNum(n){ const d=new Date(n*864e5); return d.toISOString().slice(0,10); }
function todayIso(){ const t=new Date(); return t.getFullYear()+"-"+String(t.getMonth()+1).padStart(2,"0")+"-"+String(t.getDate()).padStart(2,"0"); }
function fmtFecha(iso,{year=false,weekday=false}={}){
  if(!iso) return "Sin fecha";
  const [y,m,d]=iso.split("-").map(Number);
  let s=d+" "+MESES[m-1]+(year?" "+y:"");
  if(weekday) s=DIAS[new Date(Date.UTC(y,m-1,d)).getUTCDay()]+" "+s;
  return s;
}
function fmtNum(v,dec=1){ if(v==null||isNaN(v)) return "–"; return Number(v).toLocaleString("es-ES",{minimumFractionDigits:dec,maximumFractionDigits:dec}); }
function diaSemana(iso){ return iso? DIAS[new Date(dayNum(iso)*864e5).getUTCDay()] : ""; }

function sortKey(r){ return r.fecha || r.fechaAprox || "9999"; }
function ordenar(recs){ return [...recs].sort((a,b)=> sortKey(a)<sortKey(b)?-1: sortKey(a)>sortKey(b)?1: (a.orden||0)-(b.orden||0)); }

function tipoDe(r){
  const iny = r.dosis!=null && r.dosis!=="";
  const rev = r.revision;
  const peso = r.peso!=null && r.peso!=="";
  if(rev && iny) return "Revisión + inyección";
  if(rev) return r.lugar && /llamada|tel/i.test(r.lugar) ? "Revisión telefónica" : "Revisión";
  if(iny && peso) return "Inyección + peso";
  if(iny) return "Inyección";
  if(peso) return "Peso";
  return "Nota";
}
const esInyeccion = r => r.dosis!=null && r.dosis!=="";
const tienePeso = r => r.peso!=null && r.peso!=="" && !isNaN(r.peso);

function calcular(recsIn,cfg){
  const recs=ordenar(recsIn);
  const iny=recs.filter(r=>esInyeccion(r)&&r.fecha);
  const inicio=iny.length? iny[0].fecha : (recs.find(r=>r.fecha)||{}).fecha || null;
  const semana = iso => (iso&&inicio)? Math.floor((dayNum(iso)-dayNum(inicio))/7)+1 : null;
  const pesos=recs.filter(r=>r.fecha&&tienePeso(r));
  const talla=Number(cfg&&cfg.talla)||null;
  const imc = p => talla? p/((talla/100)**2) : null;
  const ini=pesos[0], act=pesos[pesos.length-1];
  const ultIny=iny[iny.length-1];
  const s={recs,inicio,semana,pesos,talla,imc,ini,act,ultIny,nIny:iny.length};
  if(ini&&act){
    s.perdida=ini.peso-act.peso; s.perdidaPct=s.perdida/ini.peso;
    const sem=(dayNum(act.fecha)-dayNum(ini.fecha))/7; s.ritmo= sem>0? s.perdida/sem : null;
    s.minimo=Math.min(...pesos.map(p=>p.peso));
    const ref=[...pesos].reverse().find(p=>dayNum(p.fecha)<=dayNum(act.fecha)-28);
    s.var4=ref? act.peso-ref.peso : null; s.ref4=ref;
  }
  const cint=recs.filter(r=>r.fecha&&r.cintura);
  s.cinIni=cint[0]; s.cinAct=cint[cint.length-1];
  s.dosisActual= ultIny? ultIny.dosis : null; s.clicsActual= ultIny? ultIny.clics : null;
  s.proxima= ultIny? isoFromDayNum(dayNum(ultIny.fecha)+7) : null;
  s.ultimo = recs.filter(r=>r.fecha).map(r=>r.fecha).pop() || null;
  return s;
}

/* Interpreta lo que copia el atajo (o cualquier texto pegado): "98,7 kg", "ZEPP;98.7 kg;2026-10-17T08:12:00+02:00", "17/10/2026 98,7" */
function parseZepp(txt){
  if(!txt) return null;
  let t=String(txt), fecha=null, hora=null;
  let m=t.match(/(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{1,2}):(\d{2}))?/);
  if(m){ fecha=`${m[1]}-${m[2]}-${m[3]}`; if(m[4]) hora=m[4].padStart(2,"0")+":"+m[5]; t=t.replace(m[0]," "); }
  else { m=t.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:,?\s+(\d{1,2}):(\d{2}))?/);
    if(m){ let y=m[3].length===2?"20"+m[3]:m[3]; fecha=`${y}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`; if(m[4]) hora=m[4].padStart(2,"0")+":"+m[5]; t=t.replace(m[0]," "); } }
  t=t.replace(/[+-]\d{2}:\d{2}/g," ");
  const nums=(t.match(/\d+(?:[.,]\d+)?/g)||[]).map(x=>parseFloat(x.replace(",","."))).filter(x=>x>=30&&x<=300);
  if(!nums.length) return null;
  let peso=nums[0];
  if(/\blb\b|libras/i.test(t)) peso=peso*0.45359237;
  return {peso:Math.round(peso*100)/100, fecha, hora};
}

/* ===== Excel ===== */
function excelDate(iso){ const [y,m,d]=iso.split("-").map(Number); return new Date(Date.UTC(y,m-1,d)); }

function construirExcel(ExcelJS, recsIn, cfg, chartPng){
  const s=calcular(recsIn,cfg); const recs=s.recs;
  const wb=new ExcelJS.Workbook(); wb.creator="Diario Wegovy"; wb.calcProperties.fullCalcOnLoad=true;
  const F="Arial", AZUL="FF0000FF", HFILL={type:"pattern",pattern:"solid",fgColor:{argb:"FF1F4E78"}};
  const REV={type:"pattern",pattern:"solid",fgColor:{argb:"FFFCE4D6"}}, AMA={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFFF00"}};
  const thin={style:"thin",color:{argb:"FFBFBFBF"}}, B={top:thin,left:thin,bottom:thin,right:thin};
  const nombre=(cfg&&cfg.nombre)||"Paciente";
  const head=(ws,row,labels)=>{ labels.forEach((h,i)=>{const c=ws.getCell(row,i+1); c.value=h; c.fill=HFILL; c.font={name:F,bold:true,color:{argb:"FFFFFFFF"},size:10}; c.alignment={horizontal:"center",vertical:"middle",wrapText:true}; c.border=B;}); ws.getRow(row).height=32; };

  const wsR=wb.addWorksheet("Resumen"), wsG=wb.addWorksheet("Registro"), wsP=wb.addWorksheet("Evolución peso");

  // --- Registro
  wsG.getCell("A1").value="Registro de seguimiento – Tratamiento con Wegovy (semaglutida)"; wsG.getCell("A1").font={name:F,bold:true,size:14};
  wsG.getCell("A2").value=`Paciente: ${nombre}. Exportado el ${fmtFecha(todayIso(),{year:true})} desde la app Diario Wegovy.`; wsG.getCell("A2").font={name:F,size:9};
  head(wsG,4,["Nº","Fecha","Día","Semana tto.","Tipo de registro","Lugar","Dosis (mg)","Clics","Peso (kg)","Cintura (cm)","Origen del dato","Observaciones"]);
  const r0=5, filaDe=new Map();
  recs.forEach((r,i)=>{
    const row=r0+i; filaDe.set(r,row); const tipo=tipoDe(r);
    const v=[i+1, r.fecha?excelDate(r.fecha):"Sin fecha",
      {formula:`IF(ISNUMBER(B${row}),PROPER(TEXT(B${row},"dddd")),"")`, result: r.fecha? diaSemana(r.fecha).replace(/^./,c=>c.toUpperCase()):""},
      {formula:`IF(ISNUMBER(B${row}),INT((B${row}-$B$${r0})/7)+1,"")`, result: r.fecha? s.semana(r.fecha):""},
      tipo, r.lugar|| (r.revision?"Centro de Salud":"Domicilio"),
      esInyeccion(r)?Number(r.dosis):null, esInyeccion(r)&&r.clics?Number(r.clics):null,
      tienePeso(r)?Number(r.peso):null, r.cintura?Number(r.cintura):null, r.origen||"", r.obs||""];
    v.forEach((val,j)=>{ const c=wsG.getCell(row,j+1); c.value=val; c.border=B;
      c.font={name:F,size:10,color:{argb:[1,6,7,8,9].includes(j)?AZUL:"FF000000"}};
      c.alignment={vertical:"middle",horizontal:[4,5,10,11].includes(j)?"left":"center",wrapText:j===11};
      if(r.revision) c.fill=REV; });
    wsG.getCell(row,2).numFmt="dd/mm/yyyy"; wsG.getCell(row,7).numFmt="0.00"; wsG.getCell(row,9).numFmt="0.00";
    if(!r.fecha) wsG.getCell(row,2).fill=AMA;
  });
  const last=r0+recs.length-1;
  [5,12,11,9,20,22,10,8,10,11,14,64].forEach((w,i)=>wsG.getColumn(i+1).width=w);
  wsG.views=[{state:"frozen",ySplit:4}];
  if(recs.length) wsG.autoFilter=`A4:L${last}`;
  ["Texto azul = dato anotado por el paciente; negro = calculado con fórmula.","Filas en naranja claro = revisiones sanitarias. Celda amarilla = fecha pendiente de confirmar.","Semana tto. = semana contada desde la 1ª dosis."]
    .forEach((t,i)=>{const c=wsG.getCell(last+2+i,1); c.value=t; c.font={name:F,size:9};});

  // --- Evolución peso
  wsP.getCell("A1").value="Evolución del peso y del perímetro abdominal"; wsP.getCell("A1").font={name:F,bold:true,size:14};
  wsP.getCell("A2").value=`Paciente: ${nombre}. Solo registros con fecha y peso.`; wsP.getCell("A2").font={name:F,size:9};
  head(wsP,4,["Fecha","Semana tto.","Peso (kg)","Cintura (cm)","Días desde anterior","Variación vs anterior (kg)","Pérdida acumulada (kg)","% pérdida acumulada","IMC"]);
  const pes=s.pesos; const p0=pes.length? pes[0].peso:0;
  pes.forEach((p,i)=>{
    const row=5+i, src=filaDe.get(p), prev=pes[i-1];
    const vals=[{formula:`Registro!B${src}`,result:excelDate(p.fecha)},{formula:`Registro!D${src}`,result:s.semana(p.fecha)},{formula:`Registro!I${src}`,result:p.peso},
      {formula:`IF(Registro!J${src}="","",Registro!J${src})`,result:p.cintura||""},
      i? {formula:`A${row}-A${row-1}`,result:dayNum(p.fecha)-dayNum(prev.fecha)}:"",
      i? {formula:`C${row}-C${row-1}`,result:p.peso-prev.peso}:"",
      {formula:`$C$5-C${row}`,result:p0-p.peso},{formula:`G${row}/$C$5`,result:(p0-p.peso)/p0},
      {formula:`IF(Resumen!$C$6="","",C${row}/(Resumen!$C$6/100)^2)`,result:s.talla? s.imc(p.peso):""}];
    vals.forEach((val,j)=>{const c=wsP.getCell(row,j+1); c.value=val; c.border=B; c.font={name:F,size:10,color:{argb:j<4?"FF008000":"FF000000"}}; c.alignment={horizontal:"center"};});
    wsP.getCell(row,1).numFmt="dd/mm/yyyy"; wsP.getCell(row,3).numFmt="0.00"; wsP.getCell(row,6).numFmt="+0.00;-0.00;0.00";
    wsP.getCell(row,7).numFmt="0.00"; wsP.getCell(row,8).numFmt="0.0%"; wsP.getCell(row,9).numFmt="0.0";
  });
  [12,9,10,11,11,13,13,12,8].forEach((w,i)=>wsP.getColumn(i+1).width=w);
  const wl=4+pes.length;
  wsP.getCell(wl+2,1).value="Verde = enlazado desde 'Registro'. Negro = fórmula. Gráfica generada por la app el día de la exportación."; wsP.getCell(wl+2,1).font={name:F,size:9};
  if(chartPng){ const id=wb.addImage({base64:chartPng,extension:"png"}); wsP.addImage(id,{tl:{col:10,row:3},ext:{width:760,height:400}}); }

  // --- Resumen
  wsR.getColumn(1).width=2; wsR.getColumn(2).width=38; wsR.getColumn(3).width=16; wsR.getColumn(4).width=50;
  wsR.getCell("B2").value="Resumen del tratamiento – para el enfermero/a"; wsR.getCell("B2").font={name:F,bold:true,size:14};
  wsR.getCell("B3").value=`Paciente: ${nombre}   ·   Wegovy (semaglutida) – inyección subcutánea semanal`; wsR.getCell("B3").font={name:F,bold:true,size:11};
  const P="'Evolución peso'", R="Registro", rng=c=>`${R}!${c}${r0}:${c}${Math.max(last,r0)}`, prng=c=>`${P}!${c}5:${c}${Math.max(wl,5)}`;
  const items=[
    ["Fecha de exportación", excelDate(todayIso()), null, "dd/mm/yyyy",""],
    ["Talla del paciente (cm)", s.talla||null, null, "0","Dato del paciente"],
    ["Fecha de inicio del tratamiento", {formula:`MIN(${rng("B")})`,result:s.inicio?excelDate(s.inicio):""},null,"dd/mm/yyyy",""],
    ["Fecha del último registro", {formula:`MAX(${rng("B")})`,result:s.ultimo?excelDate(s.ultimo):""},null,"dd/mm/yyyy",""],
    ["Semanas de tratamiento", {formula:"INT((C8-C7)/7)+1",result:(s.inicio&&s.ultimo)?Math.floor((dayNum(s.ultimo)-dayNum(s.inicio))/7)+1:""},null,"0",""],
    ["Nº de inyecciones registradas", {formula:`COUNT(${rng("G")})`,result:recs.filter(esInyeccion).length},null,"0",""],
    ["Dosis actual (mg)", s.dosisActual!=null?Number(s.dosisActual):null,null,"0.00", s.ultIny?`Última inyección: ${fmtFecha(s.ultIny.fecha,{year:true})}`:""],
    ["Clics por dosis actual", s.clicsActual!=null?Number(s.clicsActual):null,null,"0",""],
    ["Peso inicial (kg)", {formula:`${P}!C5`,result:s.ini?s.ini.peso:""},null,"0.00", s.ini?fmtFecha(s.ini.fecha,{year:true}):""],
    ["Peso actual (kg)", {formula:`INDEX(${prng("C")},COUNT(${prng("C")}))`,result:s.act?s.act.peso:""},null,"0.00", s.act?fmtFecha(s.act.fecha,{year:true}):""],
    ["Peso mínimo registrado (kg)", {formula:`MIN(${prng("C")})`,result:s.minimo??""},null,"0.00",""],
    ["Pérdida total (kg)", {formula:"C13-C14",result:s.perdida??""},null,"0.00",""],
    ["Pérdida total (%)", {formula:"IF(C13=0,\"\",C16/C13)",result:s.perdidaPct??""},null,"0.0%",""],
    ["Ritmo medio de pérdida (kg/semana)", {formula:`IF(C8=C7,"",C16/((INDEX(${prng("A")},COUNT(${prng("C")}))-${P}!A5)/7))`,result:s.ritmo??""},null,"0.00",""],
    ["Variación últimas 4 semanas (kg)", s.var4??null,null,"+0.00;-0.00;0.00", s.ref4?`Frente a ${fmtFecha(s.ref4.fecha,{year:true})} (negativo = pérdida)`:""],
    ["Perímetro abdominal inicial (cm)", s.cinIni?Number(s.cinIni.cintura):null,null,"0", s.cinIni?fmtFecha(s.cinIni.fecha,{year:true}):""],
    ["Perímetro abdominal último (cm)", s.cinAct?Number(s.cinAct.cintura):null,null,"0", s.cinAct?fmtFecha(s.cinAct.fecha,{year:true}):""],
    ["Reducción perímetro abdominal (cm)", {formula:"IF(OR(C20=\"\",C21=\"\"),\"\",C20-C21)",result:(s.cinIni&&s.cinAct)?s.cinIni.cintura-s.cinAct.cintura:""},null,"0",""],
    ["IMC inicial", {formula:"IF(C6=\"\",\"\",C13/(C6/100)^2)",result:(s.talla&&s.ini)?s.imc(s.ini.peso):""},null,"0.0",""],
    ["IMC actual", {formula:"IF(C6=\"\",\"\",C14/(C6/100)^2)",result:(s.talla&&s.act)?s.imc(s.act.peso):""},null,"0.0",""],
  ];
  items.forEach(([lab,val,_,fmt,nota],i)=>{ const row=5+i;
    const a=wsR.getCell(row,2); a.value=lab; a.font={name:F,bold:true,size:10}; a.border=B;
    const c=wsR.getCell(row,3); c.value=val; c.numFmt=fmt; c.font={name:F,size:10,color:{argb:(val&&val.formula)?"FF000000":AZUL}}; c.border=B; c.alignment={horizontal:"center"};
    const n=wsR.getCell(row,4); n.value=nota; n.font={name:F,size:9,color:{argb:"FF595959"}}; });
  let ob=5+items.length+1;
  wsR.getCell(ob,2).value="Observaciones anotadas"; wsR.getCell(ob,2).font={name:F,bold:true,size:12};
  recs.filter(r=>r.obs).forEach((r,i)=>{ const c=wsR.getCell(ob+1+i,2); c.value=`• ${r.fecha?fmtFecha(r.fecha,{year:true}):"Sin fecha"}: ${r.obs}`; c.font={name:F,size:10}; });
  return wb;
}

if(typeof module!=="undefined") module.exports={calcular,parseZepp,construirExcel,ordenar,tipoDe,fmtFecha,dayNum,isoFromDayNum};
