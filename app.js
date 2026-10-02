const $=s=>document.querySelector(s), app=$("#app");
const EMP=[{id:"EMP001",name:"Leonardo Espino",short:"Leo",role:"Auditor"},{id:"EMP002",name:"Darlin Tapia",short:"Darlin",role:"Auditor"},{id:"EMP003",name:"Ridy Gonzalez",short:"Ridy",role:"Auditor"},{id:"EMP004",name:"Pedro Amador",short:"Pedro",role:"Supervisor de Operadores"}];
let state={view:"home",opId:null,pendingEmployee:null};

const DB={db:null, async init(){this.db=await new Promise((ok,no)=>{let r=indexedDB.open("slc_ops",1);r.onupgradeneeded=()=>{let d=r.result;if(!d.objectStoreNames.contains("operations"))d.createObjectStore("operations",{keyPath:"id"});if(!d.objectStoreNames.contains("snapshots")){let s=d.createObjectStore("snapshots",{keyPath:"id"});s.createIndex("operationId","operationId")}if(!d.objectStoreNames.contains("settings"))d.createObjectStore("settings",{keyPath:"key"})};r.onsuccess=()=>ok(r.result);r.onerror=()=>no(r.error)});},
 store(n,m="readonly"){return this.db.transaction(n,m).objectStore(n)},
 all(n){return new Promise((ok,no)=>{let r=this.store(n).getAll();r.onsuccess=()=>ok(r.result);r.onerror=()=>no(r.error)})},
 get(n,k){return new Promise((ok,no)=>{let r=this.store(n).get(k);r.onsuccess=()=>ok(r.result);r.onerror=()=>no(r.error)})},
 put(n,v){return new Promise((ok,no)=>{let r=this.store(n,"readwrite").put(v);r.onsuccess=()=>ok(v);r.onerror=()=>no(r.error)})},
 del(n,k){return new Promise((ok,no)=>{let r=this.store(n,"readwrite").delete(k);r.onsuccess=()=>ok();r.onerror=()=>no(r.error)})}
};
const uid=p=>p+"_"+Date.now()+"_"+Math.random().toString(36).slice(2,7);
const esc=s=>String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));
const num=v=>{let n=Number(v);return Number.isFinite(n)?n:0};
const fmt=v=>v instanceof Date?v.toLocaleString("es-DO"):String(v??"—");
const dt=v=>{if(v instanceof Date)return v.toISOString(); if(typeof v==="number"&&v>20000){let d=XLSX.SSF.parse_date_code(v);return d?new Date(d.y,d.m-1,d.d,d.H,d.M,d.S).toISOString():v} return v};

function rowsFromSheet(ws){return XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});}
function norm(v){return String(v??"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s+/g," ").trim()}
function findLabel(rows, labels){
 for(let r=0;r<rows.length;r++)for(let c=0;c<rows[r].length;c++){let x=norm(rows[r][c]);if(labels.some(l=>x.includes(norm(l))))return {r,c,value:rows[r][c]}}
 return null
}
function rightValue(rows,p,max=4){if(!p)return null;for(let c=p.c+1;c<=Math.min(p.c+max,(rows[p.r]||[]).length-1);c++){let v=rows[p.r][c];if(v!==null&&v!=="")return v}return null}
function parseVesselVoyage(raw){
 let s=String(raw||"").trim(), voyage="";
 let m=s.match(/\b(?:V(?:OY(?:AGE)?)?\.?\s*)?(\d{1,4}[-\/][A-Z0-9]{1,5})\b/i);
 if(m){voyage=m[1];s=s.replace(m[0],"").replace(/[-–]\s*$/,"").trim()}
 return {vessel:s||String(raw||"Unknown Vessel"),voyage}
}
function parseHourly(rows,sheetName){
 let vesselRaw=rightValue(rows,findLabel(rows,["Vessel Name"]),5);
 let vv=parseVesselVoyage(vesselRaw);
 let report=rightValue(rows,findLabel(rows,["REPORT TIME"]),3);
 let fields={};
 [["eosp","EOSP"],["pilot","Pilot on board"],["cleared","Ship cleared"],["commenced","Commenced Ops"],["etc","ETC"],["etd","ETD"]].forEach(([k,l])=>fields[k]=dt(rightValue(rows,findLabel(rows,[l]),3)));
 let pm=findLabel(rows,["Programed Moves","Programmed Moves"]);
 let movementNames=["Import","Export FULL","Export EMPTY","Restow Discharge","Restow Load","Hatch Cover/ G. Boxes"];
 let programmed={};
 if(pm){let headerR=Math.max(0,pm.r-2);movementNames.forEach((n,i)=>programmed[n]=rows[pm.r]?.[pm.c+1+i]??null)}
 let gangHead=findLabel(rows,["# Gangs assigned"]);
 let gangs=[];
 if(gangHead){for(let r=gangHead.r+1;r<Math.min(rows.length,gangHead.r+8);r++){let name=rows[r]?.[gangHead.c];let crane=rows[r]?.[gangHead.c+1];if(!name&&!crane)continue;if(norm(crane).includes("total moves"))break;if(norm(name).startsWith("gang")||norm(crane).includes("gantry")||norm(crane).includes("crane"))gangs.push({name:name||`Gang ${gangs.length+1}`,crane,initial:rows[r]?.[gangHead.c+2],current:rows[r]?.[gangHead.c+3],completed:rows[r]?.[gangHead.c+4],hatch:rows[r]?.[gangHead.c+5]})}}
 let prod=findLabel(rows,["Productivity (hours)"]);
 let hourly=[];
 if(prod){for(let r=prod.r+1;r<rows.length;r++){let period=rows[r]?.[prod.c];if(norm(period).includes("moves completed")||norm(period).includes("remaining moves"))break;if(!period)continue;let vals=[];for(let c=prod.c+1;c<=prod.c+7;c++)vals.push(rows[r]?.[c]);let total=rows[r]?.[prod.c+7];let productivity=rows[r]?.[prod.c+8];let remark=rows[r]?.[prod.c+9];hourly.push({period,import:vals[0],exportFull:vals[1],exportEmpty:vals[2],restowDischarge:vals[3],restowLoad:vals[4],hatch:vals[5],total,productivity,remark})}}
 let mc=findLabel(rows,["Moves Completed"]), rm=findLabel(rows,["Remaining Moves"]);
 let completed={},remaining={};
 if(mc && mc.r>25){movementNames.forEach((n,i)=>completed[n]=rows[mc.r]?.[mc.c+1+i]??null)}
 if(rm && rm.r>25){movementNames.forEach((n,i)=>remaining[n]=rows[rm.r]?.[rm.c+1+i]??null)}
 let emptyHead=findLabel(rows,["Empty: Loading Breakdown"]);
 let empties=[];
 if(emptyHead){for(let r=emptyHead.r+2;r<Math.min(rows.length,emptyHead.r+12);r++){let type=rows[r]?.[emptyHead.c],qty=rows[r]?.[emptyHead.c+1],destination=rows[r]?.[emptyHead.c+2];if(type)empties.push({type,qty,destination})}}
 let remarks=hourly.filter(x=>x.remark).map(x=>({period:x.period,text:x.remark}));
 let totalProgrammed=Object.values(programmed).reduce((a,b)=>a+num(b),0);
 let totalCompleted=Object.values(completed).reduce((a,b)=>a+num(b),0);
 if(!totalCompleted) totalCompleted=hourly.reduce((a,b)=>a+num(b.total),0);
 return {sheetName,vessel:vv.vessel,voyage:vv.voyage,reportTime:dt(report),fields,programmed,completed,remaining,gangs,hourly,empties,remarks,totalProgrammed,totalCompleted,rawRows:rows};
}
function parseWorkbook(wb){
 let candidates=wb.SheetNames.map(n=>parseHourly(rowsFromSheet(wb.Sheets[n]),n)).filter(x=>x.vessel&&x.vessel!=="Unknown Vessel");
 if(!candidates.length)throw Error("No pude identificar una hoja Hourly.");
 return candidates.sort((a,b)=>num(b.totalCompleted)-num(a.totalCompleted))[0];
}
function latest(snaps){return [...snaps].sort((a,b)=>new Date(b.importedAt)-new Date(a.importedAt))[0]}
async function home(){
 let ops=(await DB.all("operations")).filter(o=>o.status==="open");
 app.innerHTML=`<h1>SLC OPS</h1><p class="sub">Operations Control System</p>
 <button class="wide" id="newOp">＋ INICIAR NUEVA OPERACIÓN</button>
 <h2>OPERACIONES ABIERTAS</h2>${ops.length?ops.map(o=>`<div class="op"><div><strong>${esc(o.vessel)}</strong><small>Voyage ${esc(o.voyage||"—")} · ${esc(o.currentEmployeeName||"Sin auditor")}</small></div><button onclick="openOp('${o.id}')">CONTINUAR</button></div>`).join(""):`<div class="card sub">No hay operaciones abiertas.</div>`}
 <h2>MÓDULOS</h2><div class="grid g2"><div class="card"><b>01 · Hourly Reports</b><p class="sub">Activo</p></div><div class="card"><b>02 · Employees</b><p class="sub">Reservado</p></div><div class="card"><b>03 · Estimado de Operaciones</b><p class="sub">Reservado</p></div></div>`;
 $("#newOp").onclick=()=>chooseEmployee("new");
}
function chooseEmployee(mode){
 app.innerHTML=`<h1>${mode==="new"?"Nueva operación":"Importar Hourly"}</h1><p class="sub">Selecciona el empleado que está vigilando las operaciones en este momento.</p>
 <div class="grid">${EMP.slice(0,3).map(e=>`<button class="wide secondary" onclick="employeeChosen('${e.id}','${mode}')">${e.name}<small style="display:block;color:#8ea1b1">${e.role}</small></button>`).join("")}</div>
 <h2>Supervisor</h2><button class="wide secondary" onclick="employeeChosen('EMP004','${mode}')">Pedro Amador<small style="display:block;color:#8ea1b1">Supervisor de Operadores</small></button>`;
}
window.employeeChosen=(id,mode)=>{state.pendingEmployee=id;state.importMode=mode;$("#excelInput").click()};
async function importExcel(file){
 if(!file)return; if(typeof XLSX==="undefined"){alert("La librería Excel aún no está disponible. Conecta a internet una vez y recarga SLC OPS.");return}
 let wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true}); let p=parseWorkbook(wb);
 let emp=EMP.find(e=>e.id===state.pendingEmployee);
 if(state.importMode==="new"){
   let op={id:uid("op"),vessel:p.vessel,voyage:p.voyage,status:"open",createdAt:new Date().toISOString(),currentEmployeeId:emp.id,currentEmployeeName:emp.name};
   await DB.put("operations",op);state.opId=op.id;
 }else{
   let op=await DB.get("operations",state.opId);
   if(op && ((p.vessel&&norm(op.vessel)!==norm(p.vessel)) || (p.voyage&&op.voyage&&norm(op.voyage)!==norm(p.voyage)))){if(!confirm(`El Excel indica ${p.vessel} / ${p.voyage}, diferente a ${op.vessel} / ${op.voyage}. ¿Importar de todos modos?`))return}
   op.currentEmployeeId=emp.id;op.currentEmployeeName=emp.name;await DB.put("operations",op);
 }
 let snap={id:uid("hr"),operationId:state.opId,employeeId:emp.id,employeeName:emp.name,importedAt:new Date().toISOString(),sourceFile:file.name,data:p,edits:{}};
 await DB.put("snapshots",snap); await dashboard();
}
window.openOp=async id=>{state.opId=id;state.view="hourly";await dashboard()};
function val(s,path,orig){return Object.prototype.hasOwnProperty.call(s.edits||{},path)?s.edits[path]:orig}
function editCell(sid,path,current){
 let nv=prompt("Editar valor",current??""); if(nv===null)return;
 DB.get("snapshots",sid).then(async s=>{s.edits=s.edits||{};s.edits[path]=nv;await DB.put("snapshots",s);dashboard()})
}
window.editCell=editCell;
async function dashboard(){
 let op=await DB.get("operations",state.opId);if(!op)return home();
 let snaps=(await DB.all("snapshots")).filter(s=>s.operationId===op.id);let s=latest(snaps),d=s?.data||{};
 let tp=num(d.totalProgrammed),tc=num(d.totalCompleted),rem=Math.max(0,tp-tc),pct=tp?Math.round(tc/tp*100):0;
 let lastH=d.hourly?.length?d.hourly[d.hourly.length-1]:null;
 let currentProd=lastH?num(lastH.total):0;
 let etc=d.fields?.etc, hours=etc?(new Date(etc)-new Date())/36e5:null, req=hours&&hours>0?rem/hours:null;
 let projected=currentProd>0?new Date(Date.now()+rem/currentProd*36e5):null;
 app.innerHTML=`<div class="row between"><div><h1>${esc(op.vessel)}</h1><p class="sub">Voyage ${esc(op.voyage||"—")} · ${snaps.length} Hourly${snaps.length===1?"":"s"} · ${esc(s?.employeeName||"—")}</p></div><span class="tag">OPERACIÓN ACTUAL</span></div>
 <div class="row"><button id="addHourly">＋ IMPORTAR NUEVO HOURLY</button><button class="ghost" id="historyBtn">Historial</button><button class="danger" id="finishBtn">FINALIZAR OPS</button></div>
 <h2>ESTADO ACTUAL</h2><div class="grid g4">
 ${kpi("Avance",pct+"%",`${tc} / ${tp||"—"} movimientos`)}
 ${kpi("Restantes",rem||"—","movimientos")}
 ${kpi("Última hora",currentProd||"—",lastH?lastH.period:"Sin dato")}
 ${kpi("Prod. requerida",req?req.toFixed(1):"—","mov/h para ETC")}
 </div>
 <h2>ETC</h2><div class="grid g2">${kpi("ETC Hourly",etc?new Date(etc).toLocaleString("es-DO"):"—","Dato fuente Excel")}${kpi("ETC proyectado SLC",projected?projected.toLocaleString("es-DO"):"—",currentProd?`A ${currentProd} mov/h`:"Sin productividad")}</div>
 <h2>MOVIMIENTOS</h2><div class="scroll"><table><thead><tr><th>Tipo</th><th>Plan</th><th>Completado</th><th>Restante</th><th>%</th></tr></thead><tbody>
 ${Object.keys(d.programmed||{}).map(n=>{let p=num(d.programmed[n]),c=num(d.completed?.[n]),r=d.remaining?.[n]!=null?num(d.remaining[n]):Math.max(0,p-c);return `<tr><td>${esc(n)}</td><td>${p}</td><td>${c}</td><td>${r}</td><td>${p?Math.round(c/p*100):0}%</td></tr>`}).join("")}</tbody></table></div>
 <h2>GANGS / GRÚAS</h2><div class="scroll"><table><thead><tr><th>Gang</th><th>Grúa</th><th>Initial Split</th><th>Current Split</th><th>Completed</th></tr></thead><tbody>
 ${(d.gangs||[]).map((g,i)=>`<tr><td>${esc(g.name)}</td><td>${esc(g.crane)}</td><td>${fmt(g.initial)}</td><td>${fmt(g.current)}</td><td>${fmt(g.completed)}</td></tr>`).join("")}</tbody></table></div>
 <h2>PRODUCTIVIDAD POR HORA</h2><div class="scroll"><table><thead><tr><th>Hora</th><th>Import</th><th>Exp Full</th><th>Exp Empty</th><th>Restow D</th><th>Restow L</th><th>Hatch</th><th>Total</th><th>Remark</th></tr></thead><tbody>
 ${(d.hourly||[]).map((h,i)=>`<tr><td>${esc(h.period)}</td><td>${fmt(h.import)}</td><td>${fmt(h.exportFull)}</td><td>${fmt(h.exportEmpty)}</td><td>${fmt(h.restowDischarge)}</td><td>${fmt(h.restowLoad)}</td><td>${fmt(h.hatch)}</td><td><b>${fmt(h.total)}</b></td><td>${esc(h.remark||"")}</td></tr>`).join("")}</tbody></table></div>
 <h2>EMPTY LOADING BREAKDOWN</h2><div class="scroll"><table><thead><tr><th>Type</th><th>Qty</th><th>Destination</th></tr></thead><tbody>${(d.empties||[]).map(x=>`<tr><td>${esc(x.type)}</td><td>${fmt(x.qty)}</td><td>${esc(x.destination||"")}</td></tr>`).join("")}</tbody></table></div>
 <h2>DATOS DEL HOURLY · EDITABLES</h2><div class="card">${editableRow(s,"vessel","Vessel",d.vessel)}${editableRow(s,"voyage","Voyage",d.voyage)}${editableRow(s,"reportTime","Report Time",d.reportTime)}${Object.entries(d.fields||{}).map(([k,v])=>editableRow(s,"fields."+k,k.toUpperCase(),v)).join("")}<p class="sub">Al editar, SLC OPS conserva internamente el valor original importado.</p></div>`;
 $("#addHourly").onclick=()=>chooseEmployee("continue");$("#finishBtn").onclick=finishOp;$("#historyBtn").onclick=()=>history(op,snaps);
}
function kpi(a,b,c){return `<div class="kpi"><label>${a}</label><b>${b}</b><small>${c}</small></div>`}
function editableRow(s,path,label,orig){let cur=val(s,path,orig),edited=Object.prototype.hasOwnProperty.call(s.edits||{},path);return `<div class="op editable ${edited?"edited":""}" onclick='editCell(${JSON.stringify(s.id)},${JSON.stringify(path)},${JSON.stringify(fmt(cur))})'><div><small>${esc(label)}</small><strong>${esc(fmt(cur))}</strong></div><span>✎</span></div>`}
function history(op,snaps){app.innerHTML=`<h1>Historial Hourly</h1><p class="sub">${esc(op.vessel)} · ${esc(op.voyage||"—")}</p>${[...snaps].sort((a,b)=>new Date(b.importedAt)-new Date(a.importedAt)).map(s=>`<div class="op"><div><strong>${esc(s.data.reportTime?new Date(s.data.reportTime).toLocaleString("es-DO"):new Date(s.importedAt).toLocaleString("es-DO"))}</strong><small>${esc(s.employeeName)} · ${esc(s.sourceFile)}</small></div><span class="tag">${s.data.totalCompleted||0} moves</span></div>`).join("")}<button class="wide" onclick="dashboard()">VOLVER AL DASHBOARD</button>`}
window.dashboard=dashboard;
async function finishOp(){if(!confirm("¿Finalizar esta operación? Dejará de aparecer entre las operaciones abiertas."))return;let o=await DB.get("operations",state.opId);o.status="closed";o.closedAt=new Date().toISOString();await DB.put("operations",o);state.opId=null;home()}
async function settings(){
 let ops=await DB.all("operations"),sn=await DB.all("snapshots");
 app.innerHTML=`<h1>Configuración</h1><p class="sub">Datos maestros y respaldo de SLC OPS.</p>
 <h2>Datos maestros</h2><div class="card">Barcos · Rutas/Servicios · Itinerarios · Puertos/Destinos<div class="sub" style="margin-top:8px">La administración detallada de estos catálogos se incorporará en la siguiente iteración.</div></div>
 <h2>Base de datos</h2><div class="grid g2"><button id="exportDB">Exportar base de datos</button><button class="secondary" id="importDB">Importar base de datos</button></div><p class="sub">${ops.length} operaciones · ${sn.length} Hourly snapshots</p>`;
 $("#exportDB").onclick=backup;$("#importDB").onclick=()=>$("#jsonInput").click();
}
async function backup(){let data={format:"SLC_OPS_DATABASE",version:1,exportedAt:new Date().toISOString(),employees:EMP,operations:await DB.all("operations"),snapshots:await DB.all("snapshots"),settings:await DB.all("settings")};let a=document.createElement("a");a.href=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:"application/json"}));a.download=`SLC_OPS_DATABASE_${new Date().toISOString().slice(0,10)}.json`;a.click();URL.revokeObjectURL(a.href)}
async function restore(file){let d=JSON.parse(await file.text());if(d.format!=="SLC_OPS_DATABASE")return alert("Archivo no reconocido.");if(!confirm("Esto agregará/actualizará los datos del backup en SLC OPS. ¿Continuar?"))return;for(let x of d.operations||[])await DB.put("operations",x);for(let x of d.snapshots||[])await DB.put("snapshots",x);for(let x of d.settings||[])await DB.put("settings",x);alert("Base de datos importada.");home()}
function placeholder(title,n){app.innerHTML=`<h1>${title}</h1><div class="card"><b>Módulo ${n}</b><p class="sub">Reservado para desarrollo posterior. La arquitectura de SLC OPS ya mantiene este módulo separado de Hourly Reports.</p></div>`}
document.querySelectorAll("[data-nav]").forEach(b=>b.onclick=()=>{document.querySelectorAll("[data-nav]").forEach(x=>x.classList.remove("active"));b.classList.add("active");let v=b.dataset.nav;if(v==="home")home();else if(v==="hourly"){if(state.opId)dashboard();else home()}else if(v==="employees")placeholder("Employees","02");else placeholder("Estimado de Operaciones","03")});
$("#settingsBtn").onclick=settings;$("#excelInput").onchange=e=>{importExcel(e.target.files[0]);e.target.value=""};$("#jsonInput").onchange=e=>{restore(e.target.files[0]);e.target.value=""};
(async()=>{await DB.init();if("serviceWorker"in navigator)navigator.serviceWorker.register("./sw.js").catch(()=>{});home()})();