/* =========================
   STATE
========================= */
const state = {
headers: [],
rawData: [],
matrix: [],
top: [],
chart: null
};

/* =========================
   FILE LOAD (EXCEL SAFE)
========================= */
document.getElementById("fileInput").onchange = e=>{
loadFile(e.target.files[0]);
};

function loadFile(file){

let reader=new FileReader();

reader.onload=e=>{

let wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
let ws=wb.Sheets[wb.SheetNames[0]];
let data=XLSX.utils.sheet_to_json(ws,{header:1});

/* KPI (C2 → right) */
state.headers = (data[1]||[])
.slice(2)
.map(x=>String(x).trim())
.filter(Boolean);

/* RAW (B3 → down) */
state.rawData = (data||[])
.slice(2)
.map(r=>(r||[]).slice(1));

document.getElementById("aiBox").innerText =
`📊 ${state.headers.length} KPI | ${state.rawData.length} Satır`;
};

reader.readAsArrayBuffer(file);
}

/* =========================
   EXCEL SAFE NUMBER PARSER
========================= */
function toNumber(v){

if(v === null || v === undefined) return null;

if(typeof v === "number") return v;

if(typeof v === "string"){
v = v.replace(",",".").trim();
}

let n = Number(v);

return isNaN(n) ? null : n;
}

/* =========================
   EXCEL TOOLPAK CORREL (FIXED)
========================= */
function corrExcel(a,b){

let pairs=[];

for(let i=0;i<a.length;i++){

let x = toNumber(a[i]);
let y = toNumber(b[i]);

if(x===null || y===null) continue;

pairs.push([x,y]);
}

if(pairs.length<2) return 0;

let mx = pairs.reduce((s,p)=>s+p[0],0)/pairs.length;
let my = pairs.reduce((s,p)=>s+p[1],0)/pairs.length;

let num=0,dx=0,dy=0;

for(let [x,y] of pairs){

let a=x-mx;
let b=y-my;

num+=a*b;
dx+=a*a;
dy+=b*b;
}

return dx&&dy ? num/Math.sqrt(dx*dy) : 0;
}

/* =========================
   MATRIX (NO FILTER)
========================= */
function matrix(){

const cols = state.headers.map((_,i)=>
state.rawData.map(r=>r[i])
);

state.matrix = state.headers.map((_,i)=>
state.headers.map((_,j)=>
i===j?null:corrExcel(cols[i],cols[j])
));
}

/* =========================
   TOP ANALYSIS
========================= */
function getTop(){

let arr=[];

for(let i=0;i<state.matrix.length;i++){
for(let j=0;j<state.matrix.length;j++){

if(i===j) continue;

arr.push({
name:`${state.headers[i]} ↔ ${state.headers[j]}`,
value:state.matrix[i][j]
});
}
}

state.top = arr
.sort((a,b)=>Math.abs(b.value)-Math.abs(a.value))
.slice(0,10);
}

/* =========================
   AI ENGINE
========================= */
function ai(){

if(!state.top.length) return "AI: no data";

let best = state.top[0];

let avg = state.top.reduce((a,b)=>a+Math.abs(b.value),0)/state.top.length;

let trend =
avg>0.7?"🔴 güçlü korelasyon sistemi":
avg>0.4?"🟡 orta sistem":
"🟢 zayıf sistem";

return `
🧠 AI ENGINE v14

📊 System:
${trend}

🔥 Strongest:
${best.name} (${best.value.toFixed(3)})

🧭 Drivers:
${state.top.slice(0,3).map(x=>x.name).join(" | ")}
`;
}

/* =========================
   TABLE
========================= */
function drawTable(){

let html=`<table><tr><th></th>`;

state.headers.forEach(h=>{
html+=`<th>${h}</th>`;
});

html+=`</tr>`;

state.matrix.forEach((row,i)=>{

html+=`<tr><td>${state.headers[i]}</td>`;

row.forEach((v,j)=>{

if(i===j){
html+=`<td></td>`;
return;
}

let c =
Math.abs(v)>0.7?'#00ffd5':
Math.abs(v)>0.4?'#63f3e5':'#1f2937';

html+=`<td style="background:${c};color:#fff;font-size:10px">
${v?.toFixed?.(2)??""}
</td>`;
});

html+=`</tr>`;
});

html+=`</table>`;

document.getElementById("table").innerHTML=html;
}

/* =========================
   CHART
========================= */
function drawChart(){

if(state.chart) state.chart.destroy();

state.chart = new Chart(document.getElementById("chart"),{
type:"bar",
data:{
labels:state.top.map(x=>""),
datasets:[{
data:state.top.map(x=>x.value),
backgroundColor:"#00ffd5"
}]
},
options:{
plugins:{
legend:{display:false},
tooltip:{
callbacks:{
label:(ctx)=>{
let t=state.top[ctx.dataIndex];
return `${t.name}: ${t.value.toFixed(3)}`;
}
}
}
},
scales:{x:{display:false}}
}
});
}

/* =========================
   RUN PIPELINE
========================= */
function run(){

if(!state.headers.length) return;

matrix();
getTop();

document.getElementById("aiBox").innerText = ai();

drawTable();
drawChart();
}

/* =========================
   EXPORT
========================= */
function exportExcel(){

let project =
document.getElementById("projectName").value || "Project";

let wb = XLSX.utils.book_new();

/* MATRIX */
let ws = XLSX.utils.aoa_to_sheet([
["",...state.headers],
...state.matrix.map((r,i)=>[state.headers[i],...r])
]);

XLSX.utils.book_append_sheet(wb,ws,"Matrix");

/* AI */
let wsAI = XLSX.utils.aoa_to_sheet([
["AI REPORT"],
[ai()]
]);

XLSX.utils.book_append_sheet(wb,wsAI,"AI");

XLSX.writeFile(wb,`${project}_TP_Korelasyon.xlsx`);
}