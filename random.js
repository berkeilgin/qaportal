const drop = document.getElementById("drop");
const fileInput = document.getElementById("fileInput");

const randomBtn = document.getElementById("randomBtn");
const exportBtn = document.getElementById("exportBtn");

const statusEl = document.getElementById("status");
const randomCountInput = document.getElementById("randomCount");

const manualBtn = document.getElementById("manualBtn");
const manualResult = document.getElementById("manualResult");
const maxNumberInput = document.getElementById("maxNumber");

const logList = document.getElementById("logList");
const clearLogBtn = document.getElementById("clearLogBtn");

const excelPageBtn = document.getElementById("excelPageBtn");
const manualPageBtn = document.getElementById("manualPageBtn");

const excelPage = document.getElementById("excelPage");
const manualPage = document.getElementById("manualPage");

const toggle = document.getElementById("themeToggle");
const themeText = document.getElementById("themeText");

let uploadedData = [];
let selectedRandomData = [];
let manualLogs = [];
let usedNumbers = new Set();

/* THEME */
toggle.addEventListener("change", () => {
  document.body.classList.toggle("light");
  themeText.innerText = document.body.classList.contains("light") ? "☀ Light" : "🌙 Dark";
});

/* BACK */
document.getElementById("backBtnPlain").addEventListener("click", () => history.back());

/* PAGE SWITCH (FIXED) */
excelPageBtn.addEventListener("click", () => {
  excelPage.classList.remove("hidden-page");
  manualPage.classList.add("hidden-page");
  excelPageBtn.classList.add("active");
  manualPageBtn.classList.remove("active");
});

manualPageBtn.addEventListener("click", () => {
  manualPage.classList.remove("hidden-page");
  excelPage.classList.add("hidden-page");
  manualPageBtn.classList.add("active");
  excelPageBtn.classList.remove("active");
});

/* DROP */
drop.addEventListener("click", () => fileInput.click());

drop.addEventListener("dragover", e => {
  e.preventDefault();
  drop.classList.add("drag");
});

drop.addEventListener("dragleave", () => drop.classList.remove("drag"));

drop.addEventListener("drop", e => {
  e.preventDefault();
  drop.classList.remove("drag");
  handleFile(e.dataTransfer.files[0]);
});

fileInput.addEventListener("change", e => handleFile(e.target.files[0]));

function handleFile(file){
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(new Uint8Array(e.target.result), {type:"array"});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    uploadedData = XLSX.utils.sheet_to_json(sheet);
    statusEl.innerHTML = `✅ ${uploadedData.length} kayıt yüklendi`;
  };
  reader.readAsArrayBuffer(file);
}

/* EXCEL */
randomBtn.addEventListener("click", () => {
  if(!uploadedData.length) return alert("Önce dosya yükleyin");
  const count = parseInt(randomCountInput.value);
  if(!count || count < 1) return alert("Geçerli sayı girin");
  selectedRandomData = shuffle(uploadedData).slice(0,count);
  statusEl.innerHTML = `🎯 ${selectedRandomData.length} kayıt seçildi`;
});

exportBtn.addEventListener("click", () => {
  if(!selectedRandomData.length) return alert("Önce seçim yapın");
  const ws = XLSX.utils.json_to_sheet(selectedRandomData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Random");
  XLSX.writeFile(wb,"random.xlsx");
});

/* MANUAL FIXED UNIQUE */
manualBtn.addEventListener("click", () => {
  const max = parseInt(maxNumberInput.value);
  if(!max || max < 1) return alert("Geçerli maksimum girin");

  if(usedNumbers.size >= max){
    alert("Tüm sayılar kullanıldı!");
    return;
  }

  let rand;
  do {
    rand = Math.floor(Math.random()*max)+1;
  } while(usedNumbers.has(rand));

  usedNumbers.add(rand);

  manualResult.innerHTML = `🎧 ${rand}`;

  const time = new Date().toLocaleTimeString("tr-TR");
  manualLogs.unshift(`🎧 ${rand} — ${time}`);
  renderLogs();
});

clearLogBtn.addEventListener("click", () => {
  manualLogs=[];
  usedNumbers.clear();
  renderLogs();
});

function renderLogs(){
  logList.innerHTML = manualLogs.length
    ? manualLogs.map(x=>`<div class="log-item">${x}</div>`).join("")
    : "Henüz seçim yapılmadı.";
}

function shuffle(arr){
  const a=[...arr];
  for(let i=a.length-1;i>0;i--){
    const j=Math.floor(Math.random()*(i+1));
    [a[i],a[j]]=[a[j],a[i]];
  }
  return a;
}