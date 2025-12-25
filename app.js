
const $ = (id) => document.getElementById(id);
const worker = new Worker("worker.js?v=18");
let lastResults = null;

function setupTabs(){
  document.querySelectorAll(".tab").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
      document.querySelectorAll(".tabpane").forEach(x=>x.classList.remove("active"));
      btn.classList.add("active");
      $(btn.dataset.tab).classList.add("active");
    });
  });
}
setupTabs();

function renderTable(containerId, columns, rows){
  const wrap=$(containerId);
  if (!rows || rows.length===0){ wrap.innerHTML=""; return; }
  let html="<table><thead><tr>";
  for (const c of columns) html+=`<th>${c}</th>`;
  html+="</tr></thead><tbody>";
  for (const r of rows){
    html+="<tr>";
    for (const c of columns) html+=`<td>${r[c] ?? ""}</td>`;
    html+="</tr>";
  }
  html+="</tbody></table>";
  wrap.innerHTML=html;
}

function downloadCSV(filename, rows){
  if (!rows || rows.length===0) return;
  const cols=Object.keys(rows[0]);
  const esc=(s)=>`"${(s??"").toString().replaceAll('"','""')}"`;
  const lines=[cols.map(esc).join(",")];
  for (const r of rows) lines.push(cols.map(c=>esc(r[c])).join(","));
  const blob=new Blob([lines.join("\n")],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=filename;
  document.body.appendChild(a); a.click(); a.remove();
}

function setStatus(msg){ $("status").textContent = msg || ""; }
function updateRunEnabled(){
  const f1=$("file1").files[0], f2=$("file2").files[0];
  $("btnRun").disabled = !(f1 && f2);
}
$("file1").addEventListener("change", updateRunEnabled);
$("file2").addEventListener("change", updateRunEnabled);

$("btnRun").addEventListener("click", async ()=>{
  const f1=$("file1").files[0], f2=$("file2").files[0];
  const teamFilter=$("teamFilter").value || "";
  const monthPick=$("monthPick").value || "";
  if (!f1 || !f2) return;

  $("btnRun").disabled = true;
  setStatus("Reading...");
  const [buf1, buf2] = await Promise.all([f1.arrayBuffer(), f2.arrayBuffer()]);
  setStatus("Analyzing...");
  worker.postMessage({type:"analyze", buf1, buf2, teamFilter, monthPick}, [buf1, buf2]);
});

worker.onmessage = (ev)=>{
  const data = ev.data || {};
  if (data.type === "progress"){ setStatus(data.message || ""); return; }
  if (data.type === "error"){ setStatus(data.message || "Error"); $("btnRun").disabled=false; return; }
  if (data.type !== "result") return;

  setStatus("");
  $("btnRun").disabled=false;
  lastResults = data.payload;

  $("t1meta").textContent = `T1 end (month): ${lastResults.t1End || "-"}`;
  $("t2meta").textContent = `T2 end (month): ${lastResults.t2End || "-"}`;

  $("resultsCard").style.display = "block";

  $("summary").innerHTML = `
    <div class="pill">Month: <b>${lastResults.monthLabel}</b></div>
    <div class="pill">T1 end: <b>${lastResults.t1End || "-"}</b></div>
    <div class="pill">T2 end: <b>${lastResults.t2End || "-"}</b></div>
    <div class="pill">Export base: <b>${lastResults.exportBase}</b></div>
  `;

  const colsEA = [
    "team","ea","total_records",
    "month_connected","month_rate",
    "added_connected",
    "month_first","month_follow_up",
    "remark"
  ];
  renderTable("eaTable", colsEA, lastResults.eaOverview);

  const colsPool = [
    "team","pool","ea","total_records",
    "month_connected","month_rate",
    "added_connected",
    "month_first","month_follow_up"
  ];
  renderTable("poolTable", colsPool, lastResults.poolOverview);

  const colsRec = ["team","ea","learnerId","family_id","pool","lastConn","lastMonthCons","thisMonthCons","remaining","reason"];
  renderTable("recommendTable", colsRec, lastResults.recommendations);

  $("exportEA").onclick=()=>downloadCSV(`${lastResults.exportBase}-EA.csv`, lastResults.eaOverview);
  $("exportPool").onclick=()=>downloadCSV(`${lastResults.exportBase}-POOL.csv`, lastResults.poolOverview);
  $("exportRecommend").onclick=()=>downloadCSV(`${lastResults.exportBase}.csv`, lastResults.recommendations);
};


worker.onerror = (e)=>{
  setStatus(`Worker error: ${e.message || e.type || 'unknown'}`);
  $("btnRun").disabled=false;
};
