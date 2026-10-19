let sessionId = null;
let currentResult = null;
let currentUpload = null;
const $ = (id) => document.getElementById(id);
function toast(message) {
  const el = $("toast"); el.textContent = message; el.classList.add("show");
  clearTimeout(toast.timer); toast.timer = setTimeout(() => el.classList.remove("show"), 3500);
}
function setBusy(busy) {
  $("verifyBtn").disabled = busy || !sessionId;
  $("verifyBtn").textContent = busy ? "Verifying…" : "Run verification →";
}
function show(el) { el.classList.remove("hidden"); }
function hide(el) { el.classList.add("hidden"); }
function esc(v) {
  return String(v ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}
async function api(url, options={}) {
  const res = await fetch(url, options);
  let body = {};
  try { body = await res.json(); } catch {}
  if (!res.ok) throw new Error(body.detail || `Request failed (${res.status})`);
  return body;
}
async function upload(file) {
  if (!file) return;
  const ext = file.name.split(".").pop().toLowerCase();
  if (!["xlsx","xls","csv"].includes(ext)) { toast("Please choose an .xlsx, .xls, or .csv file."); return; }
  $("uploadProgress").classList.remove("hidden");
  $("verifyBtn").disabled = true;
  const form = new FormData(); form.append("file", file);
  try {
    const data = await api("/api/upload", {method:"POST", body:form});
    sessionId = data.session_id; currentUpload = data;
    $("fileMeta").innerHTML = `<span class="file-type">${esc(ext.toUpperCase())}</span><div><strong>${esc(data.filename)}</strong><div class="muted">${data.rows.toLocaleString()} data rows · ${data.columns} columns</div></div>`;
    show($("fileMeta")); show($("config")); show($("previewWrap"));
    $("sheetSelect").innerHTML = data.sheets.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
    $("sheetSelect").disabled = ext === "csv";
    $("previewCount").textContent = `Showing ${data.preview.length} of ${data.rows} rows`;
    renderPreview(data.headers, data.preview);
    $("verifyHint").textContent = "Choose optional rules, then run verification.";
    $("verifyBtn").disabled = false;
    hide($("results")); hide($("downloads"));
    toast("File uploaded and parsed successfully.");
  } catch (e) {
    toast(e.message);
  } finally { hide($("uploadProgress")); }
}
function renderPreview(headers, rows) {
  const table = $("previewTable");
  table.innerHTML = `<thead><tr>${headers.map(h=>`<th>${esc(h || "(blank header)")}</th>`).join("")}</tr></thead><tbody>${rows.slice(0,100).map(row=>`<tr>${row.map(v=>`<td title="${esc(v)}">${esc(v)}</td>`).join("")}</tr>`).join("")}</tbody>`;
}
function renderIssues() {
  if (!currentResult) return;
  const q = $("issueSearch").value.toLowerCase().trim();
  const sev = $("severityFilter").value;
  const issues = currentResult.issues.filter(i => (sev === "all" || i.severity === sev) &&
    [i.row,i.column,i.issue,i.explanation,i.original_value,i.suggestion,i.category].join(" ").toLowerCase().includes(q));
  $("issuesBody").innerHTML = issues.map((i, n) => {
    const editable = i.row >= 2 && currentResult.columns.includes(i.column);
    return `<tr>
      <td>${i.row}</td><td>${esc(i.column)}</td><td><span class="severity ${esc(i.severity)}">${esc(i.severity)}</span></td>
      <td><strong>${esc(i.issue)}</strong><div class="muted">${esc(i.explanation)}</div></td>
      <td>${esc(i.original_value ?? "—")}</td><td>${esc(i.suggestion ?? "—")}</td>
      <td>${editable ? `<label class="correction-control"><input class="correction-check" type="checkbox" data-idx="${n}" aria-label="Approve correction for row ${i.row} ${esc(i.column)}"><input class="correction-value" data-idx="${n}" data-row="${i.row}" data-column="${esc(i.column)}" placeholder="Enter corrected value" aria-label="Proposed value for row ${i.row} ${esc(i.column)}"></label>` : `<span class="muted">Not cell-editable</span>`}</td>
    </tr>`;
  }).join("");
  // Keep the currently displayed issue objects available to the approval handler.
  window.visibleIssues = issues;
  $("noIssues").classList.toggle("hidden", issues.length !== 0);
  $("issueSummary").textContent = `${issues.length} shown · ${currentResult.error_count} errors · ${currentResult.warning_count} warnings`;
}
async function verify() {
  if (!sessionId) return;
  setBusy(true);
  const parse = id => $(id).value.split(",").map(s=>s.trim()).filter(Boolean);
  try {
    currentResult = await api("/api/verify", {method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({session_id:sessionId,sheet:$("sheetSelect").value || null,
        required_columns:parse("requiredCols"),unique_columns:parse("uniqueCols")})});
    $("statusBadge").textContent = currentResult.status;
    $("statusBadge").className = "status-badge " + (currentResult.status === "FAIL" ? "fail" : currentResult.status === "PASS WITH WARNINGS" ? "warning" : "pass");
    $("rowsChecked").textContent = currentResult.data_rows_checked.toLocaleString();
    $("passedRows").textContent = currentResult.passed_rows.toLocaleString();
    $("failedRows").textContent = currentResult.failed_rows.toLocaleString();
    $("warningCount").textContent = currentResult.warning_count.toLocaleString();
    $("warningRows").textContent = `${currentResult.rows_with_warnings} rows affected`;
    $("resultSub").textContent = currentResult.verification_scope;
    renderIssues(); show($("results")); show($("downloads"));
    $("verifyHint").textContent = "Verification completed.";
    $("results").scrollIntoView({behavior:"smooth",block:"start"});
    toast("Verification completed.");
  } catch(e) { toast(e.message); }
  finally { setBusy(false); }
}
$("fileInput").addEventListener("change", e => upload(e.target.files[0]));
$("verifyBtn").addEventListener("click", verify);
$("issueSearch").addEventListener("input", renderIssues);
$("severityFilter").addEventListener("change", renderIssues);
$("reportXlsx").addEventListener("click", () => { if(sessionId && currentResult) location.href=`/api/report/${sessionId}.xlsx`; });
$("reportCsv").addEventListener("click", () => { if(sessionId && currentResult) location.href=`/api/report/${sessionId}.csv`; });
$("correctedBtn").addEventListener("click", () => { if(sessionId && currentResult) location.href=`/api/corrected/${sessionId}`; });
$("deleteBtn").addEventListener("click", async () => {
  if (!sessionId || !confirm("Delete this upload session and its temporary file?")) return;
  try { await api(`/api/session/${sessionId}`, {method:"DELETE"}); sessionId=null; currentResult=null; currentUpload=null;
    hide($("fileMeta")); hide($("config")); hide($("previewWrap")); hide($("results")); hide($("downloads"));
    $("fileInput").value=""; $("verifyBtn").disabled=true; $("verifyHint").textContent="Upload a file to begin.";
    toast("Session and temporary file deleted.");
  } catch(e) { toast(e.message); }
});
$("applyCorrectionsBtn").addEventListener("click", async () => {
  if (!sessionId || !currentResult) return;
  const corrections = [];
  document.querySelectorAll(".correction-check:checked").forEach(check => {
    const idx = Number(check.dataset.idx);
    const issue = (window.visibleIssues || [])[idx];
    const input = document.querySelector(`.correction-value[data-idx="${idx}"]`);
    if (!issue || !input || !input.value.trim()) return;
    corrections.push({session_id:sessionId,row:issue.row,column:issue.column,new_value:input.value,approved:true});
  });
  if (!corrections.length) { toast("Select a correction and enter its new value first."); return; }
  if (!confirm(`Apply ${corrections.length} approved change(s) to a separate working copy and re-run validation?`)) return;
  try {
    const data = await api("/api/corrections", {method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({session_id:sessionId,corrections})});
    currentResult = data.result;
    renderIssues();
    $("statusBadge").textContent = currentResult.status;
    $("statusBadge").className = "status-badge " + (currentResult.status === "FAIL" ? "fail" : currentResult.status === "PASS WITH WARNINGS" ? "warning" : "pass");
    $("rowsChecked").textContent = currentResult.data_rows_checked.toLocaleString();
    $("passedRows").textContent = currentResult.passed_rows.toLocaleString();
    $("failedRows").textContent = currentResult.failed_rows.toLocaleString();
    $("warningCount").textContent = currentResult.warning_count.toLocaleString();
    $("warningRows").textContent = `${currentResult.rows_with_warnings} rows affected`;
    toast(`${data.applied_count} approved change(s) applied; validation rerun.`);
  } catch(e) { toast(e.message); }
});
const dz = $("dropzone");
dz.addEventListener("dragover", e=>{e.preventDefault();dz.classList.add("drag")});
dz.addEventListener("dragleave", ()=>dz.classList.remove("drag"));
dz.addEventListener("drop", e=>{e.preventDefault();dz.classList.remove("drag");upload(e.dataTransfer.files[0])});
