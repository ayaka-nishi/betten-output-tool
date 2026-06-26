const govSelect    = document.getElementById("govSelect");
const mayorName     = document.getElementById("mayorName");
const govHint       = document.getElementById("govHint");
const statusBar      = document.getElementById("statusBar");
const statusText     = document.getElementById("statusText");
const dropzone       = document.getElementById("dropzone");
const dropzoneIcon   = document.getElementById("dropzoneIcon");
const dropzoneTitle  = document.getElementById("dropzoneTitle");
const dropzoneSub    = document.getElementById("dropzoneSub");
const fileInput      = document.getElementById("fileInput");
const btnRun         = document.getElementById("btnRun");
const btnDownload    = document.getElementById("btnDownload");
const btnClear       = document.getElementById("btnClear");
const spinner        = document.getElementById("spinner");
const logBody        = document.getElementById("logBody");

let selectedFile  = null;
let downloadToken = null;

function setFile(file) {
  if (file && !file.name.toLowerCase().endsWith(".xlsx")) {
    alert("xlsxファイルを選択してください");
    return;
  }
  selectedFile = file;
  if (file) {
    dropzone.classList.add("has-file");
    dropzoneIcon.textContent = "✅";
    dropzoneTitle.textContent = file.name;
    dropzoneSub.textContent = "クリックして別のファイルに変更";
  } else {
    dropzone.classList.remove("has-file");
    dropzoneIcon.textContent = "📁";
    dropzoneTitle.textContent = "ここにxlsxファイルをドラッグ＆ドロップ";
    dropzoneSub.textContent = "またはクリックして選択";
  }
  refreshState();
}

function refreshState() {
  const hasFile = selectedFile !== null;
  btnRun.disabled = !hasFile;
  statusText.textContent = hasFile
    ? "準備完了！「生成」ボタンを押してください"
    : "まずは対象ファイル（xlsx）を選択してください";
}

dropzone.addEventListener("click", () => fileInput.click());

["dragenter", "dragover"].forEach(evt =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  })
);
["dragleave", "drop"].forEach(evt =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  })
);
dropzone.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files[0];
  if (file) setFile(file);
});

fileInput.addEventListener("change", () => {
  setFile(fileInput.files[0] || null);
});

async function updateGovHint() {
  const params = new URLSearchParams({
    gov_label: govSelect.value,
    mayor_name: mayorName.value,
  });
  try {
    const res = await fetch(`/gov_hint?${params}`);
    const data = await res.json();
    govHint.textContent = data.text || "";
  } catch {
    govHint.textContent = "";
  }
}

govSelect.addEventListener("change", updateGovHint);
mayorName.addEventListener("input", updateGovHint);

function appendLog(text, cls) {
  const span = document.createElement("span");
  if (cls) span.className = cls;
  span.textContent = text + "\n";
  logBody.appendChild(span);
  logBody.scrollTop = logBody.scrollHeight;
}

function clearLog() {
  logBody.textContent = "";
}

function setPrimaryRole(isRunPrimary) {
  btnRun.classList.toggle("is-secondary-role", !isRunPrimary);
  btnDownload.classList.toggle("is-primary-role", !isRunPrimary);
}

btnRun.addEventListener("click", async () => {
  if (!selectedFile) return;

  setPrimaryRole(true);
  btnRun.disabled = true;
  btnDownload.disabled = true;
  spinner.hidden = false;
  statusText.textContent = "生成中です。しばらくお待ちください…";
  clearLog();
  appendLog("=== 提出済み（○）商品の別添1・別添2 生成 ===\n");

  const formData = new FormData();
  formData.append("xlsx_file", selectedFile);
  formData.append("gov_label", govSelect.value);
  formData.append("mayor_name", mayorName.value);

  try {
    const res = await fetch("/generate", { method: "POST", body: formData });
    const data = await res.json();

    if (!res.ok) {
      appendLog(`⚠ ${data.error || "エラーが発生しました"}`, "error");
      statusText.textContent = "生成できませんでした。ログを確認してください";
      return;
    }

    for (const line of data.logs) {
      let cls = "";
      if (line.startsWith("✓")) cls = "success";
      else if (line.startsWith("⚠")) cls = "warn";
      appendLog(line, cls);
    }

    downloadToken = data.download_token;
    btnDownload.disabled = false;
    setPrimaryRole(false);
    statusText.textContent = `完了：${data.n_companies}社 / ${data.n_products}件 を出力しました 🎉`;
  } catch (e) {
    appendLog(`⚠ 通信エラーが発生しました: ${e}`, "error");
    statusText.textContent = "生成できませんでした。ログを確認してください";
  } finally {
    btnRun.disabled = false;
    spinner.hidden = true;
  }
});

btnDownload.addEventListener("click", () => {
  if (!downloadToken) return;
  window.location.href = `/download/${downloadToken}`;
});

btnClear.addEventListener("click", () => {
  downloadToken = null;
  fileInput.value = "";
  setFile(null);
  mayorName.value = "";
  govSelect.selectedIndex = 0;
  btnDownload.disabled = true;
  setPrimaryRole(true);
  clearLog();
  logBody.textContent = "準備が整うとここに進行状況が表示されます。";
  updateGovHint();
  refreshState();
});

updateGovHint();
refreshState();
