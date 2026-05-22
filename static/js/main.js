const form = document.getElementById("convertForm");
const button = document.getElementById("convertButton");
const alertBox = document.getElementById("alert");
const resultPanel = document.getElementById("resultPanel");
const resultList = document.getElementById("resultList");
const bundleLink = document.getElementById("bundleLink");
const serverStatus = document.getElementById("serverStatus");
const modeSelect = document.getElementById("mode");
const dpiField = document.getElementById("dpiField");

checkHealth();
syncModeFields();

modeSelect.addEventListener("change", syncModeFields);

form.addEventListener("submit", async (event) => {
    event.preventDefault();

    const files = document.getElementById("files").files;
    if (!files.length) {
        showAlert("error", "请选择 PDF 文件。");
        return;
    }

    const data = new FormData(form);
    setBusy(true);
    showAlert("success", "正在转换，请稍候。");
    resultPanel.hidden = true;
    resultList.innerHTML = "";
    bundleLink.hidden = true;

    try {
        const response = await fetch("/api/convert", {
            method: "POST",
            body: data,
        });

        const payload = await response.json();
        if (!response.ok) {
            throw new Error(payload.error || "转换失败。");
        }

        renderResults(payload);
        showAlert("success", "转换完成。");
    } catch (error) {
        showAlert("error", error.message || "转换失败。");
    } finally {
        setBusy(false);
    }
});

async function checkHealth() {
    try {
        const response = await fetch("/api/health");
        if (!response.ok) {
            throw new Error("bad status");
        }
        serverStatus.textContent = "服务正常";
        serverStatus.classList.add("ok");
    } catch {
        serverStatus.textContent = "服务异常";
        serverStatus.classList.remove("ok");
    }
}

function renderResults(payload) {
    resultPanel.hidden = false;
    resultList.innerHTML = "";

    if (payload.bundle) {
        bundleLink.href = payload.bundle.download_url;
        bundleLink.download = payload.bundle.filename;
        bundleLink.hidden = false;
    }

    payload.results.forEach((item) => {
        const row = document.createElement("div");
        row.className = "result-item";

        const info = document.createElement("div");
        const name = document.createElement("div");
        name.className = "result-name";
        name.textContent = item.filename;

        const meta = document.createElement("div");
        meta.className = "result-meta";
        const modeText = item.mode === "image" ? `${item.dpi} DPI 图片版` : "可编辑文字版";
        meta.textContent = `${item.pages} 页 · ${modeText} · ${formatSize(item.size)}`;

        info.append(name, meta);

        const link = document.createElement("a");
        link.className = "button";
        link.href = item.download_url;
        link.download = item.filename;
        link.textContent = "下载";

        row.append(info, link);
        resultList.append(row);
    });
}

function syncModeFields() {
    dpiField.hidden = modeSelect.value !== "image";
}

function setBusy(isBusy) {
    button.disabled = isBusy;
    button.textContent = isBusy ? "转换中" : "开始转换";
}

function showAlert(type, message) {
    alertBox.hidden = false;
    alertBox.className = `alert ${type}`;
    alertBox.textContent = message;
}

function formatSize(bytes) {
    if (bytes < 1024) {
        return `${bytes} B`;
    }
    if (bytes < 1024 * 1024) {
        return `${(bytes / 1024).toFixed(1)} KB`;
    }
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}
