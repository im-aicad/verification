const API_BASE = window.location.origin;

const state = {
  selectedFile: null,
  running: false,
  completed: false,
  hasResultFile: false,
  lastResult: null,
  captureStream: null,
  resultSummaryPrinted: false,
};

const els = {
  footerStatus: document.getElementById("footerStatus"),
  footerMeta: document.getElementById("footerMeta"),
  refreshButton: document.getElementById("refreshButton"),
  clearLogButton: document.getElementById("clearLogButton"),
  catpartInput: document.getElementById("catpartInput"),
  uploadedFileName: document.getElementById("uploadedFileName"),
  startButton: document.getElementById("startButton"),
  resetButton: document.getElementById("resetButton"),
  captureButton: document.getElementById("captureButton"),
  stopCaptureButton: document.getElementById("stopCaptureButton"),
  downloadButton: document.getElementById("downloadButton"),
  viewport: document.getElementById("viewport"),
  viewportPlaceholder: document.getElementById("viewportPlaceholder"),
  vpLive: document.getElementById("vpLive"),
  vpInfo: document.getElementById("vpInfo"),
  barTitle: document.getElementById("barTitle"),
  vpAngle: document.getElementById("vpAngle"),
  connectionState: document.getElementById("connectionState"),
  outputSubtitle: document.getElementById("outputSubtitle"),
  sourcePath: document.getElementById("sourcePath"),
  resultState: document.getElementById("resultState"),
  logOutput: document.getElementById("logOutput"),
};

function basename(path) {
  return String(path || "").split(/[\\/]/).pop() || "-";
}

function setBusy(isBusy) {
  state.running = isBusy;
  for (const button of [
    els.startButton,
    els.resetButton,
  ]) {
    button.disabled = isBusy;
  }
  els.footerStatus.textContent = isBusy ? "执行中" : state.completed ? "已完成" : "待机";
  els.footerMeta.textContent = isBusy
    ? "CATIA 自动化执行中，请等待结果"
    : state.completed
      ? "检测完成，可下载结果文件"
      : "等待上传或使用默认 CATPart";
}

function clearEmptyState() {
  const empty = els.logOutput.querySelector(".empty-state, .empty-message-placeholder");
  if (empty) empty.remove();
}

function escapeHTML(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function levelDotClass(level) {
  if (level === "success") return "dot-success";
  if (level === "error") return "dot-error";
  if (level === "warn") return "dot-warn";
  if (level === "system") return "dot-system";
  return "dot-info";
}

function addLog(message, level = "info", time = new Date().toLocaleTimeString()) {
  clearEmptyState();
  const row = document.createElement("div");
  row.className = "message-item system";
  row.innerHTML = `
    <div class="message-bubble">${escapeHTML(message)}</div>
    <div class="message-meta">
      <span class="message-dot ${levelDotClass(level)}"></span>
      <span>${escapeHTML(time || "")}</span>
    </div>
  `;
  els.logOutput.appendChild(row);
  els.logOutput.scrollTop = els.logOutput.scrollHeight;
}

function addResultSummary(result) {
  if (result.summary_text) {
    addLog(result.summary_text, result.success ? "info" : "error");
    return;
  }

  const temporaryCount =
    result.temporary_reflection_points?.total_reflection_point_count ?? "-";
  const reflectionCount =
    result.regulation_reflection_points?.reflection_point_count ?? "-";
  const gapMirrors = result.gap_check?.mirrors || {};
  const gapText =
    Object.entries(gapMirrors)
      .map(([side, info]) => {
        if (info.skipped) return `${side}镜片: 跳过`;
        return `${side}镜片: ${info.all_points_pass ? "通过" : "未通过"}`;
      })
      .join(" / ") || "-";
  const savedPath = result.saved_as_path || "-";

  addLog(
    [
      "检测结果摘要",
      `临时反射取点: ${temporaryCount}`,
      `正式反射取点: ${reflectionCount}`,
      `间隙校验: ${gapText}`,
      `另存文件: ${savedPath}`,
    ].join("\n"),
    result.success ? "info" : "error",
  );
}

async function api(path, options = {}) {
  const response = await fetch(`${API_BASE}${path}`, options);
  const contentType = response.headers.get("content-type") || "";
  const data = contentType.includes("application/json")
    ? await response.json()
    : await response.text();
  if (!response.ok) {
    throw new Error(data?.detail || data || `HTTP ${response.status}`);
  }
  return data;
}

function updateStatus(payload) {
  state.running = !!payload.running;
  state.completed = !!payload.completed;
  state.hasResultFile = !!payload.has_result_file;
  els.sourcePath.textContent = payload.active_file || "-";
  els.uploadedFileName.textContent = payload.uploaded_file_name || basename(payload.active_file);
  els.downloadButton.disabled = !state.hasResultFile;

  if (payload.last_error) {
    els.resultState.textContent = "失败";
    els.resultState.dataset.state = "error";
  } else if (payload.running) {
    els.resultState.textContent = "运行中";
    els.resultState.dataset.state = "running";
  } else if (payload.completed) {
    els.resultState.textContent = "完成";
    els.resultState.dataset.state = "done";
  } else {
    els.resultState.textContent = "暂无结果";
    els.resultState.dataset.state = "idle";
  }
  setBusy(state.running);
}

function updateSummary(result) {
  if (!result || Object.keys(result).length === 0) return;
  state.lastResult = result;
  if (!state.resultSummaryPrinted) {
    addResultSummary(result);
    state.resultSummaryPrinted = true;
  }
}

async function refreshStatus() {
  const payload = await api("/api/status");
  updateStatus(payload);
  if (payload.completed || payload.last_error) {
    try {
      updateSummary(await api("/api/result"));
    } catch {
      // No result available yet.
    }
  }
}

async function uploadSelectedFile() {
  if (!state.selectedFile) {
    addLog("请选择 CATPart 文件", "warn");
    return;
  }
  const form = new FormData();
  form.append("file", state.selectedFile);
  setBusy(true);
  try {
    const result = await api("/api/upload", { method: "POST", body: form });
    addLog(`上传完成: ${result.filename}`);
    await refreshStatus();
  } catch (error) {
    addLog(`上传失败: ${error.message}`, "error");
  } finally {
    setBusy(false);
  }
}

async function startDetection(useDefaultFile = false) {
  setBusy(true);
  state.resultSummaryPrinted = false;
  try {
    await api("/api/start", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ use_default_file: useDefaultFile }),
    });
    addLog(useDefaultFile ? "已使用默认文件启动检测" : "已启动检测");
    await refreshStatus();
  } catch (error) {
    addLog(`启动失败: ${error.message}`, "error");
    setBusy(false);
  }
}

async function resetState() {
  try {
    await api("/api/reset", { method: "POST" });
    state.selectedFile = null;
    els.catpartInput.value = "";
    els.uploadedFileName.textContent = "未选择文件";
    state.resultSummaryPrinted = false;
    await refreshStatus();
  } catch (error) {
    addLog(`重置失败: ${error.message}`, "error");
  }
}

async function startCapture() {
  if (!navigator.mediaDevices?.getDisplayMedia) {
    addLog("当前浏览器不支持应用窗口捕获", "error");
    return;
  }
  try {
    const stream = await navigator.mediaDevices.getDisplayMedia({
      video: {
        displaySurface: "window",
        frameRate: 30,
      },
      audio: false,
    });
    state.captureStream = stream;
    els.vpLive.srcObject = stream;
    els.viewport.classList.add("live");
    els.viewportPlaceholder.classList.add("hidden");
    els.vpInfo.innerHTML = "状态: 正在共享窗口<br>建议选择 CATIA V5 主窗口";
    addLog("应用窗口捕获已开始");
    stream.getVideoTracks()[0]?.addEventListener("ended", stopCapture);
  } catch (error) {
    addLog(`应用捕获取消或失败: ${error.message}`, "warn");
  }
}

function stopCapture() {
  if (state.captureStream) {
    for (const track of state.captureStream.getTracks()) {
      track.stop();
    }
  }
  state.captureStream = null;
  els.vpLive.srcObject = null;
  els.viewport.classList.remove("live");
  els.viewportPlaceholder.classList.remove("hidden");
  els.vpInfo.innerHTML = "状态: 等待捕获<br>请选择 CATIA 窗口";
}

function connectWebSocket() {
  const wsProtocol = window.location.protocol === "https:" ? "wss:" : "ws:";
  const socket = new WebSocket(`${wsProtocol}//${window.location.host}/ws`);
  socket.onopen = () => {
    els.connectionState.textContent = "已连接";
    els.connectionState.dataset.state = "connected";
  };
  socket.onclose = () => {
    els.connectionState.textContent = "已断开";
    els.connectionState.dataset.state = "disconnected";
    setTimeout(connectWebSocket, 1500);
  };
  socket.onerror = () => {
    els.connectionState.textContent = "连接异常";
    els.connectionState.dataset.state = "error";
  };
  socket.onmessage = async (event) => {
    const payload = JSON.parse(event.data);
    if (payload.type === "status") {
      updateStatus(payload.data || {});
      if (payload.data?.completed || payload.data?.last_error) {
        try {
          updateSummary(await api("/api/result"));
        } catch {
          // No result available.
        }
      }
    } else if (payload.type === "log") {
      const data = payload.data || {};
      addLog(data.message || "", data.level || "info", data.time);
    }
  };
}

els.catpartInput.addEventListener("change", () => {
  const file = els.catpartInput.files?.[0] || null;
  state.selectedFile = file;
  els.uploadedFileName.textContent = file ? file.name : "未选择文件";
  if (file) {
    uploadSelectedFile();
  }
});

els.startButton.addEventListener("click", () => startDetection(false));
els.resetButton.addEventListener("click", resetState);
els.refreshButton.addEventListener("click", refreshStatus);
els.clearLogButton.addEventListener("click", () => {
  els.logOutput.innerHTML = `
    <div class="empty-message-placeholder">
      等待检测流程启动<br>
      系统日志会在这里实时展开
    </div>
  `;
});
els.captureButton.addEventListener("click", startCapture);
els.stopCaptureButton.addEventListener("click", stopCapture);
els.downloadButton.addEventListener("click", () => {
  window.location.href = `${API_BASE}/api/download-result`;
});

connectWebSocket();
refreshStatus().catch((error) => addLog(`状态读取失败: ${error.message}`, "error"));
