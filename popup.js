// WhatsApp Group & Chat Exporter - Popup Controller

const statusPill = document.getElementById("statusPill");
const statusPillText = document.getElementById("statusPillText");
const chatNameEl = document.getElementById("chatName");
const chatStatsEl = document.getElementById("chatStats");

const startTimeInput = document.getElementById("startTime");
const endTimeInput = document.getElementById("endTime");
const btnNow = document.getElementById("btnNow");
const btnClearDates = document.getElementById("btnClearDates");

const btnExportExcel = document.getElementById("btnExportExcel");
const btnAutoFetch = document.getElementById("btnAutoFetch");
const btnStopScroll = document.getElementById("btnStopScroll");
const btnExportCsv = document.getElementById("btnExportCsv");

const statusCard = document.getElementById("statusCard");
const statusText = document.getElementById("statusText");
const progressBar = document.getElementById("progressBar");

let activeTabId = null;
let isExportRunning = false;

// Format Date object to local YYYY-MM-DDTHH:mm string for datetime-local input
function formatLocalDatetime(date) {
  const pad = n => String(n).padStart(2, "0");
  const y = date.getFullYear();
  const m = pad(date.getMonth() + 1);
  const d = pad(date.getDate());
  const hh = pad(date.getHours());
  const mm = pad(date.getMinutes());
  return `${y}-${m}-${d}T${hh}:${mm}`;
}

// Return current date & time formatted for input
function getNowString() {
  return formatLocalDatetime(new Date());
}

// Find active WhatsApp Web tab
async function getWhatsAppTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || !tab.url || !tab.url.startsWith("https://web.whatsapp.com/")) {
    throw new Error("Please open WhatsApp Web (https://web.whatsapp.com) and select a chat.");
  }
  return tab;
}

// Ensure content scripts are loaded even if WhatsApp tab was opened before extension installation
async function sendTabMessage(tabId, message) {
  try {
    return await chrome.tabs.sendMessage(tabId, message);
  } catch (err) {
    // If receiving end doesn't exist, inject scripts programmatically and retry
    if (chrome.scripting) {
      try {
        await chrome.scripting.executeScript({
          target: { tabId },
          files: ["xlsx.full.min.js", "content.js"]
        });
        // Brief pause after injection
        await new Promise(r => setTimeout(r, 200));
        return await chrome.tabs.sendMessage(tabId, message);
      } catch (injectionErr) {
        throw new Error("Could not connect to WhatsApp Web. Please refresh the WhatsApp Web page and try again.");
      }
    }
    throw err;
  }
}

// Update status message UI
function setStatus(text, type = "normal", showProgress = false) {
  statusText.textContent = text;
  statusCard.className = `status-card ${type}`;
  progressBar.style.display = showProgress ? "block" : "none";
}

// Update chat info display
async function refreshChatInfo() {
  try {
    const tab = await getWhatsAppTab();
    activeTabId = tab.id;

    statusPill.className = "chat-status-pill";
    statusPillText.textContent = "Connecting...";

    const res = await sendTabMessage(tab.id, { action: "get-chat-info" });
    if (!res || !res.ok) {
      statusPill.className = "chat-status-pill";
      statusPillText.textContent = "Not connected";
      chatNameEl.innerHTML = "<span>No conversation selected</span>";
      chatStatsEl.textContent = "Click on a group or chat in WhatsApp Web.";
      return;
    }

    if (res.isOpen) {
      statusPill.className = "chat-status-pill active";
      statusPillText.textContent = res.isGroup ? "Group Active" : "Chat Active";
      const icon = res.isGroup ? "👥" : "👤";
      chatNameEl.innerHTML = `<span>${icon} ${escapeHtml(res.chatName)}</span>`;
      
      let stats = `${res.count} messages currently loaded`;
      if (res.earliest && res.latest) {
        stats += ` (${res.earliest} → ${res.latest})`;
      }
      chatStatsEl.textContent = stats;
    } else {
      statusPill.className = "chat-status-pill";
      statusPillText.textContent = "No chat open";
      chatNameEl.innerHTML = "<span>Please select a chat</span>";
      chatStatsEl.textContent = "Open a WhatsApp group or direct conversation.";
    }
  } catch (err) {
    statusPill.className = "chat-status-pill";
    statusPillText.textContent = "Not connected";
    chatNameEl.innerHTML = `<span>WhatsApp Web not open</span>`;
    chatStatsEl.textContent = err.message;
    setStatus(err.message, "error");
  }
}

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Execute export
async function executeExport(actionName, autoScroll = false) {
  if (isExportRunning) return;

  const startVal = startTimeInput.value;
  const endVal = endTimeInput.value;

  // Validation: start time cannot be greater than end time
  if (startVal && endVal && new Date(startVal) > new Date(endVal)) {
    setStatus("Starting Time cannot be later than Ending Time. Please adjust your range.", "error");
    return;
  }

  isExportRunning = true;
  btnExportExcel.disabled = true;
  btnAutoFetch.disabled = true;
  btnExportCsv.disabled = true;

  if (autoScroll) {
    btnStopScroll.style.display = "block";
    setStatus("🔄 Auto-fetching older chat history... Please keep WhatsApp Web tab open.", "loading", true);
  } else {
    setStatus("⏳ Extracting messages and building Excel file...", "loading", true);
  }

  try {
    const tab = await getWhatsAppTab();
    const res = await sendTabMessage(tab.id, {
      action: actionName,
      startTime: startVal || null,
      endTime: endVal || null,
      autoScroll
    });

    if (!res || !res.ok) {
      throw new Error(res?.error || "Export failed. Please check WhatsApp Web.");
    }

    const fileLabel = actionName === "export-excel" ? "Excel (.xlsx)" : "CSV";
    setStatus(`✅ Successfully exported ${res.count} messages to ${fileLabel}!\nFile: ${res.filename}`, "success");
    await refreshChatInfo();
  } catch (err) {
    setStatus(`❌ Error: ${err.message}`, "error");
  } finally {
    isExportRunning = false;
    btnExportExcel.disabled = false;
    btnAutoFetch.disabled = false;
    btnExportCsv.disabled = false;
    btnStopScroll.style.display = "none";
    progressBar.style.display = "none";
  }
}

// Stop scrolling button
btnStopScroll.addEventListener("click", async () => {
  btnStopScroll.textContent = "Stopping and preparing download...";
  try {
    const tab = await getWhatsAppTab();
    await sendTabMessage(tab.id, { action: "stop-scroll" });
  } catch (err) {
    console.error(err);
  }
});

// "Now" button for End Time
btnNow.addEventListener("click", () => {
  endTimeInput.value = getNowString();
  setStatus("End Time set to current time (Now).", "normal");
});

// Clear dates button
btnClearDates.addEventListener("click", () => {
  startTimeInput.value = "";
  endTimeInput.value = "";
  document.querySelectorAll(".preset-pill").forEach(p => p.classList.remove("active"));
  setStatus("Date/time filters cleared. All messages will be exported.", "normal");
});

// Presets handling
document.querySelectorAll(".preset-pill").forEach(pill => {
  pill.addEventListener("click", () => {
    document.querySelectorAll(".preset-pill").forEach(p => p.classList.remove("active"));
    pill.classList.add("active");

    const preset = pill.getAttribute("data-preset");
    const now = new Date();

    if (preset === "1h") {
      const past1h = new Date(now.getTime() - 60 * 60 * 1000);
      startTimeInput.value = formatLocalDatetime(past1h);
      endTimeInput.value = formatLocalDatetime(now);
      setStatus("Range set: Past 1 hour to Now.", "normal");
    } else if (preset === "today") {
      const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
      startTimeInput.value = formatLocalDatetime(startOfToday);
      endTimeInput.value = formatLocalDatetime(now);
      setStatus("Range set: Today (00:00) to Now.", "normal");
    } else if (preset === "24h") {
      const past24h = new Date(now.getTime() - 24 * 60 * 60 * 1000);
      startTimeInput.value = formatLocalDatetime(past24h);
      endTimeInput.value = formatLocalDatetime(now);
      setStatus("Range set: Past 24 hours to Now.", "normal");
    } else if (preset === "7d") {
      const past7d = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
      startTimeInput.value = formatLocalDatetime(past7d);
      endTimeInput.value = formatLocalDatetime(now);
      setStatus("Range set: Past 7 days to Now.", "normal");
    } else if (preset === "all") {
      startTimeInput.value = "";
      endTimeInput.value = "";
      setStatus("Range cleared: All messages.", "normal");
    }
  });
});

// Listen to progress updates from content script during auto-scroll
chrome.runtime.onMessage.addListener(msg => {
  if (msg.action === "scan-progress") {
    let text = `🔄 Scanning history: ${msg.count} messages collected...`;
    if (msg.earliest) {
      text += `\nEarliest date reached: ${msg.earliest}`;
    }
    if (msg.reachedTarget) {
      text += `\nTarget start date reached! Preparing export...`;
    }
    setStatus(text, "loading", true);
  }
});

// Button event listeners
btnExportExcel.addEventListener("click", () => executeExport("export-excel", false));
btnAutoFetch.addEventListener("click", () => executeExport("export-excel", true));
btnExportCsv.addEventListener("click", () => executeExport("export-csv", false));

// Initialize popup on open
document.addEventListener("DOMContentLoaded", () => {
  refreshChatInfo();
});
