/**
 * WhatsApp Group & Chat Exporter
 * Runs locally inside WhatsApp Web to export messages with date, time, sender, type, and content to Excel (.xlsx) or CSV.
 */
(() => {
  if (window.__waChatExporterInstalled) return;
  window.__waChatExporterInstalled = true;

  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

  // Global state for auto-scroll operations
  let isScrollingActive = false;
  let shouldStopScroll = false;

  // Clean hidden Unicode formatting characters often inserted by WhatsApp
  function cleanUnicode(str) {
    if (!str) return "";
    return str.replace(/[\u200e\u200f\u202a-\u202e\u200b\u00a0]/g, " ").trim();
  }

  // Parse time & date from WhatsApp's data-pre-plain-text (e.g. "[10:20 AM, 22/10/2026] Sender: ")
  function parseWhatsAppMeta(metaStr) {
    if (!metaStr) return null;
    const clean = cleanUnicode(metaStr);
    const match = clean.match(/^\[(.*?)\]\s*(.*?):\s*$/);
    if (!match) return null;

    const rawDateTime = match[1];
    const sender = match[2].trim() || "Unknown";

    const parts = rawDateTime.split(",").map(s => s.trim());
    let timeStr = "";
    let dateStr = "";

    for (const part of parts) {
      if (/\d{1,2}:\d{2}/.test(part)) {
        timeStr = part;
      } else if (/\d{1,4}[-/. ]\d{1,2}[-/. ]\d{1,4}/.test(part) || /[a-zA-Z]{3,}/.test(part)) {
        dateStr = part;
      }
    }

    if (!timeStr || !dateStr) {
      const timeMatch = rawDateTime.match(/(\d{1,2}:\d{2}(?::\d{2})?\s*(?:[ap]\.?m\.?)?)/i);
      if (timeMatch) timeStr = timeMatch[1].trim();
      dateStr = rawDateTime.replace(timeStr, "").replace(/,/g, "").trim();
    }

    const parsedDate = parseDateTime(dateStr, timeStr);

    return {
      sender,
      dateStr,
      timeStr,
      timestamp: parsedDate ? parsedDate.getTime() : null,
      formattedDate: parsedDate ? formatDate(parsedDate) : dateStr,
      formattedTime: parsedDate ? formatTime(parsedDate) : timeStr
    };
  }

  function parseDateTime(dateStr, timeStr) {
    if (!dateStr && !timeStr) return null;

    let hours = 0, minutes = 0, seconds = 0;
    if (timeStr) {
      const tMatch = timeStr.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap]\.?m\.?)?/i);
      if (tMatch) {
        hours = parseInt(tMatch[1], 10);
        minutes = parseInt(tMatch[2], 10);
        seconds = tMatch[3] ? parseInt(tMatch[3], 10) : 0;
        const ampm = tMatch[4] ? tMatch[4].toLowerCase().replace(/\./g, "") : null;
        if (ampm === "pm" && hours < 12) hours += 12;
        if (ampm === "am" && hours === 12) hours = 0;
      }
    }

    let year = new Date().getFullYear();
    let month = new Date().getMonth();
    let day = new Date().getDate();

    if (dateStr) {
      const isoMatch = dateStr.match(/^(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})$/);
      if (isoMatch) {
        year = parseInt(isoMatch[1], 10);
        month = parseInt(isoMatch[2], 10) - 1;
        day = parseInt(isoMatch[3], 10);
      } else {
        const numMatch = dateStr.match(/^(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{2,4})$/);
        if (numMatch) {
          const p1 = parseInt(numMatch[1], 10);
          const p2 = parseInt(numMatch[2], 10);
          let y = parseInt(numMatch[3], 10);
          if (y < 100) y += 2000;
          year = y;

          if (p1 > 12) {
            day = p1;
            month = p2 - 1;
          } else if (p2 > 12) {
            month = p1 - 1;
            day = p2;
          } else {
            const isUS = (typeof navigator !== "undefined" && navigator.language === "en-US");
            if (isUS) {
              month = p1 - 1;
              day = p2;
            } else {
              day = p1;
              month = p2 - 1;
            }
          }
        } else {
          const d = new Date(dateStr);
          if (!isNaN(d.getTime())) {
            year = d.getFullYear();
            month = d.getMonth();
            day = d.getDate();
          }
        }
      }
    }

    const result = new Date(year, month, day, hours, minutes, seconds);
    return isNaN(result.getTime()) ? null : result;
  }

  function formatDate(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  function formatTime(d) {
    const h = String(d.getHours()).padStart(2, "0");
    const min = String(d.getMinutes()).padStart(2, "0");
    const s = String(d.getSeconds()).padStart(2, "0");
    return `${h}:${min}:${s}`;
  }

  // Find scrollable chat container
  function getChatPanel() {
    const candidates = [
      document.querySelector('[data-testid="conversation-panel-messages"]'),
      document.querySelector('#main [role="application"]'),
      document.querySelector('#main div[tabindex="-1"]'),
      document.querySelector('#main')
    ].filter(Boolean);

    for (const el of candidates) {
      const scroller = findScrollable(el);
      if (scroller) return scroller;
    }
    return [...document.querySelectorAll('#main *')]
      .filter(el => el.scrollHeight > el.clientHeight + 100 && el.clientHeight > 150)
      .sort((a, b) => b.clientHeight - a.clientHeight)[0] || null;
  }

  function findScrollable(root) {
    if (!root) return null;
    const all = [root, ...root.querySelectorAll('*')];
    return all.find(el => {
      const s = window.getComputedStyle(el);
      return el.scrollHeight > el.clientHeight + 50 &&
        /(auto|scroll)/.test(s.overflowY) &&
        el.clientHeight > 120;
    }) || null;
  }

  // Get current active chat name and whether it is a group
  function getCurrentChatInfo() {
    const main = document.querySelector('#main');
    if (!main) return { isOpen: false, chatName: "No chat open", isGroup: false };

    const header = main.querySelector('header');
    let chatName = "WhatsApp Chat";
    let isGroup = false;

    if (header) {
      const titleEl = header.querySelector('[title], span[dir="auto"]');
      if (titleEl) {
        chatName = (titleEl.getAttribute('title') || titleEl.textContent || "WhatsApp Chat").trim().slice(0, 100);
      }
      // Check if group: subtitle often contains member count or comma-separated names
      const subtitleEl = header.querySelector('span[title], div[role="button"] span');
      const subtitleText = subtitleEl?.textContent?.toLowerCase() || "";
      if (subtitleText.includes(',') || subtitleText.includes('group') || subtitleText.includes('click here for group')) {
        isGroup = true;
      }
    }

    return { isOpen: true, chatName, isGroup };
  }

  // Parse a single message DOM node into structured data
  function parseMessageNode(node, lastKnownDate) {
    const dataId = node.getAttribute('data-id') || node.closest('[data-id]')?.getAttribute('data-id') || "";

    // Determine direction
    const isOutgoing = node.classList.contains('message-out') ||
      Boolean(node.closest('.message-out')) ||
      dataId.startsWith('true_') ||
      Boolean(node.querySelector('[data-testid="tail-out"], [data-icon="msg-dblcheck"], [data-icon="msg-check"]'));

    const isSystem = (node.getAttribute('data-testid') === 'chat-system-message') ||
      Boolean(node.querySelector('[data-testid="chat-system-message"]')) ||
      node.classList.contains('system-message');

    // Extract meta if present
    const metaEl = node.querySelector('[data-pre-plain-text]') ||
      (node.hasAttribute('data-pre-plain-text') ? node : null) ||
      node.closest('[data-pre-plain-text]');
    const metaStr = metaEl ? metaEl.getAttribute('data-pre-plain-text') : "";

    const parsedMeta = parseWhatsAppMeta(metaStr);

    // Resolve Sender Name
    let sender = "";
    if (isSystem) {
      sender = "System";
    } else if (isOutgoing) {
      sender = "You";
    } else if (parsedMeta?.sender && parsedMeta.sender !== "You") {
      sender = parsedMeta.sender;
    } else {
      // Check author elements inside group chat message bubble
      const authorEl = node.querySelector('[data-testid="msg-author"]') ||
        node.querySelector('span[data-testid="msg-author"]') ||
        node.querySelector('span[dir="auto"][aria-label]');
      if (authorEl && authorEl.textContent.trim()) {
        sender = authorEl.textContent.trim();
      } else {
        // Fallback: check if sender phone number is encoded in data-id (e.g. false_GROUP@g.us_ID_SENDER@c.us)
        const phoneMatch = dataId.match(/_([0-9a-zA-Z\.\-]+)@c\.us$/);
        if (phoneMatch) {
          sender = "+" + phoneMatch[1];
        } else {
          sender = isOutgoing ? "You" : "Unknown";
        }
      }
    }

    // Resolve Message Type & Body
    let type = "Text";
    let body = "";

    const textNodes = [...node.querySelectorAll('.selectable-text, [data-testid="selectable-text"]')];
    const textContent = textNodes.map(el => (el.innerText || el.textContent || "").trim()).filter(Boolean).join("\n").trim();

    if (node.querySelector('[data-testid="audio-player"], [data-testid="voice-note"], audio')) {
      type = "Voice Note";
      const duration = node.querySelector('[data-testid="audio-duration"]')?.textContent?.trim() || "";
      body = duration ? `[Voice Note (${duration})]` : "[Voice Note]";
      if (textContent) body += ` ${textContent}`;
    } else if (node.querySelector('[data-testid="image-thumb"], img[src*="blob:"]')) {
      type = "Image";
      body = textContent ? `[Image] ${textContent}` : "[Image]";
    } else if (node.querySelector('[data-testid="video-thumb"], video')) {
      type = "Video";
      body = textContent ? `[Video] ${textContent}` : "[Video]";
    } else if (node.querySelector('[data-testid="document-thumb"], [data-testid="file-download"]')) {
      type = "Document";
      const docTitle = node.querySelector('[title], [data-testid="document-thumb"] span[dir="auto"]')?.textContent?.trim() || "";
      body = docTitle ? `[Document: ${docTitle}]` : "[Document]";
      if (textContent) body += ` ${textContent}`;
    } else if (node.querySelector('[data-testid="sticker"], [data-testid="custom-sticker"], img[src*="sticker"]')) {
      type = "Sticker";
      body = "[Sticker]";
    } else if (node.querySelector('[data-testid="location"]')) {
      type = "Location";
      body = "[Location]";
    } else if (node.querySelector('[data-testid="contact-card"]')) {
      type = "Contact";
      body = `[Contact Card: ${textContent || "Contact"}]`;
    } else if (node.querySelector('[data-testid="poll-bubble"]')) {
      type = "Poll";
      body = `[Poll] ${textContent}`;
    } else if (/message was deleted|deleted this message/i.test(node.innerText || "")) {
      type = "Deleted";
      body = "[Deleted Message]";
    } else {
      type = isSystem ? "System" : "Text";
      body = textContent || (node.innerText || "").trim();
    }

    if (!body) return null;

    // Check for Quoted / Replied message
    let quotedSender = "";
    let quotedMessage = "";
    const quoteEl = node.querySelector('[data-testid="quoted-message"]');
    if (quoteEl) {
      const qAuthorEl = quoteEl.querySelector('[data-testid="msg-author"], span[dir="auto"]');
      quotedSender = qAuthorEl?.textContent?.trim() || "";
      quotedMessage = quoteEl.innerText?.replace(quotedSender, "")?.trim() || "";
    }

    // Resolve Date & Time
    let date = parsedMeta?.formattedDate || "";
    let time = parsedMeta?.formattedTime || "";
    let timestamp = parsedMeta?.timestamp || null;

    if (!time) {
      const metaTimeEl = node.querySelector('[data-testid="msg-meta"], [data-testid="status-time"]');
      const timeMatch = (metaTimeEl?.textContent || "").match(/(\d{1,2}:\d{2}(?::\d{2})?\s*(?:[ap]\.?m\.?)?)/i);
      if (timeMatch) {
        time = timeMatch[1].trim();
      }
    }

    if (!date && lastKnownDate) {
      date = lastKnownDate;
    }

    if (!timestamp && date && time) {
      const dt = parseDateTime(date, time);
      if (dt) timestamp = dt.getTime();
    }

    // Create unique key for deduplication
    const uniqueKey = dataId || `${date}_${time}_${sender}_${body.slice(0, 40)}`;

    return {
      id: uniqueKey,
      date,
      time,
      timestamp,
      sender,
      type,
      message: body,
      quotedSender,
      quotedMessage,
      isOutgoing
    };
  }

  // Extract all currently visible messages from the DOM
  function extractCurrentDOMMessages(store = new Map()) {
    const main = document.querySelector('#main');
    if (!main) return store;

    let nodes = [...main.querySelectorAll('[data-testid="msg-container"]')];
    if (!nodes.length) {
      nodes = [...main.querySelectorAll('div[data-id]')].filter(el =>
        el.querySelector('.selectable-text, [data-pre-plain-text], [data-testid="selectable-text"], [data-testid="msg-meta"]')
      );
    }
    if (!nodes.length) {
      nodes = [...main.querySelectorAll('.message-in, .message-out')];
    }

    let runningDate = "";

    // Track running date from date separator bubbles if present
    const dateDividers = [...main.querySelectorAll('[data-testid="chat-system-message"], .focusable-list-item')];
    for (const divider of dateDividers) {
      const text = divider.textContent?.trim() || "";
      if (/today|yesterday|\d{1,2}\s+[a-zA-Z]+|\d{1,4}[-/. ]\d{1,2}[-/. ]\d{1,4}/i.test(text)) {
        const d = parseDividerDate(text);
        if (d) runningDate = formatDate(d);
      }
    }

    for (const node of nodes) {
      const msg = parseMessageNode(node, runningDate);
      if (msg) {
        if (msg.date) runningDate = msg.date;
        store.set(msg.id, msg);
      }
    }
    return store;
  }

  function parseDividerDate(text) {
    if (!text) return null;
    const clean = text.trim().toUpperCase();
    const now = new Date();
    if (clean === "TODAY") return new Date(now.getFullYear(), now.getMonth(), now.getDate());
    if (clean === "YESTERDAY") {
      const y = new Date(now);
      y.setDate(now.getDate() - 1);
      return new Date(y.getFullYear(), y.getMonth(), y.getDate());
    }
    const d = new Date(text);
    return isNaN(d.getTime()) ? null : d;
  }

  // Filter messages by start and end timestamps
  function filterMessages(messages, startIso, endIso) {
    const startMs = startIso ? new Date(startIso).getTime() : null;
    const endMs = endIso ? new Date(endIso).getTime() : null;

    return messages.filter(msg => {
      if (!msg.timestamp) return true; // Keep if timestamp couldn't be parsed
      if (startMs !== null && msg.timestamp < startMs) return false;
      if (endMs !== null && msg.timestamp > endMs) return false;
      return true;
    });
  }

  // Safe file download helper
  function downloadBlob(name, blob) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  function sanitizeFilename(name) {
    return (name || "WhatsApp-Export").replace(/[\\/:*?"<>|]+/g, "_").trim().slice(0, 80) || "WhatsApp-Export";
  }

  // Generate Excel (.xlsx) file using SheetJS
  function generateExcelFile(messages, chatName, startIso, endIso) {
    if (typeof XLSX === "undefined") {
      throw new Error("Excel generator library (SheetJS) is not loaded. Refresh WhatsApp Web and try again.");
    }

    const data = messages.map((m, idx) => ({
      "#": idx + 1,
      "Date": m.date || "",
      "Time": m.time || "",
      "Sender": m.sender || "",
      "Type": m.type || "Text",
      "Message": m.message || "",
      "Quoted Sender": m.quotedSender || "",
      "Quoted Message": m.quotedMessage || ""
    }));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);

    // Set column widths
    ws['!cols'] = [
      { wch: 6 },   // #
      { wch: 13 },  // Date
      { wch: 12 },  // Time
      { wch: 22 },  // Sender
      { wch: 13 },  // Type
      { wch: 65 },  // Message
      { wch: 20 },  // Quoted Sender
      { wch: 35 }   // Quoted Message
    ];

    const safeSheetName = (chatName || "Messages")
      .replace(/[\\/?*\[\]:]/g, "")
      .trim()
      .slice(0, 30) || "Chat Export";

    XLSX.utils.book_append_sheet(wb, ws, safeSheetName);

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const dateSuffix = new Date().toISOString().slice(0, 10);
    const filename = `${sanitizeFilename(chatName)}_${dateSuffix}.xlsx`;

    downloadBlob(filename, blob);
    return filename;
  }

  // Generate CSV file fallback
  function generateCsvFile(messages, chatName) {
    function csvEscape(val) {
      const s = String(val ?? "");
      return '"' + s.replace(/"/g, '""') + '"';
    }

    const headers = ["#", "Date", "Time", "Sender", "Type", "Message", "Quoted Sender", "Quoted Message"];
    const rows = messages.map((m, idx) => [
      idx + 1,
      m.date,
      m.time,
      m.sender,
      m.type,
      m.message,
      m.quotedSender,
      m.quotedMessage
    ]);

    const csvContent = [
      headers.map(csvEscape).join(","),
      ...rows.map(r => r.map(csvEscape).join(","))
    ].join("\r\n");

    const blob = new Blob(["\ufeff", csvContent], { type: "text/csv;charset=utf-8" });
    const filename = `${sanitizeFilename(chatName)}_${new Date().toISOString().slice(0, 10)}.csv`;
    downloadBlob(filename, blob);
    return filename;
  }

  // Auto-scroll backward to load chat history up to startTime
  async function autoFetchHistory(startTimeIso, onProgress) {
    const scroller = getChatPanel();
    if (!scroller) throw new Error("Could not locate WhatsApp chat panel. Open a chat and try again.");

    const store = new Map();
    extractCurrentDOMMessages(store);

    const targetStartMs = startTimeIso ? new Date(startTimeIso).getTime() : null;

    let rounds = 0;
    const maxRounds = 350;
    let stableRounds = 0;
    let previousHeight = -1;
    let previousTop = -1;

    isScrollingActive = true;
    shouldStopScroll = false;

    while (rounds < maxRounds && !shouldStopScroll) {
      rounds++;

      // Check earliest timestamp loaded so far
      let earliestMs = Infinity;
      let earliestStr = "";
      for (const msg of store.values()) {
        if (msg.timestamp && msg.timestamp < earliestMs) {
          earliestMs = msg.timestamp;
          earliestStr = `${msg.date} ${msg.time}`;
        }
      }

      // If we've reached messages older than target start time, we can safely stop!
      if (targetStartMs && earliestMs < targetStartMs - 60000) {
        if (onProgress) onProgress({ count: store.size, earliest: earliestStr, reachedTarget: true });
        break;
      }

      if (onProgress && rounds % 2 === 0) {
        onProgress({ count: store.size, earliest: earliestStr, round: rounds });
      }

      // Scroll up toward older messages
      const beforeTop = scroller.scrollTop;
      const beforeHeight = scroller.scrollHeight;

      scroller.scrollTop = 0;
      scroller.dispatchEvent(new Event("scroll", { bubbles: true }));

      await sleep(400);

      extractCurrentDOMMessages(store);

      const afterHeight = scroller.scrollHeight;
      const afterTop = scroller.scrollTop;

      if (afterHeight === previousHeight && afterHeight === beforeHeight && Math.abs(afterTop - beforeTop) < 2) {
        stableRounds++;
        if (stableRounds >= 5) {
          // Reached the top of the chat (no more older messages)
          break;
        }
      } else {
        stableRounds = 0;
      }

      previousHeight = afterHeight;
      previousTop = afterTop;

      if (rounds % 10 === 0) await sleep(300);
    }

    isScrollingActive = false;
    return store;
  }

  // Handle messages from popup
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    (async () => {
      try {
        if (request.action === "get-chat-info") {
          const info = getCurrentChatInfo();
          const store = new Map();
          if (info.isOpen) {
            extractCurrentDOMMessages(store);
          }
          let earliest = "";
          let latest = "";
          const list = [...store.values()].filter(m => m.timestamp).sort((a, b) => a.timestamp - b.timestamp);
          if (list.length > 0) {
            earliest = `${list[0].date} ${list[0].time}`;
            latest = `${list[list.length - 1].date} ${list[list.length - 1].time}`;
          }
          sendResponse({
            ok: true,
            chatName: info.chatName,
            isGroup: info.isGroup,
            isOpen: info.isOpen,
            count: store.size,
            earliest,
            latest
          });

        } else if (request.action === "stop-scroll") {
          shouldStopScroll = true;
          sendResponse({ ok: true, message: "Stopping history scan..." });

        } else if (request.action === "export-excel" || request.action === "export-csv") {
          const info = getCurrentChatInfo();
          if (!info.isOpen) {
            throw new Error("No chat is open. Please click on a WhatsApp group or chat first.");
          }

          let store = new Map();

          if (request.autoScroll) {
            // Auto scroll to load history
            store = await autoFetchHistory(request.startTime, progress => {
              chrome.runtime.sendMessage({
                action: "scan-progress",
                ...progress
              }).catch(() => {});
            });
          } else {
            // Extract currently loaded messages
            store = extractCurrentDOMMessages(new Map());
          }

          const allMessages = [...store.values()].sort((a, b) => (a.timestamp || 0) - (b.timestamp || 0));
          if (!allMessages.length) {
            throw new Error("No messages found in the loaded chat. Make sure the conversation pane is visible.");
          }

          // Filter by Start Time and End Time
          const filtered = filterMessages(allMessages, request.startTime, request.endTime);

          if (!filtered.length) {
            throw new Error(`Found ${allMessages.length} messages, but 0 matched the date/time filter. Try broadening the date range.`);
          }

          let filename = "";
          if (request.action === "export-excel") {
            filename = generateExcelFile(filtered, info.chatName, request.startTime, request.endTime);
          } else {
            filename = generateCsvFile(filtered, info.chatName);
          }

          sendResponse({
            ok: true,
            count: filtered.length,
            totalFound: allMessages.length,
            filename,
            chatName: info.chatName
          });

        } else {
          sendResponse({ ok: false, error: "Unknown action." });
        }
      } catch (err) {
        sendResponse({ ok: false, error: err.message || String(err) });
      }
    })();
    return true; // Keep message channel open for async response
  });
})();