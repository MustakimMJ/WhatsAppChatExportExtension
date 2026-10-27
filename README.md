# WhatsApp Group & Chat Exporter to Excel (.xlsx)

A powerful, privacy-first Manifest V3 Chrome extension designed to export WhatsApp Web group chats and direct messages to real Excel (`.xlsx`) spreadsheets and CSV with full metadata.

## ✨ Features

- **📊 Native Excel (.xlsx) Export**: Exports directly to formatted Microsoft Excel `.xlsx` spreadsheets (powered by SheetJS) with clean columns, auto-adjusted column widths, and proper cell types.
- **👥 Full Group Chat Support**: Accurately extracts **Sender Name** or Phone Number for every participant in WhatsApp groups, identifies your own outgoing messages as `"You"`, and flags system notices.
- **🕒 Date & Time Filtering**:
  - **Starting Time (with date)**: Filter messages sent on or after a specific date and time.
  - **Ending Time (with date)**: Filter messages sent on or before a specific date and time.
  - **⏱️ "Now" Button**: 1-click option to instantly set the Ending Time to the current local date and time.
  - **⚡ Quick Presets**: Instant 1-click range buttons: **Past 1h**, **Today (from 00:00)**, **Past 24h**, **Past 7d**, and **All Messages**.
- **🔄 Smart History Auto-Scroller**:
  - Automatically scrolls upward in WhatsApp Web to fetch past chat history until your chosen **Start Time** is reached.
  - Continuously accumulates and deduplicates virtualized messages without losing loaded history.
  - Live progress display with an instant **"Stop & Export"** option.
- **📋 Rich Exported Columns**:
  1. `#` (Index)
  2. `Date` (YYYY-MM-DD)
  3. `Time` (HH:MM:SS)
  4. `Sender` (Group member name, contact name, phone number, or "You")
  5. `Type` (Text, Image, Video, Voice Note / Audio, Document, Sticker, Location, Contact, Poll, Deleted)
  6. `Message` (Full text content, media caption, or document title)
  7. `Quoted Sender` (Participant being replied to)
  8. `Quoted Message` (Snippet of the quoted message)
- **🔒 100% Private & Local**: Runs completely offline in your browser. No third-party servers, tracking, or data collection.

---

## 🚀 How to Install in Chrome / Brave / Edge

1. Open your browser and navigate to `chrome://extensions` (or `edge://extensions`).
2. Toggle **Developer mode** on (top-right corner).
3. Click **Load unpacked** (top-left).
4. Select this extension directory:
   `c:\Users\Lenovo\Desktop\Whatsapp Extension`
5. The extension **"WhatsApp Group & Chat Exporter to Excel"** will appear in your extensions list!

---

## 📖 How to Use

1. Open **[WhatsApp Web](https://web.whatsapp.com)** in your browser and log in.
2. Click on the group chat or individual conversation you want to export.
3. Click the extension icon in your browser toolbar (you can pin it for convenience).
4. In the extension popup:
   - Verify the active chat title displayed at the top.
   - (Optional) Set your **Starting Time** and **Ending Time**:
     - Click **⏱️ Now** to automatically set the Ending Time to the current moment.
     - Or click one of the quick presets like **Today** or **Past 24h**.
5. Choose your export method:
   - **📊 Export to Excel (.xlsx)**: Instantly exports currently loaded messages matching your date range.
   - **🔄 Auto-Scroll & Export to Excel**: Intelligently scrolls backward in time to load older history until your chosen Starting Time is reached, then downloads the `.xlsx` file.
6. The Excel `.xlsx` file will be downloaded directly to your computer!

---

## 📁 Technical Overview

- `manifest.json`: Manifest V3 specification with `activeTab`, `downloads`, and `scripting` permissions.
- `content.js`: Injected into WhatsApp Web to parse group sender names, DOM metadata (`data-pre-plain-text`), timestamps, and manage history auto-scrolling.
- `popup.html` & `popup.js`: Modern user interface with datetime-local pickers, "Now" button, preset pills, live progress tracking, and SheetJS integration.
- `xlsx.full.min.js`: Bundled standalone SheetJS library for generating native Excel files client-side.
