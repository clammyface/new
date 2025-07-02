# 🧠 Likhai Assistant for Google Sheets

Elevate your spreadsheet with the power of GPT. **Likhai Assistant** is a smart Google Sheets add-on that transforms your workflow through AI-driven insights, contextual chat, table generation, and visual dashboards—seamlessly embedded into your familiar Sheets interface.

---

## 🌟 Key Features

🔹 **Smart Chat Sidebar**

> Ask anything. Get answers with spreadsheet context. AI remembers your conversation.

🔹 **Inline AI Formula**

> Use `=LIKHAIFORMULA(...)` to extract insights, classify rows, or analyze patterns.

🔹 **One-Click Dashboards**

> Turn raw data into clean, modern dashboards with KPIs, charts, and summaries.

🔹 **Support Ticket Classifier**

> Classify product issues with rule-based logic (Product, Category, Subcategory).

🔹 **Markdown Table Generator**

> Generate full tables from plain-text prompts.

🔹 **Query Logs**

> All chat queries and responses are saved in a `Likhai Logs` sheet.

🔒 **Secure API Calls**

> Integrated with [LiteLLM](https://docs.litellm.ai/) using your own key and endpoint.

---

## 🚀 Getting Started

### 1. Setup

* Open the [Apps Script Project](https://script.google.com/u/0/home/projects/1XVMSvdFSd5RZItUqSECZuv6k3Ub1Ky-N-c0sQD9OHJsGMckGXRV5Sq90/edit)
* Link to your copy of [this sample Google Sheet](https://docs.google.com/spreadsheets/d/1p71o9dSEz51RxC9H6aoyByBA16OlDjIdwnRS0-AdedI/edit?gid=243015641#gid=243015641)

### 2. File Structure

```
├── Code.gs        # All backend logic & functions
└── chat.html      # Clean chat interface with loading, memory & input controls
```

### 3. Feature Overview

* `LIKHAI(messages [, forceJson])`: Core message handler to LiteLLM
* `LIKHAIFORMULA(prompt, range)`: Formula-style data analysis
* `processLikhaiQuery(query)`: Chat sidebar logic and history management
* `classifyTickets()`: Rule-based ticket classification
* `createTableFromUserPrompt()`: Markdown table to Sheet
* `buildVideoStyleDashboard()`: Layout plan + chart insertion
* `onOpen()` / `showLikhaiSidebar()`: UI menu + sidebar trigger

---

## 💻 Sidebar Chat UI

Elegant, fast, and context-aware:

* Memory of last 10 interactions
* Auto scroll + clean assistant vs user bubble display
* Loading spinner and disabled state for async handling

---

## ⚙️ Configuration

Set your LiteLLM API Key and endpoint in `Code.gs`:

```js
const LITE_LLM_KEY = "sk-xxxx";
const LITE_LLM_ENDPOINT = "https://your-litellm-endpoint/chat/completions";
```

---

## 💡 Example Use Cases

* 💬 "Summarize this table into 3 bullet insights"
* 📑 `=LIKHAIFORMULA("Identify trends", A1:F10)`
* 🛠️ Run `classifyTickets()` on support logs
* 📊 Auto-build a dashboard from your selected sheet
* 📋 Convert "Plan a 4-week study schedule" into a usable table

---

## 📁 Audit Trail & Logging

All prompts and responses are saved in the `Likhai Logs` tab for audit and tracking.

---

## 🔧 Contributing & Extensions

This is a prototype framework. Ideas to extend:

* Switch between GPT models
* Enable persistent chat history
* Add advanced chart types or filters
* Embed auto-refresh functionality

PRs and forks welcome.

---

## 📜 License

MIT-style. Use freely for educational, internal, or experimental purposes. Not affiliated with Google or OpenAI.

---

## 🙌 Credits

Built by \[Akash Rathi] using Google Apps Script, OpenAI APIs, and LiteLLM proxy.
