🧠 Likhai Assistant for Google Sheets

A powerful Google Sheets add-on that brings generative AI capabilities directly into your spreadsheet using a chat sidebar and in-sheet functions. Likhai Assistant helps you analyze tables, classify support tickets, auto-generate dashboards, and more using GPT-4o-mini through the LiteLLM proxy.

📦 Features

💬 Chat Sidebar with context-aware responses

📄 =LIKHAIFORMULA(...) for AI-powered in-sheet insights

📊 Auto-generated Dashboards from selected data

🏷️ Keyword-based Ticket Classifier (Product, Category, etc.)

📋 Markdown Table Generator from plain prompts

📜 Interaction Logging to the "Likhai Logs" sheet

🔐 API key-based secure call to LiteLLM endpoint

🚀 Getting Started

1. Setup

Open your Google Apps Script Project

Deploy or link to this sample Sheet

2. File Structure

Code.gs: All backend logic

chat.html: Sidebar chat UI with styling and chat logic

3. Functions Overview

➤ LIKHAI(messages [, forceJson])

Sends chat messages to the LLM and returns the assistant's response.

➤ LIKHAIFORMULA(prompt, range)

Formula-style function for spreadsheet users to analyze selected table data using prompts.

➤ processLikhaiQuery(query)

Handles chat sidebar interaction using context and history.

➤ classifyTickets()

Classifies selected ticket rows into product, category, subcategory, etc., using hardcoded keyword rules.

➤ createTableFromUserPrompt()

Turns a plain-language prompt into a well-formatted Markdown table inserted into a new sheet.

➤ buildVideoStyleDashboard()

Generates a modern dashboard layout using selected table data and AI-planned visualization strategy.

➤ showLikhaiSidebar()

Displays the chat-based assistant in the Google Sheets sidebar.

➤ onOpen()

Registers the Likhai Assistant menu in the Sheet UI.

💬 Chat Sidebar UI

Built using chat.html

Automatically maintains chat history and displays alternating roles

Responsive buttons and loading indicators

⚙️ Configuration

Update the following line with your LiteLLM API Key:

const LITE_LLM_KEY = "sk-xxxx";

And optionally your LiteLLM endpoint:

const LITE_LLM_ENDPOINT = "https://your-litellm-domain/chat/completions";

📓 Example Use Cases

Ask: "Create a table of monthly goals for Q3 2025"

Highlight a support ticket row → classify it via classifyTickets()

Select table + =LIKHAIFORMULA("Give top 3 insights", A1:D10)

Click "Build Dashboard" from menu to visualize KPIs

📁 Logs

All interactions with the model (queries + responses) are saved in the Likhai Logs sheet for later review.

🛠️ Contributions

Feel free to fork and adapt this for your use-case. Recommended improvements:

Add model switcher

Implement persistent sidebar memory

Add rich chart types (bar, scatter, etc.)

🧪 License

This project is for internal prototype and educational use. Not affiliated with OpenAI or Google. Use at your own risk.

🙏 Credits

Developed by [Akash Rathi] using Google Apps Script, OpenAI, and LiteLLM.
