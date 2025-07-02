// ========== Code.gs ========== //

const LITE_LLM_ENDPOINT = "https://dev-litellm.leadschool.in/chat/completions";
const LITE_LLM_KEY = "sk-5BWbp70TbFmOu8SXa9VzFQ";

function LIKHAI(messages, forceJson = false) {
  if (!Array.isArray(messages) || messages.length === 0 || !messages[0].content) {
    return "❌ Error: Prompt is empty or invalid.";
  }

  const payload = {
    model: "gpt-4o-mini",
    messages: messages.map(msg => ({
      role: msg.role,
      content: msg.content || "Hello"
    })),
    temperature: 0.7
  };

  if (forceJson) {
    payload.response_format = { "type": "json_object" };
  }

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${LITE_LLM_KEY}`
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(LITE_LLM_ENDPOINT, options);
    const jsonText = response.getContentText();
    Logger.log("Raw Response: " + jsonText);

    const json = JSON.parse(jsonText);
    if (!json.choices || !json.choices[0] || !json.choices[0].message) {
      const errorContent = `{"error": "Unexpected response structure: ${jsonText}"}`;
      return forceJson ? errorContent : "❌ Error: Unexpected response structure.";
    }
    
    return json.choices[0].message.content.trim();
  } catch (e) {
    const errorContent = `{"error": "Failed to fetch or parse AI response: ${e.message}"}`;
    return forceJson ? errorContent : `❌ Error: ${e.message}`;
  }
}

function classifyTickets() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Prompt user to specify classification columns
  var productColumn = Browser.inputBox('Enter the column name for Product (e.g., GPT Product)');
  var categoryColumn = Browser.inputBox('Enter the column name for Category (e.g., GPT Category)');
  var subcategoryColumn = Browser.inputBox('Enter the column name for Subcategory (e.g., GPT Subcategory)');
  var statusColumn = Browser.inputBox('Enter the column name for Status (e.g., GPT Status)');
  var notesColumn = Browser.inputBox('Enter the column name for Notes (e.g., GPT Notes)');
  
  // Define the mapping of keywords to products
  var products = {
    "Teacher App": ["Dayplan", "Resources", "SD Card", "SD cards not detecting", "Content", "Login", "Device", "Device lock", "One Click", "Marks Entry", "Remediation", "Attendance", "Casting"],
    "ASM": ["ASM", "Assessments", "delivery", "Mistakes", "Service Requests", "Question Paper (QP)", "answer key(AK)", "Builder", "Generator"],
    "Nucleus": ["ELGA Allocation", "Student Data", "Class allocation", "Teacher allocation", "Board Change", "Implementation", "Teacher timelyness", "Report", "Report card", "Login Credentials", "forgot password", "Promotions", "Student detail", "Teacher detail"],
    "Student App": ["Parent App Login", "Student APP login", "Homework", "Quizzes"]
  };
  
  // Categories and Subcategories
  var categories = {
    "Teacher App": ["Queries", "Requests", "Issues"],
    "ASM": ["Queries", "Requests", "Issues"],
    "Nucleus": ["Queries", "Requests", "Issues"],
    "K12 / Pearson": ["Queries", "Requests", "Issues"]
  };

  // Get selected rows
  var range = sheet.getActiveRange();
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    var ticketText = values[i][0]; // Assuming ticket text is in the first column
    var matchedProduct = "Others";
    var matchedCategory = "";
    var matchedSubcategory = "";
    
    // Step 1: Classify Product based on keywords
    for (var product in products) {
      for (var keyword of products[product]) {
        if (ticketText.includes(keyword)) {
          matchedProduct = product;
          break;
        }
      }
    }

    // Step 2: Classify Category and Subcategory
    if (matchedProduct in categories) {
      matchedCategory = categories[matchedProduct][0]; // Default to Queries if we don't specify
      matchedSubcategory = categories[matchedProduct][0]; // Default to Queries if no specific subcategory is needed
    }
    
    // Step 3: Update columns
    var productIndex = sheet.getRange(1, sheet.getRange(productColumn + '1').getColumn()).getColumn();
    var categoryIndex = sheet.getRange(1, sheet.getRange(categoryColumn + '1').getColumn()).getColumn();
    var subcategoryIndex = sheet.getRange(1, sheet.getRange(subcategoryColumn + '1').getColumn()).getColumn();
    var statusIndex = sheet.getRange(1, sheet.getRange(statusColumn + '1').getColumn()).getColumn();
    var notesIndex = sheet.getRange(1, sheet.getRange(notesColumn + '1').getColumn()).getColumn();
    
    // Set the mapped values to the respective columns
    sheet.getRange(i + 2, productIndex).setValue(matchedProduct);
    sheet.getRange(i + 2, categoryIndex).setValue(matchedCategory);
    sheet.getRange(i + 2, subcategoryIndex).setValue(matchedSubcategory);
    sheet.getRange(i + 2, statusIndex).setValue("Open"); // You can add logic here if you need dynamic status
    sheet.getRange(i + 2, notesIndex).setValue("Classified based on keywords");
  }
}


function processLikhaiQuery(query) {
  Logger.log("Query received: " + query);
  if (typeof query !== "string" || query.trim().length < 2) {
    return "❌ Error: Please enter a valid prompt.";
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  const tableText = values.map(row => row.join('\t')).join('\n');

  let history = PropertiesService.getUserProperties().getProperty("chatHistory");
  history = history ? JSON.parse(history) : [];

  const contextIntro = "Here is the selected table data:\n" + tableText;
  history.push({ role: "user", content: `${query}\n\n${contextIntro}` });
  if (history.length > 10) history.shift();
  
  const response = LIKHAI(history);
  history.push({ role: "assistant", content: response });
  PropertiesService.getUserProperties().setProperty("chatHistory", JSON.stringify(history));
  logInteraction(query, response);
  return response;
}

function logInteraction(userPrompt, aiResponse) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Likhai Logs");
  if (!logSheet) logSheet = ss.insertSheet("Likhai Logs");
  logSheet.appendRow([new Date(), userPrompt, aiResponse]);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧠 Likhai Assistant')
    .addItem('Open Chat Sidebar', 'showLikhaiSidebar')
    .addItem('Generate Table from Prompt', 'createTableFromUserPrompt')
    .addItem('Create Modern Dashboard', 'buildVideoStyleDashboard')
    .addToUi();
}

function showLikhaiSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("chat").setTitle("Likhai Assistant");
  SpreadsheetApp.getUi().showSidebar(html);
}

function createTableFromUserPrompt() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Enter your query for a table (e.g., "6-week study plan")').getResponseText();
  if (!prompt || prompt.trim() === "") {
    ui.alert("❌ Prompt is required.");
    return;
  }
  const sheetNamePrompt = ui.prompt('Enter a name for the new sheet where the table should go (e.g., "Study Plan")');
  const sheetName = sheetNamePrompt.getResponseText().trim();
  if (!sheetName) {
    ui.alert("❌ Sheet name is required.");
    return;
  }
  const messages = [
    { role: "system", content: "You are a table generator. Respond only in Markdown table format with proper headers and rows." },
    { role: "user", content: prompt }
  ];
  
  const markdownTable = LIKHAI(messages);
  Logger.log("GPT Table Output:\n" + markdownTable);

  const lines = markdownTable.trim().split('\n');
  if (lines.length < 3) {
    ui.alert("⚠️ Could not extract a valid table from the AI response.");
    return;
  }

  const header = lines[0];
  const dataLines = lines.slice(2);
  const rows = [header, ...dataLines].map(row =>
    row.split('|').slice(1, -1).map(cell => cell.trim())
  );

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = spreadsheet.getSheetByName(sheetName);
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  const newSheet = spreadsheet.insertSheet(sheetName);
  newSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}

function LIKHAIFORMULA(prompt, range) {
  if (!prompt || typeof prompt !== "string" || prompt.trim().length === 0) {
    return "❌ Invalid prompt.";
  }

  // Safely parse range content
  let tableData = "No data selected.";
  if (range && typeof range.getValues === "function") {
    const values = range.getValues();
    tableData = values.map(row => row.join('\t')).join('\n');
  } else if (Array.isArray(range)) {
    tableData = range.map(row => row.join('\t')).join('\n');
  }

  const messages = [
    {
      role: "system",
      content: "You are a helpful assistant analyzing spreadsheet data."
    },
    {
      role: "user",
      content: `${prompt}\n\nHere is the selected table data:\n${tableData}`
    }
  ];

  return LIKHAI(messages);
}


// ========================================================================= //
// === Likhai Assistant: Smart Dashboard Builder (VERIFIED FINAL VERSION) === //
// ========================================================================= //
function buildVideoStyleDashboard() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getActiveSheet();

  if (sourceSheet.getName() === "Dashboard" && sourceSheet.getLastRow() < 2) {
    ui.alert("Please select a sheet with real data before generating the dashboard.");
    return;
  }

  const dataRange = sourceSheet.getDataRange();
  if (!dataRange || dataRange.isBlank()) {
    ui.alert("❌ The active sheet is empty.");
    return;
  }

  const promptResponse = ui.prompt("📌 Dashboard Prompt", "Enter instructions for dashboard (e.g., 'show subject summary and exclude names')", ui.ButtonSet.OK_CANCEL);
  const userPrompt = promptResponse.getResponseText()?.trim();
  if (!userPrompt) return ui.alert("❌ Prompt is required.");

  const dashboardSheetName = "Dashboard";
  let sheet = spreadsheet.getSheetByName(dashboardSheetName);
  if (sheet) spreadsheet.deleteSheet(sheet);
  sheet = spreadsheet.insertSheet(dashboardSheetName, 0);
  spreadsheet.setActiveSheet(sheet);

  const rawValues = dataRange.getValues();
  const headers = rawValues[0];
  const dataRows = rawValues.slice(1).filter(row => row[0] && row[0].toString().trim() !== ""); // skip empty rows
  const tableText = [headers, ...dataRows].map(r => r.join('\t')).join('\n');

  // === Compute all numeric KPIs (filtered rows only) ===
  const numericKPIs = [];
  headers.forEach((header, colIndex) => {
    if (typeof header !== 'string') return;
    const cleanedHeader = header.trim();
    if (!cleanedHeader || colIndex === 0) return;

    let sum = 0;
    let count = 0;

    dataRows.forEach(row => {
      const val = parseFloat(row[colIndex]);
      if (!isNaN(val)) {
        sum += val;
        count++;
      }
    });

    if (count > 0) {
      let value = sum.toLocaleString();
      if (cleanedHeader.toLowerCase().includes("average") || cleanedHeader.toLowerCase().includes("score")) {
        value = (sum / count).toFixed(1);
      }

      numericKPIs.push({
        label: cleanedHeader,
        value: value,
        icon: "📈"
      });
    }
  });

  // === Ask GPT for layout plan (for chart layout only) ===
  const messages = [{
    role: "system",
    content: `You are a dashboard generator AI. Given data and user prompt, return JSON with:
{
  "kpis": [ {"label": "...", "value": "...", "icon": "..."}, ... ],
  "lineChart": {"title": "...", "data": [["X", "Y1", "Y2"], ["Jan", 10, 20]]},
  "taskChart": {"title": "...", "data": [["Status", "Count"], ["Done", 60], ["Pending", 40]]}
}`
  }, {
    role: "user",
    content: `${userPrompt}\n\nHere is the selected data:\n${tableText}`
  }];

  let plan;
  try {
    const json = LIKHAI(messages, true);
    Logger.log("🎨 Dashboard Plan:\n" + json);
    plan = JSON.parse(json);
    if (plan.error) throw new Error(plan.error);
  } catch (e) {
    ui.alert("Failed to generate dashboard plan:\n\n" + e.message);
    return;
  }

  // === Remove GPT duplicates from KPI list
  const existingLabels = new Set(numericKPIs.map(kpi => kpi.label.trim().toLowerCase()));
  const filteredGPTKpis = (plan.kpis || []).filter(kpi => {
    const label = kpi.label?.trim().toLowerCase();
    return label && !existingLabels.has(label);
  });
  const kpis = [...numericKPIs, ...filteredGPTKpis];

  // === Sheet Styling ===
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.getRange("A1:Z50").setBackground("#f5f7fa");

  // === Sidebar ===
  sheet.getRange("A1:B30").setBackground("#ffffff");
  sheet.getRange("A1:B1").merge().setValue("🔷 Admin Panel").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

  const menu = ["📊 Dashboard", "👤 User", "📬 Inbox", "📅 Calendar", "📁 Files"];
  menu.forEach((label, i) => sheet.getRange(i + 3, 1, 1, 2).merge().setValue(label).setFontSize(10));

  // === KPI Cards ===
  kpis.forEach((kpi, i) => {
    const col = 4 + (i * 4);
    sheet.getRange(2, col, 4, 3).setBackground("#ffffff").setBorder(true, true, true, true, true, true);
    sheet.getRange(3, col, 1, 3).merge().setValue(kpi.icon || "❓").setFontSize(20).setHorizontalAlignment("center");
    sheet.getRange(4, col, 1, 3).merge().setValue(kpi.value || "").setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center");
    sheet.getRange(5, col, 1, 3).merge().setValue(kpi.label || "").setFontSize(9).setHorizontalAlignment("center");
  });

  // === Line Chart ===
  const lineChartData = plan.lineChart?.data;
  if (lineChartData && lineChartData.length > 1) {
    const tempLineSheet = spreadsheet.insertSheet("line_temp_" + Date.now());
    tempLineSheet.getRange(1, 1, lineChartData.length, lineChartData[0].length).setValues(lineChartData);
    const lineChartRange = tempLineSheet.getDataRange();

    const lineChart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(lineChartRange)
      .setPosition(8, 4, 0, 0)
      .setOption("title", plan.lineChart?.title || "Trends")
      .setOption("legend", { position: "top" })
      .setOption("width", 700)
      .setOption("height", 300)
      .build();

    sheet.insertChart(lineChart);
    tempLineSheet.hideSheet();
  }

  // === Pie Chart (Open vs Closed Tickets - filtered) ===
  const openCol = headers.findIndex(h => h.toLowerCase().includes("open ticket"));
  const closedCol = headers.findIndex(h => h.toLowerCase().includes("closed ticket"));

  let openSum = 0;
  let closedSum = 0;

  dataRows.forEach(row => {
    const openVal = parseFloat(row[openCol]);
    const closedVal = parseFloat(row[closedCol]);
    if (!isNaN(openVal)) openSum += openVal;
    if (!isNaN(closedVal)) closedSum += closedVal;
  });

  const pieData = [
    ["Status", "Count"],
    ["🟥 Open Tickets", openSum],
    ["🟩 Closed Tickets", closedSum]
  ];

  const tempPieSheet = spreadsheet.insertSheet("pie_temp_" + Date.now());
  tempPieSheet.getRange(1, 1, pieData.length, 2).setValues(pieData);
  const chartRange = tempPieSheet.getDataRange();

  const pieChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(chartRange)
    .setPosition(8, 14, 0, 0)
    .setOption("title", "Ticket Status Overview")
    .setOption("pieHole", 0.6)
    .setOption("legend", { position: "right", textStyle: { fontSize: 10 } })
    .setOption("slices", {
      0: { color: '#ea4335' },
      1: { color: '#34a853' }
    })
    .setOption("width", 350)
    .setOption("height", 300)
    .build();

  sheet.insertChart(pieChart);
  tempPieSheet.hideSheet();

  ui.alert("✅ Dashboard generated on the new 'Dashboard' sheet.");
}
