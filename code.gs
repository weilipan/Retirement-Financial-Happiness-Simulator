
/**
 * 在 Google Sheets 建立自訂功能表
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("退休模擬工具")
    .addItem("📘 模擬簡化版（以儲蓄為主）", "simulateSimplifiedSavingPlan")
    .addItem("🧮 自動反推年投資金額（維持正值）", "findMinPositiveInvestment")
    .addItem("📝 產出報告並寄送","generateAndSendRetirementReport")
    .addToUi();
}

/**
 * 模擬退休財務計畫（以儲蓄/投資為主，支出已內含於儲蓄計算中）
 */

function simulateSimplifiedSavingPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("輸入資料");
  let resultSheet = ss.getSheetByName("模擬結果");

  if (!resultSheet) {
    resultSheet = ss.insertSheet("模擬結果");
  } else {
    resultSheet.clear();
    const charts = resultSheet.getCharts();
    charts.forEach(chart => resultSheet.removeChart(chart)); // 清除圖表
  }

  const getVal = (label) => {
    const range = inputSheet.getRange("A1:A100").getValues();
    for (let i = 0; i < range.length; i++) {
      if (range[i][0].toString().trim() === label) {
        return inputSheet.getRange(i + 1, 2).getValue();
      }
    }
    return null;
  };

  const currentAge = parseInt(getVal("目前年齡"));
  const retireAge = parseInt(getVal("預計退休年齡"));
  const lifespan = parseInt(getVal("預期壽命"));
  const annualSaving = parseFloat(getVal("每年儲蓄金額（不參與投資）")) || 0;
  const annualInvest = parseFloat(getVal("每年投入投資金額（參與複利）")) || 0;
  const retiredInvest = parseFloat(getVal("退休後每年持續投入投資金額")) || 0;
  const initialCash = parseFloat(getVal("現金儲蓄總額")) || 0;
  const hasPension = getVal("是否有退休金（Y/N）")?.toString().toLowerCase() === "y";
  const pension = hasPension ? parseFloat(getVal("每年退休金金額")) || 0 : 0;
  const returnRate = parseFloat(getVal("預期年投資報酬率（%）")) / 100;
  const inflRate = parseFloat(getVal("支出年成長率（%）")) / 100;
  const retireExpense = parseFloat(getVal("預期退休後每年支出")) || 0;

  let year = new Date().getFullYear();
  let age = currentAge;
  let cashAsset = initialCash;
  let investment = 0;
  let investmentPrincipal = 0;
  let breakYear = null;
  let peakAsset = 0;
  let peakYear = year;
  let totalWithdrawFromInvestment = 0;

  resultSheet.appendRow([
    "年份", "年齡", "非投資資產", "投資本金", "投資收益", "總資產",
    "收入", "支出", "實際動用資產金額", "年度淨變化",
    "當年動用投資收益", "累積動用投資收益"
  ]);

  while (age <= lifespan) {
    investment *= (1 + returnRate);

    if (age < retireAge) {
      investment += annualInvest;
      investmentPrincipal += annualInvest;
      cashAsset += annualSaving;
    } else {
      investment += retiredInvest;
      investmentPrincipal += retiredInvest;
    }

    const income = age < retireAge ? 0 : pension;
    const expense = age < retireAge ? 0 : retireExpense * Math.pow(1 + inflRate, age - retireAge);
    cashAsset += income - expense;

    let withdrawFromInvestment = 0;

    if (cashAsset < 0) {
      const deficit = Math.abs(cashAsset);
      const available = investment - investmentPrincipal;
      withdrawFromInvestment = Math.min(deficit, available);
      investment -= withdrawFromInvestment;
      cashAsset += withdrawFromInvestment;
      totalWithdrawFromInvestment += withdrawFromInvestment;
    }

    const investmentProfit = investment - investmentPrincipal;
    const totalAsset = cashAsset + investment;
    const withdraw = Math.max(0, expense - income);
    const net = (age < retireAge ? annualSaving + annualInvest : retiredInvest) + income - expense;

    if (totalAsset > peakAsset) {
      peakAsset = totalAsset;
      peakYear = year;
    }

    if (totalAsset < 0 && breakYear === null) {
      breakYear = year;
    }

    resultSheet.appendRow([
      year, age,
      Math.round(cashAsset),
      Math.round(investmentPrincipal),
      Math.round(investmentProfit),
      Math.round(totalAsset),
      Math.round(income),
      Math.round(expense),
      Math.round(withdraw),
      Math.round(net),
      Math.round(withdrawFromInvestment),
      Math.round(totalWithdrawFromInvestment)
    ]);

    age++;
    year++;
  }

  drawChart(resultSheet);
  writeSummary(resultSheet, breakYear, peakAsset, peakYear);
  formatSimulationResults();
  SpreadsheetApp.getUi().alert("模擬完成（已自動清除舊資料與圖表），請查看「模擬結果」。");
}

function formatSimulationResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("模擬結果");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const numberCols = [3, 4, 5, 6, 7, 8, 9, 10];

  numberCols.forEach(col => {
    const range = sheet.getRange(2, col, lastRow - 1);
    range.setNumberFormat("#,##0");

    const values = range.getValues();
    const backgrounds = range.getBackgrounds();

    for (let i = 0; i < values.length; i++) {
      backgrounds[i][0] = values[i][0] < 0 ? "#ffe6e6" : "white";
    }
    range.setBackgrounds(backgrounds);
  });

  const titleRange = sheet.getRange("A1:J1");
  titleRange.setFontWeight("bold").setBackground("#f1f1f1");
}


/**
 * 繪製資產變化圖表
 */
function drawChart(sheet) {
  const lastRow = sheet.getLastRow();

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange("A2:A" + lastRow)) // 年份
    .addRange(sheet.getRange("C2:C" + lastRow)) // 非投資資產
    .addRange(sheet.getRange("D2:D" + lastRow)) // 投資本金
    .addRange(sheet.getRange("E2:E" + lastRow)) // 投資收益
    .addRange(sheet.getRange("F2:F" + lastRow)) // 總資產
    .addRange(sheet.getRange("L2:L" + lastRow)) // 當年動用投資收益
    .addRange(sheet.getRange("M2:M" + lastRow)) // 累積動用投資收益
    .setPosition(2, 12, 0, 0)
    .setOption("title", "資產變化趨勢與投資收益動用")
    .setOption("hAxis", { title: "年份" })
    .setOption("vAxis", { title: "金額" })
    .setOption("legend", { position: "right" })
    .setOption("series", {
      0: { labelInLegend: "非投資資產" },
      1: { labelInLegend: "投資本金" },
      2: { labelInLegend: "投資收益" },
      3: { labelInLegend: "總資產" },
      4: { labelInLegend: "當年動用投資收益" },
      5: { labelInLegend: "累積動用投資收益" }
    })
    .build();

  sheet.insertChart(chart);
}

function writeSummary(sheet, breakYear, peakAsset, peakYear) {
  const row = 2;
  const col = 20;

  const summary = breakYear
    ? `⚠️ 注意：你的資產將在 ${breakYear} 年出現負值，可能無法支撐至壽命終點。`
    : `✅ 恭喜！你的資產可支撐至預期壽命，且最高點出現在 ${peakYear} 年，資產約為 NT$${Math.round(peakAsset).toLocaleString()}`;

  const lastRow = sheet.getLastRow();
  const cumulativeWithdraw = sheet.getRange(lastRow, 12).getValue(); // 第 13 欄是累積動用投資收益

  // 掃描退休後每年現金佔總資產比
  const startRow = 2;
  const cashRatios = [];
  for (let i = startRow; i <= lastRow; i++) {
    const age = sheet.getRange(i, 2).getValue();
    const cash = sheet.getRange(i, 3).getValue();
    const total = sheet.getRange(i, 6).getValue();
    if (age >= getRetireAgeFromInput_()) {
      const ratio = total > 0 ? cash / total : 0;
      cashRatios.push(ratio);
    }
  }

  const overcashCount = cashRatios.filter(r => r >= 0.6).length;

  const tips = [
    "📌 若「累積動用投資收益」數字過高，代表現金流不足，可能需降低支出或增加投資報酬率。",
    "📈 每年動用投資收益應維持穩定，若呈快速上升，恐將影響長期投資資產。",
    `💰 模擬結束時「累積動用投資收益」總額為：NT$${Math.round(cumulativeWithdraw).toLocaleString()}`,
    "🧾 若此值趨近投資收益總額，建議保守調整退休支出或延後退休時間。"
  ];

  if (overcashCount > 3) {
    tips.push("📊 提醒：退休後現金資產占比過高，建議適度提高投資比例以優化資產運用效率。");
  }

  sheet.getRange(row, col).setValue("📊 模擬總結");
  sheet.getRange(row + 1, col).setValue(summary);
  sheet.getRange(row + 3, col).setValue("🧠 財務診斷建議");
  for (let i = 0; i < tips.length; i++) {
    sheet.getRange(row + 4 + i, col).setValue(tips[i]);
  }
}

// 從輸入資料中取得退休年齡
function getRetireAgeFromInput_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("輸入資料");
  const values = sheet.getRange("A1:A100").getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "預計退休年齡") {
      return parseInt(sheet.getRange(i + 1, 2).getValue());
    }
  }
  return 65;
}



/**
 * 自動反推最低年投資金額，使模擬過程中年度淨變化皆為正值
 */
/**
 * 修改後的反推函式：總資產不曾低於 0
 */
function findMinPositiveInvestment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("輸入資料");

  const getVal = (label) => {
    const range = sheet.getRange("A1:A100").getValues();
    for (let i = 0; i < range.length; i++) {
      if (range[i][0] === label) {
        return sheet.getRange(i + 1, 2).getValue();
      }
    }
    return null;
  };

  const setVal = (label, value) => {
    const range = sheet.getRange("A1:A100").getValues();
    for (let i = 0; i < range.length; i++) {
      if (range[i][0] === label) {
        sheet.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
  };

  const originalValue = parseFloat(getVal("每年投入投資金額（參與複利）")) || 0;

  let low = 0;
  let high = 2000000;
  let answer = -1;

  while (low <= high) {
    let mid = Math.floor((low + high) / 2);
    setVal("每年投入投資金額（參與複利）", mid);
    const result = simulateAndCheckNoBankruptcy(); // use new logic
    if (result) {
      answer = mid;
      high = mid - 1;
    } else {
      low = mid + 1;
    }
  }

  if (answer !== -1) {
    setVal("每年投入投資金額（參與複利）", answer);
    SpreadsheetApp.getUi().alert("✅ 自動反推完成（依總資產是否破產）：最低每年投資金額為 NT$" + answer.toLocaleString() + "\n請重新執行模擬查看詳情。");
  } else {
    SpreadsheetApp.getUi().alert("❌ 無法找到符合條件的金額（總資產仍會歸零）。");
    setVal("每年投入投資金額（參與複利）", originalValue);
  }
}

/**
 * 判斷資產是否在存活期內從未破產（總資產 >= 0）
 */
function simulateAndCheckNoBankruptcy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("輸入資料");

  const getVal = (label) => {
    const range = sheet.getRange("A1:A100").getValues();
    for (let i = 0; i < range.length; i++) {
      if (range[i][0] === label) {
        return sheet.getRange(i + 1, 2).getValue();
      }
    }
    return null;
  };

  const currentAge = parseInt(getVal("目前年齡"));
  const retireAge = parseInt(getVal("預計退休年齡"));
  const lifespan = parseInt(getVal("預期壽命"));
  const annualSaving = parseFloat(getVal("每年儲蓄金額（不參與投資）")) || 0;
  const annualInvest = parseFloat(getVal("每年投入投資金額（參與複利）")) || 0;
  const retiredInvest = parseFloat(getVal("退休後每年持續投入投資金額")) || 0;
  const initialCash = parseFloat(getVal("現金儲蓄總額")) || 0;
  const hasPension = getVal("是否有退休金（Y/N）")?.toString().toLowerCase() === "y";
  const pension = hasPension ? parseFloat(getVal("每年退休金金額")) || 0 : 0;
  const returnRate = parseFloat(getVal("預期年投資報酬率（%）")) / 100;
  const inflRate = parseFloat(getVal("支出年成長率（%）")) / 100;
  const retireExpense = parseFloat(getVal("預期退休後每年支出")) || 0;

  let cashAsset = initialCash;
  let investment = 0;
  let investmentPrincipal = 0;
  let age = currentAge;

  while (age <= lifespan) {
    investment *= (1 + returnRate);

    if (age < retireAge) {
      investment += annualInvest;
      investmentPrincipal += annualInvest;
      cashAsset += annualSaving;
    } else {
      investment += retiredInvest;
      investmentPrincipal += retiredInvest;
    }

    const income = age < retireAge ? 0 : pension;
    const expense = age < retireAge ? 0 : retireExpense * Math.pow(1 + inflRate, age - retireAge);

    cashAsset += income - expense;

    const totalAsset = cashAsset + investment;
    if (totalAsset < 0) return false; // ❗新邏輯：只要總資產變成負數，就算失敗

    age++;
  }

  return true;
}

function generateAndSendRetirementReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("輸入資料");
  const resultSheet = ss.getSheetByName("模擬結果");
  const url = ss.getUrl();
  const name = getVal(inputSheet, "使用者姓名") || "使用者";
  const email = getVal(inputSheet, "Email 收件人") || Session.getActiveUser().getEmail();
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const docTitle = `退休模擬報告 - ${name} - ${date}`;
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  body.appendParagraph("📘 退休財務自由模擬報告").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`👤 使用者姓名：${name}`);
  body.appendParagraph(`📅 產出日期：${date}`);
  body.appendParagraph("");

  // 🗂️ 模擬條件
  body.appendParagraph("🗂️ 模擬輸入條件").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const inputData = inputSheet.getRange("A1:B50").getValues().filter(row => row[0]);
  inputData.forEach(([label, value]) => {
    body.appendParagraph(`• ${label}: ${value}`);
  });
  body.appendParagraph("");

  // 🧠 診斷建議
  const diagnosisTips = resultSheet.getRange("T4:T10").getValues().flat().filter(String);
  body.appendParagraph("🧠 財務診斷建議").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  diagnosisTips.forEach(tip => body.appendParagraph("• " + tip));
  body.appendParagraph("");

  // 📊 精簡模擬表格（含格式化）
  const lastRow = resultSheet.getLastRow();
  const shortTableData = [["年份", "年齡", "投資本金", "投資收益", "總資產"]];
  for (let i = 2; i <= lastRow; i++) {
    const row = resultSheet.getRange(i, 1, 1, 6).getValues()[0];
    const formattedRow = [
      parseInt(row[0]),
      parseInt(row[1]),
      formatThousands(row[3]),
      formatThousands(row[4]),
      formatThousands(row[5])
    ];
    shortTableData.push(formattedRow);
  }

  const table = body.appendTable(shortTableData);
  table.setBorderWidth(0);

  // 美化表格
  const headerRow = table.getRow(0);
  for (let j = 0; j < headerRow.getNumCells(); j++) {
    const cell = headerRow.getCell(j);
    cell.setBold(true);
    cell.setBackgroundColor("#f1f3f4");
    cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(4);
    cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }

  for (let r = 1; r < table.getNumRows(); r++) {
    for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
      const cell = table.getRow(r).getCell(c);
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(4);
      const para = cell.getChild(0).asParagraph();
      if (c >= 2) {
        para.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      } else {
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    }
  }

  body.appendParagraph("");

  // 📈 插入圖表
  const charts = resultSheet.getCharts();
  if (charts.length > 0) {
    const blob = charts[0].getAs('image/png');
    body.appendParagraph("📈 資產與投資收益圖表").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendImage(blob).setWidth(650);
  }

  // 🔗 Google Sheets 連結
  body.appendParagraph("");
  body.appendParagraph("🔗 若需檢視完整模擬資料，請點選下方連結：");
  body.appendParagraph(url).setLinkUrl(url);

  doc.saveAndClose();

  const pdf = DriveApp.getFileById(doc.getId()).getAs("application/pdf");
  const subject = `退休模擬報告 - ${name} - ${date}`;
  const bodyText = `您好，

請見附件的退休模擬報告（含輸入條件、圖表與格式化表格），如需完整資料請參考 Google Sheets。

祝順心
退休模擬系統`;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: bodyText,
    attachments: [pdf]
  });

  SpreadsheetApp.getUi().alert(`✅ 美化報告已成功產出並寄送給：${email}`);
}

function getVal(sheet, label) {
  const range = sheet.getRange("A1:A100").getValues();
  for (let i = 0; i < range.length; i++) {
    if (range[i][0] === label) {
      return sheet.getRange(i + 1, 2).getValue();
    }
  }
  return null;
}

function formatThousands(num) {
  return typeof num === "number" ? num.toLocaleString("en-US", { maximumFractionDigits: 0 }) : num;
}

