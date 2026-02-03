/**
 * ==================================================
 * 圖書股長自動排班系統 (v5.0 高三停勤特化版)
 * Update Highlights:
 * 1. [新增設定] Config B4 欄位：高三停勤開始日。
 * 2. [排班邏輯] 超過停勤日後，高三自動強制免勤 (無須手動打勾)。
 * 3. [表格結構] 配合新設定，排除列表下移至第 8 列開始。
 * ==================================================
 */

// --- 全域變數 ---
const SHEET_CONFIG = "Config";
const SHEET_LIB = "Librarians";
const SHEET_RESULT = "Result";
const SHEET_STATS = "Stats";

const TOTAL_CLASSES = 28;
const SKIP_CLASS = 24;

// 設定資料開始的列數 (因為上面多了 B4 設定，所以標題移到第 7 列，資料從第 8 列開始)
const EXCLUSION_START_ROW = 8; 

/**
 * 1. 建立試算表上方選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📚 圖書股長系統')
    .addItem('🚀 1. 系統初次建置 (含高三停勤設定)', 'firstTimeSetup')
    .addSeparator()
    .addItem('2. 重新產生假日時程', 'initializeSemesterSetup')
    .addItem('3. 產生/更新值勤表', 'generateDutyRoster')
    .addSeparator()
    .addItem('📊 4. 期末結算統計', 'generateStats')
    .addItem('📧 5. 寄送明日提醒信', 'sendDailyReminders')
    .addToUi();
}

/**
 * 【功能 1】系統初次建置 (新增高三停勤日詢問)
 */
function firstTimeSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // --- Step A: 建立 Config 工作表 ---
  let configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (configSheet) ss.deleteSheet(configSheet); // 強制重建以確保格式正確
  configSheet = ss.insertSheet(SHEET_CONFIG);
  
  const headers = [
    ["參數設定", "", "", "", ""],
    ["學期開始", "", "(系統自動填入)", "", ""], 
    ["學期結束", "", "(系統自動填入)", "", ""],
    ["高三停勤開始日", "", "(在此日期(含)之後，高三全面免勤)", "", ""], // New B4
    ["", "", "", "", ""], // Row 5 空白
    ["特殊日期排除設定", "", "高一免勤", "高二免勤", "高三免勤"], // Row 6 標題
    ["日期", "事由", "(打勾=免勤)", "(打勾=免勤)", "(打勾=免勤)"]  // Row 7 欄位名
  ];
  
  configSheet.getRange(1, 1, 7, 5).setValues(headers);
  
  // 美化
  configSheet.getRange("A1:E1").setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
  configSheet.getRange("A6:E7").setBackground("#cfe2f3").setFontWeight("bold");
  configSheet.getRange("A4").setFontColor("#cc0000").setFontWeight("bold"); // 高三設定特別標示
  configSheet.setColumnWidth(1, 120); 
  configSheet.setColumnWidth(2, 150); 
  configSheet.deleteRows(8, configSheet.getMaxRows() - 7); 

  // --- Step B: 建立 Librarians 工作表 ---
  let libSheet = ss.getSheetByName(SHEET_LIB);
  if (!libSheet) {
    libSheet = ss.insertSheet(SHEET_LIB);
    libSheet.getRange(1, 1, 1, 4).setValues([["年級", "班級", "姓名", "Email"]]);
    libSheet.getRange("A1:D1").setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
    let classList = [];
    for (let g = 1; g <= 3; g++) {
      for (let c = 1; c <= TOTAL_CLASSES; c++) {
        if (c !== SKIP_CLASS) classList.push([g, c, "", ""]);
      }
    }
    libSheet.getRange(2, 1, classList.length, 4).setValues(classList);
    libSheet.setFrozenRows(1);
  }

  // --- Step C: 建立 Result 空白表 ---
  if (!ss.getSheetByName(SHEET_RESULT)) ss.insertSheet(SHEET_RESULT);

  // --- Step D: 對話框詢問 (三連問) ---
  let defaultDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  // Q1. 開始
  let r1 = ui.prompt('1/3 設定學期開始', `格式: YYYY/MM/DD (例: ${defaultDate})`, ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  let dStart = r1.getResponseText();

  // Q2. 結束
  let r2 = ui.prompt('2/3 設定學期結束', `格式: YYYY/MM/DD`, ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  let dEnd = r2.getResponseText();

  // Q3. 高三停勤 (New)
  let r3 = ui.prompt('3/3 設定高三停勤開始日', 
    `從哪一天開始高三不用值勤？(通常是統測或畢業前)\n若不確定或全學期皆要值勤，請直接按確定(留白)即可。`, 
    ui.ButtonSet.OK);
  let dStopG3 = r3.getResponseText();

  // --- Step E: 寫入與初始化 ---
  if (!isValidDate(dStart) || !isValidDate(dEnd)) {
    Browser.msgBox("錯誤：起訖日期格式不正確。");
    return;
  }

  configSheet.getRange("B2").setValue(dStart);
  configSheet.getRange("B3").setValue(dEnd);
  
  // 如果有填寫高三停勤日，且格式正確
  if (dStopG3 && isValidDate(dStopG3)) {
    configSheet.getRange("B4").setValue(dStopG3);
  } else {
    configSheet.getRange("B4").clearContent(); // 留白代表無停勤
  }

  // 呼叫初始化
  initializeSemesterSetup(true);
}

/**
 * 【功能 2】初始化學期設定
 */
function initializeSemesterSetup(isAutoRun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) return;

  const startDate = configSheet.getRange("B2").getValue();
  const endDate = configSheet.getRange("B3").getValue();
  
  if (!(startDate instanceof Date) || !(endDate instanceof Date)) {
    if (!isAutoRun) Browser.msgBox("請檢查 B2, B3 日期設定。");
    return;
  }

  // 取得現有資料 (從第8列開始)
  let existingKeys = new Set();
  const lastRow = configSheet.getLastRow();
  if (lastRow >= EXCLUSION_START_ROW) {
    const data = configSheet.getRange(EXCLUSION_START_ROW, 1, lastRow - EXCLUSION_START_ROW + 1, 2).getValues();
    data.forEach(r => {
      let d = (r[0] instanceof Date) ? formatDateKey(r[0]) : "BLANK";
      existingKeys.add(d + "_" + r[1]);
    });
  }

  let newRows = [];
  let currentDate = new Date(startDate);
  const end = new Date(endDate);
  
  // A. 六日
  while (currentDate <= end) {
    let day = currentDate.getDay();
    let dateKey = formatDateKey(currentDate);
    if (day === 0 || day === 6) {
      let name = day === 0 ? "週日" : "週六";
      if (!existingKeys.has(dateKey + "_" + name)) {
        newRows.push([new Date(currentDate), name, true, true, true]);
      }
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }

  // B. 段考
  let hasExam = Array.from(existingKeys).some(k => k.includes("段考"));
  if (!hasExam) {
    const exams = ["第一次段考", "第二次段考", "第三次段考"];
    const days = ["(Day1)", "(Day2)"];
    exams.forEach(exam => {
      days.forEach(day => {
        newRows.push(["", `${exam} ${day}`, true, true, true]);
      });
    });
  }

  // C. 寫入
  if (newRows.length > 0) {
    let startRow = configSheet.getLastRow() + 1;
    // 如果表格還是空的(剛建立)，從 EXCLUSION_START_ROW 開始
    if (startRow < EXCLUSION_START_ROW) startRow = EXCLUSION_START_ROW;
    
    configSheet.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
    configSheet.getRange(startRow, 3, newRows.length, 3).insertCheckboxes();
    
    // 排序 (從第 8 列開始排)
    const sortRange = configSheet.getRange(EXCLUSION_START_ROW, 1, configSheet.getLastRow() - EXCLUSION_START_ROW + 1, 5);
    sortRange.sort({column: 1, ascending: true});
    
    Browser.msgBox(`設定完成！已更新假日與考試欄位。`);
  } else {
    if (!isAutoRun) Browser.msgBox("無新增項目。");
  }
}

/**
 * 【功能 3】產生值勤表 (核心：高三停勤邏輯)
 */
function generateDutyRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEET_CONFIG);
  const libSheet = ss.getSheetByName(SHEET_LIB);
  let resultSheet = ss.getSheetByName(SHEET_RESULT);

  if (!configSheet || !libSheet) { Browser.msgBox("請先執行「系統初次建置」。"); return; }
  if (!resultSheet) resultSheet = ss.insertSheet(SHEET_RESULT);

  try { resultSheet.showColumns(1, 20); } catch(e) {}
  resultSheet.clear();

  // 1. 讀取名單
  let librarianMap = new Map();
  const libRows = libSheet.getLastRow();
  if (libRows > 1) {
    const data = libSheet.getRange(2, 1, libRows - 1, 4).getValues();
    data.forEach(row => {
      librarianMap.set(`${row[0]}-${row[1]}`, { name: row[2], email: row[3] });
    });
  }

  // 2. 讀取設定 (包含高三停勤日)
  const startDate = configSheet.getRange("B2").getValue();
  const endDate = configSheet.getRange("B3").getValue();
  const stopDateG3Raw = configSheet.getRange("B4").getValue(); // 讀取 B4
  
  let stopDateG3 = null;
  if (stopDateG3Raw instanceof Date) {
    stopDateG3 = stopDateG3Raw;
  }
  
  // 3. 讀取排除清單 (從第8列開始)
  let exclusionMap = new Map();
  const configLastRow = configSheet.getLastRow();
  if (configLastRow >= EXCLUSION_START_ROW) {
    const exData = configSheet.getRange(EXCLUSION_START_ROW, 1, configLastRow - EXCLUSION_START_ROW + 1, 5).getValues();
    exData.forEach(row => {
      let d = row[0];
      if (d instanceof Date && !isNaN(d)) {
        let key = formatDateKey(d);
        let current = exclusionMap.get(key) || [false, false, false];
        exclusionMap.set(key, [
          current[0] || row[2] === true,
          current[1] || row[3] === true,
          current[2] || row[4] === true
        ]);
      }
    });
  }

  // 4. 排班
  let classes = [];
  for (let i = 1; i <= TOTAL_CLASSES; i++) {
    if (i !== SKIP_CLASS) classes.push(i);
  }
  let idxG1 = 0, idxG2 = 0, idxG3 = 0;

  let outputData = [[
    "日期", "星期", 
    "高一值勤", "高一簽到", 
    "高二值勤", "高二簽到", 
    "高三值勤", "高三簽到", 
    "Sys_Email_1", "Sys_Email_2", "Sys_Email_3"
  ]];
  const weekDayZh = ["日", "一", "二", "三", "四", "五", "六"];

  let currentDate = new Date(startDate);
  const end = new Date(endDate);

  while (currentDate <= end) {
    let day = currentDate.getDay();
    let dateStr = formatDateKey(currentDate);
    
    // 取得原本設定的排除狀態
    let exclusions = exclusionMap.get(dateStr) || [false, false, false];
    
    // ★ 高三停勤邏輯：如果今天 >= 停勤日，強制將高三設為免勤 (True)
    if (stopDateG3 && currentDate >= stopDateG3) {
      exclusions[2] = true; 
    }

    let rowData = [
      Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy/MM/dd"),
      weekDayZh[day]
    ];
    let emailData = [];

    function processGrade(grade, idx, isExcluded, tracker) {
      // 邏輯優化：如果是高三且是因為停勤日而免勤，可以顯示不同文字 (這裡統一顯示免勤保持簡潔)
      if (isExcluded) {
        rowData.push("免勤", ""); 
        emailData.push("");
        return tracker;
      } else {
        let cls = classes[idx];
        let info = librarianMap.get(`${grade}-${cls}`);
        let txt = `${grade}年${cls}班`;
        if (info && info.name && isNaN(info.name) && info.name.toString().trim() !== "") {
          txt += `\n(${info.name})`;
        }
        rowData.push(txt, "");
        emailData.push(info ? info.email : "");
        return (tracker + 1) % classes.length;
      }
    }

    idxG1 = processGrade(1, idxG1, exclusions[0], idxG1);
    idxG2 = processGrade(2, idxG2, exclusions[1], idxG2);
    idxG3 = processGrade(3, idxG3, exclusions[2], idxG3);

    rowData = rowData.concat(emailData);
    outputData.push(rowData);
    currentDate.setDate(currentDate.getDate() + 1);
  }

  // 5. 寫入
  if (outputData.length > 1) {
    resultSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    let rng = resultSheet.getDataRange();
    rng.setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    resultSheet.getRange("A1:K1").setBackground("#cfe2f3").setFontWeight("bold");
    
    resultSheet.setColumnWidth(1, 90); resultSheet.setColumnWidth(2, 40);
    [3, 5, 7].forEach(c => resultSheet.setColumnWidth(c, 110));
    [4, 6, 8].forEach(c => resultSheet.setColumnWidth(c, 70));
    resultSheet.hideColumns(9, 3);
    
    let rule = SpreadsheetApp.newConditionalFormatRule().whenTextContains("免勤").setBackground("#E0E0E0").setFontColor("#888888").setRanges([
        resultSheet.getRange(2, 3, outputData.length, 1), resultSheet.getRange(2, 5, outputData.length, 1), resultSheet.getRange(2, 7, outputData.length, 1)
      ]).build();
    resultSheet.setConditionalFormatRules([rule]);
  }
}

/**
 * 【功能 4】期末結算統計 (含截止日期過濾版)
 * Update:
 * 1. 跳出視窗詢問「統計截止日期」。
 * 2. 只計算該日期(含)以前的排班紀錄。
 * 3. 標題自動標註統計截止日。
 */
function generateStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const libSheet = ss.getSheetByName(SHEET_LIB);
  const resultSheet = ss.getSheetByName(SHEET_RESULT);
  let statsSheet = ss.getSheetByName(SHEET_STATS);

  if (!libSheet || !resultSheet) {
    Browser.msgBox("資料不足，無法統計。請確認 Librarians 和 Result 表都已存在。");
    return;
  }

  // --- Step 1: 詢問截止日期 ---
  let today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  let response = ui.prompt(
    '設定統計截止日期',
    `只統計此日期 (含) 以前的資料。\n預設為今天：${today}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    Browser.msgBox("已取消統計。");
    return;
  }

  let limitDateStr = response.getResponseText();
  let limitDate = new Date(limitDateStr);

  if (isNaN(limitDate.getTime())) {
    Browser.msgBox("日期格式錯誤，請輸入 YYYY/MM/DD");
    return;
  }

  // --- Step 2: 準備統計表 ---
  if (!statsSheet) statsSheet = ss.insertSheet(SHEET_STATS);
  statsSheet.clear();

  // --- Step 3: 初始化人員名單 ---
  // Key: "Grade-Class", Value: Object
  let statsMap = new Map();
  const libData = libSheet.getRange(2, 1, libSheet.getLastRow() - 1, 3).getValues();
  libData.forEach(r => {
    statsMap.set(`${r[0]}-${r[1]}`, { 
      g: r[0], c: r[1], name: r[2], 
      scheduled: 0, actual: 0 
    });
  });

  // --- Step 4: 掃描 Result 表並過濾日期 ---
  const resData = resultSheet.getDataRange().getValues();
  // 欄位索引: 高一(C=2, 簽=3), 高二(E=4, 簽=5), 高三(G=6, 簽=7)
  const pairs = [[2, 3], [4, 5], [6, 7]];

  // 從第 2 列 (index 1) 開始掃描
  for (let i = 1; i < resData.length; i++) {
    let rowDateRaw = resData[i][0];
    
    // 檢查日期是否有效
    if (!(rowDateRaw instanceof Date)) continue;

    // ★ 關鍵過濾邏輯：如果該行日期 > 截止日期，直接跳過不統計
    if (rowDateRaw > limitDate) continue;

    pairs.forEach(pair => {
      let cellText = resData[i][pair[0]].toString(); // 排班內容
      let signText = resData[i][pair[1]].toString().trim(); // 簽到內容

      // 檢查是否為排班 (排除"免勤")
      let match = cellText.match(/^(\d+)年(\d+)班/);
      if (match) {
        let key = `${match[1]}-${match[2]}`;
        if (statsMap.has(key)) {
          let rec = statsMap.get(key);
          rec.scheduled += 1; // 應到 +1
          
          // 只要簽到欄有字，就算實到
          if (signText !== "") {
            rec.actual += 1; // 實到 +1
          }
        }
      }
    });
  }

  // --- Step 5: 輸出報表 ---
  // 標題列
  let titleStr = `圖書股長值勤統計 (截至 ${limitDateStr})`;
  let header = ["年級", "班級", "姓名", "應值勤次數", "實簽到次數", "出勤百分比"];
  let output = [header];
  
  // 轉陣列並排序 (先年級再班級)
  let list = Array.from(statsMap.values()).sort((a, b) => {
    if (a.g !== b.g) return a.g - b.g;
    return a.c - b.c;
  });

  list.forEach(item => {
    let percent = 0;
    if (item.scheduled > 0) {
      percent = item.actual / item.scheduled;
    }
    output.push([
      item.g, item.c, item.name, 
      item.scheduled, item.actual, percent
    ]);
  });

  // 寫入資料
  statsSheet.getRange(2, 1, output.length, 6).setValues(output);
  
  // 設定大標題 (在第一列合併儲存格顯示截止日)
  statsSheet.getRange("A1:F1").merge().setValue(titleStr)
    .setBackground("#4a86e8").setFontColor("white")
    .setFontWeight("bold").setHorizontalAlignment("center");
  
  // 設定欄位標題樣式 (第二列)
  statsSheet.getRange("A2:F2").setBackground("#e06666").setFontColor("white").setFontWeight("bold");

  // 表格框線與對齊
  let dataRange = statsSheet.getRange(2, 1, output.length, 6);
  dataRange.setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);

  // 設定百分比格式 (F欄)
  statsSheet.getRange(3, 6, output.length - 1, 1).setNumberFormat("0%");

  // 加上資料條 (Data Bar)
  let rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint("#57bb8a") // 綠色
    .setGradientMinpoint("#ffffff") // 白色
    .setRanges([statsSheet.getRange(3, 6, output.length - 1, 1)])
    .build();
  statsSheet.setConditionalFormatRules([rule]);
  
  statsSheet.activate();
  Browser.msgBox(`統計完成！\n統計區間：學期開始 ~ ${limitDateStr}`);
}
/**
 * 【功能 5】寄信 (邏輯不變)
 */
function sendDailyReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName(SHEET_RESULT);
  if (!resultSheet) return;
  const data = resultSheet.getDataRange().getValues();
  let tomorrow = new Date(); tomorrow.setDate(tomorrow.getDate() + 1);
  let tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  for (let i = 1; i < data.length; i++) {
    let rowDate = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "yyyy/MM/dd") : data[i][0];
    if (rowDate === tomorrowStr) {
      let emails = [data[i][8], data[i][9], data[i][10]];
      let classes = [data[i][2], data[i][4], data[i][6]];
      emails.forEach((email, idx) => {
        if (email && email.toString().includes("@")) {
          MailApp.sendEmail(email, `【圖書館通知】明日值勤提醒 (${tomorrowStr})`, `同學您好，明日 ${tomorrowStr} 輪到您 (${classes[idx]}) 值勤，請記得準時簽到。`);
        }
      });
      break;
    }
  }
}

function formatDateKey(date) { return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd"); }
function isValidDate(dateString) { return !isNaN(Date.parse(dateString)); }