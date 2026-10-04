// ==================== 🛠️ 欄位名稱與設定維護區 ====================
// 1. 工作表（頁籤）名稱設定
const SHEET_CALC     = "計算";
const SHEET_CALENDAR = "行事曆";
const SHEET_MASTER   = "主檔";

// 2. 儲存格與列號設定
const CELL_TARGET_DATE     = "B1";  // 「計算」頁籤中，目標日期所在的儲存格
const ROW_CALC_HEADER      = 2;     // 「計算」頁籤中，標題列所在的列號 (第 2 列)
const ROW_CALC_DATA_START  = 3;     // 「計算」頁籤中，資料開始寫入的列號 (第 3 列)
const ROW_MASTER_HEADER    = 1;     // 「主檔」頁籤中，標題列所在的列號 (第 1 列)
const ROW_CALENDAR_HEADER  = 1;     // 「行事曆」頁籤中，標題列所在的列號 (第 1 列)

// 3. 「行事曆」頁籤的關鍵欄位名稱
const HDR_CAL_DATE       = "日期";
const HDR_CAL_IS_WORKDAY = "是否為上班日";
const HDR_CAL_SHIFT      = "輪班班次";

// 4. 「主檔」與核心邏輯判斷的關鍵欄位名稱
const HDR_MASTER_SHIFT   = "班別";  // 用來判定 0、A、B 班別的欄位名稱
// =================================================================

/**
 * 1. 點擊選單時觸發此函數，負責彈出日曆選擇視窗
 */
function getWorkingEmployees() {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  const todayStr = `${yyyy}-${mm}-${dd}`; // 自動產生今天的日期格式，如 "2026-07-02"

  // 💡【已修正】：將原先的 \${todayStr} 改為 ${todayStr}，現在能完美預帶今天日期了！
  const htmlString = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: "Microsoft JhengHei", Arial, sans-serif; margin: 15px; color: #333; background-color: #f9f9f9; }
        .title { font-size: 14px; font-weight: bold; margin-bottom: 12px; color: #1a73e8; }
        input[type="date"] { padding: 8px; font-size: 14px; width: 90%; margin-bottom: 18px; border: 1px solid #ccc; border-radius: 4px; box-shadow: inset 0 1px 3px rgba(0,0,0,0.1); }
        .btn-group { display: flex; justify-content: flex-end; gap: 10px; width: 95%; }
        button { padding: 7px 14px; font-size: 13px; cursor: pointer; border: none; border-radius: 4px; font-weight: bold; }
        .btn-ok { background-color: #1a73e8; color: white; }
        .btn-cancel { background-color: #e0e0e0; color: #333; }
        button:hover { opacity: 0.9; }
      </style>
    </head>
    <body>
      <div class="title">📅 請選擇要計算出勤的日期：</div>
      <input type="date" id="selDate" value="${todayStr}">
      <div class="btn-group">
        <button class="btn-cancel" onclick="google.script.host.close()">取消</button>
        <button class="btn-ok" onclick="submit()">確定計算</button>
      </div>

      <script>
        function submit() {
          var dateVal = document.getElementById('selDate').value;
          if (!dateVal) {
            alert('請選擇有效日期！');
            return;
          }
          google.script.run.executeAttendanceWorkflow(dateVal);
          google.script.host.close(); // 秒閃機制：立刻關閉視窗
        }
      </script>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlString)
      .setWidth(340)
      .setHeight(150);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '排班與便當管理系統');
}

/**
 * 2. 接收到日曆日期後，寫入 B1 並執行排班計算
 */
function executeAttendanceWorkflow(selectedDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const calcSheet = ss.getSheetByName(SHEET_CALC);
  const calSheet = ss.getSheetByName(SHEET_CALENDAR);
  const masterSheet = ss.getSheetByName(SHEET_MASTER);
  
  if (!calcSheet || !calSheet || !masterSheet) {
    SpreadsheetApp.getUi().alert("錯誤：找不到指定的工作表，請確認下方頁籤名稱是否有被更改！");
    return false; // 🎯【熔斷機制】：改為回傳 false
  }

  // 將 HTML 選定日期寫入 B1
  const parts = selectedDateStr.split('-');
  const dateObj = new Date(parts[0], parts[1] - 1, parts[2]);
  calcSheet.getRange(CELL_TARGET_DATE).setValue(dateObj);
  
  // 1. 取得並格式化目標日期
  const targetDate = calcSheet.getRange(CELL_TARGET_DATE).getValue();
  if (!targetDate || !(targetDate instanceof Date)) {
    SpreadsheetApp.getUi().alert("請在「" + SHEET_CALC + "」工作表的 " + CELL_TARGET_DATE + " 儲存格輸入正確的日期！");
    return false; // 🎯【熔斷機制】：改為回傳 false
  }
  
  const timeZone = Session.getScriptTimeZone();
  const targetDateStr = Utilities.formatDate(targetDate, timeZone, "yyyy/MM/dd");
  
  // 2. 自動定位「行事曆」欄位，並找出目標日期的排班狀態
  const calData = calSheet.getDataRange().getValues();
  const calHeaders = calData[ROW_CALENDAR_HEADER - 1]; 
  
  const idxCalDate = calHeaders.indexOf(HDR_CAL_DATE);
  const idxCalIsWorkday = calHeaders.indexOf(HDR_CAL_IS_WORKDAY);
  const idxCalShift = calHeaders.indexOf(HDR_CAL_SHIFT);
  
  if (idxCalDate === -1 || idxCalIsWorkday === -1 || idxCalShift === -1) {
    SpreadsheetApp.getUi().alert("「" + SHEET_CALENDAR + "」缺少必要欄位！");
    return false; // 🎯【熔斷機制】：改為回傳 false
  }
  
  let isWorkday = false;
  let shiftToday = "";
  let dateFound = false;
  
  for (let i = ROW_CALENDAR_HEADER; i < calData.length; i++) {
    let rowDate = calData[i][idxCalDate];
    if (rowDate instanceof Date) {
      let rowDateStr = Utilities.formatDate(rowDate, timeZone, "yyyy/MM/dd");
      if (rowDateStr === targetDateStr) {
        let valWorkday = calData[i][idxCalIsWorkday];
        isWorkday = (valWorkday === "是");
        shiftToday = String(calData[i][idxCalShift]).trim();
        dateFound = true;
        break;
      }
    }
  }
  
  if (!dateFound) {
    SpreadsheetApp.getUi().alert("在「" + SHEET_CALENDAR + "」中找不到日期：" + targetDateStr + " 的設定。");
    return false; // 🎯【熔斷機制】：改為回傳 false
  }
  
  // 3. 自動對應「主檔」與「計算」的欄位位置
  const masterData = masterSheet.getDataRange().getValues();
  const masterHeaders = masterData[ROW_MASTER_HEADER - 1]; 
  const idxMasterShift = masterHeaders.indexOf(HDR_MASTER_SHIFT);
  
  if (idxMasterShift === -1) {
    SpreadsheetApp.getUi().alert("「" + SHEET_MASTER + "」找不到「" + HDR_MASTER_SHIFT + "」欄位！");
    return false; // 🎯【熔斷機制】：改為回傳 false
  }
  
  // 抓取「計算」頁籤第 2 列目前所有的標題名稱
  const calcHeaders = calcSheet.getRange(ROW_CALC_HEADER, 1, 1, calcSheet.getLastColumn()).getValues()[0];
  
  // 🚧【核心優化鎖】動態尋找「加班」欄位在哪裡，鎖定資料處理範圍
  const idxOvertimeCol = calcHeaders.indexOf("加班");
  const colsToProcess = idxOvertimeCol !== -1 ? idxOvertimeCol + 1 : calcHeaders.length;
  
  // 讓欄位對照僅處理到「加班」這一欄為止（忽略右側統計區）
  const processHeaders = calcHeaders.slice(0, colsToProcess);
  const columnMapping = processHeaders.map(calcHeaderName => {
    return masterHeaders.indexOf(calcHeaderName); 
  });
  
  // 4. 執行篩選邏輯
  let resultData = [];
  for (let i = ROW_MASTER_HEADER; i < masterData.length; i++) {
    let row = masterData[i];
    let empShift = String(row[idxMasterShift]).trim();
    if (empShift === "") continue; 
    
    let isWorking = false;
    if (empShift === "0") {
      if (isWorkday) isWorking = true;
    } else if (empShift === "A" || empShift === "B") {
      if (empShift === shiftToday) isWorking = true;
    }
    
    if (isWorking) {
      let calcRow = columnMapping.map(masterIdx => {
        return masterIdx !== -1 ? row[masterIdx] : ""; 
      });
      resultData.push(calcRow);
    }
  }
  
  // ---------------------------------------------------------------
  // 5. 寫入「計算」頁籤（已加封鎖區，不干涉右側欄位）
  // ---------------------------------------------------------------
  let lastRowCalc = calcSheet.getLastRow();
  
  // 清空舊資料時，寬度只限定在 colsToProcess (欄位 A 到 加班欄)
  if (lastRowCalc >= ROW_CALC_DATA_START) {
    calcSheet.getRange(ROW_CALC_DATA_START, 1, lastRowCalc - ROW_CALC_DATA_START + 1, colsToProcess).clearContent();
  }
  
  // 寫入新資料時，同樣將寬度嚴格限制在 colsToProcess
  if (resultData.length > 0) {
    calcSheet.getRange(ROW_CALC_DATA_START, 1, resultData.length, colsToProcess).setValues(resultData);
    SpreadsheetApp.getUi().alert("✅ 執行成功！\n日期：" + targetDateStr + "\n當日班次：" + shiftToday + "\n共帶入 " + resultData.length + " 位上班人員。");
  } else {
    SpreadsheetApp.getUi().alert("提示：本日無人需要上班，或查無符合條件的資料。");
  }
  
  return true; // 🎯【熔斷機制】：成功結束，回傳 true
}