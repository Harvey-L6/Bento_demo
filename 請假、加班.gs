// ==================== 🛠️ 加班請假功能維護區 ====================
const SHEET_LEAVE_OT     = "請假/加班";
const SHEET_CALC_MAIN    = "計算";
const SHEET_MASTER_MAIN  = "主檔";

// 「計算」與「主檔」頁籤的關鍵欄位名稱
const HDR_C_EMPID        = "員工編號";
const HDR_C_NAME         = "員工姓名";
const HDR_C_LEAVE        = "請假";
const HDR_C_OVERTIME     = "加班";
// =================================================================

function processLeaveAndOvertime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName(SHEET_CALC_MAIN);
  const masterSheet = ss.getSheetByName(SHEET_MASTER_MAIN);
  const leaveOtSheet = ss.getSheetByName(SHEET_LEAVE_OT);
  
  if (!calcSheet || !masterSheet || !leaveOtSheet) {
    SpreadsheetApp.getUi().alert("錯誤：找不到必要的頁籤（計算、主檔 或 請假/加班），請檢查名稱！");
    return false; // 🎯【熔斷機制】：改為回傳 false[cite: 2]
  }
  
  // 取得目標日期用於提示訊息
  const targetDate = calcSheet.getRange("B1").getValue();
  const timeZone = Session.getScriptTimeZone();
  const targetDateStr = targetDate instanceof Date ? Utilities.formatDate(targetDate, timeZone, "yyyy/MM/dd") : String(targetDate);

  // ---------------------------------------------------------------
  // 🧭 預備作業：讀取主檔並建立「存在性檢查集合」與「雙向對照表」
  // ---------------------------------------------------------------
  const masterData = masterSheet.getDataRange().getValues();
  const masterHeaders = masterData[0]; // 主檔標題在第 1 列
  const idxMasterId = masterHeaders.indexOf(HDR_C_EMPID);
  const idxMasterName = masterHeaders.indexOf(HDR_C_NAME);
  
  if (idxMasterId === -1 || idxMasterName === -1) {
    SpreadsheetApp.getUi().alert("「" + SHEET_MASTER_MAIN + "」頁籤缺少 員工編號 或 員工姓名 欄位！");
    return false; // 🎯【熔斷機制】：改為回傳 false[cite: 2]
  }
  
  let masterIdSet = new Set();
  let masterNameSet = new Set();
  let idToNameMap = {};
  let nameToIdMap = {};
  
  for (let j = 1; j < masterData.length; j++) {
    let mId = String(masterData[j][idxMasterId] || "").trim();
    let mName = String(masterData[j][idxMasterName] || "").trim();
    
    if (mId) {
      masterIdSet.add(mId);
      idToNameMap[mId] = mName;
    }
    if (mName) {
      masterNameSet.add(mName);
      nameToIdMap[mName] = mId;
    }
  }
  
  let leaveOtLastRow = leaveOtSheet.getLastRow();
  if (leaveOtLastRow < 3) {
    SpreadsheetApp.getUi().alert("提示：「" + SHEET_LEAVE_OT + "」頁籤內沒有任何填寫資料。");
    return true; // 💡【放行機制】：今天沒人請假或加班是合法的正常現象，直接回傳 true 放行！[cite: 2]
  }
  
  // 讀取「請假/加班」原始填寫資料 (從第 3 列開始，共 4 欄：A~D)
  const leaveOtRange = leaveOtSheet.getRange(3, 1, leaveOtLastRow - 2, 4);
  const leaveOtValues = leaveOtRange.getValues();
  
  // ---------------------------------------------------------------
  // 🚨 🛑 核心防呆零：【在補漏之前】檢查填寫的人員是否存在於主檔
  // ---------------------------------------------------------------
  let notFoundInMaster = []; // 收集查無此人的輸入值
  
  for (let i = 0; i < leaveOtValues.length; i++) {
    let leaveId   = String(leaveOtValues[i][0] || "").trim();
    let leaveName = String(leaveOtValues[i][1] || "").trim();
    let otId      = String(leaveOtValues[i][2] || "").trim();
    let otName    = String(leaveOtValues[i][3] || "").trim();
    
    // 檢查請假區
    if (leaveId && !masterIdSet.has(leaveId)) {
      notFoundInMaster.push("請假欄工號 [" + leaveId + "]");
    }
    if (leaveName && !masterNameSet.has(leaveName)) {
      notFoundInMaster.push("請假欄姓名 [" + leaveName + "]");
    }
    
    // 檢查加班區
    if (otId && !masterIdSet.has(otId)) {
      notFoundInMaster.push("加班欄工號 [" + otId + "]");
    }
    if (otName && !masterNameSet.has(otName)) {
      notFoundInMaster.push("加班欄姓名 [" + otName + "]");
    }
  }
  
  // 🔥【第一道熔斷】若有任何人不在主檔中，直接中斷，不進行自動補齊與寫入！
  if (notFoundInMaster.length > 0) {
    let errorMsg = "⚠️【資料檢查發現嚴重錯誤，已停止執行】\n\n";
    errorMsg += "❌ 填寫的人員資料「不存在」於主檔中（請確認是否打錯字或輸入錯誤）：\n";
    errorMsg += "👉 " + notFoundInMaster.join("\n👉 ") + "\n\n";
    errorMsg += "請您再次確認主檔資料，或修正錯誤填寫後重新執行。";
    
    SpreadsheetApp.getUi().alert(errorMsg);
    return false; // 🎯【熔斷機制】：改為回傳 false[cite: 2]
  }

  // ---------------------------------------------------------------
  // 🧭 步驟 1：【通過存在檢查，開始補漏】雙向自動補齊
  // ---------------------------------------------------------------
  for (let i = 0; i < leaveOtValues.length; i++) {
    let leaveId   = String(leaveOtValues[i][0] || "").trim();
    let leaveName = String(leaveOtValues[i][1] || "").trim();
    let otId      = String(leaveOtValues[i][2] || "").trim();
    let otName    = String(leaveOtValues[i][3] || "").trim();
    
    // 補齊請假人員
    if (!leaveId && leaveName) {
      leaveOtValues[i][0] = nameToIdMap[leaveName];
    } else if (leaveId && !leaveName) {
      leaveOtValues[i][1] = idToNameMap[leaveId];
    }
    
    // 補齊加班人員
    if (!otId && otName) {
      leaveOtValues[i][2] = nameToIdMap[otName];
    } else if (otId && !otName) {
      leaveOtValues[i][3] = idToNameMap[otId];
    }
  }
  // 將補齊後的完整資料「立刻」寫回工作表畫面
  leaveOtRange.setValues(leaveOtValues);
  
  // 彙整已完成補齊的名單，並在此時同步偵測「自身清單重複填寫」
  let leavePeople = [];
  let otPeople = [];
  let leaveSeenIds = new Set();
  let otSeenIds = new Set();
  
  let dupLeaveInList = []; // 請假自身重複
  let dupOtInList = [];    // 加班自身重複
  
  for (let i = 0; i < leaveOtValues.length; i++) {
    let leaveId   = String(leaveOtValues[i][0] || "").trim();
    let leaveName = String(leaveOtValues[i][1] || "").trim();
    let otId      = String(leaveOtValues[i][2] || "").trim();
    let otName    = String(leaveOtValues[i][3] || "").trim();
    
    if (leaveId) {
      if (leaveSeenIds.has(leaveId)) {
        if (!dupLeaveInList.includes(leaveName)) dupLeaveInList.push(leaveName);
      }
      leaveSeenIds.add(leaveId);
      leavePeople.push({ id: leaveId, name: leaveName });
    }
    
    if (otId) {
      if (otSeenIds.has(otId)) {
        if (!dupOtInList.includes(otName)) dupOtInList.push(otName);
      }
      otSeenIds.add(otId);
      otPeople.push({ id: otId, name: otName });
    }
  }
  
  // ---------------------------------------------------------------
  // 🧭 步驟 2：讀取「計算」頁籤目前由常規排班抓出的「當日基本出勤名單」
  // ---------------------------------------------------------------
  const calcHeaders = calcSheet.getRange(2, 1, 1, calcSheet.getLastColumn()).getValues()[0];
  const idxCalcId    = calcHeaders.indexOf(HDR_C_EMPID);
  const idxCalcName  = calcHeaders.indexOf(HDR_C_NAME);
  const idxCalcLeave = calcHeaders.indexOf(HDR_C_LEAVE);
  const idxCalcOt    = calcHeaders.indexOf(HDR_C_OVERTIME);
  
  if ([idxCalcId, idxCalcName, idxCalcLeave, idxCalcOt].includes(-1)) {
    SpreadsheetApp.getUi().alert("「" + SHEET_CALC_MAIN + "」頁籤缺少必要欄位（員工編號/員工姓名/請假/加班）！");
    return false; // 🎯【熔斷機制】：改為回傳 false[cite: 2]
  }
  
  let calcLastRow = calcSheet.getLastRow();
  let regularEmpIds = new Set();
  let calcData = [];
  let calcRange = null;
  
  if (calcLastRow >= 3) {
    calcRange = calcSheet.getRange(3, 1, calcLastRow - 2, calcHeaders.length);
    calcData = calcRange.getValues();
    for (let i = 0; i < calcData.length; i++) {
      let empId = String(calcData[i][idxCalcId] || "").trim();
      if (empId) regularEmpIds.add(empId);
    }
  }
  
  // ---------------------------------------------------------------
  // 🧭 步驟 3：🔥 執行全面防呆比對（多重邏輯交叉檢查）
  // ---------------------------------------------------------------
  let crossDupErrors = []; // 請假與加班重複
  let leaveErrors = [];    // 無須出勤卻請假
  let otErrors = [];       // 本來就要出勤卻送加班
  
  // 防呆 A：檢查「請假」與「加班」之間是否有重複人員
  for (let lp of leavePeople) {
    if (otSeenIds.has(lp.id)) {
      if (!crossDupErrors.includes(lp.name)) crossDupErrors.push(lp.name);
    }
  }
  
  // 防呆 B：請假防呆（如果當天沒在常規出勤名單內，代表無須出勤，不應請假）
  for (let lp of leavePeople) {
    if (!regularEmpIds.has(lp.id)) {
      if (!leaveErrors.includes(lp.name)) leaveErrors.push(lp.name);
    }
  }
  
  // 防呆 C：加班防呆（如果當天本來就在常規出勤名單內，代表本來就要出勤，不應重複送加班）
  for (let op of otPeople) {
    if (regularEmpIds.has(op.id)) {
      if (!otErrors.includes(op.name)) otErrors.push(op.name);
    }
  }
  
  // 🚨【第二道熔斷】若有任何邏輯衝突錯誤，立即顯示警告並終止程式
  if (dupLeaveInList.length > 0 || dupOtInList.length > 0 || crossDupErrors.length > 0 || leaveErrors.length > 0 || otErrors.length > 0) {
    let errorMsg = "⚠️【考勤邏輯檢查發現異常，已停止執行】\n\n";
    
    if (dupLeaveInList.length > 0) {
      errorMsg += "❌ 請假名單有人員重複輸入多筆：" + dupLeaveInList.join("、") + "\n";
    }
    if (dupOtInList.length > 0) {
      errorMsg += "❌ 加班名單有人員重複輸入多筆：" + dupOtInList.join("、") + "\n";
    }
    if (crossDupErrors.length > 0) {
      errorMsg += "❌ 人員：" + crossDupErrors.join("、") + " 重複出現在請假與加班名單中。\n";
    }
    if (leaveErrors.length > 0) {
      errorMsg += "❌ 日期-" + targetDateStr + " 人員 " + leaveErrors.join("、") + " 無須出勤，不應請假。\n";
    }
    if (otErrors.length > 0) {
      errorMsg += "❌ 人員 " + otErrors.join("、") + " 當天本來就要出勤，不能加班。\n";
    }
    errorMsg += "\n請您再次確認，修正名單後重新執行。";
    
    SpreadsheetApp.getUi().alert(errorMsg);
    return false; // 🎯【熔斷機制】：改為回傳 false[cite: 2]
  }
  
  // ---------------------------------------------------------------
  // 🧭 步驟 4：【正式執行】通過所有安全關卡，開始變更「計算」頁籤
  // ---------------------------------------------------------------
  // A. 將常規出勤名單中的請假者標記為 "Y"
  if (calcLastRow >= 3) {
    for (let i = 0; i < calcData.length; i++) {
      let empId = String(calcData[i][idxCalcId]).trim();
      if (leaveSeenIds.has(empId)) {
        calcData[i][idxCalcLeave] = "Y";
      }
    }
    calcRange.setValues(calcData);
  }
  
  // B. 撈取主檔完整欄位，將加班者追加至計算表最下方
  const columnMapping = calcHeaders.map(h => masterHeaders.indexOf(h));
  let appendRows = [];
  
  for (let otUser of otPeople) {
    let foundMasterRow = null;
    for (let j = 1; j < masterData.length; j++) {
      let mId = String(masterData[j][idxMasterId]).trim();
      if (otUser.id === mId) {
        foundMasterRow = masterData[j];
        break;
      }
    }
    
    if (foundMasterRow) {
      let newRow = columnMapping.map((masterIdx, idx) => {
        if (idx === idxCalcLeave) return "";      
        if (idx === idxCalcOt) return "Y";        
        if (masterIdx !== -1) return foundMasterRow[masterIdx]; 
        return "";
      });
      appendRows.push(newRow);
    }
  }
  
  // 寫入加班人員
  if (appendRows.length > 0) {
    let currentLastRow = calcSheet.getLastRow();
    let insertRowStart = currentLastRow < 3 ? 3 : currentLastRow + 1;
    calcSheet.getRange(insertRowStart, 1, appendRows.length, calcHeaders.length).setValues(appendRows);
  }
  
  SpreadsheetApp.getUi().alert("✅ 恭喜！資料完全正確。\n請假標記與加班名單已成功匯入！\n接下來請點選「3. 計算便當數量」。");
  
  return true; // 🎯【熔斷機制】：成功結束，回傳 true[cite: 2]
}