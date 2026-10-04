/**
 * 🔥 終極全自動工作流 (獨立放在 自動化.gs)
 * 核心邏輯：自動讀取 B1 日期 ➔ 帶入出勤 ➔ 計算狀態 ➔ 統計便當 ➔ 生成訊息
 * 🔐 已啟用熔斷安全機制：任一關卡防呆卡住，後續流程立刻中止。
 */
function runFullBentoWorkflow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("計算"); // 💡 請確保與您的計算工作表名稱一致
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("錯誤：找不到「計算」工作表！");
    return;
  }
  
  // 1️⃣ 自動讀取 B1 儲存格目前顯示的日期
  const targetDate = sheet.getRange("B1").getValue();
  if (!targetDate || !(targetDate instanceof Date)) {
    SpreadsheetApp.getUi().alert("❌ 自動化失敗：請確保「計算」頁籤的 B1 儲存格內有正確的日期！");
    return;
  }
  
  // 2️⃣ 將 B1 日期轉換為第一支程式所需的 yyyy-mm-dd 字串格式
  const yyyy = targetDate.getFullYear();
  const mm = String(targetDate.getMonth() + 1).padStart(2, '0');
  const dd = String(targetDate.getDate()).padStart(2, '0');
  const dateStr = `${yyyy}-${mm}-${dd}`;
  
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 3️⃣ ⚙️ 開始依序在背景執行您的四支核心主程式
    
    // 步驟 1：直接呼叫後台邏輯，帶入當日出勤人員 (跳過前端選日曆視窗)[cite: 2]
    if (executeAttendanceWorkflow(dateStr) === false) {
      console.log("[熔斷] 步驟 1 出錯或被防呆攔截，全自動中止。");
      return;
    }
    
    // 步驟 2：計算加班、請假狀態[cite: 2]
    if (processLeaveAndOvertime() === false) {
      console.log("[熔斷] 步驟 2 出錯或被防呆攔截，全自動中止。");
      return;
    }
    
    // 步驟 3：計算便當數量 (含分廠、外部加總與牛奶)[cite: 2]
    if (calculateBentoTotals() === false) {
      console.log("[熔斷] 步驟 3 出錯或被防呆攔截，全自動中止。");
      return;
    }
    
    // 步驟 4：自動產生訂購訊息範本 (寫入 R11、R12)[cite: 2]
    if (generateOrderMessages() === false) {
      console.log("[熔斷] 步驟 4 出錯或被防呆攔截，全自動中止。");
      return;
    }
    
    // 4️⃣ 🏁 完美大結局提示 (只有四關都綠燈才會走到這)[cite: 2]
    ui.alert("🏁 全自動刷新完畢！", "✨ 所有數據已更新！\n👉 請直接至 R11 與 R12 儲存格複製 LINE 訂購訊息。", ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert("❌ 全自動執行過程中發生系統崩潰錯誤：", error.toString(), ui.ButtonSet.OK);
  }
}