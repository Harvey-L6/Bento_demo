/**
 * Google Sheets 開啟時自動執行的內建觸發器
 * 用於建立自訂的上方功能表選單
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🍱 排班與便當管理系統')
    .addItem('🔥 一鍵全自動：刷新今日便當', 'runFullBentoWorkflow') // 👈 只要在您的選單裡加上這行就行了！
    .addSeparator() // 加條分隔線更好看
    .addItem('1. 🔄 帶入當日出勤人員', 'getWorkingEmployees')
    .addItem('2. 📝 計算加班、請假狀態', 'processLeaveAndOvertime')
    .addItem('3. 📊 計算便當數量', 'calculateBentoTotals')
    .addItem('4. ✉️ 產生訂購訊息', 'generateOrderMessages')
    .addToUi();
}