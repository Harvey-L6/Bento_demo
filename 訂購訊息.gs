// ==================== 🛠️ 訂購訊息生成欄位設定維護區 ====================
const SHEET_NAME          = "計算";    // 工作表名稱
const CELL_DATE           = "B1";      // 日期所在的儲存格

// 🍱 便當統計表格位置 (列號)
const ROW_F1_DATA         = 3;         // 一廠資料在第 3 列
const ROW_F2_DATA         = 4;         // 二廠資料在第 4 列
const ROW_EXTERNAL_DATA   = 6;         // 外部(算一廠)資料在第 6 列

// 🍙 側邊附餐與價格位置 (儲存格)
const CELL_F1_RICE        = "N11";     // 一廠白飯數量
const CELL_F1_SOUP        = "O11";     // 一廠鹹湯數量
const CELL_F2_RICE        = "N12";     // 二廠白飯數量
const CELL_F2_SOUP        = "O12";     // 二廠鹹湯數量
const CELL_BENTO_PRICE    = "N15";     // 便當價格儲存格

// 📝 訊息寫入目標格子
const CELL_MSG_BENTO      = "R11";     // 便當訊息寫入位置
const CELL_MSG_MILK       = "R12";     // 牛奶訊息寫入位置
// =================================================================

function generateOrderMessages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("錯誤：找不到「" + SHEET_NAME + "」工作表！");
    return false; // 🎯【熔斷機制】：改為回傳 false[cite: 2]
  }
  
  // 1️⃣ 讀取日期並轉換格式 (例如：7/2 (四))
  const rawDate = sheet.getRange(CELL_DATE).getValue();
  const dateStr = formatOrderDate(rawDate);
  
  // 2️⃣ 讀取便當價格、白飯、鹹湯
  const bentoPrice = sheet.getRange(CELL_BENTO_PRICE).getValue() || 0;
  const f1Rice     = sheet.getRange(CELL_F1_RICE).getValue() || 0;
  const f1Soup     = sheet.getRange(CELL_F1_SOUP).getValue() || 0;
  const f2Rice     = sheet.getRange(CELL_F2_RICE).getValue() || 0;
  const f2Soup     = sheet.getRange(CELL_F2_SOUP).getValue() || 0;
  
  // 3️⃣ 尋找動態欄位索引 (防呆，避免您未來抽換 N~S 欄順序)
  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxMeat     = headers.indexOf("葷");
  const idxMeatNoR  = headers.indexOf("葷-不飯");
  const idxVeg      = headers.indexOf("素");
  const idxMilk     = headers.indexOf("牛奶");
  
  if ([idxMeat, idxMeatNoR, idxVeg, idxMilk].includes(-1)) {
    SpreadsheetApp.getUi().alert("錯誤：找不到 葷 / 葷-不飯 / 素 / 牛奶 的標題欄位，請確認第 2 列標題未被修改！");
    return false; // 🎯【熔斷機制】：改為回傳 false[cite: 2]
  }
  
  // 4️⃣ 讀取各列真實數據 (欄位索引 + 1 轉為真實欄號)
  const getNum = (row, idx) => Number(sheet.getRange(row, idx + 1).getValue()) || 0;
  
  let r3Meat = getNum(ROW_F1_DATA, idxMeat), r3MeatNoR = getNum(ROW_F1_DATA, idxMeatNoR), r3Veg = getNum(ROW_F1_DATA, idxVeg), r3Milk = getNum(ROW_F1_DATA, idxMilk);
  let r4Meat = getNum(ROW_F2_DATA, idxMeat), r4MeatNoR = getNum(ROW_F2_DATA, idxMeatNoR), r4Veg = getNum(ROW_F2_DATA, idxVeg), r4Milk = getNum(ROW_F2_DATA, idxMilk);
  let r6Meat = getNum(ROW_EXTERNAL_DATA, idxMeat), r6MeatNoR = getNum(ROW_EXTERNAL_DATA, idxMeatNoR), r6Veg = getNum(ROW_EXTERNAL_DATA, idxVeg), r6Milk = getNum(ROW_EXTERNAL_DATA, idxMilk);
  
  // 5️⃣ 核心商業邏輯：將「外部」數量直接併入「一廠」
  const f1Normal = r3Meat + r6Meat;        // 葷 ➔ 《正常》
  const f1NoRice = r3MeatNoR + r6MeatNoR;  // 葷-不飯 ➔ 《不飯》
  const f1Veg    = r3Veg + r6Veg;          // 素 ➔ 《素食》
  const f1Milk   = r3Milk + r6Milk;        // 一廠牛奶總計
  const f1BentoTotal = f1Normal + f1NoRice + f1Veg;
  
  const f2Normal = r4Meat;
  const f2NoRice = r4MeatNoR;
  const f2Veg    = r4Veg;
  const f2Milk   = r4Milk;
  const f2BentoTotal = f2Normal + f2NoRice + f2Veg;
  
  const grandTotalBento = f1BentoTotal + f2BentoTotal;
  const grandTotalMilk  = f1Milk + f2Milk;
  
  // 6️⃣ 🛠️ 建立【便當訂購訊息】文字範本
  let bentoMsg = `${dateStr} 中餐 ${grandTotalBento}個*${bentoPrice}元\n10：50前送到-早到OK\n\n`;
  
  // 一廠區塊生成
  bentoMsg += `❶【一廠】合計 ${f1BentoTotal}\n`;
  if (f1Veg > 0)    bentoMsg += `《素食》* ${f1Veg} 個\n`;
  if (f1Normal > 0) bentoMsg += `《正常》* ${f1Normal} 個\n`;
  if (f1NoRice > 0) bentoMsg += `《不飯》* ${f1NoRice} 個\n`;
  if (f1Rice > 0)   bentoMsg += `🍙 白飯* ${f1Rice} 包\n`;
  if (f1Soup > 0)   bentoMsg += `🍲 鹹湯* ${f1Soup} 碗\n`;
  
  bentoMsg += `...........................................\n\n`;
  
  // 二廠區塊生成
  bentoMsg += `❷【二廠】合計 ${f2BentoTotal}\n`;
  if (f2Veg > 0)    bentoMsg += `《素食》* ${f2Veg} 個\n`;
  if (f2Normal > 0) bentoMsg += `《正常》* ${f2Normal} 個\n`;
  if (f2NoRice > 0) bentoMsg += `《不飯》* ${f2NoRice} 個\n`;
  if (f2Rice > 0)   bentoMsg += `🍙 白飯* ${f2Rice} 包\n`;
  if (f2Soup > 0)   bentoMsg += `🍲 鹹湯* ${f2Soup} 碗\n`;
  
  bentoMsg = bentoMsg.trim(); // 去除末尾空白換行
  
  // 7️⃣ 🛠️ 建立【牛奶訂購訊息】文字範本
  let milkMsg = `${dateStr}\n` +
                `一廠數量：${f1Milk}\n` +
                `二廠數量：${f2Milk}\n` +
                `合計：${grandTotalMilk}`;
  
  // 8️⃣ 🚀 精準回寫至您指定的 R11 與 R12 儲存格
  sheet.getRange(CELL_MSG_BENTO).setValue(bentoMsg);
  sheet.getRange(CELL_MSG_MILK).setValue(milkMsg);
  
  // 提示視窗
  SpreadsheetApp.getUi().alert("✨ 訂購訊息範本已自動生成！\n👉 請至 R11(便當) 與 R12(牛奶) 複製使用。");

  return true; // 🎯【熔斷機制】：成功結束，回傳 true[cite: 2]
}

/**
 * 📅 日期格式化小助手：將試算表日期轉換為 M/D (星期)
 */
function formatOrderDate(dateVal) {
  if (!dateVal) return "7/2 (四)";
  const d = new Date(dateVal);
  if (isNaN(d.getTime())) return String(dateVal);
  
  const month = d.getMonth() + 1;
  const date = d.getDate();
  const weekdays = ["日", "一", "二", "三", "四", "五", "六"];
  const dayName = weekdays[d.getDay()];
  
  return `${month}/${date} (${dayName})`;
}