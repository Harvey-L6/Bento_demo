// ==================== 🛠️ 便當與牛奶計算欄位與設定維護區 ====================
const SHEET_CALC_BENTO     = "計算";
const ROW_CALC_HEADER_B    = 2;     // 標題列所在的列號 (第 2 列)
const ROW_CALC_DATA_START_B = 3;    // 資料開始尋找的列號 (第 3 列)

// 🎯 右側數量計算表格的固定列號設定 (對應 N ~ S 欄的列號)
const ROW_FACTORY_1        = 3;     // 「一廠」數據寫入的列號
const ROW_FACTORY_2        = 4;     // 「二廠」數據寫入的列號
const ROW_INIT_TOTAL       = 5;     // 「初始加總」所在的列號 (一廠 + 二廠)
const ROW_EXTRA_VALUE      = 6;     // 「外部(算一廠)」所在的列號 (供人員手動填寫)
const ROW_TOTAL_VALUE      = 7;     // 「Total」所在的列號 (初始加總 + 外部)
  
// 判斷便當與牛奶的來源欄位名稱
const HDR_B_FACTORY        = "廠別";     
const HDR_B_EMP_ID         = "員工編號"; 
const HDR_B_NIGHT_SHIFT    = "夜班";
const HDR_B_VEGETARIAN     = "素食";
const HDR_B_NO_RICE        = "不要飯";
const HDR_B_NO_MILK        = "不要牛奶"; 
const HDR_B_LEAVE          = "請假";     

// 統計結果要寫入的目標欄位名稱 (會自動動態對應最新 N ~ S 欄)
const HDR_TOTAL_MEAT       = "葷";
const HDR_TOTAL_MEAT_NO_R  = "葷-不飯";
const HDR_TOTAL_VEG        = "素";
const HDR_TOTAL_VEG_NO_R   = "素-不飯";
const HDR_TOTAL_BENTO_SUM  = "便當合計"; // ✨ 便當合計欄位名稱
const HDR_TOTAL_MILK       = "牛奶";     
// =================================================================

function calculateBentoTotals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName(SHEET_CALC_BENTO);
  
  if (!calcSheet) {
    SpreadsheetApp.getUi().alert("錯誤：找不到「" + SHEET_CALC_BENTO + "」工作表！");
    return false; // 🎯【熔斷機制】：改為回傳 false[cite: 2]
  }
  
  const lastRow = calcSheet.getLastRow();
  if (lastRow < ROW_CALC_DATA_START_B) {
    SpreadsheetApp.getUi().alert("目前「計算」頁籤中沒有員工資料，無法計算。");
    return false; // 🎯【熔斷機制】：改為回傳 false[cite: 2]
  }
  
  const maxColumns = calcSheet.getLastColumn();
  const headers = calcSheet.getRange(ROW_CALC_HEADER_B, 1, 1, maxColumns).getValues()[0];
  
  // 尋找來源欄位索引
  const idxFactory  = headers.indexOf(HDR_B_FACTORY); 
  const idxEmpId    = headers.indexOf(HDR_B_EMP_ID);
  const idxNight    = headers.indexOf(HDR_B_NIGHT_SHIFT);
  const idxVeg      = headers.indexOf(HDR_B_VEGETARIAN);
  const idxNoRice   = headers.indexOf(HDR_B_NO_RICE);
  const idxNoMilk   = headers.indexOf(HDR_B_NO_MILK); 
  const idxLeave    = headers.indexOf(HDR_B_LEAVE); 
  
  // 尋找寫入欄位索引
  const idxTMeat     = headers.indexOf(HDR_TOTAL_MEAT);
  const idxTMeatNoR  = headers.indexOf(HDR_TOTAL_MEAT_NO_R);
  const idxTVeg      = headers.indexOf(HDR_TOTAL_VEG);
  const idxTVegNoR   = headers.indexOf(HDR_TOTAL_VEG_NO_R);
  const idxTTotalBento = headers.indexOf(HDR_TOTAL_BENTO_SUM); // 新增便當合計索引
  const idxTMilk     = headers.indexOf(HDR_TOTAL_MILK); 
  
  if ([idxFactory, idxEmpId, idxNight, idxVeg, idxNoRice, idxNoMilk, idxLeave, idxTMeat, idxTMeatNoR, idxTVeg, idxTVegNoR, idxTTotalBento, idxTMilk].includes(-1)) {
    SpreadsheetApp.getUi().alert("錯誤：計算頁籤的標題名稱不完整，請檢查維護區的欄位名稱（如便當合計、牛奶）是否與工作表第 2 列一致！");
    return false; // 🎯【熔斷機制】：改為回傳 false[cite: 2]
  }
  
  // 抓取整張表的資料進行統計
  const data = calcSheet.getRange(ROW_CALC_DATA_START_B, 1, lastRow - ROW_CALC_DATA_START_B + 1, maxColumns).getValues();
  
  // 初始化一廠與二廠的計數器物件
  let f1 = { meat: 0, meatNoRice: 0, veg: 0, vegNoRice: 0, milk: 0 };
  let f2 = { meat: 0, meatNoRice: 0, veg: 0, vegNoRice: 0, milk: 0 };
  
  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    if (!row[idxEmpId]) continue; // 沒有員工編號就跳過
    
    let factoryStr   = String(row[idxFactory]).trim(); 
    let isNightShift = (row[idxNight] === "Y" || row[idxNight] === "是" || row[idxNight] === true);
    let isVegetarian = (row[idxVeg] === "Y" || row[idxVeg] === "是" || row[idxVeg] === true);
    let isNoRice     = (row[idxNoRice] === "Y" || row[idxNoRice] === "是" || row[idxNoRice] === true);
    let isNoMilk     = (row[idxNoMilk] === "Y" || row[idxNoMilk] === "是" || row[idxNoMilk] === true); 
    let isLeave      = (row[idxLeave] === "Y" || row[idxLeave] === "是" || row[idxLeave] === true); 
    
    // 牛奶判定：在名單內、沒請假，且沒有註記「不要牛奶」
    let needMilk = (!isLeave && !isNoMilk);
    
    // 便當判定：沒請假且不是夜班
    let needBento = (!isNightShift && !isLeave);
    
    // 分流計算：區分一廠與二廠
    if (factoryStr === "1") {
      if (needMilk) f1.milk++;
      if (needBento) {
        if (isVegetarian) {
          if (isNoRice) f1.vegNoRice++; else f1.veg++;
        } else {
          if (isNoRice) f1.meatNoRice++; else f1.meat++;
        }
      }
    } else if (factoryStr === "2") {
      if (needMilk) f2.milk++;
      if (needBento) {
        if (isVegetarian) {
          if (isNoRice) f2.vegNoRice++; else f2.veg++;
        } else {
          if (isNoRice) f2.meatNoRice++; else f2.meat++;
        }
      }
    }
  }
  
  // 🏪 讀取管理員手動在黃色區塊填寫的「外部(算一廠)」各分項數值 (第 6 列)
  let extraMeat     = Number(calcSheet.getRange(ROW_EXTRA_VALUE, idxTMeat + 1).getValue()) || 0;
  let extraMeatNoR  = Number(calcSheet.getRange(ROW_EXTRA_VALUE, idxTMeatNoR + 1).getValue()) || 0;
  let extraVeg      = Number(calcSheet.getRange(ROW_EXTRA_VALUE, idxTVeg + 1).getValue()) || 0;
  let extraVegNoR   = Number(calcSheet.getRange(ROW_EXTRA_VALUE, idxTVegNoR + 1).getValue()) || 0;
  let extraMilk     = Number(calcSheet.getRange(ROW_EXTRA_VALUE, idxTMilk + 1).getValue()) || 0;

  // 🧮 【智慧加總鎖】自動將外部手動輸入的 4 項便當相加，得到外部的「便當合計」
  let extraBentoTotal = extraMeat + extraMeatNoR + extraVeg + extraVegNoR;
  // 自動回寫外部的便當合計到第 6 列 R 欄，省去人員手動計算
  calcSheet.getRange(ROW_EXTRA_VALUE, idxTTotalBento + 1).setValue(extraBentoTotal);

  // 🎯 統整 6 大項目（含新成員：便當合計）的橫向與直向加總對應表
  const targetColumns = [
    { idx: idxTMeat,    f1Val: f1.meat,       f2Val: f2.meat,       extra: extraMeat },
    { idx: idxTMeatNoR, f1Val: f1.meatNoRice, f2Val: f2.meatNoRice, extra: extraMeatNoR },
    { idx: idxTVeg,    f1Val: f1.veg,        f2Val: f2.veg,        extra: extraVeg },
    { idx: idxTVegNoR,  f1Val: f1.vegNoRice,  f2Val: f2.vegNoRice,  extra: extraVegNoR },
    // 便當合計項目的基礎值，直接由前 4 項的加總衍生而來
    { idx: idxTTotalBento, 
      f1Val: f1.meat + f1.meatNoRice + f1.veg + f1.vegNoRice,       
      f2Val: f2.meat + f2.meatNoRice + f2.veg + f2.vegNoRice,       
      extra: extraBentoTotal 
    },
    { idx: idxTMilk,    f1Val: f1.milk,       f2Val: f2.milk,       extra: extraMilk }
  ];
  
  // 🚀 執行各欄位的縱向與橫向自動填入
  targetColumns.forEach(item => {
    let colNum = item.idx + 1; // 轉換為真實欄位序號 (N ~ S 欄)
    
    // 1. 寫入「一廠」基礎值 (第 3 列)
    calcSheet.getRange(ROW_FACTORY_1, colNum).setValue(item.f1Val);
    
    // 2. 寫入「二廠」基礎值 (第 4 列)
    calcSheet.getRange(ROW_FACTORY_2, colNum).setValue(item.f2Val);
    
    // 3. 計算並寫入「初始加總」 (第 5 列 = 一廠 + 二廠)
    let initTotal = item.f1Val + item.f2Val;
    calcSheet.getRange(ROW_INIT_TOTAL, colNum).setValue(initTotal);
    
    // 4. 計算並寫入最終的「Total」 (第 7 列 = 初始加總 + 外部外加)
    let totalVal = initTotal + item.extra;
    calcSheet.getRange(ROW_TOTAL_VALUE, colNum).setValue(totalVal);
  });
  
  // 彈出成功提示
  SpreadsheetApp.getUi().alert(
    "🍱 便當與 🥛 牛奶分廠統計完成！\n" +
    "-------------------------\n" +
    "✨ 本次更新亮點：\n" +
    "▪ 已新增「便當合計」直橫向交叉運算。\n" +
    "▪ 偵測到「外部(算一廠)」手動輸入，已自動為其產生便當合計：" + extraBentoTotal + " 個。\n" +
    "▪ 初始加總與最下方 Total 數據已同步更新完畢！"
  );

  return true; // 🎯【熔斷機制】：成功結束，回傳 true[cite: 2]
}