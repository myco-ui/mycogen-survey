/**
 * ══════════════════════════════════════════════════════════════
 *  Mycogen × Mycodrops — Google Sheets 後台接收程式
 *  Google Apps Script Web App
 * ══════════════════════════════════════════════════════════════
 *
 *  設定步驟（5 分鐘完成）：
 *
 *  1. 開啟 Google 試算表（Sheets），建立兩個分頁：
 *       - 「購後回饋」
 *       - 「新訪客探索」
 *
 *  2. 點選上方選單：「擴充功能」→「Apps Script」
 *
 *  3. 將本檔案的全部內容貼入 Apps Script 編輯器，取代預設內容
 *
 *  4. 點選「部署」→「新增部署」→ 類型選「網頁應用程式」
 *       - 執行身份：我（您的 Google 帳號）
 *       - 存取對象：任何人
 *     → 複製產生的「網頁應用程式 URL」
 *
 *  5. 打開 mycogen-survey.html，找到這一行：
 *       const SHEET_URL = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE';
 *     將引號內的文字換成您複製的 URL。
 *
 *  完成！每次問卷送出，資料即自動寫入對應的分頁。
 * ══════════════════════════════════════════════════════════════
 */

// ── 設定：填入您的 Google 試算表 ID ─────────────────────────────
// 從試算表網址取得：
// https://docs.google.com/spreadsheets/d/【這裡是 SPREADSHEET_ID】/edit
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

// ── 問卷一欄位順序 ───────────────────────────────────────────────
const S1_HEADERS = [
  'timestamp',
  'nickname',
  'email',
  'age',
  'gender',
  'occupation',
  'source',
  's1_product',
  's1_duration',
  's1_frequency',
  's1_effects',
  's1_satisfaction',
  's1_taste',
  's1_taste_comment',
  's1_improve',
  's1_improve_other',
  's1_repurchase',
  's1_nps',
  's1_extra'
];

// ── 問卷二欄位順序 ───────────────────────────────────────────────
const S2_HEADERS = [
  'timestamp',
  'nickname',
  'email',
  'age',
  'gender',
  'source',
  's2_concern',
  's2_lifestyle',
  's2_sleep',
  's2_goal',
  's2_supplement',
  's2_recommended',
  's2_scores'
];

// ── 主接收函式 ───────────────────────────────────────────────────
function doPost(e) {
  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (data.survey === '購後回饋') {
      writeRow(ss, '購後回饋', S1_HEADERS, data);
    } else if (data.survey === '新訪客探索') {
      writeRow(ss, '新訪客探索', S2_HEADERS, data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Error: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 輔助：將資料寫入對應分頁 ────────────────────────────────────
function writeRow(ss, sheetName, headers, data) {
  let sheet = ss.getSheetByName(sheetName);

  // 若分頁不存在則自動建立並加上標題列
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers.map(h => columnLabel(h)));
    // 凍結標題列
    sheet.setFrozenRows(1);
    // 標題列底色
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground('#8E6B8E')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
  }

  // 將多選陣列轉為逗號字串
  const row = headers.map(h => {
    const v = data[h];
    if (Array.isArray(v)) return v.join(', ');
    return v !== undefined ? v : '';
  });

  sheet.appendRow(row);
}

// ── 欄位中文標籤對照 ────────────────────────────────────────────
function columnLabel(key) {
  const map = {
    timestamp:         '填寫時間',
    nickname:          '暱稱',
    email:             '電子信箱',
    age:               '年齡範圍',
    gender:            '生理性別',
    occupation:        '職業',
    source:            '來源管道',
    s1_product:        '購買產品',
    s1_duration:       '使用時長',
    s1_frequency:      '使用頻率',
    s1_effects:        '感受到的改善',
    s1_satisfaction:   '整體滿意度（星）',
    s1_taste:          '口感評分（星）',
    s1_taste_comment:  '口感評語',
    s1_improve:        '期望改善方向',
    s1_improve_other:  '其他建議',
    s1_repurchase:     '回購意願',
    s1_nps:            'NPS 推薦分數',
    s1_extra:          '其他留言',
    s2_concern:        '主要生活困擾',
    s2_lifestyle:      '生活型態',
    s2_sleep:          '睡眠品質',
    s2_goal:           '最期待的改變',
    s2_supplement:     '保健品習慣',
    s2_recommended:    '推薦結果',
    s2_scores:         '各菇吻合度'
  };
  return map[key] || key;
}

// ── GET 請求（測試用）──────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput('Mycogen Survey Backend — Active ✅')
    .setMimeType(ContentService.MimeType.TEXT);
}
