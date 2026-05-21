// GETリクエストでWebアプリの画面（index.html）を表示する関数
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('令和8年度森の工場事業実施計画 調査フォーム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // iframeでの表示を許可（プレビュー用）
}

// ヘッダーの日本語変換用マップ
var headerMap = {
  targetYear: '対象年度',
  office: '事務所',
  organizationName: '事業体名',
  submissionTiming: '計画書 提出時期',
  new_districtName: '【新規・更新】地区名',
  new_area: '【新規・更新】面積(ha)',
  new_period: '【新規・更新】計画期間',
  change_districtName: '【変更(拡大・縮小)】地区名',
  change_existingArea: '【変更(拡大・縮小)】既存面積(ha)',
  change_modifiedArea: '【変更(拡大・縮小)】変更面積(ha)',
  change_totalArea: '【変更(拡大・縮小)】合計面積(ha)',
  change_period: '【変更(拡大・縮小)】計画期間',
  clearCut_districtName: '【皆伐・再造林】地区名',
  clearCut_area: '【皆伐・再造林】面積(ha)',
  clearCut_period: '【皆伐・再造林】計画期間'
};

// スプレッドシートに書き込む確実な順序（キーの配列）を定義
var columnOrder = [
  'targetYear',
  'office',
  'organizationName',
  'submissionTiming',
  'new_districtName',
  'new_area',
  'new_period',
  'change_districtName',
  'change_period',
  'change_existingArea',
  'change_modifiedArea',
  'change_totalArea',
  'clearCut_districtName',
  'clearCut_area',
  'clearCut_period'
];

// フォームから送信されたデータを受け取り、スプレッドシートに保存する関数
function processForm(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  try {
    // シートが真っ白な場合、指定した順序で1行目に項目名（見出し）を作成
    if (sheet.getLastRow() === 0) {
      var japaneseHeaders = columnOrder.map(function(key) {
        return headerMap[key] || key;
      });
      sheet.appendRow(japaneseHeaders);
    }
    
    // 指定した順序に従ってデータを配列化し、シートに追加
    var rowData = columnOrder.map(function(key) {
      return data[key] !== undefined ? data[key] : ""; // データが存在しない場合は空文字
    });
    
    sheet.appendRow(rowData);
    
    return "Success";
  } catch(error) {
    return "Error: " + error.toString();
  }
}

// スプレッドシートから既存データを読み取り、候補リストとして返す関数
function getSuggestions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // データ行がない（ヘッダーのみ、または空）場合は空のリストを返す
  if (lastRow <= 1) {
    return { offices: [], organizations: [], districts: [] };
  }
  
  // ヘッダー行とデータ行を取得
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  // 各列のインデックス（何列目か）を特定
  var colIdx = {
    office: headers.indexOf(headerMap['office']),
    org: headers.indexOf(headerMap['organizationName']),
    distNew: headers.indexOf(headerMap['new_districtName']),
    distChange: headers.indexOf(headerMap['change_districtName']),
    distClear: headers.indexOf(headerMap['clearCut_districtName'])
  };
  
  // Setを使って重複を省きながらリストを作成
  var offices = new Set();
  var organizations = new Set();
  var districts = new Set(); // 地区名は3箇所すべてから集約する
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (colIdx.office > -1 && row[colIdx.office]) offices.add(String(row[colIdx.office]).trim());
    if (colIdx.org > -1 && row[colIdx.org]) organizations.add(String(row[colIdx.org]).trim());
    
    if (colIdx.distNew > -1 && row[colIdx.distNew]) districts.add(String(row[colIdx.distNew]).trim());
    if (colIdx.distChange > -1 && row[colIdx.distChange]) districts.add(String(row[colIdx.distChange]).trim());
    if (colIdx.distClear > -1 && row[colIdx.distClear]) districts.add(String(row[colIdx.distClear]).trim());
  }
  
  // Setを配列に戻してフロントエンドに返す
  return {
    offices: Array.from(offices),
    organizations: Array.from(organizations),
    districts: Array.from(districts)
  };
}