/**
 * 自動販売機アンケート（Apps Script）
 *
 * - スプレッドシートの各カテゴリーシートから商品情報を取得
 * - ドライブフォルダから自販機メーカー画像を取得
 * - 回答を Responses シートへ保存（メールアドレスはバックエンドで取得）
 */

const SPREADSHEET_ID = '1xkg8vNscpcWTA6GA0VPxGTJCAH6LyvsYhq7VhOlDcXg';
const VENDOR_FOLDER_ID = '1nuVlneWO0PbmOapb_dTYoMNIWQFP45Gg';
const RESPONSE_SHEET_NAME = 'Responses';

const CATEGORY_SHEETS = [
  { key: 'コーヒー', sheetName: 'コーヒー' },
  { key: 'エナドリ', sheetName: 'エナドリ' },
  { key: '水', sheetName: '水' },
  { key: 'お茶', sheetName: 'お茶' },
  { key: '炭酸', sheetName: '炭酸' },
  { key: 'スポドリ', sheetName: 'スポドリ' },
  { key: 'その他(果汁等)', sheetName: 'その他(果汁等)' },
];

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('自動販売機アンケート')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInitialData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const categories = CATEGORY_SHEETS.map((cat) => ({
    name: cat.key,
    sheetName: cat.sheetName,
    items: readCategorySheet(ss, cat.sheetName),
  }));

  return {
    categories,
    vendorImages: getVendorImages(),
    email: Session.getActiveUser().getEmail() || '',
  };
}

function submitResponse(payload) {
  if (!payload) throw new Error('送信データが見つかりません。');
  const products = Array.isArray(payload.selectedProducts) ? payload.selectedProducts : [];
  if (products.length === 0) throw new Error('飲みたい商品を１つ以上選択してください。');
  if (products.length > 3) throw new Error('選択できる商品は３つまでです。');

  const vendor = payload.vendorImage || {};
  if (!vendor.id) throw new Error('自販機メーカーを１つ選択してください。');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME) || ss.insertSheet(RESPONSE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'タイムスタンプ',
      'メールアドレス',
      '選択商品（カテゴリ:メーカー/商品名/価格）',
      '自販機メーカー画像ID',
      '自販機メーカー名',
      '自由記入',
    ]);
  }

  const email = Session.getActiveUser().getEmail() || payload.email || 'anonymous';
  const productText = products
    .map((p) => {
      const maker = p.maker ? `${p.maker}/` : '';
      const price = p.price ? ` ${p.price}` : '';
      return `${p.category || ''}: ${maker}${p.product}${price}`;
    })
    .join('\n');

  sheet.appendRow([
    new Date(),
    email,
    productText,
    vendor.id,
    vendor.name || '',
    payload.freeText || '',
  ]);

  return { ok: true };
}

function readCategorySheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const rows = sheet.getDataRange().getValues();
  return rows
    .map((row) => {
      const product = row[0] ? String(row[0]).trim() : '';
      if (!product) return null;
      return {
        product,
        maker: row[1] ? String(row[1]).trim() : '',
        price: row[2] ? String(row[2]).trim() : '',
        imageUrl: buildDriveViewUrl(extractDriveId(String(row[3] || ''))) || '',
        category: sheetName,
      };
    })
    .filter(Boolean);
}

function getVendorImages() {
  const folder = DriveApp.getFolderById(VENDOR_FOLDER_ID);
  const iterator = folder.getFiles();
  const images = [];

  while (iterator.hasNext()) {
    const file = iterator.next();
    const mime = file.getMimeType() || '';
    if (!mime.startsWith('image/')) continue;
    images.push({
      id: file.getId(),
      name: file.getName(),
      url: buildDriveViewUrl(file.getId()),
    });
  }

  return images;
}

function buildDriveViewUrl(fileId) {
  if (!fileId) return '';
  return `https://drive.google.com/uc?export=view&id=${fileId}`;
}

function extractDriveId(url) {
  if (!url) return '';
  const idMatch = url.match(/[-\\w]{25,}/);
  return idMatch ? idMatch[0] : url;
}
