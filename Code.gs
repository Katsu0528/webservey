/**
 * 自動販売機アンケート（Apps Script）
 *
 * - スプレッドシートの各カテゴリーシートから商品情報を取得
 * - ドライブフォルダから自販機メーカー画像を取得
 * - 回答を Responses シートへ保存（メールアドレスはバックエンドで取得）
 */

const SPREADSHEET_ID = '1xkg8vNscpcWTA6GA0VPxGTJCAH6LyvsYhq7VhOlDcXg';
const PRODUCT_FOLDER_ID = '18fA4HRavIBTM2aPL-OqVaWhjRRgBhlKg';
const VENDOR_FOLDER_ID = '1AJd4BTFTVrLNep44PDz1AuwSF_5TFxdx';
const RESPONSE_SHEET_NAME = 'Responses';

const makerFolderCache = {};
const makerImageCache = {};
const imageDataCache = {};
let productRootFolder;

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
  const dataRows = rows.length && isHeaderRow(rows[0]) ? rows.slice(1) : rows;

  return dataRows
    .map((row) => {
      const product = row[0] ? String(row[0]).trim() : '';
      if (!product) return null;
      const maker = row[1] ? String(row[1]).trim() : '';
      const price = row[2] ? String(row[2]).trim() : '';

      const imageUrlFromSheet = normalizeImageUrl(row[3]);
      const imageUrl = imageUrlFromSheet || getProductImageUrlFromDrive(maker, product);
      return {
        product,
        maker,
        price,
        imageUrl,
        category: sheetName,
      };
    })
    .filter(Boolean);
}

function isHeaderRow(row) {
  const first = String(row[0] || '').trim();
  const second = String(row[1] || '').trim();
  const third = String(row[2] || '').trim();
  const headerWords = ['商品名', 'メーカー', 'メーカー名', '価格'];
  return headerWords.some((word) => first === word || second === word || third === word);
}

function getVendorImages() {
  const folder = DriveApp.getFolderById(VENDOR_FOLDER_ID);
  const iterator = folder.getFiles();
  const images = [];

  while (iterator.hasNext()) {
    const file = iterator.next();
    const mime = file.getMimeType() || '';
    if (!mime.startsWith('image/')) continue;
    const url = buildDriveImageUrl(file.getId()) || buildDriveViewUrl(file.getId());
    if (!url) continue;
    images.push({
      id: file.getId(),
      name: file.getName(),
      url,
    });
  }

  return images;
}

function buildDriveViewUrl(fileId) {
  if (!fileId) return '';
  return `https://drive.google.com/uc?export=view&id=${fileId}`;
}

function getProductImageUrlFromDrive(rawMaker, product) {
  const maker = normalizeMakerKey(rawMaker);
  if (!PRODUCT_FOLDER_ID || !maker || !product) return '';
  const normalizedProduct = normalizeProductKey(product);
  const folder = getMakerFolder(maker);
  if (!folder) return '';

  if (!makerImageCache[maker]) {
    makerImageCache[maker] = buildMakerImageMap(folder);
  }

  const imageMap = makerImageCache[maker];
  const fileId = imageMap[normalizedProduct];
  return fileId ? buildDriveImageUrl(fileId) || buildDriveViewUrl(fileId) : '';
}

function buildMakerImageMap(folder) {
  const map = {};
  const iterator = folder.getFiles();
  while (iterator.hasNext()) {
    const file = iterator.next();
    const mime = file.getMimeType() || '';
    if (!mime.startsWith('image/')) continue;
    const nameKey = normalizeProductKey(file.getName());
    if (!nameKey) continue;
    map[nameKey] = file.getId();
  }
  return map;
}

function getMakerFolder(maker) {
  if (makerFolderCache[maker]) return makerFolderCache[maker];
  const root = getProductRootFolder();
  if (!root) return null;

  const iterator = root.getFoldersByName(maker);
  makerFolderCache[maker] = iterator.hasNext() ? iterator.next() : null;
  return makerFolderCache[maker];
}

function getProductRootFolder() {
  if (productRootFolder !== undefined) return productRootFolder;
  if (!PRODUCT_FOLDER_ID) {
    productRootFolder = null;
    return null;
  }
  try {
    productRootFolder = DriveApp.getFolderById(PRODUCT_FOLDER_ID);
    return productRootFolder;
  } catch (e) {
    productRootFolder = null;
    return null;
  }
}

function normalizeProductKey(name) {
  if (!name) return '';
  return String(name)
    .trim()
    .replace(/\.[^.]+$/, '')
    .toLowerCase();
}

function normalizeMakerKey(name) {
  if (!name) return '';
  return String(name).trim();
}

function normalizeImageUrl(rawValue) {
  if (!rawValue) return '';
  const value = String(rawValue).trim();

  // data URL はそのまま返却
  if (value.startsWith('data:')) return value;

  // 通常の http(s) 画像 URL はそのまま利用
  if (new RegExp('^https?://', 'i').test(value) && !value.includes('drive.google.com')) {
    return value;
  }

  // Drive の共有 URL / ID を正規化
  const driveId = extractDriveId(value);
  if (driveId) return buildDriveImageUrl(driveId) || buildDriveViewUrl(driveId);

  return '';
}

function extractDriveId(url) {
  if (!url) return '';
  const idMatch = url.match(/[-\w]{25,}/);
  return idMatch ? idMatch[0] : '';
}

function buildDriveImageUrl(fileId) {
  if (!fileId) return '';
  return getDriveImageDataUrl(fileId);
}

function getDriveImageDataUrl(fileId) {
  if (!fileId) return '';
  if (imageDataCache[fileId] !== undefined) return imageDataCache[fileId];
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const mimeType = blob.getContentType() || 'image/png';
    const dataUrl = `data:${mimeType};base64,${Utilities.base64Encode(blob.getBytes())}`;
    imageDataCache[fileId] = dataUrl;
    return dataUrl;
  } catch (e) {
    imageDataCache[fileId] = '';
    return '';
  }
}
