/**
 * 自動販売機アンケート（Apps Script）
 *
 * - スプレッドシートの各カテゴリーシートから商品情報を取得
 * - 回答を「回答」シートへ保存（メールアドレスはバックエンドで取得）
 */

const SPREADSHEET_ID = '1xkg8vNscpcWTA6GA0VPxGTJCAH6LyvsYhq7VhOlDcXg';
const PRODUCT_FOLDER_ID = '18fA4HRavIBTM2aPL-OqVaWhjRRgBhlKg';
const RESPONSE_SHEET_NAME = '回答';
const AGGREGATE_SHEET_NAME = '集計';
const RANK_POINTS = [3, 2, 1];

const categoryFolderCache = {};
const makerFolderCache = {};
const makerImageCache = {};
let productRootFolder;
let productImageMap;

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
  const sheetDataMap = buildSheetDataMap(ss);
  const categories = CATEGORY_SHEETS.map((cat) => ({
    name: cat.key,
    sheetName: cat.sheetName,
    items: buildCategoryItems(cat.sheetName, cat.key, sheetDataMap[cat.sheetName]),
  }));

  return {
    categories,
    email: Session.getActiveUser().getEmail() || '',
  };
}

function submitResponse(payload) {
  if (!payload) throw new Error('送信データが見つかりません。');
  const products = Array.isArray(payload.selectedProducts)
    ? payload.selectedProducts.filter((p) => p && p.product).slice(0, 3)
    : [];
  if (products.length !== 3) throw new Error('推しドリンクのTOP3を1位から3位まで選択してください。');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const email = Session.getActiveUser().getEmail() || payload.email || 'anonymous';

  appendResponseRow(ss, email, products, payload.freeText);
  appendAggregateRows(ss, email, products, payload.freeText);

  return { ok: true };
}

function appendResponseRow(ss, email, rankedProducts, freeText) {
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME) || ss.insertSheet(RESPONSE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'タイムスタンプ',
      'メールアドレス',
      '1位（カテゴリ/メーカー/商品名/価格）',
      '2位（カテゴリ/メーカー/商品名/価格）',
      '3位（カテゴリ/メーカー/商品名/価格）',
      '自由記入',
    ]);
  }

  const formatProduct = (product) => {
    if (!product) return '';
    const maker = product.maker ? `${product.maker}/` : '';
    const price = product.price ? ` ${product.price}` : '';
    return `${product.category || ''}: ${maker}${product.product || ''}${price}`.trim();
  };

  sheet.appendRow([
    new Date(),
    email,
    formatProduct(rankedProducts[0]),
    formatProduct(rankedProducts[1]),
    formatProduct(rankedProducts[2]),
    freeText || '',
  ]);
}

function appendAggregateRows(ss, email, rankedProducts, freeText) {
  const sheet = ss.getSheetByName(AGGREGATE_SHEET_NAME) || ss.insertSheet(AGGREGATE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'タイムスタンプ',
      'メールアドレス',
      '順位',
      'ポイント',
      'カテゴリ',
      'メーカー',
      '商品名',
      '価格',
      '自由記入',
    ]);
  }

  rankedProducts.forEach((product, idx) => {
    sheet.appendRow([
      new Date(),
      email,
      `${idx + 1}位`,
      RANK_POINTS[idx] || 0,
      (product && product.category) || '',
      (product && product.maker) || '',
      (product && product.product) || '',
      (product && product.price) || '',
      idx === 0 ? freeText || '' : '',
    ]);
  });
}

function buildCategoryItems(sheetName, folderName, sheetData) {
  const data = sheetData || {};
  const items = [];
  const matchedKeys = new Set();

  const folder = getCategoryFolder(folderName);
  if (folder) {
    collectCategoryFolderItems(folder, sheetName, data, matchedKeys, items);
  }

  Object.keys(data).forEach((key) => {
    if (matchedKeys.has(key)) return;
    const row = data[key];
    const imageUrl = normalizeImageUrl(row.imageUrl) || getProductImageUrlFromDrive(row.maker, row.product, sheetName);
    items.push({
      product: row.product,
      maker: row.maker,
      price: row.price,
      imageUrl,
      category: sheetName,
    });
  });

  return items;
}

function collectCategoryFolderItems(folder, sheetName, sheetData, matchedKeys, items, makerName) {
  const fileIterator = folder.getFiles();
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    const mime = file.getMimeType() || '';
    if (!mime.startsWith('image/')) continue;
    const displayName = stripExtension(file.getName());
    const key = normalizeProductKey(displayName);
    if (!key || matchedKeys.has(key)) continue;

    const row = sheetData[key];
    matchedKeys.add(key);

    items.push({
      product: (row && row.product) || displayName,
      maker: (row && row.maker) || makerName || '',
      price: (row && row.price) || '',
      imageUrl: buildDriveImageUrl(file.getId()) || buildDriveViewUrl(file.getId()),
      category: sheetName,
    });
  }

  const folderIterator = folder.getFolders();
  while (folderIterator.hasNext()) {
    const child = folderIterator.next();
    const childMakerName = child.getName() || makerName;
    collectCategoryFolderItems(child, sheetName, sheetData, matchedKeys, items, childMakerName);
  }
}

function buildSheetDataMap(ss) {
  return CATEGORY_SHEETS.reduce((acc, cat) => {
    acc[cat.sheetName] = buildSheetRowMap(ss, cat.sheetName);
    return acc;
  }, {});
}

function buildSheetRowMap(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {};

  const rows = sheet.getDataRange().getValues();
  const dataRows = rows.length && isHeaderRow(rows[0]) ? rows.slice(1) : rows;

  const map = {};
  dataRows.forEach((row) => {
    const product = row[0] ? String(row[0]).trim() : '';
    if (!product) return;
    const maker = row[1] ? String(row[1]).trim() : '';
    const price = row[2] ? String(row[2]).trim() : '';
    const imageUrl = row[3] ? String(row[3]).trim() : '';
    const key = normalizeProductKey(product);
    if (!key) return;
    map[key] = { product, maker, price, imageUrl };
  });
  return map;
}

function isHeaderRow(row) {
  const first = String(row[0] || '').trim();
  const second = String(row[1] || '').trim();
  const third = String(row[2] || '').trim();
  const headerWords = ['商品名', 'メーカー', 'メーカー名', '価格'];
  return headerWords.some((word) => first === word || second === word || third === word);
}

function getProductImageUrlFromDrive(rawMaker, product, categoryName) {
  if (!PRODUCT_FOLDER_ID || !product) return '';
  const normalizedProduct = normalizeProductKey(product);

  const rootImageMap = getProductImageMap();
  const rootFileId = findMatchingImageId(rootImageMap, normalizedProduct);
  if (rootFileId) return buildDriveImageUrl(rootFileId) || buildDriveViewUrl(rootFileId);

  const maker = normalizeMakerKey(rawMaker);
  if (!maker) return '';

  const cacheKey = buildMakerCacheKey(maker, categoryName);
  const folder = getMakerFolder(maker, categoryName);
  if (!folder) return '';

  if (!makerImageCache[cacheKey]) {
    makerImageCache[cacheKey] = buildMakerImageMap(folder);
  }

  const imageMap = makerImageCache[cacheKey];
  const fileId = findMatchingImageId(imageMap, normalizedProduct);
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

function getProductImageMap() {
  if (productImageMap !== undefined) return productImageMap;
  const root = getProductRootFolder();
  if (!root) {
    productImageMap = null;
    return null;
  }

  const map = {};
  const iterator = root.getFiles();
  while (iterator.hasNext()) {
    const file = iterator.next();
    const mime = file.getMimeType() || '';
    if (!mime.startsWith('image/')) continue;
    const nameKey = normalizeProductKey(file.getName());
    if (!nameKey) continue;
    map[nameKey] = file.getId();
  }

  productImageMap = map;
  return productImageMap;
}

function findMatchingImageId(imageMap, normalizedProduct) {
  if (!imageMap || !normalizedProduct) return '';

  if (imageMap[normalizedProduct]) return imageMap[normalizedProduct];

  const keys = Object.keys(imageMap);

  const forwardMatch = keys
    .filter((key) => key.startsWith(normalizedProduct))
    .sort((a, b) => a.length - b.length);
  if (forwardMatch.length) return imageMap[forwardMatch[0]];

  const reverseMatch = keys
    .filter((key) => normalizedProduct.startsWith(key))
    .sort((a, b) => b.length - a.length);
  if (reverseMatch.length) return imageMap[reverseMatch[0]];

  return '';
}

function getMakerFolder(maker, categoryName) {
  if (!maker) return null;
  const cacheKey = buildMakerCacheKey(maker, categoryName);
  if (makerFolderCache.hasOwnProperty(cacheKey)) return makerFolderCache[cacheKey];

  const root = getProductRootFolder();
  if (!root) {
    makerFolderCache[cacheKey] = null;
    return null;
  }

  const rootIterator = root.getFoldersByName(maker);
  if (rootIterator.hasNext()) {
    makerFolderCache[cacheKey] = rootIterator.next();
    return makerFolderCache[cacheKey];
  }

  const candidates = [];
  if (categoryName) {
    const directMatch = getCategoryFolder(categoryName);
    if (directMatch) candidates.push(directMatch);
  }
  CATEGORY_SHEETS.forEach((cat) => {
    if (categoryName && cat.key === categoryName) return;
    const folder = getCategoryFolder(cat.key);
    if (folder) candidates.push(folder);
  });

  for (let i = 0; i < candidates.length; i++) {
    const categoryFolder = candidates[i];
    if (!categoryFolder) continue;
    const nestedIterator = categoryFolder.getFoldersByName(maker);
    if (nestedIterator.hasNext()) {
      makerFolderCache[cacheKey] = nestedIterator.next();
      return makerFolderCache[cacheKey];
    }
  }

  makerFolderCache[cacheKey] = null;
  return makerFolderCache[cacheKey];
}

function buildMakerCacheKey(maker, categoryName) {
  return `${categoryName || 'root'}::${maker || ''}`;
}

function getCategoryFolder(categoryName) {
  if (!categoryName) return null;
  if (categoryFolderCache[categoryName]) return categoryFolderCache[categoryName];
  const root = getProductRootFolder();
  if (!root) return null;

  const iterator = root.getFoldersByName(categoryName);
  categoryFolderCache[categoryName] = iterator.hasNext() ? iterator.next() : null;
  return categoryFolderCache[categoryName];
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

function stripExtension(name) {
  if (!name) return '';
  return String(name).replace(/\.[^.]+$/, '').trim();
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

function buildDrivePreviewUrl(fileId) {
  if (!fileId) return '';
  return `https://drive.google.com/file/d/${fileId}/preview`;
}

function buildDriveViewUrl(fileId) {
  if (!fileId) return '';
  return `https://drive.google.com/uc?export=view&id=${fileId}`;
}

function buildDriveImageUrl(fileId) {
  if (!fileId) return '';
  // `thumbnail` エンドポイントは閲覧権限が無いと 403 になるケースが増えたため、
  // 公開設定が有効なファイルならブラウザから直接参照できる preview リンクを優先する。
  return buildDrivePreviewUrl(fileId) || buildDriveViewUrl(fileId);
}
