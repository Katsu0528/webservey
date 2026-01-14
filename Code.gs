/**
 * 自動販売機アンケート（Apps Script）
 *
 * - スプレッドシートの各カテゴリーシートから商品情報を取得
 * - 回答を「回答」シートへ保存（メールアドレスはバックエンドで取得）
 */

const SPREADSHEET_ID = '13MM299vAlOK437M7qP3BhFILD2_-rh8KG8doveVZcFQ';
const PRODUCT_FOLDER_ID = '18fA4HRavIBTM2aPL-OqVaWhjRRgBhlKg';
const MAKER_BADGE_FOLDER_ID = '1AJd4BTFTVrLNep44PDz1AuwSF_5TFxdx';
const RESPONSE_SHEET_NAME = '回答';
const AGGREGATE_SHEET_NAME = '集計';
const FREE_TEXT_SHEET_NAME = '自由記入';
const RANK_POINTS = [3, 2, 1];

const categoryFolderCache = {};
const makerFolderCache = {};
const makerImageCache = {};
let productRootFolder;
let productImageMap;
let makerBadgeFolder;
let makerBadgeImageMap;

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
  const categories = buildCategoriesFromDrive();

  return {
    categories,
    email: Session.getActiveUser().getEmail() || '',
  };
}

function submitResponse(payload) {
  if (!payload) throw new Error('送信データが見つかりません。');
  const ranking = normalizeRanking(payload.selectedProducts);
  if (!ranking) throw new Error('1位から3位まで選択してください。');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const email = Session.getActiveUser().getEmail() || payload.email || 'anonymous';

  appendResponseRow(ss, email, ranking, payload.freeText);
  appendAggregateRows(ss, email, ranking, payload.freeText);
  appendFreeTextRow(ss, email, payload.freeText);

  return { ok: true };
}

function normalizeRanking(list) {
  const items = Array.isArray(list) ? list.filter((p) => p && p.product).slice(0, 3) : [];
  if (items.length !== 3) return null;
  return items;
}

function appendResponseRow(ss, email, ranking, freeText) {
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME) || ss.insertSheet(RESPONSE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'タイムスタンプ',
      'メールアドレス',
      '1位カテゴリ',
      '1位メーカー',
      '1位商品名',
      '1位価格',
      '2位カテゴリ',
      '2位メーカー',
      '2位商品名',
      '2位価格',
      '3位カテゴリ',
      '3位メーカー',
      '3位商品名',
      '3位価格',
      '自由記入',
    ]);
  }

  const buildProductCells = (product) => {
    if (!product) return ['', '', '', ''];
    return [product.category || '', product.maker || '', product.product || '', product.price || ''];
  };

  sheet.appendRow([
    new Date(),
    email,
    ...buildProductCells(ranking[0]),
    ...buildProductCells(ranking[1]),
    ...buildProductCells(ranking[2]),
    freeText || '',
  ]);
}

function appendAggregateRows(ss, email, ranking, freeText) {
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

  ranking.forEach((product, idx) => {
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

  refreshPointSummary(sheet);
}

function appendFreeTextRow(ss, email, freeText) {
  const text = typeof freeText === 'string' ? freeText.trim() : '';
  if (!text) return;

  const sheet = ss.getSheetByName(FREE_TEXT_SHEET_NAME) || ss.insertSheet(FREE_TEXT_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['タイムスタンプ', 'メールアドレス', '自由記入']);
  }

  sheet.appendRow([new Date(), email, text]);
}

function refreshPointSummary(sheet) {
  const DATA_COLUMNS = 9;
  const SUMMARY_START_COL = 11; // Column K
  const SUMMARY_TITLE_CELL = sheet.getRange(1, SUMMARY_START_COL);
  const SUMMARY_START_ROW = 2;
  const headers = ['カテゴリ', 'メーカー', '商品名', '価格', '合計Pt', '票数'];

  SUMMARY_TITLE_CELL.setValue('ポイント集計');

  const lastRow = sheet.getLastRow();
  const dataRowCount = Math.max(lastRow - 1, 0);
  if (dataRowCount === 0) {
    sheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL, sheet.getMaxRows() - SUMMARY_START_ROW + 1, headers.length).clearContent();
    return;
  }

  const data = sheet.getRange(2, 1, dataRowCount, DATA_COLUMNS).getValues();
  const summaryMap = {};

  data.forEach((row) => {
    const timestamp = row[0];
    if (!timestamp) return;
    const points = Number(row[4]) || 0;
    const category = row[5] || '';
    const maker = row[6] || '';
    const product = row[7] || '';
    const price = row[8] || '';
    const key = [category, maker, product, price].join('||');

    if (!summaryMap[key]) {
      summaryMap[key] = { category, maker, product, price, points: 0, count: 0 };
    }

    summaryMap[key].points += points;
    summaryMap[key].count += 1;
  });

  const summaryRows = Object.values(summaryMap).sort((a, b) => {
    if (b.points !== a.points) return b.points - a.points;
    if (a.category !== b.category) return String(a.category).localeCompare(String(b.category), 'ja');
    if (a.maker !== b.maker) return String(a.maker).localeCompare(String(b.maker), 'ja');
    return String(a.product).localeCompare(String(b.product), 'ja');
  });

  const output = [headers, ...summaryRows.map((item) => [item.category, item.maker, item.product, item.price, item.points, item.count])];
  const clearHeight = Math.max(sheet.getLastRow(), SUMMARY_START_ROW + output.length) - SUMMARY_START_ROW + 1;
  sheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL, clearHeight, headers.length).clearContent();
  sheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL, output.length, headers.length).setValues(output);
}

function getRankingSummary() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(AGGREGATE_SHEET_NAME);
  if (!sheet) return { products: [] };

  const lastRow = sheet.getLastRow();
  const DATA_COLUMNS = 9;
  if (lastRow <= 1) return { products: [] };

  const rows = sheet.getRange(2, 1, lastRow - 1, DATA_COLUMNS).getValues();
  const productMap = buildRankingMap(rows);
  const products = buildProductRanking(productMap).slice(0, 15);
  const enrichedProducts = enrichRankingImages(products);

  return { products: enrichedProducts };
}

function enrichRankingImages(products) {
  if (!Array.isArray(products) || !products.length) return products || [];

  let fallbackMap;
  return products.map((item) => {
    if (!item) return item;
    if (item.imageUrl) return item;
    const key = String(item.product || '').trim();
    if (!key) return item;
    if (!fallbackMap) fallbackMap = buildProductImageLookup();
    const fallbackUrl = fallbackMap[key];
    return fallbackUrl ? { ...item, imageUrl: fallbackUrl } : item;
  });
}

function buildProductImageLookup() {
  const map = {};
  const categories = buildCategoriesFromDrive();
  categories.forEach((category) => {
    (category.items || []).forEach((item) => {
      const key = String(item.product || '').trim();
      if (key && item.imageUrl) {
        map[key] = item.imageUrl;
      }
    });
  });
  return map;
}

function buildRankingMap(rows) {
  const productMap = {};

  rows.forEach((row) => {
    const timestamp = row[0];
    if (!timestamp) return;
    const points = Number(row[3]) || 0;
    const category = row[4] || '';
    const maker = row[5] || '';
    const product = row[6] || '';
    const price = row[7] || '';
    const productKey = [category, maker, product, price].join('||');

    if (!productMap[productKey]) {
      productMap[productKey] = { category, maker, product, price, points: 0, count: 0 };
    }
    productMap[productKey].points += points;
    productMap[productKey].count += 1;
  });

  return productMap;
}

function buildProductRanking(productMap) {
  if (!productMap) return [];

  return Object.values(productMap)
    .sort((a, b) => {
      if (b.points !== a.points) return b.points - a.points;
      if (a.category !== b.category) return String(a.category).localeCompare(String(b.category), 'ja');
      if (a.maker !== b.maker) return String(a.maker).localeCompare(String(b.maker), 'ja');
      return String(a.product).localeCompare(String(b.product), 'ja');
    })
    .map((item) => ({
      ...item,
      imageUrl: getProductImageUrlFromDrive(item.maker, item.product, item.category) || '',
    }));
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

function getMakerBadgeFolder() {
  if (makerBadgeFolder !== undefined) return makerBadgeFolder;
  if (!MAKER_BADGE_FOLDER_ID) {
    makerBadgeFolder = null;
    return null;
  }
  try {
    makerBadgeFolder = DriveApp.getFolderById(MAKER_BADGE_FOLDER_ID);
    return makerBadgeFolder;
  } catch (e) {
    makerBadgeFolder = null;
    return null;
  }
}

function getMakerBadgeImageMap() {
  if (makerBadgeImageMap !== undefined) return makerBadgeImageMap;
  const folder = getMakerBadgeFolder();
  if (!folder) {
    makerBadgeImageMap = null;
    return null;
  }

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

  makerBadgeImageMap = map;
  return makerBadgeImageMap;
}

function getMakerBadgeImageUrl(makerName) {
  const imageMap = getMakerBadgeImageMap();
  const normalizedMaker = normalizeProductKey(makerName);
  const fileId = findMatchingImageId(imageMap, normalizedMaker);
  return fileId ? buildDriveImageUrl(fileId) || buildDriveViewUrl(fileId) : '';
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
  return `https://lh3.googleusercontent.com/d/${fileId}`;
}

function buildDriveImageUrl(fileId) {
  if (!fileId) return '';
  // 画像タグで直接表示できる googleusercontent リンクを優先し、フォールバックとして preview を返す。
  return buildDriveViewUrl(fileId) || buildDrivePreviewUrl(fileId);
}

function buildCategoriesFromDrive() {
  const root = getProductRootFolder();
  if (!root) return [];

  const categories = [];
  const iterator = root.getFolders();
  while (iterator.hasNext()) {
    const categoryFolder = iterator.next();
    const categoryName = categoryFolder.getName();
    const items = collectCategoryItemsFromDrive(categoryFolder, categoryName);
    categories.push({
      name: categoryName,
      sheetName: categoryName,
      items,
    });
  }

  categories.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
  categories.forEach((cat) => {
    cat.items.sort((a, b) => {
      const makerDiff = (a.maker || '').localeCompare(b.maker || '', 'ja');
      if (makerDiff !== 0) return makerDiff;
      const priceDiff = (a.price || '').localeCompare(b.price || '', 'ja', { numeric: true });
      if (priceDiff !== 0) return priceDiff;
      return (a.product || '').localeCompare(b.product || '', 'ja');
    });
  });

  return categories;
}

function collectCategoryItemsFromDrive(categoryFolder, categoryName) {
  const items = [];
  const makerIterator = categoryFolder.getFolders();

  while (makerIterator.hasNext()) {
    const makerFolder = makerIterator.next();
    const makerName = makerFolder.getName();
    let hasPriceFolder = false;

    const priceIterator = makerFolder.getFolders();
    while (priceIterator.hasNext()) {
      hasPriceFolder = true;
      const priceFolder = priceIterator.next();
      const priceLabel = priceFolder.getName();
      items.push(...collectImageItems(priceFolder, { category: categoryName, maker: makerName, price: priceLabel }));
    }

    if (!hasPriceFolder) {
      items.push(...collectImageItems(makerFolder, { category: categoryName, maker: makerName, price: '' }));
    }
  }

  // 価格・メーカー階層以外に直下に画像がある場合のフォールバック
  items.push(...collectImageItems(categoryFolder, { category: categoryName, maker: '', price: '' }));

  return items;
}

function collectImageItems(folder, meta) {
  const results = [];
  const iterator = folder.getFiles();

  while (iterator.hasNext()) {
    const file = iterator.next();
    const mime = file.getMimeType() || '';
    if (!mime.startsWith('image/')) continue;
    const productName = stripExtension(file.getName());
    results.push({
      product: productName,
      maker: meta.maker || '',
      price: meta.price || '',
      imageUrl: buildDriveImageUrl(file.getId()) || buildDriveViewUrl(file.getId()),
      category: meta.category || '',
    });
  }

  return results;
}
