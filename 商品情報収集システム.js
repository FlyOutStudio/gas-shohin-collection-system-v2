/**
 * =====================================================================
 *  Google Apps Script – EC 商品情報収集
 *    - 楽天/Yahoo検索 → シート出力（既存機能）
 *    - ScreenshotOneでフルページスクショ保存（Microlink撤去）
 *    - Diffbot（URLベース）で商品詳細の自動抽出 → D_列に出力
 * =====================================================================
 *
 * ▼ Script Properties に設定しておくキー
 *   - RAKUTEN_APP_ID            : 楽天 API アプリ ID（必須）
 *   - YAHOO_APP_ID              : Yahoo API アプリ ID（必須）
 *   - GOOGLE_API_KEY            : Google Custom Search API キー（必須）
 *   - GOOGLE_SEARCH_ENGINE_ID   : Google 検索エンジン ID（必須）
 *   - SCREENSHOTONE_ACCESS_KEY  : ScreenshotOne アクセスキー（必須）
 *   - SCREENSHOT_FOLDER_ID      : Drive 保存先フォルダ ID（任意）
 *   - DIFFBOT_TOKEN             : Diffbot アクセストークン（必須）
 */

const CONFIG = {
  RAKUTEN_APP_ID             : PropertiesService.getScriptProperties().getProperty('RAKUTEN_APP_ID'),
  YAHOO_APP_ID               : PropertiesService.getScriptProperties().getProperty('YAHOO_APP_ID'),
  GOOGLE_API_KEY             : PropertiesService.getScriptProperties().getProperty('GOOGLE_API_KEY'),
  GOOGLE_SEARCH_ENGINE_ID    : PropertiesService.getScriptProperties().getProperty('GOOGLE_SEARCH_ENGINE_ID'),
  SCREENSHOTONE_ACCESS_KEY   : PropertiesService.getScriptProperties().getProperty('SCREENSHOTONE_ACCESS_KEY'),
  SCREENSHOT_FOLDER_ID       : PropertiesService.getScriptProperties().getProperty('SCREENSHOT_FOLDER_ID'),
  DIFFBOT_TOKEN              : PropertiesService.getScriptProperties().getProperty('DIFFBOT_TOKEN'),

  RAKUTEN_SEARCH_API         : 'https://app.rakuten.co.jp/services/api/IchibaItem/Search/20220601',
  YAHOO_ITEM_API             : 'https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch',
  GOOGLE_SEARCH_API          : 'https://www.googleapis.com/customsearch/v1'
};

// 取得件数（楽天/Yahoo共通）
const TOP_N = 10;

/* ------------------------------------------------------------------ */
/* 0. カスタムメニュー                                                */
/* ------------------------------------------------------------------ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('商品調査ツール')
    .addItem('商品情報を収集（楽天・Yahoo・Google）', 'showSearchDialog')
    .addSeparator()
    .addItem('各モール上位3件のPDFスクショを取得', 'captureScreenshots')
    .addItem('Google検索結果をDiffbotで詳細分析', 'enrichByDiffbot')
    .addItem('Diffbot詳細レポートをPDF出力', 'createDiffbotDetailsPdf')
    .addToUi();
}

/* ------------------------------------------------------------------ */
/* 1. ダイアログ表示                                                  */
/* ------------------------------------------------------------------ */
function showSearchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SearchDialog')
                .setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, '商品キーワード入力');
}

/* ------------------------------------------------------------------ */
/* 2. メイン処理：商品情報を取得してシート出力                       */
/* ------------------------------------------------------------------ */
function searchProducts(keyword) {
  if (!keyword) throw new Error('キーワードが空です');

  const { rakuten, yahoo, google } = getProducts(keyword.trim());

  const sheetName = `結果_${keyword}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);

  // ヘッダー
  const headers = [
    '収集日時','モール','ランキング順位','商品名','価格','URL',
    '販売者','レビュー数','評価','スクショURL','商品説明（抜粋）'
  ];
  sheet.appendRow(headers);

  // データ行（楽天・Yahoo・Googleの順で出力）
  [...rakuten, ...yahoo, ...google].forEach(p =>
    sheet.appendRow([
      p.collectedAt, p.platform, p.rank, p.name, p.price,
      p.url, p.shopName, p.reviewCount, p.reviewAvg, '', p.description
    ])
  );

  return `検索完了：${rakuten.length + yahoo.length + google.length} 件取得`;
}

/* ------------------------------------------------------------------ */
/* 3. API 呼び出しラッパ                                             */
/* ------------------------------------------------------------------ */
const getProducts = (keyword) => ({
  rakuten: fetchRakutenItems(keyword, TOP_N),
  yahoo  : fetchYahooItems(keyword, TOP_N),
  google : fetchGoogleItems(keyword, TOP_N)
});

/* -------- 楽天：rank 付与・URL そのまま --------------------------- */
const fetchRakutenItems = (keyword, hits = 10) => {
  const params = toQuery({
    applicationId : CONFIG.RAKUTEN_APP_ID,
    keyword,
    hits,
    sort          : 'standard',
    formatVersion : 2,
    elements      : [
      'itemName','itemPrice','itemUrl','shopName',
      'reviewCount','reviewAverage','itemCaption'
    ].join(',')
  });

  const raw   = UrlFetchApp.fetch(`${CONFIG.RAKUTEN_SEARCH_API}?${params}`).getContentText();
  const items = JSON.parse(raw).Items ?? [];

  return items.map((it, idx) => ({
    rank        : idx + 1,
    collectedAt : new Date(),
    platform    : '楽天市場',
    name        : it.itemName,
    price       : it.itemPrice,
    url         : it.itemUrl,
    shopName    : it.shopName,
    reviewCount : it.reviewCount,
    reviewAvg   : it.reviewAverage,
    description : it.itemCaption ?? ''
  }));
};

/* -------- Google：rank 付与 ------------------------------------- */
const fetchGoogleItems = (keyword, hits = 10) => {
  if (!CONFIG.GOOGLE_API_KEY || !CONFIG.GOOGLE_SEARCH_ENGINE_ID) {
    console.warn('Google Custom Search API の設定が不完全です');
    return [];
  }

  const params = toQuery({
    key: CONFIG.GOOGLE_API_KEY,
    cx: CONFIG.GOOGLE_SEARCH_ENGINE_ID,
    q: keyword,
    num: Math.min(hits, 10), // Google APIは最大10件
    lr: 'lang_ja', // 日本語検索
    safe: 'medium'
  });

  try {
    const raw = UrlFetchApp.fetch(`${CONFIG.GOOGLE_SEARCH_API}?${params}`).getContentText();
    const data = JSON.parse(raw);
    const items = data.items || [];

    return items.map((it, idx) => ({
      rank        : idx + 1,
      collectedAt : new Date(),
      platform    : 'Google検索',
      name        : it.title,
      price       : '', // Google検索には価格情報なし
      url         : it.link,
      shopName    : extractDomain(it.link),
      reviewCount : '',
      reviewAvg   : '',
      description : it.snippet || ''
    }));
  } catch (e) {
    console.error('Google検索エラー:', e.message);
    return [];
  }
};

// URLからドメイン名を抽出
const extractDomain = (url) => {
  try {
    return new URL(url).hostname.replace('www.', '');
  } catch (e) {
    return url;
  }
};

/* -------- Yahoo!：rank 付与 -------------------------------------- */
const fetchYahooItems = (keyword, hits = 10) => {
  const params = toQuery({
    appid   : CONFIG.YAHOO_APP_ID,
    query   : keyword,
    results : hits,  // Yahoo APIではresultsパラメータを使用
    sort    : '-score'
  });

  const raw   = UrlFetchApp.fetch(`${CONFIG.YAHOO_ITEM_API}?${params}`).getContentText();
  const items = JSON.parse(raw).hits ?? [];

  return items.map((it, idx) => ({
    rank        : idx + 1,
    collectedAt : new Date(),
    platform    : 'Yahoo!ショッピング',
    name        : it.name,
    price       : it.price,
    url         : it.url,
    shopName    : it.seller?.name,
    reviewCount : it.review?.count,
    reviewAvg   : it.review?.rate,
    description : it.description ?? it.explanation ?? ''
  }));
};

/* ------------------------------------------------------------------ */
/* 4. 共通ユーティリティ                                              */
/* ------------------------------------------------------------------ */
const toQuery = obj =>
  Object.entries(obj)
    .filter(([_, v]) => v !== undefined && v !== null && v !== '')
    .map(([k, v]) => `${k}=${encodeURIComponent(v)}`)
    .join('&');

function idxOf_(header, name) {
  const n = header.indexOf(name);
  return n >= 0 ? n + 1 : 0; // 1-based
}

/** 見出しがなければ末尾に追加し、その列番号(1-based)を返す */
function ensureCol_(sh, header, name) {
  let idx = header.indexOf(name);
  if (idx >= 0) return idx + 1;
  const col = header.length + 1;
  sh.getRange(1, col).setValue(name);
  header.push(name);
  return col;
}

function toNum_(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = Number(v);
  return Number.isFinite(n) ? n : '';
}
function trim_(s, n) {
  return (s || '').toString().slice(0, n);
}

/* ------------------------------------------------------------------ */
/* 5. スクリーンショット（ScreenshotOne 専用）                       */
/* ------------------------------------------------------------------ */
/**
 * ScreenshotOne でフルページ JPEG を取得 → Drive 保存 → 共有URLを返す
 * - 無料枠: 100枚/月
 * - 必須: Script Property SCREENSHOTONE_ACCESS_KEY
 * - 対象: Google検索結果のみ
 */
const captureWithScreenshotOne_ = (rawUrl) => {
  const key = CONFIG.SCREENSHOTONE_ACCESS_KEY;
  if (!key) throw new Error('SCREENSHOTONE_ACCESS_KEY が未設定です');

  // 余計なクエリを落として転送を減らす（安定化）
  const url = rawUrl.split('?')[0];

  const endpoint = 'https://api.screenshotone.com/take';
  const qs = {
    access_key      : key,
    url,
    full_page       : true,
    format          : 'pdf',
    block_ads       : true,
    wait_until      : 'networkidle2',  // ECサイトに最適：2つ以下の接続で待機
    timeout         : 60,              // 60秒タイムアウト（公式推奨デフォルト）
    navigation_timeout: 30,            // サイト応答待機30秒
    delay           : 3,               // 3秒待機でJS・画像読み込み保証（秒単位）
    viewport_width  : 1280,            // 標準的なデスクトップサイズ
    viewport_height : 1024,            // 適度な高さで重い処理を回避
    response_type   : 'json'
  };

  const res  = UrlFetchApp.fetch(`${endpoint}?${toQuery(qs)}`, { muteHttpExceptions: true, method: 'get' });
  const code = res.getResponseCode();
  if (code !== 200) {
    const body = res.getContentText();
    throw new Error(`ScreenshotOne ${code}: ${body.slice(0, 200)}`);
  }

  const data = JSON.parse(res.getContentText() || '{}');
  const shotUrl = data.screenshot_url || data.url || (data.data?.screenshot?.url);
  if (!shotUrl) throw new Error('ScreenshotOne: screenshot URL not found');

  const blob = UrlFetchApp.fetch(shotUrl).getBlob().setName(`shot_${Date.now()}.pdf`);
  const folder = CONFIG.SCREENSHOT_FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.SCREENSHOT_FOLDER_ID)
    : DriveApp.getRootFolder();
  const file = folder.createFile(blob);

  return file.getUrl();
};

/* ------------------------------------------------------------------ */
/* 6. URL 列のスクショを一括取得して「スクショURL」列へ書込み         */
/* ------------------------------------------------------------------ */
/**
 * アクティブシートの URL 列から各モールの上位3件をスクリーンショット取得
 *   ・楽天市場 ランキング順位 1〜3
 *   ・Yahoo!ショッピング ランキング順位 1〜3
 *   ・Google検索 ランキング順位 1〜3
 * ScreenshotOne でフルページ撮影し「スクショURL」列へ保存後、PDF出力
 */
function captureScreenshots() {
  const sheet  = SpreadsheetApp.getActiveSheet();
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const col = {
    url   : idxOf_(header, 'URL'),
    rank  : idxOf_(header, 'ランキング順位'),
    store : idxOf_(header, 'モール'),
    shot  : idxOf_(header, 'スクショURL'),
    name  : idxOf_(header, '商品名')
  };
  if (col.url === 0 || col.rank === 0 || col.store === 0) {
    SpreadsheetApp.getUi().alert('ヘッダーに必要な列（URL／モール／ランキング順位）が見つかりません。');
    return;
  }
  if (col.shot === 0) {                 // 「スクショURL」列がなければ追加
    col.shot = header.length + 1;
    sheet.getRange(1, col.shot).setValue('スクショURL');
  }

  const lastRow = sheet.getLastRow();
  const rows    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // 各モールの上位3件を処理
  const quota = { '楽天市場': 0, 'Yahoo!ショッピング': 0, 'Google検索': 0 };

  rows.forEach((r, idx) => {
    const store = r[col.store - 1];
    const rank  = Number(r[col.rank - 1]);
    const url   = r[col.url - 1];

    if ((store === '楽天市場' || store === 'Yahoo!ショッピング' || store === 'Google検索') &&
        rank >= 1 && rank <= 3 && quota[store] < 3 && 
        typeof url === 'string' && url.startsWith('http')) {
      try {
        const shotUrl = captureWithScreenshotOne_(url);
        sheet.getRange(idx + 2, col.shot).setValue(shotUrl);
        quota[store]++;
        Utilities.sleep(2000); // より長めの待機（タイムアウト対策）
      } catch (e) {
        console.error(`スクリーンショット失敗: ${url} - ${e.message}`);
        sheet.getRange(idx + 2, col.shot).setValue(`SKIP: タイムアウト`);
        // タイムアウト時は次のURLに進む（処理を止めない）
      }
    }
  });

  const totalCount = Object.values(quota).reduce((sum, count) => sum + count, 0);
  SpreadsheetApp.getActiveSpreadsheet()
                .toast(`各モール上位3件×3（計${totalCount}件）のPDFスクリーンショット取得が完了しました`,
                       'ScreenshotOne', 5);
}


/* ------------------------------------------------------------------ */
/* 7. Diffbot Product API 連携（URL→詳細を D_ 系列へ出力）            */
/* ------------------------------------------------------------------ */

const DIFFBOT_ENDPOINT = 'https://api.diffbot.com/v3/product';
// Freeは約5RPM。安全側で12.5秒スリープ
const DIFFBOT_SLEEP_MS = 12500;

/**
 * 現在シートの各行の「URL」をDiffbotに投げて詳細抽出。
 * Google検索結果のみ対象。直近7日以内に取得済みの行はスキップ。
 * 結果は D_ プレフィックス列に出力。
 */
function enrichByDiffbot() {
  const token = CONFIG.DIFFBOT_TOKEN;
  if (!token) throw new Error('Script Property DIFFBOT_TOKEN を設定してください。');

  const sh = SpreadsheetApp.getActiveSheet();
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const col = {
    url   : idxOf_(header, 'URL'),
    d_col : sh.getLastColumn() + 1 // 最後の列の次から開始（D列相当）
  };

  const lastRow = sh.getLastRow();
  if (lastRow < 2 || col.url === 0) {
    SpreadsheetApp.getUi().alert('ヘッダー「URL」が見つからないか、データがありません。');
    return;
  }

  const rows  = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const now   = new Date();

  for (let i = 0; i < rows.length; i++) {
    const rowIndex = i + 2;
    const url = rows[i][col.url - 1];
    const store = rows[i][idxOf_(header, 'モール') - 1];
    
    // Google検索結果のみ処理
    if (store !== 'Google検索') continue;
    if (!url || typeof url !== 'string' || !url.startsWith('http')) continue;

    // 7日以内取得済みチェック（L列 = FetchedAt）
    const fetchedAtCol = col.d_col + 8; // L列（9列目）
    const fetchedAt = sh.getRange(rowIndex, fetchedAtCol).getValue();
    if (fetchedAt && (now - new Date(fetchedAt)) / (1000 * 3600 * 24) < 7) continue;

    try {
      const prod = fetchProductFromDiffbot_(url, token);

      // ヘッダー行に項目名を設定（初回のみ）
      if (rowIndex === 2) {
        const headers = ['Title', 'Price', 'Brand', 'Rating', 'ReviewCount', 'Availability', 'Category', 'MainImage', 'FetchedAt'];
        headers.forEach((header, idx) => {
          sh.getRange(1, col.d_col + idx).setValue(header);
        });
      }

      // 数値・計算
      const price = toNum_(prod.offerPrice ?? prod.price);
      const mainImage = Array.isArray(prod.images) ? (prod.images[0] || '') : '';
      const category = (prod.category || prod.breadcrumb || []).toString();

      // D〜L列（9列）にデータを出力
      const diffbotData = [
        prod.title || '',                              // D列: Title
        price || '',                                   // E列: Price  
        prod.brand || '',                              // F列: Brand
        prod?.aggregateRating?.value || '',            // G列: Rating
        prod?.aggregateRating?.count || '',            // H列: ReviewCount
        prod.availability || '',                       // I列: Availability
        category,                                      // J列: Category
        mainImage,                                     // K列: MainImage
        new Date()                                     // L列: FetchedAt
      ];

      // D〜L列に一括書き込み
      sh.getRange(rowIndex, col.d_col, 1, diffbotData.length).setValues([diffbotData]);

    } catch (e) {
      // エラー時はD列にエラーメッセージ
      sh.getRange(rowIndex, col.d_col).setValue(`ERR: ${String(e).slice(0, 100)}`);
    }

    // Freeレート対策
    Utilities.sleep(DIFFBOT_SLEEP_MS);
  }

  SpreadsheetApp.getActiveSpreadsheet()
    .toast('Google検索結果のDiffbot詳細取得が完了しました', 'Diffbot', 5);
}

/* ------------------------------------------------------------------ */
/* 8. Diffbot結果PDF出力機能                                          */
/* ------------------------------------------------------------------ */
/**
 * DiffbotでGoogle検索結果の詳細情報を取得してPDF化
 */
function createDiffbotDetailsPdf() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Google検索結果のみ抽出
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('データがありません。');
    return;
  }
  
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const googleResults = rows.filter(row => row[idxOf_(header, 'モール') - 1] === 'Google検索');
  
  if (googleResults.length === 0) {
    SpreadsheetApp.getUi().alert('Google検索結果が見つかりません。');
    return;
  }
  
  // PDF作成
  const keyword = sheet.getName().split('_')[1] || 'search';
  const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
  const docName = `Diffbot詳細レポート_${keyword}_${timestamp}`;
  
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();
  
  // タイトル・日時
  body.appendParagraph(`商品詳細分析レポート: ${keyword}`)
      .setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(`作成日時: ${Utilities.formatDate(new Date(), 'JST', 'yyyy年MM月dd日 HH:mm:ss')}`)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(`分析対象: Google検索結果 ${googleResults.length}件`)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendHorizontalRule();
  
  // 各商品の詳細情報
  googleResults.forEach((row, index) => {
    const rank = row[idxOf_(header, 'ランキング順位') - 1];
    const title = row[idxOf_(header, '商品名') - 1];
    const url = row[idxOf_(header, 'URL') - 1];
    
    // 商品タイトル
    body.appendParagraph(`${rank}位: ${title}`)
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    // URL
    body.appendParagraph(`URL: ${url}`)
        .setFontSize(10)
        .setForegroundColor('#0066cc');
    
    // Diffbot詳細情報があるかチェック
    const diffbotDataFound = checkDiffbotData_(row, header);
    
    if (diffbotDataFound.hasData) {
      // テーブル形式で詳細情報を表示
      const table = body.appendTable();
      
      // 基本情報
      if (diffbotDataFound.title) table.appendTableRow(['商品名', diffbotDataFound.title]);
      if (diffbotDataFound.price) table.appendTableRow(['価格', `${diffbotDataFound.price} ${diffbotDataFound.currency || ''}`]);
      if (diffbotDataFound.brand) table.appendTableRow(['ブランド', diffbotDataFound.brand]);
      if (diffbotDataFound.rating) table.appendTableRow(['評価', `${diffbotDataFound.rating}/5 (${diffbotDataFound.reviewCount || 0}件)`]);
      if (diffbotDataFound.availability) table.appendTableRow(['在庫状況', diffbotDataFound.availability]);
      if (diffbotDataFound.category) table.appendTableRow(['カテゴリ', diffbotDataFound.category]);
      
      // テーブルスタイル
      table.setBorderWidth(1);
      table.setColumnWidth(0, 100);
      
    } else {
      body.appendParagraph('Diffbot詳細情報: 未取得')
          .setItalic(true)
          .setForegroundColor('#666666');
    }
    
    body.appendParagraph(''); // 空行
    
    if (index < googleResults.length - 1) {
      body.appendHorizontalRule();
    }
  });
  
  doc.saveAndClose();
  
  // PDF作成
  const folder = CONFIG.SCREENSHOT_FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.SCREENSHOT_FOLDER_ID)
    : DriveApp.getRootFolder();
  
  const pdfBlob = doc.getAs('application/pdf');
  pdfBlob.setName(`${docName}.pdf`);
  const pdfFile = folder.createFile(pdfBlob);
  
  // 元のGoogleドキュメントを削除
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  
  SpreadsheetApp.getActiveSpreadsheet()
    .toast(`PDF作成完了: ${pdfFile.getName()}`, 'Diffbot PDF', 10);
  
  return pdfFile;
}

/** Diffbotデータの存在チェックと取得 */
function checkDiffbotData_(row, header) {
  const getData = (colName) => {
    const idx = idxOf_(header, colName);
    return idx > 0 ? row[idx - 1] : '';
  };
  
  return {
    hasData: getData('D_Title') || getData('D_Price') || getData('D_Brand'),
    title: getData('D_Title'),
    price: getData('D_Price'),
    currency: getData('D_Currency'),
    brand: getData('D_Brand'),
    rating: getData('D_Rating'),
    reviewCount: getData('D_ReviewCount'),
    availability: getData('D_Availability'),
    category: getData('D_Category')
  };
}

/** Diffbot呼び出し（必要フィールドだけに絞って軽量化） */
function fetchProductFromDiffbot_(url, token) {
  const fields = [
    'title','price','offerPrice','offerPriceDetails','regularPrice','priceCurrency',
    'availability','brand','sku','seller',
    'images','variants','category','breadcrumb',
    'aggregateRating','reviews'
  ].join(',');

  const qs = toQuery({ token, url, fields });
  const res = UrlFetchApp.fetch(`${DIFFBOT_ENDPOINT}?${qs}`, { muteHttpExceptions: true, method: 'get' });
  const code = res.getResponseCode();

  if (code === 429) throw new Error('Diffbot 429: Quota/Rate limit exceeded');
  if (code >= 400) throw new Error(`Diffbot ${code}: ${res.getContentText().slice(0, 500)}`);

  const body = JSON.parse(res.getContentText() || '{}');
  const obj = body.objects && body.objects[0];
  if (!obj) throw new Error('Diffbot: No product object parsed');
  return obj;
}