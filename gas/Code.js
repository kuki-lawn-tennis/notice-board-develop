/*************************************************
 * テニスクラブ掲示板
 * Google Apps Script
 *************************************************/

const DEFAULT_BANNER_MESSAGE = '新しいお知らせがありません。';

const SHEET_NAMES = {
  SETTINGS: '設定',
  NOTICE: 'お知らせ',
  EVENT: 'イベント',
  SCHEDULE: 'スケジュール',
  ACCESS_LOG: 'アクセスログ',
  ACCESS_SUMMARY: 'アクセス集計'
};

const CONTENT_SHEET_HEADERS = {
  [SHEET_NAMES.NOTICE]: ['ID', '日付', 'タイトル', '内容', '重要度', '公開', '添付ファイル'],
  [SHEET_NAMES.EVENT]: ['ID', '日付', 'イベント名', '内容', '重要度', '公開', '添付ファイル'],
  [SHEET_NAMES.SCHEDULE]: ['ID', 'タイトル', '内容', '重要度', '公開', '添付ファイル']
};

// 種類ごとの設定（保存・削除・表示順変更で共通利用）
const TYPE_CONFIG = {
  notice:   { sheet: SHEET_NAMES.NOTICE,   label: 'お知らせ',     attachmentFolderProperty: 'NOTICE_ATTACHMENTS_FOLDER_ID' },
  event:    { sheet: SHEET_NAMES.EVENT,    label: 'イベント',     attachmentFolderProperty: 'EVENT_ATTACHMENTS_FOLDER_ID' },
  schedule: { sheet: SHEET_NAMES.SCHEDULE, label: 'スケジュール', attachmentFolderProperty: 'SCHEDULE_ATTACHMENTS_FOLDER_ID' }
};

function getTypeConfig_(type) {
  const cfg = TYPE_CONFIG[String(type || '')];
  if (!cfg) throw new Error('データの種類が正しくありません。');
  return cfg;
}

/*************************************************
 * Webアプリ
 *************************************************/
function doGet(e) {
  const parameters = e && e.parameter || {};
  const allowedOrigins = getApiBridgeAllowedOrigins_();
  const template = HtmlService.createTemplateFromFile('ApiBridge.html');
  template.allowedOrigins = allowedOrigins;
  template.parentOrigin = allowedOrigins.indexOf(String(parameters.parentOrigin || '')) !== -1
    ? String(parameters.parentOrigin)
    : '';
  template.channelNonce = /^\d+-[a-z0-9]+$/.test(String(parameters.nonce || ''))
    ? String(parameters.nonce)
    : '';
  return template.evaluate()
    .setTitle('掲示板 API')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getApiBridgeAllowedOrigins_() {
  const value = PropertiesService.getScriptProperties().getProperty('GITHUB_PAGES_ORIGINS') || '';
  const origins = value.split(',').map(function(origin) { return origin.trim(); }).filter(Boolean);
  const validOrigin = /^https:\/\/[a-z0-9.-]+(?::\d+)?$|^http:\/\/localhost(?::\d+)?$/i;
  const invalidOrigins = origins.filter(function(origin) { return !validOrigin.test(origin); });
  if (invalidOrigins.length) {
    throw new Error('GITHUB_PAGES_ORIGINSにはオリジンだけをカンマ区切りで設定してください。');
  }
  return Array.from(new Set(origins));
}

/*************************************************
 * 初期セットアップ
 *
 * 最初に1回だけ実行してください
 *************************************************/
function setupSheets() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // -------------------------
  // 設定
  // -------------------------
  createSheetIfNotExists_(ss, SHEET_NAMES.SETTINGS, ['項目', '値']);

  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);

  const defaultSettings = [
    ['クラブ名', '久喜ローンテニスクラブ'],
    ['キャッチコピー', 'みんなで楽しくテニスを！！'],
    ['連絡先', 'メール：example@example.com'],
    ['連絡先説明', 'お問い合わせはこちらからお願いします。'],
    ['更新日時', ''],
    ['バナーメッセージ', '新しいお知らせがありません。'],
    ['バナー表示期限日', '']
  ];
  const existingKeys = new Set(
    settingsSheet.getDataRange().getValues().slice(1).map(row => String(row[0] || ''))
  );

  // 既存設定は保持し、不足している初期項目だけ追加
  defaultSettings.forEach(row => {
    if (!existingKeys.has(row[0])) settingsSheet.appendRow(row);
  });

  // -------------------------
  // お知らせ・イベント・スケジュール
  // -------------------------
  Object.keys(CONTENT_SHEET_HEADERS).forEach(function(sheetName) {
    createSheetIfNotExists_(ss, sheetName, CONTENT_SHEET_HEADERS[sheetName]);
    // 登録日時・更新日時・表示順の管理列を追加（画面には表示しない）
    ensureTimestampColumns_(sheetName, CONTENT_SHEET_HEADERS[sheetName]);
  });

  // -------------------------
  // アクセスログ / アクセス集計
  // （ACCESS_LOG_SPREADSHEET_ID のスプレッドシートに作成。未設定ならこのスプレッドシート）
  // -------------------------
  const logSs = getAccessLogSpreadsheet_();
  createSheetIfNotExists_(
    logSs,
    SHEET_NAMES.ACCESS_LOG,
    ['日時', 'ページ', '端末', 'ブラウザー種類', 'ユーザーエージェント']
  );
  createSheetIfNotExists_(logSs, SHEET_NAMES.ACCESS_SUMMARY, ['日付', 'アクセス数']);

  consolidateAccessSummary_();
  createAccessLogCleanupTrigger_();
  cleanupOldAccessLogs();
}

/*************************************************
 * シート作成
 *************************************************/
function createSheetIfNotExists_(ss, sheetName, headers) {

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

/*************************************************
 * 管理者ログイン
 * パスワードそのものはブラウザーに保存せず、
 * 30分で失効するランダムなセッショントークンを返す。
 *************************************************/
function checkAdminPassword(password, userAgent) {
  const expectedPassword = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');

  if (!expectedPassword) {
    throw new Error(
      '管理者パスワードが設定されていません。Script Propertiesに ADMIN_PASSWORD を設定してください。'
    );
  }

  if (typeof password !== 'string' || password !== expectedPassword) {
    recordAdminActionAccessLog_('ログイン失敗', userAgent);
    // 連続試行を少し遅らせる（本格的な試行回数制限の代替ではありません）
    Utilities.sleep(1000);
    return null;
  }

  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('ADMIN_SESSION_' + token, '1', 1800);
  recordAdminActionAccessLog_('ログイン成功', userAgent);
  return token;
}

/*************************************************
 * 管理者セッション確認
 *************************************************/
function requireAdmin_(token) {
  if (typeof token !== 'string' || !token ||
      CacheService.getScriptCache().get('ADMIN_SESSION_' + token) !== '1') {
    throw new Error('管理者セッションが無効、または期限切れです。もう一度ログインしてください。');
  }
}

/*************************************************
 * 管理者ログアウト
 *************************************************/
function logoutAdmin(token, userAgent) {
  if (typeof token === 'string' && token) {
    const cache = CacheService.getScriptCache();
    const key = 'ADMIN_SESSION_' + token;
    // 有効なセッションをログアウトしたときだけ記録する（無効なトークンでの呼び出しは記録しない）
    const wasValid = cache.get(key) === '1';
    cache.remove(key);
    if (wasValid) recordAdminActionAccessLog_('ログアウト', userAgent);
  }
  return true;
}

/*************************************************
 * 設定取得
 *************************************************/
function getSettings_() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.SETTINGS);

  if (!sheet) {
    throw new Error('設定シートがありません。setupSheets()を実行してください。');
  }

  const values = sheet.getDataRange().getValues();
  const settings = {};

  for (let i = 1; i < values.length; i++) {

    const key = values[i][0];
    const value = values[i][1];

    if (key) {
      if (value instanceof Date) {
        // Date型はそのままブラウザへ返せないため、文字列にする
        settings[key] = String(key).trim() === '更新日時'
          ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss')
          : toDateKey_(value);
      } else {
        settings[key] = value;
      }
    }
  }

  // トップ画像IDはスプレッドシートではなくスクリプトプロパティ
  // TOP_PAGE_IMAGE_ID から取得する（シート上の同名項目は無視）
  const topImageId = getTopImageId_();

  settings['トップ画像ID'] = topImageId;

  // 予備のGoogle Drive画像URL。実際の表示は getTopImageDataUrl() を優先する。
  settings['トップ画像URL'] = topImageId
    ? 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(topImageId) + '&sz=w1600'
    : '';

  return settings;
}

/*************************************************
 * トップ画像IDを取得
 *
 * スクリプトプロパティ TOP_PAGE_IMAGE_ID に設定した
 * DriveのファイルIDを使用（URLを貼り付けても可）
 *************************************************/
function getTopImageId_() {

  const value = PropertiesService
    .getScriptProperties()
    .getProperty('TOP_PAGE_IMAGE_ID');

  return normalizeDriveFileId_(value);
}

/*************************************************
 * DriveファイルIDを正規化
 *
 * IDだけでも、DriveのURLを貼り付けても使用可能
 *************************************************/
function normalizeDriveFileId_(value) {

  if (!value) {
    return '';
  }

  const text = String(value).trim();
  const match = text.match(/[-\w]{25,}/);

  return match ? match[0] : text;
}

/*************************************************
 * トップ画像をData URLとして取得
 *
 * Webアプリ利用者側のGoogle Drive公開設定に依存せず、
 * GASの実行ユーザーが参照できるDrive画像をブラウザへ
 * 直接渡します。
 *************************************************/
function getTopImageDataUrl() {

  // セキュリティのため、引数は受け取らずスクリプトプロパティのIDのみを使用
  const normalizedId = getTopImageId_();

  if (!normalizedId) {
    return {
      success: false,
      message: 'トップ画像IDが設定されていません。スクリプトプロパティ TOP_PAGE_IMAGE_ID にDriveのファイルIDを設定してください。'
    };
  }

  try {

    const file = DriveApp.getFileById(normalizedId);
    const mimeType = file.getMimeType();

    if (!mimeType || mimeType.indexOf('image/') !== 0) {
      return {
        success: false,
        message: '指定されたファイルが画像ではありません。'
      };
    }

    // iPhoneなどのモバイルでは、数MBの元画像（特にHEIC）は表示できないことがあるため、
    // Driveが生成する縮小版（JPEG/PNG）を優先する。
    let blob = null;

    try {
      blob = fetchResizedImageBlob_(normalizedId);
    } catch (e) {
      blob = null;
    }

    if (!blob) {

      // 縮小版が取れない場合は元画像を使う。ただし大きすぎる/HEIC系は返さない
      const originalSize = file.getSize();

      if (originalSize > 1.5 * 1024 * 1024 || /heic|heif/i.test(mimeType)) {
        return {
          success: false,
          message: '画像が大きすぎるか形式が非対応です。1MB以下のJPEG/PNGに変換して差し替えてください。',
          detail: mimeType + ' / ' + Math.round(originalSize / 1024) + 'KB'
        };
      }

      blob = file.getBlob();
    }

    const bytes = blob.getBytes();
    const outMime = blob.getContentType() || mimeType;

    return {
      success: true,
      mimeType: outMime,
      dataUrl: 'data:' + outMime + ';base64,' + Utilities.base64Encode(bytes),
      fileName: file.getName(),
      size: bytes.length
    };

  } catch (error) {

    return {
      success: false,
      message: '画像を読み込めませんでした。DriveファイルIDとGAS実行ユーザーの権限を確認してください。',
      detail: error.message
    };
  }
}

/*************************************************
 * Driveの縮小版画像を取得（幅1200px・JPEG/PNG）
 *
 * 取得できない場合は null を返す
 *************************************************/
function fetchResizedImageBlob_(fileId) {

  const url = 'https://drive.google.com/thumbnail?id=' +
    encodeURIComponent(fileId) +
    '&sz=w1200';

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true,
    followRedirects: true
  });

  if (response.getResponseCode() !== 200) {
    return null;
  }

  const blob = response.getBlob();
  const type = blob.getContentType();

  return type && type.indexOf('image/') === 0 ? blob : null;
}

/*************************************************
 * 設定保存
 *************************************************/
function saveSettings(token, settings, userAgent) {

  requireAdmin_(token);

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.SETTINGS);

  if (!sheet) {
    throw new Error('設定シートがありません。');
  }

  if (Object.prototype.hasOwnProperty.call(settings, 'バナーメッセージ')) {
    const message = String(settings['バナーメッセージ'] == null ? '' : settings['バナーメッセージ']).trim();
    if (message.length > 100) throw new Error('バナーメッセージは100文字以内で入力してください。');
  }

  // 表示期限日は yyyy-MM-dd に統一して保存する（空欄は期限なし）
  if (Object.prototype.hasOwnProperty.call(settings, 'バナー表示期限日')) {
    const rawEndDate = settings['バナー表示期限日'];
    const hasValue = rawEndDate !== null && rawEndDate !== undefined && String(rawEndDate).trim() !== '';
    const endDate = hasValue ? toDateKey_(rawEndDate) : '';
    if (hasValue && !endDate) throw new Error('バナー表示期限日は正しい日付で入力してください。');
    settings['バナー表示期限日'] = endDate;
  }

  const values = sheet.getDataRange().getValues();

  Object.keys(settings).forEach(key => {

    // トップ画像IDのみスクリプトプロパティで管理
    if (key === 'トップ画像ID') {
      return;
    }

    let row = 0;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === key) {
        row = i + 1;
        break;
      }
    }

    if (row) {
      sheet.getRange(row, 2).setValue(settings[key]);
    } else {
      sheet.appendRow([key, settings[key]]);
    }
  });

  updateSettingsTimestamp_();
  const savedSettings = getSettings_();
  recordAdminActionAccessLog_('サイト設定保存完了', userAgent);
  return savedSettings;
}

/*************************************************
 * 設定シートの更新日時を更新
 * 設定内容の保存時に呼び出す
 *************************************************/
function updateSettingsTimestamp_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    throw new Error('設定シートがありません。');
  }

  const lastRow = sheet.getLastRow();
  const values = lastRow > 0
    ? sheet.getRange(1, 1, lastRow, 2).getValues()
    : [];
  let targetRow = 0;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === '更新日時') {
      targetRow = i + 1;
      break;
    }
  }

  if (!targetRow) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1).setValue('更新日時');
  }

  sheet.getRange(targetRow, 2).setValue(new Date());
  sheet.getRange(targetRow, 2).setNumberFormat('yyyy/MM/dd HH:mm:ss');
}

/*************************************************
 * バナー情報取得
 * 管理用「設定」シートからメッセージと表示期限日を読み取る
 * 戻り値: { message: 本文, endDate: 'yyyy-MM-dd'（未設定なら空） }
 *************************************************/
function getBanner_() {
  const banner = { message: '', endDate: '' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return banner;

  const range = sheet.getDataRange();
  const values = range.getValues();
  const displayValues = range.getDisplayValues();

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key === 'バナーメッセージ') {
      banner.message = String(displayValues[i][1] || '').trim();
    } else if (key === 'バナー表示期限日') {
      banner.endDate = toDateKey_(values[i][1]);
    }
  }
  return banner;
}

/*************************************************
 * バナーが表示期限を過ぎているか
 * 表示期限日の後日よりは表示しない（スクリプトのタイムゾーンで判定）
 *************************************************/
function isBannerExpired_(endDate) {
  if (!endDate) return false;
  return formatDate_(new Date()) > endDate;
}

/*************************************************
 * 公開データ取得
 *************************************************/
function getPublicData() {
  const banner = getBanner_();
  return {
    settings: getPublicSettings_(),
    // メッセージが空白、または表示期限切れの場合はデフォルト文言を表示
    bannerMessage: (!banner.message || isBannerExpired_(banner.endDate))
      ? DEFAULT_BANNER_MESSAGE
      : banner.message,
    notices: getSheetData_(SHEET_NAMES.NOTICE, false),
    events: getSheetData_(SHEET_NAMES.EVENT, false),
    schedules: getSheetData_(SHEET_NAMES.SCHEDULE, false)
  };
}

/*************************************************
 * 公開用設定：公開画面で必要な項目だけ返す
 *************************************************/
function getPublicSettings_() {
  const all = getSettings_();
  const allowedKeys = [
    'クラブ名',
    'キャッチコピー',
    '連絡先',
    '連絡先説明',
    'トップ画像ID',
    'トップ画像URL'
  ];
  const result = {};
  allowedKeys.forEach(function(key) {
    if (Object.prototype.hasOwnProperty.call(all, key)) {
      result[key] = all[key];
    }
  });
  return result;
}

/*************************************************
 * 管理画面用データ
 *************************************************/
function getAdminData(token) {
  requireAdmin_(token);
  const banner = getBanner_();
  return {
    settings: getSettings_(),
    bannerMessage: banner.message,
    bannerEndDate: banner.endDate,
    notices: getSheetData_(SHEET_NAMES.NOTICE, true),
    events: getSheetData_(SHEET_NAMES.EVENT, true),
    schedules: getSheetData_(SHEET_NAMES.SCHEDULE, true)
  };
}

/*************************************************
 * シートデータ取得
 *
 * includePrivate = true（管理画面）のときだけ、
 * IDが空の行へIDを自動で書き込む。
 * 公開画面の閲覧ではシートへ書き込まない。
 *************************************************/
function getSheetData_(sheetName, includePrivate) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetName);

  if (!sheet) {
    return [];
  }

  if (CONTENT_SHEET_HEADERS[sheetName]) {
    const headerValues = sheet.getLastColumn()
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      : [];
    if (headerValues.indexOf('重要度') === -1) {
      ensureTimestampColumns_(sheetName, CONTENT_SHEET_HEADERS[sheetName]);
    }
  }

  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  const headers = values[0];
  const tz = Session.getScriptTimeZone();
  const result = [];

  for (let i = 1; i < values.length; i++) {

    const row = values[i];

    // 行全体が空ならスキップ
    const isEmptyRow = row.every(function(value) {
      return value === '' || value === null;
    });

    if (isEmptyRow) {
      continue;
    }

    const obj = {};

    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });
    obj['添付ファイル'] = parseBoardAttachments_(obj['添付ファイル']);

    // 公開データでは非公開を除外
    if (!includePrivate) {
      const published = obj['公開'];
      if (published !== true && String(published).toUpperCase() !== 'TRUE') {
        continue;
      }
    }

    // IDが空なら自動生成（管理画面のときだけシートにも書き込む）
    if (!obj['ID']) {
      obj['ID'] = createId_();
      if (includePrivate) {
        sheet.getRange(i + 1, 1).setValue(obj['ID']);
      }
    }

    // 日付を文字列に変換
    if (obj['日付'] instanceof Date) {
      obj['日付'] = formatDate_(obj['日付']);
    }

    // 登録日時・更新日時は文字列で返す（画面には表示しない）
    ['登録日時', '更新日時'].forEach(function(key) {
      if (obj[key] instanceof Date) {
        obj[key] = Utilities.formatDate(obj[key], tz, 'yyyy/MM/dd HH:mm:ss');
      }
    });

    result.push(obj);
  }

  // 表示順の昇順。同順位は同順位として扱い、空欄・数値でない値は最後に表示する。
  result.sort(function(a, b) {
    const rawA = a['表示順'];
    const rawB = b['表示順'];
    const hasA = rawA !== '' && rawA !== null && rawA !== undefined && Number.isFinite(Number(rawA));
    const hasB = rawB !== '' && rawB !== null && rawB !== undefined && Number.isFinite(Number(rawB));
    if (!hasA && !hasB) return 0;
    if (!hasA) return 1;
    if (!hasB) return -1;
    return Number(rawA) - Number(rawB);
  });

  return result;
}

/*************************************************
 * お知らせ・イベント・スケジュールの保存
 * type: notice / event / schedule
 *************************************************/
function saveItem(token, type, data, userAgent) {

  requireAdmin_(token);

  const cfg = getTypeConfig_(type);
  return saveItemWithAttachments_(cfg, data, [], userAgent);
}

function saveItemWithAttachments(token, type, data, uploads, userAgent) {
  requireAdmin_(token);
  const cfg = getTypeConfig_(type);
  return saveItemWithAttachments_(cfg, data, uploads, userAgent);
}

function saveItemWithAttachments_(cfg, data, uploads, userAgent) {
  const itemData = Object.assign({}, data || {});
  let attachmentInput = itemData['添付ファイル'];
  const currentItem = itemData.ID
    ? getSheetData_(cfg.sheet, true).find(function(item) {
      return String(item.ID) === String(itemData.ID);
    })
    : null;
  if (!Array.isArray(attachmentInput)) {
    attachmentInput = currentItem ? currentItem['添付ファイル'] : [];
  }
  const retainedIds = Array.isArray(attachmentInput)
    ? attachmentInput.map(function(file) {
        return typeof file === 'string' ? file : String(file && file.id || '');
      })
    : [];
  const previousAttachments = currentItem && Array.isArray(currentItem['添付ファイル'])
    ? currentItem['添付ファイル']
    : [];
  const previousIdSet = {};
  previousAttachments.forEach(function(file) {
    previousIdSet[String(file.id)] = true;
  });
  const retainedIdSet = {};
  retainedIds.forEach(function(id) {
    if (!id || retainedIdSet[id]) throw new Error('添付ファイルの指定が重複、または正しくありません。');
    if (!previousIdSet[id]) throw new Error('この投稿に添付されていないファイルは指定できません。');
    retainedIdSet[id] = true;
  });
  if (!Array.isArray(uploads)) throw new Error('添付ファイルデータが正しくありません。');
  const pendingUploads = uploads;
  if (pendingUploads.length > BOARD_ATTACHMENT_MAX_FILES) {
    throw new Error('一度にアップロードできるファイルは' + BOARD_ATTACHMENT_MAX_FILES + '件までです。');
  }

  const createdFiles = [];
  const trashedFiles = [];
  let result;
  try {
    const attachments = retainedIds.map(function(id) {
      return getBoardAttachmentMetadata_(cfg, id);
    });
    const folder = pendingUploads.length ? getBoardAttachmentFolder_(cfg) : null;
    let totalBytes = 0;
    pendingUploads.forEach(function(upload) {
      const file = createBoardAttachment_(folder, upload, createdFiles);
      totalBytes += file.size;
      if (totalBytes > BOARD_ATTACHMENT_MAX_BYTES) {
        throw new Error('添付ファイルの合計サイズは10MBまでです。');
      }
      attachments.push({ id: file.id, name: file.name });
    });

    previousAttachments.forEach(function(previousFile) {
      if (retainedIdSet[String(previousFile.id)]) return;
      const file = getBoardAttachmentFile_(cfg, previousFile.id);
      file.setTrashed(true);
      trashedFiles.push(file);
    });

    itemData['添付ファイル'] = JSON.stringify(attachments);
    result = saveData_(cfg.sheet, itemData);
  } catch (error) {
    trashedFiles.forEach(function(file) {
      try {
        file.setTrashed(false);
      } catch (restoreError) {
        console.error('保存失敗後に添付ファイルを復元できませんでした:', restoreError);
      }
    });
    createdFiles.forEach(function(file) {
      try {
        file.setTrashed(true);
      } catch (cleanupError) {
        console.error('保存失敗後の添付ファイル削除に失敗しました:', cleanupError);
      }
    });
    throw error;
  }
  recordAdminActionAccessLog_(cfg.label + '保存完了', userAgent);
  return result;
}

/*************************************************
 * お知らせ・イベント・スケジュールの削除
 *************************************************/
function deleteItem(token, type, id, userAgent) {

  requireAdmin_(token);

  const cfg = getTypeConfig_(type);
  const result = deleteData_(cfg.sheet, id);
  recordAdminActionAccessLog_(cfg.label + '削除完了', userAgent);
  return result;
}

/*************************************************
 * 管理画面のドラッグ操作で表示順を一括更新
 * type: notice / event / schedule
 * orderedIds: 管理リストの上から順に並んだID配列
 *************************************************/
function saveDisplayOrder(token, type, orderedIds, userAgent) {
  requireAdmin_(token);
  const cfg = getTypeConfig_(type);
  if (!Array.isArray(orderedIds)) throw new Error('表示順データが正しくありません。');

  const headers = ensureTimestampColumns_(cfg.sheet, CONTENT_SHEET_HEADERS[cfg.sheet]);
  const idColumn = headers.indexOf('ID') + 1;
  const orderColumn = headers.indexOf('表示順') + 1;
  if (idColumn < 1 || orderColumn < 1) throw new Error('ID列または表示順列がありません。');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return true;

  const idValues = sheet.getRange(2, idColumn, lastRow - 1, 1).getValues();
  const rowById = {};
  idValues.forEach(function(row, index) {
    if (row[0] !== '' && row[0] !== null) rowById[String(row[0])] = index + 2;
  });

  const seen = {};
  orderedIds.forEach(function(id, index) {
    const key = String(id);
    if (seen[key]) throw new Error('重複したデータが含まれています。');
    seen[key] = true;
    const rowNumber = rowById[key];
    if (!rowNumber) throw new Error('対象データが見つかりません。画面を再読み込みしてください。');
    sheet.getRange(rowNumber, orderColumn).setValue(index + 1);
  });

  recordAdminActionAccessLog_(cfg.label + '表示順変更完了', userAgent);
  return true;
}

/*************************************************
 * お知らせ・イベント・スケジュールの管理列を保証
 * 登録日時・更新日時・表示順を管理用列として用意する
 * 戻り値はシートの実際のヘッダー配列
 *************************************************/
function ensureTimestampColumns_(sheetName, baseHeaders) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(sheetName + ' シートがありません。setupSheets()を実行してください。');
  }

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, baseHeaders.length).setValues([baseHeaders]);
  }

  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getValues()[0].map(function(value) { return String(value || '').trim(); });

  // 空のヘッダー列を除いて、実際のデータ列を扱う
  while (headers.length && headers[headers.length - 1] === '') {
    headers.pop();
  }

  // 必須列と管理列が既存シートにない場合は末尾に追加（既存データ列は移動しない）
  (baseHeaders || []).concat(['登録日時', '更新日時', '表示順']).forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(header);
      headers.push(header);
    }
  });

  // 既存データの表示順が空欄なら、現在の行順を保つように末尾の番号を割り当てる。
  const orderColumn = headers.indexOf('表示順') + 1;
  if (orderColumn > 0 && sheet.getLastRow() >= 2) {
    const orderValues = sheet.getRange(2, orderColumn, sheet.getLastRow() - 1, 1).getValues();
    let maxOrder = getMaxDisplayOrder_(orderValues);
    let changed = false;
    orderValues.forEach(function(row) {
      if (row[0] === '' || row[0] === null) {
        maxOrder += 1;
        row[0] = maxOrder;
        changed = true;
      }
    });
    if (changed) sheet.getRange(2, orderColumn, orderValues.length, 1).setValues(orderValues);
  }

  return headers;
}

/*************************************************
 * コート当番ファイル（Googleドライブ）
 *
 * スクリプトプロパティ COURT_DUTY_FOLDER_ID に
 * 保存先フォルダのID（またはフォルダURL）を設定する。
 * すべて管理者セッションが必要。公開画面からは使えない。
 *************************************************/
const COURT_DUTY_MAX_BYTES = 10 * 1024 * 1024; // 1ファイル10MBまで
const COURT_DUTY_ALLOWED_EXT = ['pdf', 'xlsx', 'xls', 'docx', 'doc', 'csv', 'png', 'jpg', 'jpeg'];

function getCourtDutyFolder_() {
  const value = PropertiesService.getScriptProperties().getProperty('COURT_DUTY_FOLDER_ID');
  const id = normalizeDriveFileId_(value);
  if (!id) {
    throw new Error('コート当番の保存先が未設定です。スクリプトプロパティ COURT_DUTY_FOLDER_ID にDriveフォルダのIDを設定してください。');
  }
  try {
    return DriveApp.getFolderById(id);
  } catch (error) {
    const detail = error && error.message ? error.message : String(error);
    console.error('コート当番の保存先フォルダを開けません:', detail);
    throw new Error(
      'コート当番の保存先フォルダを開けません。COURT_DUTY_FOLDER_IDとGASの実行ユーザーの権限を確認してください。詳細: ' +
      detail
    );
  }
}

function verifyCourtDutyFolderAccess() {
  const folder = getCourtDutyFolder_();
  console.log('コート当番フォルダへのアクセスを確認しました:', folder.getName());
  return true;
}

const BOARD_ATTACHMENT_MAX_BYTES = 10 * 1024 * 1024;
const BOARD_ATTACHMENT_MAX_FILES = 5;
const BOARD_ATTACHMENT_ALLOWED_EXT = [
  'pdf', 'xlsx', 'xls', 'docx', 'doc', 'pptx', 'ppt', 'csv',
  'txt', 'png', 'jpg', 'jpeg', 'gif', 'webp'
];

function getBoardAttachmentFolder_(cfg) {
  const propertyName = cfg && cfg.attachmentFolderProperty;
  if (!propertyName) throw new Error('添付ファイルの種類が正しくありません。');
  const value = PropertiesService.getScriptProperties().getProperty(propertyName);
  const id = normalizeDriveFileId_(value);
  if (!id) {
    throw new Error(cfg.label + 'の保存先が未設定です。スクリプトプロパティ ' + propertyName + ' にDriveフォルダのIDを設定してください。');
  }
  try {
    return DriveApp.getFolderById(id);
  } catch (error) {
    const detail = error && error.message ? error.message : String(error);
    console.error(cfg.label + '添付ファイルの保存先フォルダを開けません:', detail);
    throw new Error(cfg.label + 'の保存先フォルダを開けません。' + propertyName + 'とGAS実行ユーザーの権限を確認してください。詳細: ' + detail);
  }
}

function verifyBoardAttachmentFolderAccess(type) {
  const cfg = getTypeConfig_(type);
  const folder = getBoardAttachmentFolder_(cfg);
  const probe = folder.createFile('board-attachment-permission-check.txt', '権限確認用の一時ファイルです。');
  let sharingError = null;
  try {
    probe.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    sharingError = error;
  }
  try {
    probe.setTrashed(true);
  } catch (cleanupError) {
    console.error('権限確認用ファイルをゴミ箱へ移動できません:', cleanupError);
    throw cleanupError;
  }
  if (sharingError) throw sharingError;
  console.log(cfg.label + '添付ファイルの保存先フォルダへのアクセスを確認しました:', folder.getName());
  return true;
}

function verifyNoticeAttachmentFolderAccess() {
  return verifyBoardAttachmentFolderAccess('notice');
}

function verifyEventAttachmentFolderAccess() {
  return verifyBoardAttachmentFolderAccess('event');
}

function verifyScheduleAttachmentFolderAccess() {
  return verifyBoardAttachmentFolderAccess('schedule');
}

function createBoardAttachment_(folder, upload, createdFiles) {
  if (!upload || typeof upload.base64 !== 'string' || !upload.base64) {
    throw new Error('添付ファイルの内容が空です。');
  }
  const name = String(upload.name || '')
    .replace(/[\\\/:*?"<>|\u0000-\u001f]/g, '_')
    .trim();
  const ext = name.indexOf('.') >= 0 ? name.split('.').pop().toLowerCase() : '';
  if (!name || BOARD_ATTACHMENT_ALLOWED_EXT.indexOf(ext) === -1) {
    throw new Error('添付できるファイル形式ではありません: ' + name);
  }

  const bytes = Utilities.base64Decode(upload.base64);
  if (bytes.length > BOARD_ATTACHMENT_MAX_BYTES) {
    throw new Error('添付ファイルは1件10MBまでです。');
  }
  const blob = Utilities.newBlob(
    bytes,
    String(upload.mimeType || 'application/octet-stream'),
    name
  );
  const file = folder.createFile(blob);
  createdFiles.push(file);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { id: file.getId(), name: file.getName(), size: bytes.length };
}

function getBoardAttachmentFile_(cfg, fileId) {
  const id = String(fileId || '');
  if (!/^[\w-]{25,}$/.test(id)) throw new Error('添付ファイルの指定が正しくありません。');

  let file;
  try {
    file = DriveApp.getFileById(id);
  } catch (error) {
    throw new Error('添付ファイルが見つかりません: ' + id);
  }
  const parents = file.getParents();
  const allowedFolderId = normalizeDriveFileId_(
    PropertiesService.getScriptProperties().getProperty(cfg.attachmentFolderProperty)
  );
  while (parents.hasNext()) {
    if (parents.next().getId() === allowedFolderId) {
      return file;
    }
  }
  throw new Error('この種類の添付用フォルダ以外のファイルは指定できません。');
}

function getBoardAttachmentMetadata_(cfg, fileId) {
  const file = getBoardAttachmentFile_(cfg, fileId);
  return { id: file.getId(), name: file.getName() };
}

function parseBoardAttachments_(value) {
  if (value === '' || value === null || value === undefined) return [];
  try {
    const parsed = JSON.parse(String(value));
    if (!Array.isArray(parsed)) throw new Error('添付ファイル情報が配列ではありません。');
    return parsed.filter(function(file) {
      return file && /^[\w-]{25,}$/.test(String(file.id || '')) && typeof file.name === 'string';
    }).map(function(file) {
      return { id: String(file.id), name: file.name };
    });
  } catch (error) {
    console.error('添付ファイル情報を読み取れません:', error);
    return [];
  }
}

// ファイル一覧（更新日時の新しい順）
function listCourtDutyFiles(token) {
  requireAdmin_(token);
  const folder = getCourtDutyFolder_();
  const tz = Session.getScriptTimeZone();
  const result = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    result.push({
      id: f.getId(),
      name: f.getName(),
      size: f.getSize(),
      updated: Utilities.formatDate(f.getLastUpdated(), tz, 'yyyy/MM/dd HH:mm'),
      time: f.getLastUpdated().getTime()
    });
  }
  result.sort(function(a, b) { return b.time - a.time; });
  return result;
}

// 保存先フォルダ内のファイルだけを返す（他フォルダのファイルは操作させない）
function getCourtDutyFile_(folder, fileId) {
  let file;
  try {
    file = DriveApp.getFileById(String(fileId || ''));
  } catch (e) {
    throw new Error('ファイルが見つかりません。');
  }

  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === folder.getId()) return file;
  }
  throw new Error('このファイルは操作できません。');
}

// ダウンロード用にファイル本体をBase64で返す（保存先フォルダ内のファイルのみ）
function getCourtDutyFileData(token, fileId) {
  requireAdmin_(token);
  const file = getCourtDutyFile_(getCourtDutyFolder_(), fileId);
  if (file.getSize() > COURT_DUTY_MAX_BYTES) throw new Error('ファイルが大きすぎます（10MBまで）。');

  const blob = file.getBlob();
  return {
    name: file.getName(),
    mimeType: blob.getContentType() || 'application/octet-stream',
    base64: Utilities.base64Encode(blob.getBytes())
  };
}

// 管理画面から選んだファイルをDriveの保存先フォルダへ保存
function uploadCourtDutyFile(token, fileName, mimeType, base64, userAgent) {
  requireAdmin_(token);

  // フォルダ区切りや制御文字を除いたファイル名にする
  const name = String(fileName || '').replace(/[\\\/:*?"<>|\u0000-\u001f]/g, '_').trim();
  const ext = name.indexOf('.') >= 0 ? name.split('.').pop().toLowerCase() : '';
  if (!name || COURT_DUTY_ALLOWED_EXT.indexOf(ext) === -1) {
    throw new Error('アップロードできるのは ' + COURT_DUTY_ALLOWED_EXT.join(' / ') + ' ファイルのみです。');
  }
  if (typeof base64 !== 'string' || !base64) throw new Error('ファイルの内容が空です。');

  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > COURT_DUTY_MAX_BYTES) throw new Error('ファイルが大きすぎます（10MBまで）。');

  const folder = getCourtDutyFolder_();
  const blob = Utilities.newBlob(bytes, String(mimeType || 'application/octet-stream'), name);
  const file = folder.createFile(blob);

  recordAdminActionAccessLog_('コート当番ファイルアップロード完了', userAgent);
  return { success: true, id: file.getId(), name: file.getName() };
}

// ファイルをDriveのごみ箱へ移す（復元可能。保存先フォルダ内のファイルのみ）
function deleteCourtDutyFile(token, fileId, userAgent) {
  requireAdmin_(token);
  const file = getCourtDutyFile_(getCourtDutyFolder_(), fileId);
  file.setTrashed(true);
  recordAdminActionAccessLog_('コート当番ファイル削除完了', userAgent);
  return { success: true };
}

/*************************************************
 * 表示順の最大値を取得（空欄・数値でない値は無視）
 * values: getValues() で取得した1列分の2次元配列
 *************************************************/
function getMaxDisplayOrder_(values) {
  let max = 0;
  values.forEach(function(row) {
    const n = Number(row[0]);
    if (row[0] !== '' && row[0] !== null && Number.isFinite(n)) max = Math.max(max, n);
  });
  return max;
}

/*************************************************
 * データ保存
 *************************************************/
function saveData_(sheetName, data) {
  const headers = ensureTimestampColumns_(sheetName, CONTENT_SHEET_HEADERS[sheetName]);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  const id = data.ID || createId_();

  // 既存データかどうかを確認
  const row = findRowById_(sheet, id);
  const now = new Date();

  // 表示順が未入力の場合、編集は既存値を維持し、新規は末尾の番号にする。
  const isValidOrder = function(v) {
    return v !== '' && v !== null && v !== undefined && Number.isFinite(Number(v));
  };
  const orderColumn = headers.indexOf('表示順') + 1;
  let displayOrder = data['表示順'];
  if (!isValidOrder(displayOrder)) {
    if (row > 0 && orderColumn > 0) {
      displayOrder = sheet.getRange(row, orderColumn).getValue();
    }
    if (!isValidOrder(displayOrder)) {
      displayOrder = 1;
      if (orderColumn > 0 && sheet.getLastRow() >= 2) {
        displayOrder += getMaxDisplayOrder_(
          sheet.getRange(2, orderColumn, sheet.getLastRow() - 1, 1).getValues()
        );
      }
    }
  }
  displayOrder = Number(displayOrder);

  // 新規登録時は登録日時・更新日時を同じ時刻にする。
  // 編集時は登録日時を維持し、更新日時だけを現在時刻にする。
  let registrationTime = now;
  if (row > 0) {
    const registrationColumn = headers.indexOf('登録日時') + 1;
    if (registrationColumn > 0) {
      const oldRegistrationTime = sheet.getRange(row, registrationColumn).getValue();
      if (oldRegistrationTime) {
        registrationTime = oldRegistrationTime;
      }
    }
  }

  const rowData = headers.map(function(header) {
    if (header === 'ID') return id;
    if (header === '公開') {
      return data[header] === true || String(data[header]).toUpperCase() === 'TRUE';
    }
    if (header === '重要度') return Number(data[header]) === 1 ? 1 : 0;
    if (header === '登録日時') return registrationTime;
    if (header === '更新日時') return now;
    if (header === '表示順') return displayOrder;
    if (header === '内容') return sanitizeCell_(sanitizeRichTextHtml_(data[header]));
    if (header === '添付ファイル') {
      return typeof data[header] === 'string' ? data[header] : JSON.stringify(data[header] || []);
    }
    return sanitizeCell_(data[header]);
  });

  if (row > 0) {
    sheet.getRange(row, 1, 1, headers.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  // 日時列は日時として表示できる書式にする（画面には表示しない）。
  const savedRow = row > 0 ? row : sheet.getLastRow();
  ['登録日時', '更新日時'].forEach(function(header) {
    const column = headers.indexOf(header) + 1;
    if (column > 0) {
      sheet.getRange(savedRow, column).setNumberFormat('yyyy/MM/dd HH:mm:ss');
    }
  });

  return {
    success: true,
    id: id
  };
}

function sanitizeRichTextHtml_(value) {
  const allowedColors = [
    '#000000', '#d32f2f', '#ef6c00', '#f9a825',
    '#2e7d32', '#1565c0', '#6a1b9a', '#ffffff'
  ];
  const allowedTags = {
    b: true, strong: true, u: true, s: true, strike: true, del: true,
    i: true, em: true, a: true, span: true, br: true, p: true,
    div: true, ul: true, ol: true, li: true
  };
  const html = String(value || '').replace(/<!--[\s\S]*?-->/g, '');
  return html.replace(/<\/?([a-z][a-z0-9]*)\b([^>]*)>/gi, function(tag, rawName, attrs) {
    const name = rawName.toLowerCase();
    if (!allowedTags[name]) return '';
    if (tag.charAt(1) === '/') return name === 'br' ? '' : '</' + name + '>';
    if (name === 'br') return '<br>';

    if (name === 'a') {
      const hrefMatch = attrs.match(/\bhref\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i);
      const href = hrefMatch && (hrefMatch[1] || hrefMatch[2] || hrefMatch[3] || '')
        .replace(/&amp;/gi, '&')
        .trim();
      if (!href || !/^(https?:\/\/|mailto:)/i.test(href)) return '';
      return '<a href="' + href.replace(/&/g, '&amp;').replace(/"/g, '&quot;') +
        '" target="_blank" rel="noopener noreferrer">';
    }

    if (name === 'span') {
      const styleMatch = attrs.match(/\bstyle\s*=\s*(?:"([^"]*)"|'([^']*)')/i);
      const style = styleMatch ? (styleMatch[1] || styleMatch[2] || '') : '';
      const safeStyles = [];
      ['color', 'background-color'].forEach(function(property) {
        const match = style.match(new RegExp('(?:^|;)\\s*' + property + '\\s*:\\s*(#[0-9a-f]{6})\\s*(?:;|$)', 'i'));
        const color = match && match[1].toLowerCase();
        if (allowedColors.indexOf(color) !== -1) safeStyles.push(property + ':' + color);
      });
      return safeStyles.length ? '<span style="' + safeStyles.join(';') + '">' : '<span>';
    }

    return '<' + name + '>';
  });
}

/*************************************************
 * セル値の無害化
 *
 * = + - @ で始まる文字列が数式として解釈されないようにする
 *************************************************/
function sanitizeCell_(value) {

  if (value === null || value === undefined || value === '') {
    return '';
  }

  const text = String(value);

  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

/*************************************************
 * データ削除
 *************************************************/
function deleteData_(sheetName, id) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(sheetName + ' シートがありません。');
  }

  const row = findRowById_(sheet, id);

  if (row > 0) {
    sheet.deleteRow(row);
  }

  return {
    success: true
  };
}

/*************************************************
 * ID検索
 *************************************************/
function findRowById_(sheet, id) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 0;
  }

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      return i + 2;
    }
  }

  return 0;
}

/*************************************************
 * ID生成
 *************************************************/
function createId_() {
  return Utilities.getUuid();
}

/*************************************************
 * 日付フォーマット
 *************************************************/
function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/*************************************************
 * 日付の値を yyyy-MM-dd の文字列にそろえる
 *
 * Date型、または「2026-11-20」「2026/11/20」形式の文字列を受け付ける。
 * 日付として正しくない値は空文字を返す。
 *************************************************/
function toDateKey_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? '' : formatDate_(value);
  }

  const match = String(value == null ? '' : value).trim().match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (!match) return '';

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const check = new Date(year, month - 1, day);

  // 2月31日のような存在しない日付を除く
  if (check.getFullYear() !== year || check.getMonth() !== month - 1 || check.getDate() !== day) {
    return '';
  }
  return match[1] + '-' + ('0' + month).slice(-2) + '-' + ('0' + day).slice(-2);
}

/*************************************************
 * アクセスログ・アクセス集計の保存先スプレッドシート
 *
 * スクリプトプロパティ ACCESS_LOG_SPREADSHEET_ID に、
 * 別のスプレッドシートのID（またはURL）を設定する。
 * 未設定の場合は、このスクリプトのスプレッドシートを使う。
 *************************************************/
function getAccessLogSpreadsheet_() {

  const value = PropertiesService.getScriptProperties().getProperty('ACCESS_LOG_SPREADSHEET_ID');
  const id = normalizeDriveFileId_(value);

  if (!id) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  try {
    return SpreadsheetApp.openById(id);
  } catch (error) {
    const detail = error && error.message ? error.message : String(error);
    console.error('アクセスログ用スプレッドシートを開けません:', detail);
    throw new Error(
      'アクセスログ用スプレッドシートを開けません。ACCESS_LOG_SPREADSHEET_IDとGASの実行ユーザーの権限を確認してください。詳細: ' +
      detail
    );
  }
}

// アクセスログ用スプレッドシートのシートを取得（なければ null）
function getAccessLogSheet_(sheetName) {
  return getAccessLogSpreadsheet_().getSheetByName(sheetName);
}

// エディタから実行して、保存先スプレッドシートを開けるか確認する
function verifyAccessLogSpreadsheetAccess() {
  const ss = getAccessLogSpreadsheet_();
  console.log('アクセスログの保存先スプレッドシートを確認しました:', ss.getName());
  return true;
}

/*************************************************
 * アクセスログ（共通処理）
 * 日時・ページ・端末・ブラウザー種類・ユーザーエージェントを1行追加
 *************************************************/
function appendAccessLog_(page, userAgent) {

  const sheet = getAccessLogSheet_(SHEET_NAMES.ACCESS_LOG);

  if (!sheet) {
    throw new Error('アクセスログシートがありません。setupSheets()を一度実行してください。');
  }

  const ua = String(userAgent || '');

  sheet.appendRow([
    new Date(),
    String(page || ''),
    detectDevice_(ua),
    detectBrowser_(ua),
    ua
  ]);
}

/*************************************************
 * ページアクセスログ
 *
 * トップを含む公開ページが表示されるたびに、
 * 同じセッション内でも毎回記録します。
 * アクセス集計には加算しません。
 *************************************************/
function savePageAccessLog(page, userAgent) {
  appendAccessLog_(page, userAgent);
  return true;
}

/*************************************************
 * 管理操作のログ
 * ログ記録だけが失敗しても、完了した保存・削除処理自体は失敗扱いにしない。
 *************************************************/
function recordAdminActionAccessLog_(action, userAgent) {
  try {
    appendAccessLog_('管理画面：' + String(action || ''), userAgent);
    return true;
  } catch (error) {
    console.error('管理操作ログの記録に失敗しました:', error);
    return false;
  }
}

/*************************************************
 * トップページの日別アクセス集計
 *
 * クライアント側のsessionStorageで
 * 「新しいセッションかどうか」を判定し、
 * 新しいセッションのトップ表示時だけ呼び出します。
 *************************************************/
function saveTopAccessSummary() {

  const sheet = getAccessLogSheet_(SHEET_NAMES.ACCESS_SUMMARY);
  if (!sheet) throw new Error('アクセス集計シートがありません。setupSheets()を一度実行してください。');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const dateString = formatDate_(new Date());
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { sheet.appendRow([dateString, 1]); return true; }

    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    let total = 1, firstRow = 0;
    const duplicates = [];
    for (let i = 0; i < values.length; i++) {
      if (toDateKey_(values[i][0]) === dateString) {
        const row = i + 2;
        total += Number(values[i][1]) || 0;
        if (!firstRow) firstRow = row; else duplicates.push(row);
      }
    }
    if (firstRow) {
      sheet.getRange(firstRow, 2).setValue(total);
      for (let i = duplicates.length - 1; i >= 0; i--) sheet.deleteRow(duplicates[i]);
    } else {
      sheet.appendRow([dateString, 1]);
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}

/*************************************************
 * 同じ日付の集計行を1行にまとめる（setupSheets()から実行）
 *************************************************/
function consolidateAccessSummary_() {
  const sheet = getAccessLogSheet_(SHEET_NAMES.ACCESS_SUMMARY);
  if (!sheet || sheet.getLastRow() < 2) return;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const totals = {};
  const first = {};
  const del = [];

  values.forEach(function(v, i) {
    const d = toDateKey_(v[0]);
    if (!d) return;

    const row = i + 2;
    const n = Number(v[1]) || 0;

    if (!(d in totals)) {
      totals[d] = n;
      first[d] = row;
    } else {
      totals[d] += n;
      del.push(row);
    }
  });

  Object.keys(totals).forEach(function(d) {
    sheet.getRange(first[d], 1, 1, 2).setValues([[d, totals[d]]]);
  });

  del.sort(function(a, b) { return b - a; });
  del.forEach(function(row) {
    sheet.deleteRow(row);
  });
}

/*************************************************
 * 30日より古いアクセスログを削除（毎日3時のトリガーで実行）
 *************************************************/
function cleanupOldAccessLogs() {
  const sheet = getAccessLogSheet_(SHEET_NAMES.ACCESS_LOG);
  if (!sheet || sheet.getLastRow() < 2) return;

  const cutoff = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const staleRows = [];

  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (!v) continue;
    const d = v instanceof Date ? v : new Date(v);
    if (!isNaN(d.getTime()) && d < cutoff) staleRows.push(i + 2);
  }

  // 連続した古い行をひとまとまりで削除し、行削除の呼び出し回数を減らす。
  let rangeStart = 0;
  let rangeEnd = 0;
  for (let i = staleRows.length - 1; i >= 0; i--) {
    const row = staleRows[i];
    if (!rangeEnd) {
      rangeStart = rangeEnd = row;
    } else if (row === rangeStart - 1) {
      rangeStart = row;
    } else {
      sheet.deleteRows(rangeStart, rangeEnd - rangeStart + 1);
      rangeStart = rangeEnd = row;
    }
  }
  if (rangeEnd) sheet.deleteRows(rangeStart, rangeEnd - rangeStart + 1);
}

function createAccessLogCleanupTrigger_() {
  const name = 'cleanupOldAccessLogs';
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === name)) {
    ScriptApp.newTrigger(name).timeBased().everyDays(1).atHour(3).create();
  }
}

/*************************************************
 * 端末判定
 *************************************************/
function detectDevice_(userAgent) {

  const ua = String(userAgent || '');

  if (/iPhone/i.test(ua)) return 'iPhone';
  if (/iPad/i.test(ua)) return 'iPad';
  if (/Android/i.test(ua)) return 'Android';
  if (/Windows/i.test(ua)) return 'Windows';
  if (/Macintosh/i.test(ua)) return 'Mac';
  if (/Linux/i.test(ua)) return 'Linux';

  return 'その他';
}

/*************************************************
 * ブラウザー判定
 *************************************************/
function detectBrowser_(userAgent) {

  const ua = String(userAgent || '');

  if (/Edg/i.test(ua)) return 'Edge';
  if (/CriOS/i.test(ua)) return 'Chrome (iOS)';
  if (/FxiOS/i.test(ua)) return 'Firefox (iOS)';
  if (/Chrome/i.test(ua)) return 'Chrome';
  if (/Firefox/i.test(ua)) return 'Firefox';
  if (/Safari/i.test(ua)) return 'Safari';

  return 'その他';
}
