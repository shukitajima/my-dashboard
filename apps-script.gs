// ================================================================
//  ダッシュボード用 Google Apps Script
//
//  【シークレットトークンの設定手順】
//  1. スクリプトエディタ上部の「プロジェクトの設定」(歯車アイコン) を開く
//  2. 「スクリプト プロパティ」→「プロパティを追加」
//  3. プロパティ名: SECRET_TOKEN  /  値: 好きな文字列（例: my-secret-2024）
//  4. 保存してから再デプロイ
//
//  【再デプロイ手順】（コードを変更した場合）
//  デプロイ → デプロイを管理 → 編集（鉛筆アイコン）
//  → バージョン: 「新バージョン」を選択 → デプロイ
//  ※ URLは変わりません
// ================================================================

// スクリプトプロパティからトークンを取得
const SECRET_TOKEN = PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN') || '';
const CONTENT_HISTORY_SHEET_NAME = 'content_stats';

function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  // トークン認証
  if (SECRET_TOKEN && (e.parameter && e.parameter.token) !== SECRET_TOKEN) {
    out.setContent(JSON.stringify({ error: 'unauthorized' }));
    return out;
  }

  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const settings = ss.getSheetByName('Settings')      || ss.insertSheet('Settings');
    const subs     = ss.getSheetByName('Subscriptions') || ss.insertSheet('Subscriptions');
    const action   = (e.parameter && e.parameter.action) || '';

    // ── getData ──────────────────────────────────────────────
    if (action === 'getData') {
      ensureSubscriptionHeader(subs);
      const rawCut  = settings.getRange('B1').getValue();
      const rawPerm = settings.getRange('B2').getValue();
      const fmt = v => v ? Utilities.formatDate(new Date(v), 'Asia/Tokyo', 'yyyy-MM-dd') : null;
      const contentStats = getContentStats_(ss);

      const lastRow = subs.getLastRow();
      let subscriptions = [];
      if (lastRow > 1) {
        subscriptions = subs.getRange(2, 1, lastRow - 1, 5).getValues()
          .filter(r => r[0])
          .map((r, idx) => {
            const rowNumber = idx + 2;
            const id = r[4] || generateSubscriptionId();
            if (!r[4]) subs.getRange(rowNumber, 5).setValue(id);
            return {
              id:          id,
              name:        r[0],
              amount:      r[1],
              billingDate: r[2],
              frequency:   r[3] || 1,  // 旧データは毎月(1)として扱う
            };
          });
      }
      out.setContent(JSON.stringify({ cutDate: fmt(rawCut), permDate: fmt(rawPerm), subscriptions, contentStats }));
    }

    // ── updateTrackerDate ────────────────────────────────────
    else if (action === 'updateTrackerDate') {
      const key  = e.parameter.key;
      const date = e.parameter.date;
      if (key === 'cut') {
        settings.getRange('A1').setValue('last_cut_date');
        settings.getRange('B1').setValue(date);
      } else if (key === 'perm') {
        settings.getRange('A2').setValue('last_perm_date');
        settings.getRange('B2').setValue(date);
      }
      out.setContent(JSON.stringify({ success: true }));
    }

    // ── addSubscription ──────────────────────────────────────
    else if (action === 'addSubscription') {
      ensureSubscriptionHeader(subs);
      subs.appendRow([
        e.parameter.name,
        Number(e.parameter.amount),
        e.parameter.billingDate,
        Number(e.parameter.frequency) || 1,
        e.parameter.id || generateSubscriptionId(),
      ]);
      out.setContent(JSON.stringify({ success: true }));
    }

    // ── deleteSubscription ───────────────────────────────────
    else if (action === 'deleteSubscription') {
      ensureSubscriptionHeader(subs);
      const id = e.parameter.id || '';
      const row = findSubscriptionRowById(subs, id);
      if (!row) throw new Error('subscription_not_found');
      subs.deleteRow(row);
      out.setContent(JSON.stringify({ success: true }));
    }

    // ── getTrainStatus ───────────────────────────────────────
    else if (action === 'getTrainStatus') {
      // 5分キャッシュ（Yahoo!路線情報への過剰アクセス防止）
      const cache    = CacheService.getScriptCache();
      const cacheKey = 'train_status_v4';
      const cached   = cache.get(cacheKey);
      if (cached) {
        out.setContent(cached);
        return out;
      }

      // 各路線の個別ページを取得（エリア一括より精度が高い）
      const targets = [
        { key: 'yamanote',   name: '山手線',      url: 'https://transit.yahoo.co.jp/traininfo/detail/1/0/' },
        { key: 'inokashira', name: '京王井の頭線', url: 'https://transit.yahoo.co.jp/traininfo/detail/31/0/' },
        { key: 'chuo',       name: '中央線',      url: 'https://transit.yahoo.co.jp/traininfo/detail/13/0/' },
        { key: 'sobu',       name: '総武線',      url: 'https://transit.yahoo.co.jp/traininfo/detail/15/0/' },
        { key: 'keio',       name: '京王線',      url: 'https://transit.yahoo.co.jp/traininfo/detail/28/0/' },
        { key: 'odakyu',     name: '小田急線',    url: 'https://transit.yahoo.co.jp/traininfo/detail/32/0/' },
      ];

      const trains = {};
      targets.forEach(function(t) {
        try {
          var res = UrlFetchApp.fetch(t.url, {
            muteHttpExceptions: true,
            headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' },
          });
          var text = res.getContentText('UTF-8').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');

          if (/平常/.test(text)) {
            trains[t.key] = { name: t.name, status: '遅延情報なし', normal: true };
          } else if (/見合わせ/.test(text)) {
            trains[t.key] = { name: t.name, status: '運転見合わせ', normal: false };
          } else if (/運休/.test(text)) {
            trains[t.key] = { name: t.name, status: '運休', normal: false };
          } else if (/遅延|乱れ/.test(text)) {
            var minuteMatch = text.match(/(\d+)\s*分[程度]*遅[延れ]/);
            trains[t.key] = { name: t.name, status: minuteMatch ? '約' + minuteMatch[1] + '分遅延' : '遅延あり', normal: false };
          } else {
            trains[t.key] = { name: t.name, status: '遅延情報なし', normal: true };
          }
        } catch(err) {
          trains[t.key] = { name: t.name, status: '取得失敗', normal: true };
        }
      });

      var payload = JSON.stringify({ trains: trains, updatedAt: new Date().toISOString() });
      cache.put(cacheKey, payload, 300); // 5分
      out.setContent(payload);
    }

    // ── getNote ──────────────────────────────────────────────
    else if (action === 'getNote') {
      const text = settings.getRange('B3').getValue();
      out.setContent(JSON.stringify({ text: text || '' }));
    }

    // ── saveNote ──────────────────────────────────────────────
    else if (action === 'saveNote') {
      settings.getRange('A3').setValue('note');
      settings.getRange('B3').setValue(e.parameter.text || '');
      out.setContent(JSON.stringify({ success: true }));
    }

    else {
      out.setContent(JSON.stringify({ error: 'unknown action: ' + action }));
    }

  } catch(err) {
    out.setContent(JSON.stringify({ error: err.message }));
  }

  return out;
}

function ensureSubscriptionHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['サービス名', '金額', '引落日', '頻度(月)', 'ID']);
    return;
  }

  if (sheet.getRange(1, 5).getValue() !== 'ID') {
    sheet.getRange(1, 5).setValue('ID');
  }
}

function findSubscriptionRowById(sheet, id) {
  if (!id || sheet.getLastRow() <= 1) return null;
  const ids = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return null;
}

function generateSubscriptionId() {
  return 'sub_' + new Date().getTime() + '_' + Math.random().toString(36).slice(2, 10);
}

function getContentStats_(ss) {
  const sheet = ss.getSheetByName(CONTENT_HISTORY_SHEET_NAME);
  if (!sheet) return null;

  const index = getContentHistoryHeaderIndex_(sheet);
  if (!index) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const dailySnapshots = {};
  const latestByKey = {};

  values.forEach(function(row) {
    const category = normalizeCategory_(row[index.category]);
    const serviceId = String(row[index.service_id] || '').trim() || 'unknown';
    const totalCount = Number(row[index.total_count]) || 0;
    const rowDate = toDateKey_(row[index.date]);
    const fetchAtDate = toDateObject_(row[index.fetch_at]);
    const timestamp = fetchAtDate ? fetchAtDate.getTime() : (rowDate ? new Date(rowDate + 'T00:00:00+09:00').getTime() : 0);

    if (!category || !rowDate || !timestamp) return;

    const item = {
      serviceId: serviceId,
      category: category,
      totalCount: totalCount,
      date: rowDate,
      fetchAt: fetchAtDate ? fetchAtDate.toISOString() : null,
      timestamp: timestamp,
    };
    const key = serviceId + '::' + category;

    if (!latestByKey[key] || latestByKey[key].timestamp < timestamp) {
      latestByKey[key] = item;
    }

    dailySnapshots[rowDate] = dailySnapshots[rowDate] || {};
    if (!dailySnapshots[rowDate][key] || dailySnapshots[rowDate][key].timestamp < timestamp) {
      dailySnapshots[rowDate][key] = item;
    }
  });

  const timeline = Object.keys(dailySnapshots).sort().map(function(dateKey) {
    const categoryTotals = {};
    let totalCount = 0;

    Object.keys(dailySnapshots[dateKey]).forEach(function(key) {
      const item = dailySnapshots[dateKey][key];
      categoryTotals[item.category] = (categoryTotals[item.category] || 0) + item.totalCount;
      totalCount += item.totalCount;
    });

    return {
      date: dateKey,
      totalCount: totalCount,
      categories: categoryTotals,
    };
  });

  if (!timeline.length) return null;

  timeline.forEach(function(point, idx) {
    const prev = idx > 0 ? timeline[idx - 1] : null;
    point.dailyDiff = prev ? point.totalCount - prev.totalCount : 0;

    const categoryDiffs = {};
    Object.keys(point.categories).forEach(function(category) {
      const prevCount = prev && prev.categories[category] ? prev.categories[category] : 0;
      categoryDiffs[category] = point.categories[category] - prevCount;
    });
    if (prev) {
      Object.keys(prev.categories).forEach(function(category) {
        if (!(category in point.categories)) {
          categoryDiffs[category] = -prev.categories[category];
        }
      });
    }
    point.categoryDiffs = categoryDiffs;
  });

  const currentCategories = {};
  let currentTotal = 0;
  let latestTimestamp = 0;

  Object.keys(latestByKey).forEach(function(key) {
    const item = latestByKey[key];
    currentCategories[item.category] = (currentCategories[item.category] || 0) + item.totalCount;
    currentTotal += item.totalCount;
    latestTimestamp = Math.max(latestTimestamp, item.timestamp);
  });

  const latestTimeline = timeline[timeline.length - 1];
  const categoryOrder = ['movie', 'drama', 'anime', 'game', 'book', 'manga'];
  const categories = Object.keys(currentCategories)
    .sort(function(a, b) {
      const ai = categoryOrder.indexOf(a);
      const bi = categoryOrder.indexOf(b);
      if (ai === -1 && bi === -1) return a.localeCompare(b);
      if (ai === -1) return 1;
      if (bi === -1) return -1;
      return ai - bi;
    })
    .map(function(category) {
      return {
        category: category,
        totalCount: currentCategories[category],
        dailyDiff: latestTimeline.categoryDiffs[category] || 0,
      };
    });

  return {
    summary: {
      totalCount: currentTotal,
      totalDiff: latestTimeline.dailyDiff || 0,
      latestDate: latestTimeline.date,
      latestFetchAt: latestTimestamp ? new Date(latestTimestamp).toISOString() : null,
    },
    categories: categories,
    timeline: timeline.slice(-400),
  };
}

function getContentHistoryHeaderIndex_(sheet) {
  const required = ['date', 'service_id', 'category', 'total_count', 'daily_diff', 'fetch_at', 'record_id'];
  if (sheet.getLastRow() < 2) return null;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(normalizeHeader_);
  const index = {};
  required.forEach(function(name) {
    index[name] = headers.indexOf(name);
  });

  return required.every(function(name) { return index[name] >= 0; }) ? index : null;
}

function normalizeHeader_(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeCategory_(value) {
  return String(value || '').trim().toLowerCase();
}

function toDateObject_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return value;
  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
}

function toDateKey_(value) {
  const date = toDateObject_(value);
  if (!date) return null;
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}
