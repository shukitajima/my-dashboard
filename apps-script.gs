// ================================================================
//  ダッシュボード用 Google Apps Script
//
//  【再デプロイ手順】（コードを変更した場合）
//  デプロイ → デプロイを管理 → 編集（鉛筆アイコン）
//  → バージョン: 「新バージョン」を選択 → デプロイ
//  ※ URLは変わりません
// ================================================================

function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const settings = ss.getSheetByName('Settings')      || ss.insertSheet('Settings');
    const subs     = ss.getSheetByName('Subscriptions') || ss.insertSheet('Subscriptions');
    const action   = (e.parameter && e.parameter.action) || '';

    // ── getData ──────────────────────────────────────────────
    if (action === 'getData') {
      const rawCut  = settings.getRange('B1').getValue();
      const rawPerm = settings.getRange('B2').getValue();
      const fmt = v => v ? Utilities.formatDate(new Date(v), 'Asia/Tokyo', 'yyyy-MM-dd') : null;

      const lastRow = subs.getLastRow();
      let subscriptions = [];
      if (lastRow > 1) {
        subscriptions = subs.getRange(2, 1, lastRow - 1, 4).getValues()
          .filter(r => r[0])
          .map(r => ({
            name:        r[0],
            amount:      r[1],
            billingDate: r[2],
            frequency:   r[3] || 1,  // 旧データは毎月(1)として扱う
          }));
      }
      out.setContent(JSON.stringify({ cutDate: fmt(rawCut), permDate: fmt(rawPerm), subscriptions }));
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
      if (subs.getLastRow() === 0) {
        subs.appendRow(['サービス名', '金額', '引落日', '頻度(月)']);
      }
      subs.appendRow([
        e.parameter.name,
        Number(e.parameter.amount),
        e.parameter.billingDate,
        Number(e.parameter.frequency) || 1,
      ]);
      out.setContent(JSON.stringify({ success: true }));
    }

    // ── deleteSubscription ───────────────────────────────────
    else if (action === 'deleteSubscription') {
      const idx = Number(e.parameter.index);
      subs.deleteRow(idx + 2); // ヘッダー行(+1) + 0始まり(+1)
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

      const response = UrlFetchApp.fetch('https://transit.yahoo.co.jp/traininfo/area/4/', {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36' },
      });
      const html = response.getContentText('UTF-8');

      const targets = [
        { key: 'yamanote',   name: '山手線',      searches: ['山手線'] },
        { key: 'inokashira', name: '京王井の頭線', searches: ['京王井の頭線', '井の頭線'] },
        { key: 'chuo',       name: '中央線',      searches: ['中央線（快速）', '中央線'] },
        { key: 'sobu',       name: '総武線',      searches: ['総武線（各停）', '総武・中央線', '総武線'] },
        { key: 'keio',       name: '京王線',      searches: ['京王線'] },
        { key: 'odakyu',     name: '小田急線',    searches: ['小田急小田原線', '小田急線'] },
      ];

      const trains = {};
      targets.forEach(function(t) {
        var sectionHtml = '';
        for (var i = 0; i < t.searches.length; i++) {
          var idx = html.indexOf(t.searches[i]);
          if (idx !== -1) {
            sectionHtml = html.substring(Math.max(0, idx - 300), idx + 800);
            break;
          }
        }

        if (!sectionHtml) {
          // ページに路線名が見つからない = 平常通り（遅延路線のみ掲載の場合）
          trains[t.key] = { name: t.name, status: '平常通り運転', normal: true };
          return;
        }

        // HTMLタグを除去してテキスト化
        var text = sectionHtml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
        var hasIssue = /遅延|見合わせ|運休|障害|乱れ/.test(text);

        if (!hasIssue) {
          trains[t.key] = { name: t.name, status: '遅延情報なし', normal: true };
        } else {
          // 分数を抽出できればそれを優先
          var minuteMatch = text.match(/(\d+)\s*分[程度]*遅[延れ]/);
          var statusText = minuteMatch ? '約' + minuteMatch[1] + '分遅延' : '遅延あり';
          // 運転見合わせ・運休を優先
          if (/見合わせ/.test(text)) statusText = '運転見合わせ';
          if (/運休/.test(text))    statusText = '運休';
          trains[t.key] = { name: t.name, status: statusText, normal: false };
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
