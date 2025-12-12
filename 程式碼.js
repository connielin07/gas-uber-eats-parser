// 擷取 Uber Eats 電子郵件收據，解析 HTML 後將結構化資料寫入追蹤試算表。
const SPREADSHEET_ID = '1STKCbgS9Wn-1tJOBhXFDIhl72-TnbYEX2bLVoKfyDhw';
const SHEET_NAME = '工作表15';

// 同步流程的主要入口：掃描 Gmail、解析訂單資料並回寫到工作表。
function fetchUberEatsReceipts_HTML() {
  const tz = 'Asia/Taipei';
  const query = 'from:(noreply@uber.com) newer_than:365d';
  const threads = GmailApp.search(query);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  Logger.log(`共找到 ${threads.length} 封 thread`);

  // 若 E1 沒標題則補上（E 欄存 MsgID）
  if (!sheet.getRange(1, 5).getValue()) {
    sheet.getRange(1, 5).setValue('MsgID'); // E1
  }

  // 載入既有 MsgID 做去重
  const existingMsgId = new Set();
  const last = sheet.getLastRow();
  if (last >= 2) {
    const ids = sheet.getRange(2, 5, last - 1, 1).getValues(); // E2:E
    ids.forEach(r => {
      const v = (r[0] || '').toString().trim();
      if (v) existingMsgId.add(v);
    });
  }

  let ok = 0, err = 0, updated = 0, failed = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const msgId = msg.getId();                 // 只用 msgID 去重
      if (existingMsgId.has(msgId)) return;      // 已處理過 → 跳過

      try {
        const html = (msg.getBody() || '')
          .replace(/\u00A0/g, ' ')
          .replace(/&(nbsp|#160|#xA0);/gi, ' ')
          .replace(/\s+/g, ' ');

        // 是否為「更新收據」信件（簡體）
        const updateHdr = html.match(
          /<td[^>]*class="[^"]*Uber18_text_p1[^"]*header_Uber18_text_p1[^"]*"[^>]*>我们更新了您的\s*([^<]*?)\s*收据。?<\/td>/i
        );
        const adjustBlock = html.match(
          /<(?:td|div)[^>]*>[\s\S]*?我們已調整了您最近[\s\S]*?訂單的費用總額。[\s\S]*?<\/(?:td|div)>/i
        );
        let adjustMerchant = '';
        if (adjustBlock) {
          const txt = stripHtml(adjustBlock[0]);
          const m = txt.match(/我們已調整了您最近\s*(.*?)\s*訂單的費用總額。?/);
          if (m) adjustMerchant = m[1];
        }
        const headerMerchant = adjustMerchant
          ? decodeHtml(adjustMerchant)
          : updateHdr
            ? decodeHtml(updateHdr[1] || '')
            : '';
        const isAdjustMail = !!adjustBlock;
        const isUpdateMail = !!(updateHdr || adjustBlock);

        // 1) 日期
        const dateSpan = html.match(/<span[^>]*class="[^"]*Uber18_text_p1[^"]*"[^>]*>(.*?)<\/span>/g);
        let dateIso = '';
        if (dateSpan) {
          const dateStr = dateSpan
            .map(s => stripHtml(s))
            .find(t => /\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日/.test(t));
          if (dateStr) dateIso = chineseDateToISO(dateStr); // yyyy-MM-dd
        }
        if (!dateIso) {
          dateIso = Utilities.formatDate(msg.getDate(), tz, 'yyyy-MM-dd');
        }

        // 2) 金額
        let amountStr = '';
        const amountMatch = html.match(
          /<span[^>]*class="[^"]*Uber18_text_p2[^"]*"[^>]*>\s*(?:NT|TWD)?\$?\s*([\d,]+(?:\.\d{1,2})?)\s*<\/span>/i
        );
        if (amountMatch) {
          amountStr = amountMatch[1].replace(/,/g, '');
        } else {
          const fareTd = html.match(
            /<td[^>]*class="[^"]*total-fare-amount[^"]*"[^>]*>\s*(?:NT|TWD)?\$?\s*([\d,]+(?:\.\d{1,2})?)\s*<\/td>/i
          );
          if (fareTd) amountStr = fareTd[1].replace(/,/g, '');
        }
        const amount = amountStr ? Number(amountStr) : NaN;

        // 3) 來源
        const sourceDisplay = normalizeSource(headerMerchant || extractMerchant(html) || 'Uber Eats');
        const sourceKey = canonicalSource(sourceDisplay);

        // 4) 狀態
        const canParse = (!isNaN(amount) && amount >= 0 && dateIso);
        if (!canParse) {
          appendRecord(sheet, [dateIso, amountStr || '', sourceDisplay || 'Uber Eats', 'PARSE_ERROR', msgId]); // +E
          existingMsgId.add(msgId);
          err++;
          return;
        }

        if (isUpdateMail) {
          // 更新模式：新版 class 直接覆寫上一筆，舊版維持依來源比對
          let rowIdx = isAdjustMail ? getLastDataRow(sheet) : findLastRowBySource(sheet, sourceKey);
          if (rowIdx > 0) {
            Logger.log(`[UPDATE] Found row ${rowIdx} for ${sourceDisplay}`);
            sheet.getRange(rowIdx, 2).setValue(amountStr);     // B: 金額
            sheet.getRange(rowIdx, 4).setValue('OK_UPDATED');  // D: 狀態
            sheet.getRange(rowIdx, 5).setValue(msgId);         // ★ E: MsgID（記錄這封更新信）
            existingMsgId.add(msgId);
            updated++;
            return;
          }
          Logger.log(`[UPDATE] No match for ${sourceDisplay}, appending new row`);
          // 找不到既有紀錄 → 新增一列
          appendRecord(sheet, [dateIso, amountStr, sourceDisplay, 'OK_UPDATED', msgId]); // +E
          existingMsgId.add(msgId);
          updated++;
        } else {
          // 一般收據：新增
          appendRecord(sheet, [dateIso, amountStr, sourceDisplay, 'OK', msgId]); // +E
          existingMsgId.add(msgId);
          ok++;
        }
      } catch (processingError) {
        failed++;
        recordProcessingError(sheet, msg, msgId, tz, processingError);
        existingMsgId.add(msgId);
      }
    });
  });

  Logger.log(`Done. OK=${ok}, UPDATED=${updated}, PARSE_ERROR=${err}, PROCESS_ERROR=${failed}`);
}

// ---- 工具函式 ----
// 移除 HTML 標籤並修剪多餘空白，讓後續正規表示式只面對純文字。
function stripHtml(s) {
  return decodeHtml(String(s).replace(/<[^>]*>/g, '')).trim();
}
// 解碼 Uber 樣板常見的 HTML 實體。
function decodeHtml(s) {
  return s
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
}
// 將 yyyy年M月D日 格式的文字轉換成 yyyy-MM-dd。
function chineseDateToISO(s) {
  const m = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
  if (!m) return '';
  const y = m[1], mo = pad2(m[2]), d = pad2(m[3]);
  return `${y}-${mo}-${d}`;
}
// 將個位數月日補零，確保 ISO 日期文字可正確排序。
function pad2(x) { x = String(x); return x.length === 1 ? '0' + x : x; }
// 清理商家顯示文字（壓縮空白並去除首尾空格）。
function normalizeSource(s) {
  return (s || '').toString().replace(/\s+/g, ' ').trim();
}
// 透過移除空白與小寫化建立標準化名稱以便比較。
function canonicalSource(s) {
  return normalizeSource(s).replace(/\s+/g, '').toLowerCase();
}
// 取店名（來源）
// 嘗試多種在地化樣式，從 Uber 郵件中還原商家／店家名稱。
function extractMerchant(html) {
  const patterns = [
    /<td[^>]*class="[^"]*Uber18_text_p1[^"]*"[^>]*>您已訂購\s*([^<]*?)\s*的餐[點点]\s*<\/td>/i,
    /<td[^>]*class="[^"]*Uber18_text_p1[^"]*"[^>]*>您在\s*([^<]*?)\s*下了?[訂订][單单]\s*<\/td>/i,
    /<td[^>]*class="[^"]*Uber18_text_p1[^"]*"[^>]*>您在\s*([^<]*?)\s*下單\s*<\/td>/i,
    /<td[^>]*class="[^"]*Uber18_text_p1[^"]*"[^>]*>(?:You ordered from|You placed an order at)\s*([^<]*?)\s*<\/td>/i,
    /<(?:td|div)[^>]*>以下是您在\s*([^<]*?)\s*訂購[^<]*<\/(?:td|div)>/i,
    /<(?:td|div)[^>]*>我們已調整了您最近\s*([^<]*?)\s*訂單的費用總額。<\/(?:td|div)>/i,
  ];
  for (const re of patterns) {
    const m = html.match(re);
    if (m) return decodeHtml(m[1]).trim();
  }
  return '';
}
// 由下往上找「來源」最後一次出現的列（第1列預設表頭：日期,金額,來源,狀態）
// 從試算表底部往上搜尋，找到該商家最新的一列。
function findLastRowBySource(sheet, sourceKey) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const target = canonicalSource(sourceKey);
  const values = sheet.getRange(2, 1, last - 1, 4).getValues(); // A2:D
  for (let i = values.length - 1; i >= 0; i--) {
    if (canonicalSource(values[i][2]) === target) {
      return i + 2; // 轉回實際列號
    }
  }
  return -1;
}

// 新增 SYSTEM 紀錄以追蹤特定 MsgID 的處理失敗。
function recordProcessingError(sheet, msg, msgId, tz, error) {
  let fallbackDate = '-';
  try {
    fallbackDate = msg && typeof msg.getDate === 'function'
      ? Utilities.formatDate(msg.getDate(), tz, 'yyyy-MM-dd')
      : '-';
  } catch (dateErr) {
    console.error('Failed to format fallback date', dateErr);
  }
  const statusNote = `PROCESS_ERROR: ${(error && error.message) || error}`;
  appendRecord(sheet, [fallbackDate, '', 'SYSTEM', statusNote, msgId || '']);
  console.error(`Process error for MsgID=${msgId || 'NA'}`, error);
}

// 集中處理 append 行為，避免動到表頭公式並方便日後追蹤。
function appendRecord(sheet, values) {
  // 追加新列以維持表頭公式不被覆寫
  sheet.appendRow(values);
  return sheet.getLastRow();
}

// 回傳有效資料列索引（>1），若無資料則回傳 -1 供呼叫端判斷。
function getLastDataRow(sheet) {
  const last = sheet.getLastRow();
  return last >= 2 ? last : -1;
}

// ---- 觸發器 ----
// 建立（若尚未存在）每日觸發器，讓 Gmail 同步自動執行。
function ensureDailySyncTrigger() {
  const handler = 'fetchUberEatsReceipts_HTML';
  const exists = ScriptApp.getProjectTriggers().some(tr => tr.getHandlerFunction() === handler);
  if (exists) {
    return '每日同步觸發器已存在';
  }

  ScriptApp.newTrigger(handler)
    .timeBased()
    .atHour(3) // 每日 03:00（台北時間）
    .everyDays(1)
    .create();
  return '已建立每日同步觸發器（03:00）';
}

// 移除所有指向同步函式的觸發器並回報刪除數量。
function removeDailySyncTriggers() {
  const handler = 'fetchUberEatsReceipts_HTML';
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  return `已移除 ${removed} 個舊觸發器`;
}
