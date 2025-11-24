const SPREADSHEET_ID = '1STKCbgS9Wn-1tJOBhXFDIhl72-TnbYEX2bLVoKfyDhw';
const SHEET_NAME = '工作表11';

function fetchUberEatsReceipts_HTML() {
  const tz = 'Asia/Taipei';
  const query = 'from:(noreply@uber.com) subject:("[LinJay的家庭]") newer_than:365d';
  const threads = GmailApp.search(query);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  Logger.log(`共找到 ${threads.length} 封 thread`);

  // ★ 若 E1 沒標題則補上（E 欄存 MsgID）
  if (!sheet.getRange(1, 5).getValue()) {
    sheet.getRange(1, 5).setValue('MsgID'); // E1
  }

  // ★ 載入既有 MsgID 做去重
  const existingMsgId = new Set();
  const last = sheet.getLastRow();
  if (last >= 2) {
    const ids = sheet.getRange(2, 5, last - 1, 1).getValues(); // E2:E
    ids.forEach(r => {
      const v = (r[0] || '').toString().trim();
      if (v) existingMsgId.add(v);
    });
  }

  let ok = 0, err = 0, updated = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const msgId = msg.getId();                 // ★ 只用 msgID 去重
      if (existingMsgId.has(msgId)) return;      // 已處理過 → 跳過

      const html = (msg.getBody() || '').replace(/\u00A0/g, ' ').replace(/\s+/g, ' ');

      // 是否為「更新收據」信件（簡體）
      const updateHdr = html.match(
        /<td[^>]*class="[^"]*Uber18_text_p1[^"]*header_Uber18_text_p1[^"]*"[^>]*>我们更新了您的\s*([^<]*?)\s*收据。?<\/td>/i
      );
      const isUpdateMail = !!updateHdr;

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
      const amountMatch = html.match(
        /<span[^>]*class="[^"]*Uber18_text_p2[^"]*"[^>]*>\s*(?:NT|TWD)?\$?\s*([\d,]+(?:\.\d{1,2})?)\s*<\/span>/i
      );
      const amountStr = amountMatch ? amountMatch[1].replace(/,/g, '') : '';
      const amount = amountStr ? Number(amountStr) : NaN;

      // 3) 來源
      let source = extractMerchant(html) || 'Uber Eats';

      // 4) 狀態
      const canParse = (!isNaN(amount) && amount > 0 && dateIso);
      if (!canParse) {
        sheet.appendRow([dateIso, amountStr || '', source || 'Uber Eats', 'PARSE_ERROR', msgId]); // +E
        existingMsgId.add(msgId);
        err++;
        return;
      }

      if (isUpdateMail) {
        // 更新模式：找最近一筆相同來源的紀錄，把金額覆寫（找不到就新增）
        const rowIdx = findLastRowBySource(sheet, source);
        if (rowIdx > 0) {
          sheet.getRange(rowIdx, 2).setValue(amountStr);     // B: 金額
          sheet.getRange(rowIdx, 4).setValue('OK_UPDATED');  // D: 狀態
          sheet.getRange(rowIdx, 5).setValue(msgId);         // ★ E: MsgID（記錄這封更新信）
          existingMsgId.add(msgId);
          updated++;
          return;
        }
        // 找不到既有紀錄 → 新增一列
        sheet.appendRow([dateIso, amountStr, source, 'OK_UPDATED', msgId]); // +E
        existingMsgId.add(msgId);
        updated++;
      } else {
        // 一般收據：新增
        sheet.appendRow([dateIso, amountStr, source, 'OK', msgId]); // +E
        existingMsgId.add(msgId);
        ok++;
      }
    });
  });

  Logger.log(`Done. OK=${ok}, UPDATED=${updated}, PARSE_ERROR=${err}`);
}

// ---- helpers ----
function stripHtml(s) {
  return decodeHtml(String(s).replace(/<[^>]*>/g, '')).trim();
}
function decodeHtml(s) {
  return s
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
}
function chineseDateToISO(s) {
  const m = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
  if (!m) return '';
  const y = m[1], mo = pad2(m[2]), d = pad2(m[3]);
  return `${y}-${mo}-${d}`;
}
function pad2(x) { x = String(x); return x.length === 1 ? '0' + x : x; }
// 取店名（來源）
function extractMerchant(html) {
  const patterns = [
    /<td[^>]*class="[^"]*Uber18_text_p1[^"]*"[^>]*>您已訂購\s*([^<]*?)\s*的餐[點点]\s*<\/td>/i,
    /<td[^>]*class="[^"]*Uber18_text_p1[^"]*"[^>]*>您在\s*([^<]*?)\s*下了?[訂订][單单]\s*<\/td>/i,
    /<td[^>]*class="[^"]*Uber18_text_p1[^"]*"[^>]*>您在\s*([^<]*?)\s*下單\s*<\/td>/i,
    /<td[^>]*class="[^"]*Uber18_text_p1[^"]*"[^>]*>(?:You ordered from|You placed an order at)\s*([^<]*?)\s*<\/td>/i,
  ];
  for (const re of patterns) {
    const m = html.match(re);
    if (m) return decodeHtml(m[1]).trim();
  }
  return '';
}
// 由下往上找「來源」最後一次出現的列（第1列預設表頭：日期,金額,來源,狀態）
function findLastRowBySource(sheet, source) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const values = sheet.getRange(2, 1, last - 1, 4).getValues(); // A2:D
  for (let i = values.length - 1; i >= 0; i--) {
    if ((values[i][2] || '').toString().trim() === source) {
      return i + 2; // 轉回實際列號
    }
  }
  return -1;
}