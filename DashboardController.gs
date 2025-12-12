// 提供 Uber Eats 儀表板 UI 以及彙整試算表資料給前端使用。
const DASHBOARD_TIMEZONE = 'Asia/Taipei';
const MONTHLY_STATS_SHEET = 'MonthlyStats';
const WEEKLY_STATS_SHEET = 'WeeklyStats';

// Apps Script doGet：回傳 Dashboard.html 模板供前端載入。
function doGet() {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Uber Eats Dashboard');
}

// 彙總主工作表資料，計算摘要、近期訂單與統計圖表所需的內容。
function getDashboardData() {
  const tz = DASHBOARD_TIMEZONE;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const emptySummary = {
    totalOrders: 0,
    totalAmount: 0,
    updatedOrders: 0,
    errorOrders: 0,
    thisWeekAmount: 0,
    thisMonthAmount: 0,
    thisYearAmount: 0,
    lastRefreshed: Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm')
  };
  if (!sheet) {
    return { summary: emptySummary, recent: [], statusMessage: '找不到指定的工作表' };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { summary: emptySummary, recent: [], statusMessage: '尚無資料' };
  }

  const rows = values.slice(1).filter(row => row.some(cell => cell !== '' && cell !== null));
  if (!rows.length) {
    return { summary: emptySummary, recent: [], statusMessage: '尚無資料' };
  }

  const now = new Date();
  const thisMonthKey = Utilities.formatDate(now, tz, 'yyyy-MM');
  const currentYear = Utilities.formatDate(now, tz, 'yyyy');
  const weekRange = getCurrentWeekRange(now, tz);

  let totalAmount = 0;
  let totalOrders = 0;
  let updatedOrders = 0;
  let errorOrders = 0;
  let thisMonthAmount = 0;
  let thisWeekAmount = 0;
  let thisYearAmount = 0;

  const parsedRows = rows.map(row => {
    const normalized = normalizeDashboardDate(row[0], tz);
    const amountNum = Number(row[1]) || 0;
    const status = (row[3] || '').toString();
    if (amountNum > 0) totalAmount += amountNum;
    totalOrders++;
    if (status === 'OK_UPDATED') updatedOrders++;
    if (status === 'PARSE_ERROR') errorOrders++;
    if (normalized.monthKey === thisMonthKey) {
      thisMonthAmount += amountNum;
    }
    if (!isNaN(normalized.dayMs) && normalized.dayMs >= weekRange.startMs && normalized.dayMs < weekRange.endMs) {
      thisWeekAmount += amountNum;
    }
    if (normalized.yearKey === currentYear) {
      thisYearAmount += amountNum;
    }

    return {
      date: normalized.display,
      timestamp: normalized.timestamp,
      amount: amountNum,
      source: (row[2] || 'Uber Eats').toString(),
      status
    };
  });

  const recent = parsedRows
    .sort((a, b) => b.timestamp - a.timestamp)
    .slice(0, 20);

  const monthlyStats = getMonthlyStatsData();
  const weeklyStats = getWeeklyStatsData();

  return {
    summary: {
      totalOrders,
      totalAmount,
      updatedOrders,
      errorOrders,
      thisWeekAmount,
      thisMonthAmount,
      thisYearAmount,
      lastRefreshed: Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm')
    },
    recent,
    statusMessage: `共 ${totalOrders} 筆紀錄`,
    monthlyStats,
    weeklyStats
  };
}

// 標準化工作表中的日期欄位，確保可排序與格式化。
function normalizeDashboardDate(value, tz) {
  let dateObj = null;
  if (value instanceof Date) {
    dateObj = value;
  } else if (typeof value === 'string' && value) {
    dateObj = new Date(value);
  } else if (typeof value === 'number') {
    dateObj = new Date(value);
  }

  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
    return { display: '-', timestamp: 0, monthKey: '', yearKey: '', dayMs: NaN };
  }

  const display = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
  const monthKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM');
  const yearKey = Utilities.formatDate(dateObj, tz, 'yyyy');
  const dayMs = new Date(display).getTime();

  return {
    display,
    timestamp: dateObj.getTime(),
    monthKey,
    yearKey,
    dayMs
  };
}

// 讀取「MonthlyStats」並轉成前端圖表可用的物件陣列。
function getMonthlyStatsData() {
  return readStatsSheet(MONTHLY_STATS_SHEET, row => {
    const monthLabel = formatMonthLabel(row[0]);
    if (!monthLabel) return null;
    return {
      month: monthLabel,
      totalAmount: toNumber(row[1]),
      orders: toNumber(row[2]),
      avgPerOrder: toNumber(row[3])
    };
  });
}

// 讀取「WeeklyStats」並轉成週統計圖表資料。
function getWeeklyStatsData() {
  return readStatsSheet(WEEKLY_STATS_SHEET, row => {
    const weekKey = formatWeekLabel(row[0]);
    if (!weekKey) return null;
    return {
      weekKey,
      totalAmount: toNumber(row[1]),
      month: formatMonthLabel(row[2])
    };
  });
}

// 泛用讀表函式：逐列套用 mapFn，忽略空資料。
function readStatsSheet(sheetName, mapFn) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const rows = values.slice(1);
  const mapped = [];
  rows.forEach(row => {
    const item = mapFn(row);
    if (item) mapped.push(item);
  });
  return mapped;
}

// 將日期或文字轉成 yyyy-MM 月份標籤。
function formatMonthLabel(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, DASHBOARD_TIMEZONE, 'yyyy-MM');
  }
  const text = value != null ? value.toString().trim() : '';
  if (!text) return '';
  return text;
}

// 產生週標籤（yyyy-MM-第幾週），或回傳原始文字。
function formatWeekLabel(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, DASHBOARD_TIMEZONE, 'yyyy-MM-') + getWeekOfMonth(value);
  }
  const text = value != null ? value.toString().trim() : '';
  if (!text) return '';
  return text;
}

// 計算指定日期在該月的第幾週。
function getWeekOfMonth(dateObj) {
  const first = new Date(dateObj.getFullYear(), dateObj.getMonth(), 1);
  const dayOfWeek = first.getDay() || 7;
  const offset = dateObj.getDate() + dayOfWeek - 1;
  return Math.ceil(offset / 7);
}

// 將不同型別的值安全地轉成數字。
function toNumber(value) {
  if (typeof value === 'number') {
    return isNaN(value) ? 0 : value;
  }
  if (typeof value === 'string') {
    const num = Number(value.replace(/,/g, ''));
    return isNaN(num) ? 0 : num;
  }
  if (value instanceof Date) {
    return 0;
  }
  return Number(value) || 0;
}

// 依目前時間與時區推算本週（一～日）的起訖毫秒。
function getCurrentWeekRange(now, tz) {
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const localDate = new Date(todayStr);
  const dayOfWeek = Number(Utilities.formatDate(now, tz, 'u')) || 1; // Monday = 1
  const start = new Date(localDate);
  start.setUTCDate(start.getUTCDate() - (dayOfWeek - 1));
  start.setUTCHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setUTCDate(end.getUTCDate() + 7);
  return {
    startMs: start.getTime(),
    endMs: end.getTime()
  };
}
