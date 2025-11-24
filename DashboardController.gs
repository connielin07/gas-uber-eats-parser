const DASHBOARD_TIMEZONE = 'Asia/Taipei';

function doGet() {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Uber Eats Dashboard');
}

function getDashboardData() {
  const tz = DASHBOARD_TIMEZONE;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const emptySummary = {
    totalOrders: 0,
    totalAmount: 0,
    updatedOrders: 0,
    errorOrders: 0,
    thisMonthAmount: 0,
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

  let totalAmount = 0;
  let totalOrders = 0;
  let updatedOrders = 0;
  let errorOrders = 0;
  let thisMonthAmount = 0;

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

  return {
    summary: {
      totalOrders,
      totalAmount,
      updatedOrders,
      errorOrders,
      thisMonthAmount,
      lastRefreshed: Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm')
    },
    recent,
    statusMessage: `共 ${totalOrders} 筆紀錄`
  };
}

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
    return { display: '-', timestamp: 0, monthKey: '' };
  }

  return {
    display: Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd'),
    timestamp: dateObj.getTime(),
    monthKey: Utilities.formatDate(dateObj, tz, 'yyyy-MM')
  };
}
