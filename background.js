const dashboardUrl = () => chrome.runtime.getURL('dashboard.html');

const openDashboard = query => chrome.tabs.create({
  url: `${dashboardUrl()}${query || ''}`
});

chrome.action.onClicked.addListener(() => openDashboard());

const nextNoon = () => {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  if (date.getTime() <= Date.now()) date.setDate(date.getDate() + 1);
  return date.getTime();
};
const scheduleRegionSync = () => chrome.alarms.create('region-config-noon', { when: nextNoon() });
const normalize = value => String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
const spreadsheetId = url => url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/)?.[1];

async function syncRegionConfigAtNoon() {
  const { targetUrl } = await chrome.storage.sync.get({ targetUrl: '' });
  const { regionTab = '' } = await chrome.storage.local.get({ regionTab: '' });
  const id = spreadsheetId(targetUrl);
  const token = await chrome.identity.getAuthToken({ interactive: false });
  if (!id || !token || !regionTab) return;
  const sheet = `'${regionTab.replaceAll("'", "''")}'`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(`${sheet}!A:C`)}?majorDimension=ROWS&valueRenderOption=UNFORMATTED_VALUE`;
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) return;
  const rows = (await response.json()).values || [];
  const { regionConfigCache } = await chrome.storage.local.get({ regionConfigCache: null });
  if (!regionConfigCache || rows.length > regionConfigCache.rowCount) {
    await chrome.storage.local.set({ regionConfigCache: { rows, rowCount: rows.length, syncedAt: Date.now() } });
  }
  await chrome.storage.local.set({ regionConfigLastCheckedAt: Date.now(), regionConfigLastCheckRows: rows.length });
}

chrome.runtime.onInstalled.addListener(scheduleRegionSync);
chrome.runtime.onStartup.addListener(scheduleRegionSync);
chrome.alarms.onAlarm.addListener(async alarm => {
  if (alarm.name !== 'region-config-noon') return;
  try { await syncRegionConfigAtNoon(); } finally { scheduleRegionSync(); }
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'OPEN_DASHBOARD') {
    openDashboard(message.query || '');
    sendResponse({ ok: true });
  }
});
