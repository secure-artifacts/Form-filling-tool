const dashboardUrl = () => chrome.runtime.getURL('dashboard.html');

// 已有控制台标签页就复用它（带上新参数导航并置前），没有才新建——
// 反复深度查询不会堆积一堆控制台标签。
const openDashboard = async (query = '') => {
  const target = `${dashboardUrl()}${query}`;
  const tabs = await chrome.tabs.query({ url: `${dashboardUrl()}*` }).catch(() => []);
  if (tabs.length) {
    const tab = tabs[0];
    await chrome.tabs.update(tab.id, query ? { url: target, active: true } : { active: true });
    if (tab.windowId) await chrome.windows.update(tab.windowId, { focused: true }).catch(() => {});
    return;
  }
  await chrome.tabs.create({ url: target });
};

chrome.action.onClicked.addListener(() => openDashboard());

const nextNoon = () => {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  if (date.getTime() <= Date.now()) date.setDate(date.getDate() + 1);
  return date.getTime();
};
const scheduleRegionSync = () => chrome.alarms.create('region-config-noon', { when: nextNoon() });
const spreadsheetId = url => url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/)?.[1];

const rowsHash = rows => {
  const text = JSON.stringify(rows || []);
  let hash = 0;
  for (let index = 0; index < text.length; index++) hash = (hash * 31 + text.charCodeAt(index)) | 0;
  return `${text.length}:${hash}`;
};

async function syncRegionConfigAtNoon() {
  const { targetUrl } = await chrome.storage.sync.get({ targetUrl: '' });
  const { regionTab = '' } = await chrome.storage.local.get({ regionTab: '' });
  const id = spreadsheetId(targetUrl);
  if (!id || !regionTab) return;
  let token = '';
  try {
    token = await chrome.identity.getAuthToken({ interactive: false });
    token = typeof token === 'string' ? token : token?.token || '';
  } catch {
    // Non-interactive auth fails when nobody is signed in; wait for the next check.
    return;
  }
  if (!token) return;
  const sheet = `'${regionTab.replaceAll("'", "''")}'`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(`${sheet}!A:C`)}?majorDimension=ROWS&valueRenderOption=UNFORMATTED_VALUE`;
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) return;
  const rows = (await response.json()).values || [];
  const { regionConfigCache } = await chrome.storage.local.get({ regionConfigCache: null });
  // Compare content, not row count, so edited/deleted config rows propagate.
  // An empty read never wipes a working cache.
  if (regionConfigCache?.rows?.length && !rows.length) {
    await chrome.storage.local.set({ regionConfigLastCheckedAt: Date.now(), regionConfigLastCheckRows: 0 });
    return;
  }
  if (!regionConfigCache || rowsHash(rows) !== regionConfigCache.contentHash) {
    await chrome.storage.local.set({ regionConfigCache: { rows, rowCount: rows.length, syncedAt: Date.now(), contentHash: rowsHash(rows) } });
  }
  await chrome.storage.local.set({ regionConfigLastCheckedAt: Date.now(), regionConfigLastCheckRows: rows.length });
}

chrome.runtime.onInstalled.addListener(scheduleRegionSync);
chrome.runtime.onStartup.addListener(scheduleRegionSync);
chrome.alarms.onAlarm.addListener(async alarm => {
  if (alarm.name !== 'region-config-noon') return;
  try { await syncRegionConfigAtNoon(); }
  catch { /* A failed noon check simply waits for the next one. */ }
  finally { scheduleRegionSync(); }
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'OPEN_DASHBOARD') {
    openDashboard(message.query || '');
    sendResponse({ ok: true });
  }
});
