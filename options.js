const url = document.querySelector('#targetUrl');
const tab = document.querySelector('#targetTab');
const status = document.querySelector('#status');

chrome.storage.sync.get({ targetUrl: '', targetTab: '' }).then(values => {
  url.value = values.targetUrl;
  tab.value = values.targetTab;
});

document.querySelector('#save').addEventListener('click', async () => {
  if (!url.value.startsWith('https://docs.google.com/spreadsheets/')) {
    status.textContent = '请输入有效的 Google 表格网址';
    status.style.color = '#c5221f';
    return;
  }
  await chrome.storage.sync.set({ targetUrl: url.value, targetTab: tab.value.trim() });
  status.textContent = '已保存';
  status.style.color = '#188038';
});
