(() => {
  const MENU_LABEL = '转交表格';
  const marker = 'data-form-transfer-item';

  const notify = (message, error = false) => {
    const node = document.createElement('div');
    node.textContent = message;
    Object.assign(node.style, {
      position: 'fixed', zIndex: 2147483647, top: '20px', left: '50%',
      transform: 'translateX(-50%)', padding: '12px 18px', borderRadius: '6px',
      color: '#fff', background: error ? '#c5221f' : '#188038',
      font: '14px Arial', boxShadow: '0 2px 8px #0005'
    });
    document.body.append(node);
    setTimeout(() => node.remove(), 3500);
  };

  const readSelection = async () => {
    // Sheets often puts a richer HTML table on the clipboard than in plain
    // text when complete rows are copied. Prefer that table so blank columns
    // and the B:BL positions are preserved.
    let text = '';
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes('text/html')) {
          const html = await (await item.getType('text/html')).text();
          const doc = new DOMParser().parseFromString(html, 'text/html');
          const cellText = cell => {
            const holder = doc.createElement('div');
            holder.innerHTML = cell.innerHTML
              .replace(/<br\s*\/?>/gi, '\n')
              .replace(/<\/(?:div|p|li)>/gi, '\n');
            return (holder.textContent || '').replace(/\r\n/g, '\n').replace(/\t/g, ' ').replace(/\n/g, '\uE000').trim();
          };
          const rows = [...doc.querySelectorAll('tr')].map(row => [...row.querySelectorAll('th,td')]
            .map(cellText).join('\t'));
          if (rows.length && rows.some(row => row.includes('\t'))) text = rows.join('\n');
        }
        if (!text && item.types.includes('text/plain')) text = await (await item.getType('text/plain')).text();
      }
    } catch {
      text = await navigator.clipboard.readText();
    }
    if (!text.trim()) throw new Error('没有读到选区内容，请先选择 B 列到 BL 列的行。');
    return text;
  };

  const transfer = async () => {
    try {
      const text = await readSelection();
      await chrome.storage.local.set({ formTransferSource: { text, createdAt: Date.now() } });
      chrome.runtime.sendMessage({ type: 'OPEN_DASHBOARD', query: '?flow=transfer' });
      notify('已把选区交给自动化控制台。');
    } catch (error) {
      notify(error.message || '转交失败。', true);
    }
  };

  const finishPendingTransfer = async pending => {
    if (!location.href.startsWith(pending.targetUrl)) return;
    if (Date.now() - pending.createdAt > 10 * 60 * 1000) {
      await chrome.storage.local.remove('formTransferPending');
      return;
    }
    if (pending.targetTab) {
      const tab = [...document.querySelectorAll('[role="tab"]')]
        .find(node => node.textContent.trim() === pending.targetTab);
      if (tab) tab.click();
    }
    await new Promise(resolve => setTimeout(resolve, 1800));
    await navigator.clipboard.writeText(pending.text);
    const panel = document.createElement('div');
    panel.innerHTML = '<b>转交数据已准备好</b><br><span>请在目标表点击首个空行，然后按 Ctrl+V 粘贴。</span>';
    const copy = document.createElement('button');
    copy.textContent = '重新复制数据';
    const done = document.createElement('button');
    done.textContent = '完成';
    [copy, done].forEach(button => Object.assign(button.style, {
      margin: '10px 8px 0 0', padding: '6px 10px', cursor: 'pointer'
    }));
    copy.onclick = async () => {
      await navigator.clipboard.writeText(pending.text);
      notify('数据已复制，请点击目标首个空行后按 Ctrl+V。');
    };
    done.onclick = async () => {
      await chrome.storage.local.remove('formTransferPending');
      panel.remove();
    };
    panel.append(copy, done);
    Object.assign(panel.style, {
      position: 'fixed', zIndex: 2147483647, top: '16px', left: '50%',
      transform: 'translateX(-50%)', width: '360px', padding: '14px 16px',
      borderRadius: '8px', color: '#202124', background: '#fff',
      font: '14px Arial', lineHeight: '1.55', boxShadow: '0 2px 12px #0005'
    });
    document.body.append(panel);
  };

  chrome.storage.local.get('formTransferPending').then(({ formTransferPending }) => {
    if (formTransferPending) finishPendingTransfer(formTransferPending).catch(error => {
      notify(error.message || '目标表粘贴失败，请手动选择首个空行后按 Ctrl+V。', true);
    });
  });

  const addMenuItem = menu => {
    const existing = menu.querySelector(`[${marker}]`);
    if (existing) {
      if (menu.firstElementChild !== existing) menu.prepend(existing);
      return;
    }
    const item = document.createElement('div');
    item.setAttribute('role', 'menuitem');
    item.setAttribute(marker, 'true');
    item.textContent = MENU_LABEL;
    Object.assign(item.style, {
      cursor: 'pointer', padding: '8px 16px', color: '#202124',
      font: '14px Arial', borderBottom: '1px solid #dadce0'
    });
    item.addEventListener('mouseenter', () => item.style.background = '#f1f3f4');
    item.addEventListener('mouseleave', () => item.style.background = '');
    item.addEventListener('click', event => { event.stopPropagation(); transfer(); });
    menu.prepend(item);
  };

  const scanMenus = () => {
    document.querySelectorAll('[role="menu"]').forEach(menu => {
      if (menu.offsetParent !== null) addMenuItem(menu);
    });
  };

  new MutationObserver(scanMenus).observe(document.documentElement, { childList: true, subtree: true });
  const scanAfterContextMenu = () => {
    let attempts = 0;
    const poll = () => {
      scanMenus();
      if (++attempts < 8) setTimeout(poll, 25);
    };
    setTimeout(poll, 0);
  };
  document.addEventListener('contextmenu', scanAfterContextMenu, true);
  document.addEventListener('mousedown', event => { if (event.button === 2) scanAfterContextMenu(); }, true);
})();
