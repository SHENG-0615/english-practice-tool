(function () {
  'use strict';
  const config = window.__IELTS_LOCAL_CONFIG__;
  const canonical = path => String(path || '').replace(/\//g, '\\').replace(/\\$/, '').toLowerCase();
  const currentFile = location.protocol === 'file:' ? decodeURIComponent(location.pathname).replace(/^\/(?=[a-z]:)/i, '') : '';
  const valid = config && /^http:\/\/127\.0\.0\.1:\d+$/.test(config.url) && canonical(config.pagePath) === canonical(currentFile);
  let available = false;
  let connectionError = '';
  async function request(route, options = {}) {
    let response;
    try {
      response = await fetch(config.url + route, {
        ...options, cache: 'no-store', signal: AbortSignal.timeout(30000),
        headers: { 'X-IELTS-Token': config.token, ...options.headers }
      });
    } catch (_) { throw new Error('本机连接已中断。请重新双击“开始练习.cmd”，再刷新页面。'); }
    if (!response.ok) {
      let message = '本机工作簿暂时无法写入。';
      try { message = (await response.json()).error || message; } catch (_) {}
      if (response.status === 409) throw new DOMException(message, 'InvalidModificationError');
      if (response.status === 423) throw new DOMException(message, 'NoModificationAllowedError');
      throw new Error(message);
    }
    return response;
  }
  const ready = !valid ? Promise.resolve(false) : request('/api/ping').then(async response => {
    const info = await response.json();
    available = info.app === 'english-practice-local-v1' && canonical(info.root) === canonical(config.root);
    if (available) config.files = info.files;
    return available;
  }).catch(() => { connectionError = '本机连接未启动或已过期。请重新双击同文件夹里的“开始练习.cmd”，再刷新此页面；待同步作答会保留在本机。'; return false; });
  function handleFor(key) {
    const file = config?.files?.[key];
    if (!available || !file) return null;
    let etag = '';
    return {
      kind: 'file', provider: 'local', localKey: key, location: file.path, name: file.name,
      async queryPermission() { return 'granted'; }, async requestPermission() { return 'granted'; },
      async isSameEntry(other) { return other?.provider === 'local' && canonical(other.location) === canonical(file.path); },
      async getFile() {
        const response = await request('/api/workbooks/' + key);
        etag = response.headers.get('ETag');
        return new File([await response.arrayBuffer()], file.name, {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          lastModified: Number(response.headers.get('X-Modified-Time'))
        });
      },
      async createWritable() {
        const expected = etag;
        let data, aborted = false;
        return {
          async write(value) { if (aborted) throw new Error('写入已取消。'); data = value; },
          async close() {
            if (aborted || !data || !expected) throw new Error('工作簿写入未准备完成。');
            await request('/api/workbooks/' + key, { method: 'PUT', headers: { 'If-Match': expected, 'Content-Type': 'application/octet-stream' }, body: data });
          },
          async abort() { aborted = true; data = null; }
        };
      }
    };
  }
  function showLocation(id, state) {
    const node = document.getElementById(id);
    if (!node) return;
    const text = state.location || (state.connected ? state.connected + '（浏览器未提供完整地址）' : '尚未连接工作簿');
    node.querySelector('[data-file-location]').textContent = text;
    node.querySelector('[data-file-location]').title = text;
    node.querySelector('[data-file-kind]').textContent = state.connectionKind === 'local' ? '本机自动接入' : state.connected ? '浏览器已连接' : '工作簿位置';
    const copy = node.querySelector('[data-copy-location]');
    copy.hidden = !state.location;
    copy.onclick = async () => {
      try { await navigator.clipboard.writeText(state.location); copy.textContent = '已复制'; }
      catch (_) { copy.textContent = '请选中地址复制'; }
      setTimeout(() => { copy.textContent = '复制地址'; }, 1800);
    };
  }
  window.DictationLocalFiles = { ready, handleFor, showLocation, status: () => ({ configured: !!valid, available, error: connectionError }) };
})();
