(function () {
  'use strict';
  const DATABASE = 'dictation-xlsx-sync-v1';
  const JOURNAL = 'dictation-xlsx-pending:';
  const ACCESS = { mode: 'readwrite' };
  const WRITEBACK_VERSION = 3;
  const retryable = new Set(['AbortError', 'InvalidModificationError', 'InvalidStateError', 'NoModificationAllowedError', 'NotReadableError']);
  const request = operation => new Promise((resolve, reject) => {
    operation.onsuccess = () => resolve(operation.result);
    operation.onerror = () => reject(operation.error);
  });
  const committed = transaction => new Promise((resolve, reject) => {
    transaction.oncomplete = resolve;
    transaction.onabort = () => reject(transaction.error || new Error('本机记录写入中断。'));
    transaction.onerror = () => reject(transaction.error || new Error('本机记录写入失败。'));
  });

  async function openDatabase(name) {
    const opening = indexedDB.open(name, 1);
    opening.onupgradeneeded = () => {
      opening.result.createObjectStore('records', { keyPath: 'id' });
      opening.result.createObjectStore('meta');
    };
    return request(opening);
  }
  function pendingJournal(prefix) {
    const rows = [];
    try {
      for (let index = 0; index < localStorage.length; index++) {
        const key = localStorage.key(index);
        if (key?.startsWith(prefix)) {
          const row = JSON.parse(localStorage.getItem(key));
          if (row?.id) rows.push(row);
        }
      }
    } catch (_) { /* The durable browser database is also used. */ }
    return rows;
  }
  function removeJournal(row, prefix) {
    try { localStorage.removeItem(prefix + row.id); } catch (_) { /* Records also have a stable database key. */ }
  }

  function create(onChange, options = {}) {
    const database = options.database || DATABASE;
    const journalPrefix = options.journal || JOURNAL;
    const writebackVersion = options.version || WRITEBACK_VERSION;
    const localKey = options.localKey || 'dictation';
    const appendWorkbook = options.append || ((bytes, rows) => window.DictationXlsxLog.append(bytes, rows));
    let supported = typeof window.showOpenFilePicker === 'function';
    let localHandle = null;
    let connectionKey = 'connection';
    let db, starting, loaded = false, connection = null, writable = false, busy = false, count = 0, error = '', timer = null;

    function notify() {
      onChange({
        supported, loaded, busy, connected: connection?.handle?.name || '', writable, error,
        location: connection?.handle?.location || '', connectionKind: connection?.handle?.provider || 'browser',
        localAvailable: !!localHandle,
        writebackReceipt: connection?.receipt || null,
        chapterBackfill: !!connection && connection.writeback !== writebackVersion,
        pending: Math.max(0, count - (connection?.written || 0)), synced: connection?.written || 0
      });
    }
    function from(store, method, ...args) {
      return request(db.transaction(store).objectStore(store)[method](...args));
    }
    function save(store, value, key) {
      const tx = db.transaction(store, 'readwrite');
      const done = committed(tx);
      const table = tx.objectStore(store);
      if (key === undefined) table.put(value);
      else table.put(value, key);
      return done;
    }
    function storedConnection(value) {
      if (value.handle?.provider !== 'local') return value;
      const { handle, ...meta } = value;
      return { ...meta, kind: 'local', localKey: handle.localKey, location: handle.location };
    }
    function schedule(delay = 8000) {
      if (!supported || !writable || !connection || busy || timer || (count <= connection.written && connection.writeback === writebackVersion)) return;
      timer = setTimeout(() => { timer = null; void sync(false); }, delay);
    }
    function clearTimer() { clearTimeout(timer); timer = null; }
    async function initialize(history) {
      db = await openDatabase(database);
      const journal = pendingJournal(journalPrefix);
      const tx = db.transaction(['meta', 'records'], 'readwrite');
      const done = committed(tx);
      const meta = tx.objectStore('meta'), records = tx.objectStore('records');
      const initialized = meta.get('initialized');
      initialized.onsuccess = () => {
        if (!initialized.result) {
          for (const row of history) records.put(row);
          meta.put(true, 'initialized');
        }
        for (const row of journal) records.put(row);
      };
      await done;
      journal.forEach(row => removeJournal(row, journalPrefix));
      let stored = await from('meta', 'get', connectionKey);
      count = await from('records', 'count');
      await window.DictationLocalFiles?.ready;
      localHandle = options.disableLocal ? null : window.DictationLocalFiles?.handleFor(localKey);
      if (localHandle) {
        connectionKey = 'local-connection:' + localHandle.location;
        stored = await from('meta', 'get', connectionKey);
      }
      if (localHandle && (!stored?.preferManual || stored.kind === 'local')) {
        supported = true;
        const same = stored?.kind === 'local' && stored.location === localHandle.location;
        connection = { handle: localHandle, written: same ? stored.written : 0, writeback: same ? stored.writeback : 0, ...(same && stored.receipt ? { receipt: stored.receipt } : {}) };
        await save('meta', storedConnection(connection), connectionKey);
      } else connection = stored?.kind === 'local' ? null : stored;
      if (!connection && window.DictationLocalFiles?.status?.().error) error = window.DictationLocalFiles.status().error;
      loaded = true;
      writable = !!(supported && connection && await connection.handle.queryPermission(ACCESS) === 'granted');
      notify();
      schedule();
    }
    function start(history) {
      starting = initialize(history).catch(() => {
        error = '暂时无法读取本机同步记录。重新打开页面后可以再试。';
        notify();
      });
    }
    function record(row) {
      try { localStorage.setItem(journalPrefix + row.id, JSON.stringify(row)); }
      catch (_) { /* Continue saving to the browser database. */ }
      void starting.then(async () => {
        if (!loaded) return;
        await save('records', row);
        removeJournal(row, journalPrefix);
        count = await from('records', 'count');
        notify();
        schedule();
      }).catch(() => {
        error = '本次作答的同步存档暂时无法保存。重新打开页面后可以再试。';
        notify();
      });
    }
    async function writeFile(handle, binding = false) {
      const work = async () => {
        if (!binding) {
          const bound = await from('meta', 'get', connectionKey);
          const boundHandle = bound?.kind === 'local' ? window.DictationLocalFiles?.handleFor(bound.localKey) : bound?.handle;
          if (boundHandle && !await handle.isSameEntry(boundHandle)) {
            connection = { ...bound, handle: boundHandle };
            writable = await boundHandle.queryPermission(ACCESS) === 'granted';
            return;
          }
        }
        const snapshot = await handle.getFile();
        const rows = await from('records', 'getAll');
        let receipt = null;
        const updated = await appendWorkbook(await snapshot.arrayBuffer(), rows, value => { receipt = value; });
        const verifyUnchanged = async () => {
          const current = await handle.getFile();
          if (current.lastModified !== snapshot.lastModified || current.size !== snapshot.size) {
            throw new DOMException('表格在同步期间有更新。请先关闭 WPS 或 Excel，再同步。', 'InvalidModificationError');
          }
        };
        if (updated) {
          const file = await handle.createWritable({ mode: 'exclusive' });
          try {
            await verifyUnchanged();
            await file.write(updated);
            await verifyUnchanged();
            await file.close();
          } catch (cause) {
            await file.abort().catch(() => {});
            throw cause;
          }
        } else {
          await verifyUnchanged();
        }
        const latest = { handle, written: rows.length, writeback: writebackVersion, preferManual: !!localHandle && handle.provider !== 'local', ...(receipt ? { receipt: { ...receipt, confirmedAt: new Date().toISOString() } } : {}) };
        await save('meta', storedConnection(latest), connectionKey);
        connection = latest;
        count = await from('records', 'count');
        writable = true;
      };
      if (navigator.locks) await navigator.locks.request(database, work);
      else await work();
    }
    function explain(cause) {
      if (cause.name === 'NotAllowedError' || cause.name === 'SecurityError') return '需要浏览器授权写入。请点击“立即同步”重新授权。';
      if (cause.name === 'NotFoundError') return '找不到已连接的工作簿。请重新选择 XLSX。';
      if (cause instanceof DOMException) return '暂时无法写入。请在 WPS 或 Excel 中关闭这份表，再点“立即同步”。';
      return cause.message || '同步没有完成，作答已保存在本机。';
    }
    async function sync(manual = true) {
      if (!loaded || !supported || !connection || busy) return;
      clearTimer();
      busy = true; error = ''; notify();
      let retryDelay = 8000, tryAgain = true;
      try {
        const handle = connection.handle;
        const allowed = manual ? await handle.requestPermission(ACCESS) : await handle.queryPermission(ACCESS);
        writable = allowed === 'granted';
        if (writable) await writeFile(handle);
      } catch (cause) {
        error = explain(cause);
        tryAgain = retryable.has(cause.name);
        retryDelay = 15000;
        if (!tryAgain) writable = false;
      } finally {
        busy = false; notify();
        if (tryAgain) schedule(retryDelay);
      }
    }
    async function pick() {
      if (!loaded || !supported || busy) return;
      clearTimer();
      busy = true; error = ''; notify();
      let selected = false;
      try {
        const [handle] = await window.showOpenFilePicker({
          multiple: false,
          types: [{ description: 'Excel 工作簿', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
        });
        selected = true;
        if (await handle.requestPermission(ACCESS) === 'granted') await writeFile(handle, true);
        else error = '尚未授权写入。作答已保存在本机，可以重新选表。';
      } catch (cause) {
        if (selected || cause.name !== 'AbortError') error = explain(cause);
      } finally {
        busy = false; notify(); schedule();
      }
    }
    async function connectLocal() {
      if (!loaded || !localHandle || busy) return;
      clearTimer(); busy = true; error = ''; notify();
      try { await writeFile(localHandle, true); }
      catch (cause) { error = explain(cause); }
      finally { busy = false; notify(); schedule(); }
    }
    async function getRecords() { await starting; return loaded ? from('records', 'getAll') : []; }
    return { start, record, pick, sync, connectLocal, getRecords };
  }

  window.DictationXlsxSync = { create };
})();
