/* Original application logic. Spreadsheet decoding is provided separately by SheetJS CE. */
(function (root) {
  'use strict';
  const text = value => String(value ?? '').trim();
  const clean = value => text(value).replace(/\s+/g, ' ');
  const DEFAULTS = Object.freeze({ accent: 'en-GB', voiceURI: '', rate: 1, repeat: 1, gap: 3, mode: 'review', display: 'both' });
  const collator = new Intl.Collator('zh-CN', { numeric: true, sensitivity: 'base' });
  const listeningBook = '兼容听写表';
  // Keep saved chapter selections and word IDs valid after correcting the source label.
  const sourceKey = source => source === listeningBook ? '剑20' : source;
  const groupKey = item => JSON.stringify([sourceKey(item.source), item.chapter]);
  const wordKey = item => JSON.stringify([sourceKey(clean(item.source)), clean(item.chapter), clean(item.word).normalize('NFKC').toLowerCase()]);
  const hash = s => { let h = 2166136261; for (let i = 0; i < s.length; i++) h = Math.imul(h ^ s.charCodeAt(i), 16777619); return (h >>> 0).toString(36); };
  const validWord = value => { const s = clean(value); return s.length > 0 && s.length <= 300 && /[a-z]/i.test(s) && !/[\u3400-\u9fff]/.test(s) && !/^[=#]/.test(s); };
  function normalizeEntry(item, fallback = {}) {
    if (!item || typeof item !== 'object' || !validWord(item.word)) return null;
    const source = clean(item.source || fallback.source || '自建词库').slice(0, 120);
    const result = {
      word: clean(item.word), meaning: text(item.meaning).slice(0, 1500),
      chapter: clean(item.chapter || fallback.chapter || '未分类').slice(0, 100),
      source: source === '剑20' ? listeningBook : source,
      origin: clean(item.origin || fallback.origin || '自行添加').slice(0, 120)
    };
    result.id = 'w-' + hash(wordKey(result));
    if (Array.isArray(item.alternatives)) {
      const alternatives = [...new Set(item.alternatives.slice(0, 5).filter(validWord).map(clean))].filter(word => word !== result.word);
      if (alternatives.length) result.alternatives = alternatives;
    }
    return result;
  }
  function mergeEntries(existing, incoming) {
    const entries = [], index = new Map(); let added = 0, updated = 0;
    const append = (raw, isNew) => {
      const item = normalizeEntry(raw); if (!item) return;
      const key = wordKey(item);
      if (index.has(key)) {
        const target = entries[index.get(key)];
        if (item.meaning && target.meaning !== item.meaning) { target.meaning = item.meaning; if (isNew) updated++; }
        if (item.alternatives) target.alternatives = [...new Set([...(target.alternatives || []), ...item.alternatives])].slice(0, 5);
        return;
      }
      index.set(key, entries.length); entries.push(item); if (isNew) added++;
    };
    (Array.isArray(existing) ? existing : []).forEach(item => append(item, false));
    (Array.isArray(incoming) ? incoming : []).forEach(item => append(item, true));
    if (entries.length > 50000) throw new Error('词库超过 50,000 条，请分开保存。');
    return { entries, added, updated };
  }
  function groupsOf(entries) {
    const groups = new Map();
    for (const item of entries) {
      const key = groupKey(item);
      if (!groups.has(key)) groups.set(key, { key, source: item.source, chapter: item.chapter, count: 0 });
      groups.get(key).count++;
    }
    return [...groups.values()].sort((a, b) => collator.compare(a.source, b.source) || collator.compare(a.chapter, b.chapter));
  }
  function orderEntries(entries) {
    return [...entries].sort((a, b) => collator.compare(a.source, b.source) || collator.compare(a.chapter, b.chapter));
  }
  function safeSettings(raw = {}) {
    if (!raw || typeof raw !== 'object') raw = {};
    const number = (key, low, high) => raw[key] != null && raw[key] !== '' && Number.isFinite(Number(raw[key])) ? Math.min(high, Math.max(low, Number(raw[key]))) : DEFAULTS[key];
    return {
      accent: ['en-GB', 'en-US'].includes(raw.accent) ? raw.accent : DEFAULTS.accent,
      voiceURI: text(raw.voiceURI).slice(0, 500), rate: number('rate', 0.5, 1.5),
      repeat: Math.round(number('repeat', 1, 5)), gap: number('gap', 0, 30),
      mode: ['practice', 'dictation'].includes(raw.mode) ? raw.mode : 'review',
      display: ['both', 'english', 'chinese'].includes(raw.display) ? raw.display : 'both'
    };
  }
  function parseDelimited(input, delimiter) {
    const rows = []; let row = [], value = '', quoted = false;
    const data = String(input).replace(/^\uFEFF/, '');
    for (let i = 0; i <= data.length; i++) {
      const c = i === data.length ? '\n' : data[i];
      if (c === '"' && (quoted || !value)) {
        if (quoted && data[i + 1] === '"') { value += '"'; i++; } else quoted = !quoted;
      } else if (!quoted && c === delimiter) { row.push(value); value = ''; }
      else if (!quoted && (c === '\r' || c === '\n')) {
        if (c === '\r' && data[i + 1] === '\n') i++;
        row.push(value); value = '';
        if (row.some(cell => text(cell))) rows.push(row);
        row = [];
      } else value += c;
    }
    if (quoted) throw new Error('文本中有未闭合的英文双引号，请检查 CSV 格式。');
    return rows;
  }
  const aliases = {
    word: /^(英文|英语|单词|词汇|英文单词|正确答案|正确单词|word|words|english|vocabulary)$/i,
    meaning: /^(中文|中文释义|释义|意思|翻译|meaning|meanings|chinese|translation)$/i,
    chapter: /^(章节|章节号|分类|分组|chapter|section|group)$/i,
    source: /^(来源|词库|书名|source|book|library)$/i
  };
  function rowsToEntries(rows, fallback = {}) {
    const firstRows = rows.slice(0, 8); let headerIndex = -1, columns = { word: 0, meaning: 1, chapter: 2, source: -1 };
    for (let r = 0; r < firstRows.length; r++) {
      const idx = firstRows[r].findIndex(v => aliases.word.test(clean(v)));
      const hasCompanionHeader = firstRows[r].some(v => aliases.meaning.test(clean(v)) || aliases.chapter.test(clean(v)) || aliases.source.test(clean(v)));
      if (idx >= 0 && (/[\u3400-\u9fff]/.test(clean(firstRows[r][idx])) || hasCompanionHeader)) {
        headerIndex = r;
        columns = { word: idx, meaning: -1, chapter: -1, source: -1 };
        for (const key of ['meaning', 'chapter', 'source']) columns[key] = firstRows[r].findIndex(v => aliases[key].test(clean(v)));
        break;
      }
    }
    const result = [];
    for (const row of rows.slice(headerIndex + 1)) {
      const entry = normalizeEntry({ word: row[columns.word], meaning: columns.meaning < 0 ? '' : row[columns.meaning], chapter: columns.chapter < 0 ? '' : row[columns.chapter], source: columns.source < 0 ? '' : row[columns.source], origin: fallback.origin }, fallback);
      if (entry) result.push(entry);
    }
    return result;
  }
  function parseText(input, fallback = {}) {
    const s = String(input).trim(); if (!s) return [];
    const delimiter = s.includes('\t') ? '\t' : s.includes('|') ? '|' : /[,，]/.test(s.split(/\r?\n/)[0]) ? (s.split(/\r?\n/)[0].includes(',') ? ',' : '，') : null;
    let rows;
    if (delimiter) rows = parseDelimited(s, delimiter);
    else rows = s.replace(/^\uFEFF/, '').split(/\r?\n/).filter(line => line.trim()).map(line => {
      const p = line.search(/[\u3400-\u9fff]/);
      return p > 0 ? [line.slice(0, p).replace(/\s*[-–—:]\s*$/, '').trim(), line.slice(p).trim()] : [line.trim()];
    });
    return rowsToEntries(rows, fallback);
  }
  function parseWorkbook(workbook, XLSX, fallback = {}) {
    const value = (sheet, row, col) => {
      const c = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
      return c && c.t !== 'e' ? text(c.v) : '';
    };
    const result = [], notes = []; let specialized = false, formulaWithoutCache = 0;
    const bookMatch = /剑\s*(\d+)/.exec(fallback.filename || '');
    const bookSource = bookMatch ? '剑' + bookMatch[1] : fallback.source || '导入词库';
    if (workbook.Sheets['错词总表']) {
      specialized = true; const sheet = workbook.Sheets['错词总表'];
      if (sheet['!ref']) {
        const range = XLSX.utils.decode_range(sheet['!ref']);
        let headerRow = range.s.r, mistakeCols = [];
        for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 7); r++) {
          const found = [];
          for (let c = range.s.c; c <= range.e.c; c++) if (/^错词\s*\d+$/.test(value(sheet, r, c))) found.push(c);
          if (found.length) { headerRow = r; mistakeCols = found; break; }
        }
        if (!mistakeCols.length) throw new Error('找到“错词总表”，但未找到“错词1、错词2…”列。请检查表头。');
        for (let r = headerRow + 1; r <= range.e.r; r++) {
          for (const c of mistakeCols) {
            const cell = sheet[XLSX.utils.encode_cell({ r, c })];
            if (cell?.f && (cell.v === undefined || cell.t === 'e')) formulaWithoutCache++;
            const entry = normalizeEntry({ word: value(sheet, r, c), meaning: value(sheet, r, 2), chapter: value(sheet, r, 3), source: bookSource, origin: '错词总表' }, fallback);
            if (entry) result.push(entry);
          }
        }
        notes.push('已读取“错词总表”保存的错词结果，保留完整章节号。');
      }
    }
    if (workbook.Sheets['错词练习']) {
      specialized = true; const sheet = workbook.Sheets['错词练习'];
      if (sheet['!ref']) {
        const range = XLSX.utils.decode_range(sheet['!ref']);
        for (let r = 1; r <= range.e.r; r++) {
          const entry = normalizeEntry({ word: value(sheet, r, 0), meaning: value(sheet, r, 1), chapter: value(sheet, r, 2), source: bookSource, origin: '错词练习' }, fallback);
          if (entry) result.push(entry);
        }
        notes.push('已同时读取“错词练习”，按它自己的章节归类。');
      }
    }
    if (!specialized) {
      for (const name of workbook.SheetNames) {
        const sheet = workbook.Sheets[name]; if (!sheet['!ref']) continue;
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: false });
        result.push(...rowsToEntries(rows, { ...fallback, chapter: fallback.chapter && fallback.chapter !== '未分类' ? fallback.chapter : name, origin: name }));
      }
      notes.push('普通词表按“英文、中文、章节”读取；有表头时自动识别列。');
    }
    if (formulaWithoutCache) notes.push('有公式缺少保存结果或含错误值。若错词不全，请先用 WPS 打开、重新计算并保存，再导入。');
    if (!result.length && specialized) notes.push('当前没有读到已保存的错词。请先在 WPS 重新计算并保存，或粘贴错词清单。');
    return { entries: mergeEntries([], result).entries, notes };
  }
  function voicesFor(voices, accent) {
    return voices.filter(v => text(v.lang).replace(/_/g, '-').toLowerCase() === accent.toLowerCase()).sort((a, b) => {
      const score = v => (/natural|neural/i.test(v.name) ? 20 : 0) + (/google|microsoft/i.test(v.name) ? 3 : 0) + (v.default ? 1 : 0);
      return score(b) - score(a) || a.name.localeCompare(b.name);
    });
  }

  function gradeAnswer(answer, expected, alternatives = []) {
    const normalize = value => clean(value).normalize('NFKC').toLocaleLowerCase('en');
    const value = text(answer).slice(0, 300), key = normalize(value);
    const spellings = [expected, ...alternatives].map(clean);
    return { answer: value, correctSpelling: spellings.join(' / '), empty: !key, correct: !!key && spellings.some(word => key === normalize(word)) };
  }
  function restoreDictationAnswers(records, entries) {
    if (!Array.isArray(records)) return [];
    const items = new Map(entries.map(item => [item.id, item])), result = new Map();
    for (const record of records.slice(0, 50000)) {
      if (!record || typeof record.id !== 'string' || typeof record.answer !== 'string' || record.answer.length > 300 || !items.has(record.id)) continue;
      const item = items.get(record.id);
      const grade = gradeAnswer(record.answer, item.word, item.alternatives); if (grade.empty) continue;
      const sequence = Number.isFinite(Number(record.sequence)) ? Math.max(0, Number(record.sequence)) : 0;
      const round = typeof record.round === 'string' && record.round.length <= 160 ? record.round : undefined;
      result.set(record.id, { id: record.id, ...grade, sequence, ...(round ? { round } : {}), ...(record.uncertain === true && grade.correct ? { uncertain: true } : {}) });
    }
    return [...result.values()];
  }

  class SpeechPlayer {
    constructor({ synth, Utterance, onState = () => {}, onError = () => {}, watchdogMs = 30000 }) {
      this.synth = synth; this.Utterance = Utterance; this.onState = onState; this.onError = onError; this.watchdogMs = watchdogMs;
      this.token = 0; this.queue = []; this.index = 0; this.mode = 'all'; this.phase = 'idle'; this.paused = false;
    }
    emit() { this.onState({ phase: this.phase, paused: this.paused, index: this.index, total: this.queue.length, item: this.queue[this.index] || null, repeatIndex: this.repeatIndex || 0, repeat: this.options?.repeat || 1, mode: this.mode, speechProgress: this.speechProgress ?? null, remaining: this.wait ? Math.max(0, this.paused ? this.wait.remaining : this.wait.deadline - Date.now()) : 0 }); }
    get active() { return ['speaking', 'waiting', 'paused'].includes(this.phase); }
    start(items, options, mode = 'all') {
      if (!items.length) return;
      if (!this.synth || !this.Utterance || !options.voice) { this.onError(new Error('当前口音没有可用语音，请先选择语音或查看语音设置。')); return; }
      this.stop(false); this.queue = [...items]; this.options = { ...options }; this.mode = mode; this.index = 0; this.paused = false;
      const token = this.token;
      this.done = this.run(token).catch(error => { if (token === this.token) { this.stop(false); this.phase = 'error'; this.emit(); this.onError(error); } });
      return this.done;
    }
    stop(notify = true) {
      this.token++; this.paused = false; this.speechProgress = null;
      if (this.cancelSpeak) this.cancelSpeak();
      if (this.wait) { clearTimeout(this.wait.timer); this.wait.resolve(false); this.wait = null; }
      if (this.pauseGate) { this.pauseGate(false); this.pauseGate = null; }
      if (this.synth) { this.synth.cancel(); this.synth.resume(); }
      this.phase = 'idle'; if (notify) this.emit();
    }
    pause() {
      if (!this.active || this.paused) return;
      this.paused = true; this.beforePause = this.phase;
      if (this.wait) { clearTimeout(this.wait.timer); this.wait.remaining = Math.max(0, this.wait.deadline - Date.now()); }
      if (this.phase === 'speaking') this.synth.pause();
      this.phase = 'paused'; this.emit();
    }
    resume() {
      if (!this.paused) return;
      this.paused = false; this.phase = this.beforePause || 'speaking'; this.synth.resume();
      if (this.wait) this.scheduleWait();
      if (this.pauseGate) { this.pauseGate(true); this.pauseGate = null; }
      this.emit();
    }
    async unpaused(token) {
      if (!this.paused) return token === this.token;
      await new Promise(resolve => { this.pauseGate = resolve; }); return token === this.token;
    }
    async run(token) {
      for (let i = 0; i < this.queue.length; i++) {
        if (token !== this.token) return;
        this.index = i;
        for (let n = 0; n < this.options.repeat; n++) {
          if (!(await this.unpaused(token))) return;
          this.repeatIndex = n + 1; this.speechProgress = null; this.phase = 'speaking'; this.emit();
          if (!(await this.speak(this.queue[i].word, token))) return;
          if (n < this.options.repeat - 1 && !(await this.delay(450, token))) return;
        }
        if (!(await this.unpaused(token))) return;
        if (i < this.queue.length - 1 && !(await this.delay(this.options.gap * 1000, token))) return;
      }
      if (token === this.token) { this.phase = 'complete'; this.emit(); }
    }
    speak(word, token) {
      return new Promise((resolve, reject) => {
        let settled = false, watchdog;
        const finish = (ok, error) => {
          if (settled) return; settled = true; clearTimeout(watchdog);
          if (this.cancelSpeak === cancel) this.cancelSpeak = null;
          if (error) reject(error); else resolve(ok);
        };
        const cancel = () => finish(false); this.cancelSpeak = cancel;
        const utterance = new this.Utterance(word);
        utterance.voice = this.options.voice; utterance.lang = this.options.voice.lang; utterance.rate = this.options.rate;
        utterance.volume = 1; utterance.pitch = 1;
        utterance.onboundary = event => {
          if (settled || token !== this.token || !Number.isFinite(event.charIndex) || event.charIndex <= 0) return;
          this.speechProgress = Math.min(1, event.charIndex / Math.max(1, word.length)); this.emit();
        };
        utterance.onend = () => { if (!settled && token === this.token) this.speechProgress = 1; finish(token === this.token); };
        utterance.onerror = event => {
          if (token !== this.token) return finish(false);
          const reasons = { 'not-allowed': '浏览器阻止了朗读，请点击播放按钮重试。', 'network': '在线语音连接失败，请检查网络或换一个本地语音。', 'audio-busy': '音频设备忙，请稍后重试。', 'voice-unavailable': '这个语音暂时不可用，请换一个声音。', 'language-unavailable': '所选口音不可用，请添加对应的英语语音。', 'synthesis-failed': '语音合成失败，请重试或换一个声音。' };
          finish(false, new Error(reasons[event.error] || '朗读中断，请重试或换一个声音。'));
        };
        const checkTimeout = () => {
          if (settled) return;
          if (this.paused) { watchdog = setTimeout(checkTimeout, 1000); return; }
          finish(false, new Error('语音没有响应，请检查网络，或换一个声音后重新播放。'));
        };
        watchdog = setTimeout(checkTimeout, this.watchdogMs);
        try { this.synth.resume(); this.synth.speak(utterance); } catch (error) { finish(false, new Error('当前浏览器无法启动朗读，请使用 Edge 或 Chrome 重试。')); }
      });
    }
    delay(ms, token) {
      if (token !== this.token) return Promise.resolve(false);
      if (ms <= 0) return Promise.resolve(true);
      this.phase = this.paused ? 'paused' : 'waiting'; this.beforePause = 'waiting';
      return new Promise(resolve => {
        this.wait = { remaining: ms, deadline: Date.now() + ms, resolve, token, timer: null };
        if (!this.paused) this.scheduleWait(); this.emit();
      });
    }
    scheduleWait() {
      if (!this.wait) return;
      const wait = this.wait; wait.deadline = Date.now() + wait.remaining;
      wait.timer = setTimeout(() => { if (this.wait === wait) this.wait = null; wait.resolve(wait.token === this.token); }, wait.remaining);
    }
  }
  const api = { DEFAULTS, clean, validWord, normalizeEntry, wordKey, groupKey, mergeEntries, groupsOf, orderEntries, safeSettings, parseDelimited, rowsToEntries, parseText, parseWorkbook, voicesFor, gradeAnswer, restoreDictationAnswers, SpeechPlayer };
  root.DictationCore = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : window);
