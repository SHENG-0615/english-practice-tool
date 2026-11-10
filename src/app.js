(function () {
  'use strict';
  const C = window.DictationCore;
  const $ = id => document.getElementById(id);
  const SEED = JSON.parse($('seedData').textContent);
  const workbookSync = window.DictationXlsxSync.create(renderWorkbookSync, { database: window.EnglishPublicData.keys.dictationSync, journal: window.EnglishPublicData.keys.dictationJournal });
  const STORAGE_KEY = window.EnglishPublicData.keys.dictation;
  const BACKUP_FORMAT = 'dictation-player-library';
  const BACKUP_VERSION = 2;
  const orderedEntries = SEED.ordered.map(row => {
    const entry = C.normalizeEntry(row);
    return { ...entry, id: 'q-' + encodeURIComponent(row.source) + '-' + row.chapter + '-' + row.number, bankId: entry.id, number: row.number };
  });
  const element = (tag, className, content) => { const node = document.createElement(tag); if (className) node.className = className; if (content !== undefined) node.textContent = content; return node; };
  const icon = name => { const node = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); const use = document.createElementNS('http://www.w3.org/2000/svg', 'use'); use.setAttribute('href', '#i-' + name); node.setAttribute('aria-hidden', 'true'); node.append(use); return node; };
  let entries, settings, scopeModes, query = '', voices = [], groups = [], words = [];
  let savedSettings = false, storageError = '', lastImport = null, toastTimer, playback = { phase: 'idle', paused: false, total: 0 };
  let hasPlayed = false, importMode = 'paste', preview = null, importToken = 0, voiceLoaded = false, activeImportWorker = null;
  const sessions = {
    ordered: { answers: new Map(), drafts: new Map(), lastId: '', focusId: '', round: 'resume-ordered' },
    mistakes: { answers: new Map(), drafts: new Map(), lastId: '', focusId: '', round: 'resume-mistakes' }
  };
  const selections = {};
  let scope = 'ordered', session = sessions.ordered, composingAnswer = false;

  function dictationSnapshot(target) {
    return {
      answers: [...target.answers.values()].map(({ id, answer, sequence, round, uncertain }) => ({ id, answer, sequence, round, ...(uncertain ? { uncertain: true } : {}) })),
      drafts: [...target.drafts].map(([id, text]) => ({ id, text })),
      lastId: target.lastId,
      focusId: target.focusId,
      round: target.round
    };
  }
  function restoreDictation(snapshot, target, availableEntries, replace) {
    if (typeof snapshot?.round === 'string' && snapshot.round.length <= 160) target.round = snapshot.round;
    if (replace) {
      target.answers.clear();
      target.drafts.clear();
      target.lastId = '';
      target.focusId = '';
    }
    const records = C.restoreDictationAnswers(snapshot?.answers, availableEntries);
    for (const record of records) target.answers.set(record.id, record);
    const ids = new Set(availableEntries.map(item => item.id));
    for (const draft of Array.isArray(snapshot?.drafts) ? snapshot.drafts.slice(0, 50000) : []) {
      if (draft && ids.has(draft.id) && typeof draft.text === 'string') target.drafts.set(draft.id, draft.text.slice(0, 300));
    }
    if (target.answers.has(snapshot?.lastId)) target.lastId = snapshot.lastId;
    if (ids.has(snapshot?.focusId)) target.focusId = snapshot.focusId;
  }
  function practiceSnapshot() {
    return {
      ordered: { selected: [...selections.ordered], mode: scopeModes.ordered, dictation: dictationSnapshot(sessions.ordered) },
      mistakes: { selected: [...selections.mistakes], mode: scopeModes.mistakes, dictation: dictationSnapshot(sessions.mistakes) }
    };
  }
  function attemptLogRow(item, grade, practice, id, time, kind) {
    return {
      id, time, kind, source: item.source, chapter: item.chapter, number: item.number, round: grade.round || sessions[practice].round,
      mode: practice === 'ordered' ? '顺序练习' : '错题练习',
      word: grade.correctSpelling, answer: grade.answer, result: grade.correct ? '答对' : '答错', meaning: item.meaning
    };
  }
  function initialWorkbookRecords() {
    const rows = SEED.entries.map(raw => {
      const item = C.normalizeEntry(raw);
      return {
        id: '1-' + item.id, kind: '原表错词', source: item.source, chapter: item.chapter,
        mode: raw.origin, word: item.word, result: '答错', meaning: item.meaning
      };
    });
    for (const practice of ['ordered', 'mistakes']) {
      const items = new Map((practice === 'ordered' ? orderedEntries : entries).map(item => [item.id, item]));
      for (const grade of sessions[practice].answers.values()) {
        const item = items.get(grade.id);
        if (item) {
          const legacyId = JSON.stringify([practice, item.id, grade.sequence, grade.answer]);
          rows.push(attemptLogRow(item, grade, practice, '2-' + legacyId, '', '旧记录'));
          if (grade.uncertain) rows.push(attemptLogRow(item, grade, practice, '4-old-' + legacyId, '', '手动加入错题'));
        }
      }
    }
    return rows;
  }
  function renderWorkbookSync(state) {
    window.DictationLocalFiles?.showLocation('dictationWorkbookLocation', state);
    $('useLocalWorkbook').hidden = !state.localAvailable || state.connectionKind === 'local';
    $('useLocalWorkbook').disabled = state.busy;
    $('xlsxConnectionHelp').textContent = state.connectionKind === 'local'
      ? '已自动连接页面显示地址的同文件夹表格。请在 WPS 或 Excel 中关闭它，作答会分批写入，也可点“立即同步”。'
      : '请先在 WPS 或 Excel 里关闭这份工作簿，再选择本地 XLSX 并授权。授权后自动分批写入，也可点击“立即同步”。';
    const { connected, supported, loaded, busy, synced, pending, writable, error, chapterBackfill } = state;
    const status = !supported ? '请用 Microsoft Edge 或 Google Chrome 打开本文件，再连接 XLSX。'
      : !loaded ? '正在读取本机作答记录。'
      : busy ? '正在连接或同步工作簿。'
      : connected && chapterBackfill ? '已连接 ' + connected + '。旧作答等待写回原章节，请点“立即同步”。'
      : connected ? '已连接 ' + connected + '。已同步 ' + synced + ' 条，待同步 ' + pending + ' 条。'
      : '还未连接工作簿。待同步 ' + pending + ' 条。';
    $('xlsxStatus').textContent = status;
    $('xlsxSyncError').textContent = error || (connected && !writable && !busy ? '浏览器需要重新授权。点“立即同步”继续。' : '');
    $('xlsxSyncError').hidden = !$('xlsxSyncError').textContent;
    $('pickWorkbook').textContent = connected ? '更换工作簿' : '选择工作簿';
    $('pickWorkbook').disabled = !supported || !loaded || busy;
    $('pickWorkbook').classList.toggle('primary', !connected);
    $('syncWorkbook').disabled = !supported || !connected || busy;
    $('syncWorkbook').hidden = !connected;
  }
  function chapterSelection(items, saved, firstOnly = false) {
    const keys = C.groupsOf(items).map(group => group.key);
    if (!Array.isArray(saved)) return new Set(firstOnly ? keys.slice(0, 1) : keys);
    const known = new Set(keys);
    return new Set(saved.filter(key => known.has(key)));
  }
  function practiceModes(data) {
    return {
      ordered: C.safeSettings({ mode: data.ordered?.mode ?? 'dictation' }).mode,
      mistakes: C.safeSettings({ mode: data.mistakes?.mode ?? data.settings?.mode ?? 'dictation' }).mode
    };
  }
  function restorePractice(data, replace = true) {
    if (replace) {
      selections.ordered = chapterSelection(orderedEntries, data.ordered?.selected, true);
      selections.mistakes = chapterSelection(entries, data.mistakes?.selected ?? data.selected);
    }
    restoreDictation(data.ordered?.dictation, sessions.ordered, orderedEntries, replace);
    // Legacy chapter and mistake practice used the same word IDs. Prefer the review result for words in both.
    restoreDictation(data.dictation, sessions.mistakes, entries, replace);
    restoreDictation(data.mistakes?.dictation, sessions.mistakes, entries, false);
  }

  function load() {
    entries = C.mergeEntries([], SEED.entries).entries;
    let data = {};
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const saved = JSON.parse(raw);
        if (saved.format !== BACKUP_FORMAT || !Array.isArray(saved.entries) || saved.entries.length > 50000) throw new Error('词库格式无法识别');
        entries = C.mergeEntries([], saved.entries).entries;
        data = saved;
        savedSettings = true;
        scope = saved.version === BACKUP_VERSION && saved.scope === 'ordered' ? 'ordered' : 'mistakes';
      }
    } catch (error) {
      storageError = '本机存档暂时无法读取，已显示初始词单。你可以导入备份恢复；建议先备份当前词库。';
      try { const raw = localStorage.getItem(STORAGE_KEY); if (raw) localStorage.setItem(STORAGE_KEY + '-recovery', raw); } catch (_) { /* Continue with in-memory data. */ }
    }
    restorePractice(data);
    if (!sessions.mistakes.focusId && typeof data.focusId === 'string') sessions.mistakes.focusId = data.focusId;
    scopeModes = practiceModes(data);
    settings = C.safeSettings({ ...data.settings, mode: scopeModes[scope] });
    session = sessions[scope];
  }
  function persist() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ format: BACKUP_FORMAT, version: BACKUP_VERSION, entries, settings, scope, ...practiceSnapshot() }));
    } catch (error) {
      storageError = '浏览器未能保存到本机。当前可以继续练习；关闭前请点击“备份词库”保存新增词汇。';
      $('storageNotice').textContent = storageError; $('storageNotice').hidden = false;
    }
  }
  function showToast(message, withUndo = false) {
    clearTimeout(toastTimer); $('toastText').textContent = message; $('undoImport').hidden = !withUndo; $('toast').hidden = false;
    toastTimer = setTimeout(() => { $('toast').hidden = true; }, withUndo ? 30000 : 9000);
  }
  function hideToast() { clearTimeout(toastTimer); $('toast').hidden = true; }
  function invalidateList() { if (player.active) { player.stop(); showToast('词汇范围已更改，播放已停止。'); } }
  function groupDescription(item) { return item ? item.source + ' / ' + item.chapter : ''; }
  function recompute() {
    const available = scope === 'ordered' ? orderedEntries : C.orderEntries(entries);
    groups = C.groupsOf(available);
    const selected = selections[scope];
    const q = query.trim().toLocaleLowerCase();
    words = available.filter(item => {
      return selected.has(C.groupKey(item)) && (!q || [item.word, ...(item.alternatives || []), item.meaning, item.chapter, item.source].some(value => value.toLocaleLowerCase().includes(q)));
    });
    const currentExists = words.some(item => item.id === session.focusId);
    const pending = settings.mode === 'dictation' ? words.find(item => !session.answers.has(item.id)) : null;
    if (pending && (scope === 'ordered' || !currentExists || session.answers.has(session.focusId))) session.focusId = pending.id;
    else if (!currentExists) session.focusId = words[0]?.id || '';
  }
  function renderSidebar() {
    const selected = selections[scope];
    const ordered = scope === 'ordered';
    const focused = document.activeElement?.dataset?.group;
    const container = $('chapterGroups'); container.replaceChildren();
    let currentSource, section;
    for (const group of groups) {
      if (group.source !== currentSource) {
        currentSource = group.source;
        const total = groups.filter(g => g.source === currentSource).reduce((sum, g) => sum + g.count, 0);
        const heading = element('div', 'source-heading'); heading.append(element('span', '', currentSource), element('span', '', total + (ordered ? ' 题' : ' 词'))); container.append(heading);
        section = element('div', 'chapter-options'); container.append(section);
      }
      const label = element('label', 'chapter-option' + (selected.has(group.key) ? ' selected' : ''));
      const input = element('input'); input.type = 'checkbox'; input.checked = selected.has(group.key); input.dataset.group = group.key; input.setAttribute('aria-label', group.source + '，章节 ' + group.chapter + '，' + group.count + (ordered ? ' 道题' : ' 个词'));
      input.addEventListener('change', () => { invalidateList(); if (input.checked) selected.add(group.key); else selected.delete(group.key); label.classList.toggle('selected', input.checked); renderContent(); persist(); });
      label.append(input, element('span', 'chapter-name', group.chapter), element('span', 'chapter-count', group.count)); section.append(label);
      if (focused === group.key) input.focus({ preventScroll: true });
    }
    $('librarySummary').textContent = (ordered ? orderedEntries.length + ' 道题，' : entries.length + ' 个词，') + groups.length + ' 个章节';
  }
  function renderScope() {
    const inMistakes = scope === 'mistakes';
    $('orderedScope').setAttribute('aria-pressed', !inMistakes);
    $('mistakesScope').setAttribute('aria-pressed', inMistakes);
    $('chapterDetails').hidden = inMistakes && !entries.length;
    $('chapterSummary').textContent = inMistakes ? '按章节筛选' : '章节选择';
    $('mistakePanel').hidden = !inMistakes;
    $('mistakeCount').textContent = entries.length + ' 个词';
    $('manageMistakes').disabled = !entries.length;
    $('chapterFooterHint').textContent = inMistakes ? '按原章节筛选错题，也可以添加词汇。' : '按本地导入词表的章节和题号练习。';
    $('addFromSidebar').hidden = !inMistakes;
    const chosen = groups.filter(g => selections[scope].has(g.key));
    $('scopeTitle').textContent = inMistakes ? '错题练习' : '顺序练习';
    const count = (query ? '筛出 ' : '共 ') + words.length + (inMistakes ? ' 个词' : ' 道题');
    $('scopeDescription').textContent = (chosen.length === 1 ? groupDescription(chosen[0]) : '已选 ' + chosen.length + ' 个章节') + '，' + count;
    $('listCount').textContent = words.length + ' 词'; $('clearSearch').hidden = !query;
    $('reviewMode').setAttribute('aria-pressed', settings.mode === 'review'); $('practiceMode').setAttribute('aria-pressed', settings.mode === 'practice'); $('dictationMode').setAttribute('aria-pressed', settings.mode === 'dictation');
    $('displayLabel').hidden = settings.mode !== 'review'; $('practiceHint').hidden = settings.mode !== 'practice'; $('display').value = settings.display;
    $('listTitle').textContent = settings.mode === 'practice' ? '练习序号' : settings.mode === 'dictation' ? '作答记录' : '本次词汇';
    $('emptyState').hidden = !!words.length;
    if (!orderedEntries.length && !entries.length) {
      $('emptyTitle').textContent = '先导入自己的词库';
      $('emptyDescription').textContent = '点击上方“导入本地词库”。词汇和作答只保存在本机。';
      $('emptyAction').textContent = '导入本地词库';
    } else if (!words.length && inMistakes && !entries.length) {
      $('emptyTitle').textContent = '错题库还没有词';
      $('emptyDescription').textContent = '顺序练习中答错的词会自动收录到这里。';
      $('emptyAction').textContent = '去顺序练习';
    } else if (!words.length) {
      $('emptyTitle').textContent = chosen.length ? '没有找到匹配的词汇' : '先选择一个章节';
      $('emptyDescription').textContent = chosen.length ? '可以搜索英文、中文或章节号，也可以清除搜索查看所选章节。' : '勾选左侧一个或多个章节，开始练习。';
      $('emptyAction').textContent = chosen.length ? '清除搜索' : '选择全部章节';
    }
  }
  function renderFocus() {
    const index = words.findIndex(item => item.id === session.focusId), item = words[index];
    const practice = settings.mode !== 'review';
    $('focusWord').classList.toggle('number-only', practice);
    $('focusWord').setAttribute('lang', practice ? 'zh-CN' : settings.display === 'chinese' ? 'zh-CN' : 'en');
    $('focusChapter').textContent = practice ? '当前练习' : groupDescription(item) || '尚未选择词汇';
    $('focusPosition').textContent = item ? '第 ' + (index + 1) + ' / ' + words.length + ' 项' : '0 项';
    if (!item) { $('focusWord').textContent = '—'; $('focusMeaning').textContent = ''; }
    else if (practice) { $('focusWord').textContent = String(index + 1).padStart(2, '0'); $('focusMeaning').textContent = ''; }
    else if (settings.display === 'chinese') { $('focusWord').textContent = item.meaning || '暂未填写中文'; $('focusMeaning').textContent = ''; }
    else { $('focusWord').textContent = item.word; $('focusMeaning').textContent = settings.display === 'both' ? item.meaning || '暂未填写中文' : ''; }
    $('focusMeaning').hidden = !$('focusMeaning').textContent;
    $('focusHint').textContent = practice ? '只保留序号。点击序号点读，或播放全部进行听写。' : '点击下方词汇旁的播放按钮，可以逐条点读。';
    $('replay').disabled = !item;
  }
  function renderList() {
    const container = $('wordList'); container.replaceChildren();
    if (!words.length) return;
    if (settings.mode === 'dictation') { renderDictationRecords(container); return; }
    if (settings.mode === 'practice') {
      const grid = element('div', 'practice-grid');
      words.forEach((item, index) => {
        const button = element('button', 'practice-tile' + (item.id === session.focusId ? ' current' : ''));
        button.dataset.entry = item.id; button.setAttribute('aria-label', '播放第 ' + (index + 1) + ' 项');
        button.append(element('span', 'ordinal', String(index + 1).padStart(2, '0')), icon('sound'));
        button.addEventListener('click', () => startSingle(item.id)); grid.append(button);
      });
      container.append(grid); return;
    }
    const table = element('table', 'word-table');
    const head = element('thead'), tr = element('tr');
    const inMistakes = scope === 'mistakes';
    const columns = [['number', '序号'], ...(settings.display !== 'chinese' ? [['word', '英文']] : []), ...(settings.display !== 'english' ? [['meaning', '中文']] : []), ['chapter', inMistakes ? '来源与章节' : '章节'], ['play', '点读']];
    if (inMistakes) columns.push(['remove', '管理']);
    columns.forEach(([name, title]) => tr.append(element('th', 'col-' + name, title))); head.append(tr); table.append(head);
    const body = element('tbody');
    words.forEach((item, index) => {
      const row = element('tr', item.id === session.focusId ? 'current' : ''); row.dataset.entry = item.id;
      row.append(element('td', 'col-number', String(index + 1).padStart(2, '0')));
      if (settings.display !== 'chinese') {
        const cell = element('td', 'col-word'), button = element('button', 'word-button', item.word); button.lang = 'en';
        button.addEventListener('click', () => setFocus(item.id)); cell.append(button); row.append(cell);
      }
      if (settings.display !== 'english') row.append(element('td', 'col-meaning', item.meaning || '—'));
      row.append(element('td', 'col-chapter', inMistakes ? groupDescription(item) : item.chapter));
      const cell = element('td', 'col-play'), play = element('button', 'play-small'); play.setAttribute('aria-label', '播放第 ' + (index + 1) + ' 项'); play.title = '点读'; play.append(icon('sound')); play.addEventListener('click', () => startSingle(item.id)); cell.append(play); row.append(cell);
      if (inMistakes) {
        const removeCell = element('td', 'col-remove');
        const remove = element('button', 'text-button remove-mistake', '移出');
        remove.setAttribute('aria-label', '移出 ' + item.word + '（' + groupDescription(item) + '）');
        remove.addEventListener('click', () => removeMistake(item.id, index));
        removeCell.append(remove);
        row.append(removeCell);
      }
      body.append(row);
    });
    table.append(body); container.append(table);
  }
  function highlight() {
    document.querySelectorAll('[data-entry]').forEach(node => node.classList.toggle('current', node.dataset.entry === session.focusId));
  }
  function renderContent() { recompute(); renderScope(); renderModeLayout(); renderFocus(); renderDictation(); renderList(); renderTransport(); }
  function setFocus(id) { session.focusId = id; renderFocus(); highlight(); renderTransport(); persist(); }
  function dictationRecords() {
    const records = [];
    words.forEach((item, index) => {
      const result = session.answers.get(item.id);
      if (result) records.push({ item, result, number: index + 1 });
    });
    return records;
  }
  function dictationIsComplete() { return !!words.length && words.every(item => session.answers.has(item.id)); }
  function renderModeLayout() {
    const active = settings.mode === 'dictation';
    document.body.classList.toggle('dictation-active', active); $('dictationPanel').hidden = !active || !words.length; $('focusStage').hidden = active || (scope === 'mistakes' && !words.length);
    if (active) { if ($('playerDock').parentElement !== $('dictationPlayerSlot')) $('dictationPlayerSlot').append($('playerDock')); }
    else if ($('playerDock').parentElement !== $('main')) $('playerHome').after($('playerDock'));
  }
  function renderDictation() {
    if (settings.mode !== 'dictation') return;
    const records = dictationRecords(), correct = records.filter(row => row.result.correct).length, complete = dictationIsComplete();
    const index = words.findIndex(item => item.id === session.focusId), item = words[index];
    $('dictationTitle').textContent = complete ? '本轮听写结果' : '第 ' + (index + 1) + ' / ' + words.length + ' 题';
    $('bookQuestionNumber').textContent = scope === 'ordered' && item && !complete ? '原表 ' + item.chapter + ' · 第 ' + item.number + ' 题' : '';
    $('bookQuestionNumber').hidden = !$('bookQuestionNumber').textContent;
    $('dictationScore').textContent = '已提交 ' + records.length + ' 题，答对 ' + correct + ' 题';
    $('dictationForm').hidden = complete || !item; $('dictationComplete').hidden = !complete;
    $('dictationCompleteScore').textContent = '共 ' + words.length + ' 题，答对 ' + correct + ' 题，答错 ' + (records.length - correct) + ' 题。';
    $('dictationAnswerLabel').textContent = '第 ' + (index + 1) + ' 题：输入听到的英文';
    const input = $('dictationAnswer');
    input.placeholder = scope === 'ordered' ? '输入后按回车提交' : '输入后按回车，立即判题并进入下一题';
    if (input.dataset.entry !== session.focusId) { input.dataset.entry = session.focusId; input.value = session.drafts.get(session.focusId) || ''; $('dictationInputError').hidden = true; }
    const last = words.find(w => w.id === session.lastId), result = last && session.answers.get(last.id), feedback = $('dictationFeedback');
    feedback.replaceChildren(); feedback.hidden = !result;
    if (result) {
      const number = words.findIndex(w => w.id === last.id) + 1;
      feedback.classList.toggle('incorrect', !result.correct); feedback.append(icon(result.correct ? 'check' : 'close'));
      const details = element('div'), heading = element('div', 'feedback-heading');
      heading.append(element('strong', '', '上一题（第 ' + number + ' 题）：' + (result.correct ? '答对了' : '答错了')));
      details.append(heading);
      details.append(element('p', '', '你的答案：' + result.answer));
      const spelling = element('p', '', '正确拼写：');
      const expected = element('span', 'correct-spelling', result.correctSpelling);
      expected.lang = 'en';
      spelling.append(expected);
      details.append(spelling, element('p', 'answer-meaning', '中文释义 ' + (last.meaning || '暂无中文释义')));
      if (!result.correct && scope === 'ordered' && entries.some(entry => entry.id === last.bankId)) details.append(element('p', '', '已收录到错题库'));
      if (result.correct) {
        const added = result.uncertain && entries.some(entry => entry.id === (last.bankId || last.id));
        const uncertain = element('button', 'button uncertain-action', added ? '已加入错题' : '不确定？加入错题');
        uncertain.type = 'button';
        uncertain.disabled = added;
        uncertain.addEventListener('click', () => markUncertain(last.id));
        heading.append(uncertain);
      }
      feedback.append(details);
    }
  }
  function renderDictationRecords(container) {
    const records = dictationRecords().sort((a, b) => b.result.sequence - a.result.sequence);
    $('listCount').textContent = records.length + ' / ' + words.length + ' 题';
    if (!records.length) { container.append(element('p', 'dictation-record-empty', '提交后，这里会保留你的答案、对错、正确拼写和释义。')); return; }
    const table = element('table', 'dictation-records'), head = element('thead'), row = element('tr');
    ['序号', '你的答案', '结果', '正确答案'].forEach((value, i) => row.append(element('th', i === 2 ? 'result-cell' : '', value))); head.append(row); table.append(head);
    const body = element('tbody');
    records.forEach(({ item, result, number }) => {
      const tr = element('tr'); tr.append(element('td', '', String(number).padStart(2, '0')));
      tr.append(element('td', result.correct ? '' : 'answer-incorrect', result.answer));
      const status = element('td', 'result-cell ' + (result.correct ? 'correct' : 'incorrect')); status.append(icon(result.correct ? 'check' : 'close'), element('span', '', result.correct ? '正确' : '错误'));
      const answer = element('td');
      const spelling = element('div', 'correct-spelling', result.correctSpelling);
      spelling.lang = 'en';
      answer.append(spelling, element('div', 'answer-meaning', item.meaning || '暂无中文释义'));
      if (result.uncertain) answer.append(element('div', 'answer-meaning', '已标记不确定'));
      tr.append(status, answer);
      body.append(tr);
    });
    table.append(body); container.append(table);
  }
  function changeMode(mode) {
    if (settings.mode === mode) return;
    hideToast();
    player.stop(false); playback = { phase: 'idle', paused: false, total: 0 }; settings.mode = mode; scopeModes[scope] = mode;
    if (mode === 'dictation') session.focusId = words.find(item => !session.answers.has(item.id))?.id || words[0]?.id || '';
    renderContent(); persist();
    if (mode === 'dictation' && words.length && !dictationIsComplete()) $('dictationAnswer').focus({ preventScroll: true });
  }
  function changeScope(nextScope) {
    $('chapterDetails').open = true;
    if (scope === nextScope) return;
    hideToast();
    player.stop(false);
    playback = { phase: 'idle', paused: false, total: 0 };
    scope = nextScope;
    session = sessions[scope];
    settings.mode = scopeModes[scope];
    query = '';
    $('search').value = '';
    $('dictationAnswer').dataset.entry = '';
    composingAnswer = false;
    renderContent(); renderSidebar();
    $('main').scrollTop = 0;
    persist();
  }
  function removeMistake(id, index) {
    invalidateList();
    entries = entries.filter(item => item.id !== id);
    const mistakes = sessions.mistakes;
    mistakes.answers.delete(id);
    mistakes.drafts.delete(id);
    if (mistakes.lastId === id) mistakes.lastId = '';
    lastImport = null;
    renderContent(); renderSidebar();
    persist();
    const remaining = $('wordList').querySelectorAll('.remove-mistake');
    (remaining[Math.min(index, remaining.length - 1)] || $('mistakesScope')).focus();
    showToast('已移出错题库。');
  }
  function collectMistake(item, resetPractice = true) {
    const bankId = item.bankId || item.id;
    const existing = entries.findIndex(entry => entry.id === bankId);
    if (existing >= 0) {
      entries[existing] = C.mergeEntries([entries[existing]], [item]).entries[0];
    } else {
      if (entries.length >= 50000) { showToast('错题库已满，移出一些词后可继续收录。'); return false; }
      const chapter = C.groupKey(item);
      if (!entries.some(entry => C.groupKey(entry) === chapter)) selections.mistakes.add(chapter);
      entries.push(C.normalizeEntry(item));
    }
    if (resetPractice) {
      const mistakes = sessions.mistakes;
      mistakes.answers.delete(bankId);
      mistakes.drafts.delete(bankId);
      if (mistakes.lastId === bankId) mistakes.lastId = '';
    }
    return true;
  }
  function markUncertain(id) {
    const item = words.find(word => word.id === id), grade = session.answers.get(id);
    if (!item || !grade?.correct) return;
    const inBank = entries.some(entry => entry.id === (item.bankId || item.id));
    if (grade.uncertain && inBank) return;
    if (!collectMistake(item, scope === 'ordered')) return;
    grade.uncertain = true;
    const now = new Date();
    const time = new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString().slice(0, 19).replace('T', ' ');
    const actionId = '4-' + now.getTime() + '-' + (window.crypto?.randomUUID?.() || Math.random().toString(36).slice(2));
    workbookSync.record(attemptLogRow(item, grade, scope, actionId, time, '手动加入错题'));
    lastImport = null;
    hideToast();
    renderScope(); renderDictation(); renderList(); persist();
    if (dictationIsComplete()) $('dictationComplete').focus({ preventScroll: true });
    else $('dictationAnswer').focus({ preventScroll: true });
  }
  function playDictationQuestion() {
    if (!readyToSpeak() || dictationIsComplete()) return;
    const item = words.find(w => w.id === session.focusId); if (!item) return;
    hasPlayed = true; player.start([item], options(), 'dictation');
  }
  function submitDictation(event) {
    event.preventDefault();
    if (settings.mode !== 'dictation' || composingAnswer || dictationIsComplete()) return;
    const item = words.find(w => w.id === session.focusId), input = $('dictationAnswer');
    if (!item || input.dataset.entry !== item.id) return;
    const grade = C.gradeAnswer(input.value, item.word, item.alternatives);
    if (grade.empty) { $('dictationInputError').textContent = '请先输入答案，再按回车提交。'; $('dictationInputError').hidden = false; input.focus(); return; }
    const now = new Date();
    const time = new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString().slice(0, 19).replace('T', ' ');
    const attemptId = '3-' + now.getTime() + '-' + (window.crypto?.randomUUID?.() || Math.random().toString(36).slice(2));
    workbookSync.record(attemptLogRow(item, grade, scope, attemptId, time, 'HTML 作答'));
    player.stop(false); playback = { phase: 'idle', paused: false, total: 0 };
    if (lastImport) { lastImport = null; hideToast(); }
    const sequence = [...session.answers.values()].reduce((n, r) => Math.max(n, r.sequence), 0) + 1;
    session.answers.set(item.id, { id: item.id, ...grade, sequence, round: session.round });
    session.lastId = item.id;
    session.drafts.delete(item.id);
    input.value = '';
    input.dataset.entry = '';
    if (!grade.correct && scope === 'ordered') collectMistake(item);
    const position = words.findIndex(w => w.id === item.id);
    const next = words.slice(position + 1).find(w => !session.answers.has(w.id)) || words.find(w => !session.answers.has(w.id));
    if (next) session.focusId = next.id;
    renderContent(); persist();
    if (next) { input.focus({ preventScroll: true }); playDictationQuestion(); }
    else $('dictationComplete').focus();
  }
  function resetDictation() {
    hideToast();
    player.stop(false); playback = { phase: 'idle', paused: false, total: 0 };
    lastImport = null;
    session.round = 'round-' + Date.now() + '-' + (window.crypto?.randomUUID?.() || Math.random().toString(36).slice(2));
    for (const item of words) { session.answers.delete(item.id); session.drafts.delete(item.id); }
    session.lastId = ''; session.focusId = words[0]?.id || ''; $('dictationAnswer').dataset.entry = ''; renderContent(); persist();
    if (words.length) { $('dictationAnswer').focus({ preventScroll: true }); playDictationQuestion(); }
  }
  function renderSettings() {
    $('british').setAttribute('aria-pressed', settings.accent === 'en-GB'); $('american').setAttribute('aria-pressed', settings.accent === 'en-US');
    if (![...$('rate').options].some(option => Number(option.value) === settings.rate)) {
      const option = element('option', '', settings.rate + ' 倍'); option.value = settings.rate; $('rate').append(option);
    }
    $('rate').value = settings.rate; $('repeat').value = settings.repeat; $('gap').value = settings.gap; $('dialogAccent').value = settings.accent;
    renderVoiceStatus();
  }
  function availableVoices() { return C.voicesFor(voices, settings.accent); }
  function currentVoice() { const list = availableVoices(); return list.find(v => v.voiceURI === settings.voiceURI) || list[0] || null; }
  function renderVoiceStatus() {
    const voice = currentVoice();
    $('voiceDot').classList.toggle('warning', !voice);
    $('voiceStatus').textContent = voice ? (settings.accent === 'en-GB' ? '英音' : '美音') + '已就绪，' + (voice.localService ? '本地声音' : '在线声音，需要联网') : voiceLoaded ? '当前浏览器暂未提供' + (settings.accent === 'en-GB' ? '英音' : '美音') + '，请检查语音设置' : '正在检查浏览器语音…';
    const select = $('voiceChoice'); select.replaceChildren();
    const list = availableVoices();
    if (!list.length) { const option = element('option', '', '没有可用的' + (settings.accent === 'en-GB' ? '英音' : '美音')); option.value = ''; select.append(option); select.disabled = true; }
    else {
      select.disabled = false;
      list.forEach(v => { const option = element('option', '', v.name + (v.localService ? '（本地）' : '（在线）')); option.value = v.voiceURI; select.append(option); });
      select.value = voice.voiceURI;
    }
    $('testVoice').disabled = !voice;
    const gb = C.voicesFor(voices, 'en-GB').length, us = C.voicesFor(voices, 'en-US').length;
    $('voiceInventory').textContent = '英音 ' + gb + ' 个，美音 ' + us + ' 个';
    $('missingVoice').hidden = !!voice;
    $('missingVoice').textContent = '当前没有可用的' + (settings.accent === 'en-GB' ? '英音' : '美音') + '。请按下方步骤添加对应语音，再刷新列表。';
    if (!voice && voiceLoaded) $('voiceHelp').open = true;
  }
  function refreshVoices() {
    voices = window.speechSynthesis?.getVoices?.() || [];
    if (voices.length) {
      voiceLoaded = true;
      if (!savedSettings && !C.voicesFor(voices, settings.accent).length && C.voicesFor(voices, 'en-US').length) {
        settings.accent = 'en-US'; settings.voiceURI = ''; savedSettings = true;
      }
    }
    renderSettings();
  }
  function setAccent(accent) {
    if (settings.accent === accent) return;
    if (player.active) player.stop(); settings.accent = accent; settings.voiceURI = ''; savedSettings = true; renderSettings(); persist();
  }
  function options() { return { ...settings, voice: currentVoice() }; }
  function readyToSpeak() {
    if (!words.length) { showToast(scope === 'mistakes' ? '错题库暂无可练习的词。' : '请先选择有词汇的章节。'); return false; }
    if (!currentVoice()) { renderVoiceStatus(); $('voiceHelp').open = true; openDialog('voiceDialog'); return false; }
    return true;
  }
  const player = new C.SpeechPlayer({
    synth: window.speechSynthesis, Utterance: window.SpeechSynthesisUtterance,
    onState(state) {
      playback = state;
      if (state.item && words.some(item => item.id === state.item.id) && (settings.mode !== 'dictation' || state.item.id === session.focusId)) { session.focusId = state.item.id; renderFocus(); highlight(); }
      renderTransport();
    },
    onError(error) { showToast(error.message); }
  });
  function startSingle(id) {
    if (!readyToSpeak()) return;
    const item = words.find(w => w.id === id); if (!item) return;
    setFocus(id); hasPlayed = true; player.start([item], options(), 'single');
  }
  function startAll(start = 0) {
    if (!readyToSpeak()) return;
    hasPlayed = true; player.start(words.slice(start), options(), 'all');
  }
  function togglePlay() { if (player.paused) player.resume(); else if (player.active) player.pause(); else if (settings.mode === 'dictation') playDictationQuestion(); else startAll(); }
  function move(delta) {
    if (settings.mode === 'dictation') return;
    if (!words.length) return;
    const current = Math.max(0, words.findIndex(w => w.id === session.focusId));
    const index = Math.min(words.length - 1, Math.max(0, current + delta));
    const continuous = player.active && player.mode === 'all';
    if (continuous) startAll(index); else startSingle(words[index].id);
  }
  function renderTransport() {
    const index = words.findIndex(w => w.id === session.focusId), phase = playback.phase;
    const dictation = settings.mode === 'dictation';
    $('previous').hidden = dictation; $('next').hidden = dictation;
    $('progress').setAttribute('aria-label', dictation ? '本题朗读进度' : '播放进度');
    if (dictation) {
      const complete = dictationIsComplete(), ownAudio = playback.mode === 'dictation' && playback.item?.id === session.focusId;
      let status = complete ? '本轮已完成' : '点击播放本题，听音后输入答案';
      if (ownAudio && phase === 'speaking') status = '正在朗读，第 ' + playback.repeatIndex + ' / ' + playback.repeat + ' 遍';
      else if (ownAudio && phase === 'paused') status = '朗读已暂停';
      else if (ownAudio && phase === 'waiting') status = '稍后重读当前词汇';
      else if (ownAudio && phase === 'complete') status = '朗读完毕，请输入答案';
      else if (phase === 'error') status = '朗读中断，可以重播本题';
      if ($('playerStatus').textContent !== status) $('playerStatus').textContent = status;
      $('playerPosition').textContent = words.length ? '第 ' + (index + 1) + ' / ' + words.length + ' 题' : '0 / 0';
      $('progress').max = 1;
      if (ownAudio && phase === 'speaking' && playback.speechProgress === null) $('progress').removeAttribute('value');
      else $('progress').value = !ownAudio ? 0 : phase === 'complete' ? 1 : Math.min(1, ((playback.repeatIndex || 1) - 1 + (playback.speechProgress || 0)) / playback.repeat);
      const button = $('playAll'); button.replaceChildren(icon(player.paused ? 'play' : player.active ? 'pause' : 'sound'), element('span', '', player.paused ? '继续朗读' : player.active ? '暂停朗读' : complete ? '本轮已完成' : '播放本题'));
      button.disabled = !words.length || complete; $('stop').disabled = !player.active; return;
    }
    let status = hasPlayed ? '已停止，可以重新开始' : '准备开始';
    if (phase === 'complete') status = playback.mode === 'all' ? '本组已播放完毕' : playback.mode === 'sample' ? '试听结束' : '点读结束';
    else if (phase === 'error') status = '播放中断，请重试或检查语音';
    else if (phase === 'paused') status = '已暂停';
    else if (phase === 'speaking') status = playback.mode === 'sample' ? '正在试听' : (playback.mode === 'all' ? '正在连播' : '正在点读') + '，第 ' + playback.repeatIndex + ' / ' + playback.repeat + ' 遍';
    else if (phase === 'waiting') {
      const remaining = player.wait ? Math.max(0, player.paused ? player.wait.remaining : player.wait.deadline - Date.now()) : 0;
      status = playback.repeatIndex < playback.repeat ? '稍后重读当前词汇' : '停顿 ' + Math.ceil(remaining / 1000) + ' 秒，随后播放下一项';
    }
    if ($('playerStatus').textContent !== status) $('playerStatus').textContent = status;
    $('playerPosition').textContent = words.length ? (index + 1) + ' / ' + words.length : '0 / 0';
    $('progress').max = Math.max(1, words.length); $('progress').value = phase === 'complete' && playback.mode === 'all' ? words.length : Math.max(0, index);
    const button = $('playAll'); button.replaceChildren(icon(player.paused ? 'play' : player.active ? 'pause' : 'play'), element('span', '', player.paused ? '继续播放' : player.active ? '暂停播放' : '播放全部'));
    button.disabled = !words.length; $('previous').disabled = index <= 0; $('next').disabled = index < 0 || index >= words.length - 1; $('stop').disabled = !player.active;
  }

  function openDialog(id) { const dialog = $(id); if (!dialog.open) dialog.showModal(); }
  function closeDialog(id) { $(id).close(); }
  function invalidatePreview() {
    importToken++; activeImportWorker?.abort(); preview = null; $('importPreview').hidden = true; $('commitImport').disabled = true; $('previewImport').disabled = false;
    $('importFeedback').textContent = ''; $('importFeedback').classList.remove('error');
  }
  function setImportMode(mode) {
    importMode = mode; invalidatePreview(); $('pasteTab').setAttribute('aria-pressed', mode === 'paste'); $('fileTab').setAttribute('aria-pressed', mode === 'file'); $('pastePane').hidden = mode !== 'paste'; $('filePane').hidden = mode !== 'file';
  }
  function openImport() { invalidatePreview(); openDialog('importDialog'); }
  function readExcel(buffer, fallback) {
    if (typeof Worker === 'undefined') return Promise.reject(new Error('当前浏览器无法后台读取 Excel，请使用最新版 Edge 或 Chrome，或先导入 CSV。'));
    const code = $('xlsxRuntime').textContent + '\n' + $('coreRuntime').textContent + `
      self.onmessage = function(event) {
        try {
          var buffer = event.data.buffer, fallback = event.data.fallback;
          var names = XLSX.read(buffer, {type:'array',bookSheets:true}).SheetNames;
          var relevant = names.filter(function(name) { return name === '错词总表' || name === '错词练习'; });
          var options = {type:'array',cellFormula:true,cellHTML:false,cellStyles:false};
          if(relevant.length) options.sheets = relevant;
          var book = XLSX.read(buffer, options);
          self.postMessage({ok:true,result:DictationCore.parseWorkbook(book,XLSX,fallback)});
        } catch(error) { self.postMessage({ok:false,error:error.message || '无法读取这个 Excel 文件。'}); }
      };`;
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(new Blob([code], { type: 'application/javascript' }));
      let worker;
      try { worker = new Worker(url); } catch (error) { URL.revokeObjectURL(url); reject(new Error('浏览器阻止了 Excel 读取，请使用 Edge 或 Chrome 重试。')); return; }
      const handle = { abort: () => { cleanup(); reject(new Error('读取已取消。')); } };
      const cleanup = () => { worker.terminate(); URL.revokeObjectURL(url); if (activeImportWorker === handle) activeImportWorker = null; };
      activeImportWorker = handle;
      worker.onmessage = event => { cleanup(); if (event.data.ok) resolve(event.data.result); else reject(new Error(event.data.error)); };
      worker.onerror = () => { cleanup(); reject(new Error('Excel 读取没有完成。请使用 Edge 或 Chrome 重试，或先在 WPS 另存为 CSV。')); };
      worker.postMessage({ buffer, fallback }, [buffer]);
    });
  }
  async function prepareImport() {
    const token = ++importToken; preview = null; $('commitImport').disabled = true; $('importPreview').hidden = true; $('previewImport').disabled = true; $('importFeedback').classList.remove('error'); $('importFeedback').textContent = '正在读取…';
    await new Promise(resolve => requestAnimationFrame(resolve));
    try {
      const fallback = { source: C.clean($('importSource').value) || '自建词库', chapter: C.clean($('importChapter').value) || '未分类', origin: '粘贴文本' };
      let result = { entries: [], notes: [] }, backupSettings = null, backupPractice = null;
      if (importMode === 'paste') result.entries = C.parseText($('pasteText').value, fallback);
      else {
        const file = $('importFile').files[0];
        if (!file) throw new Error('请先选择一个文件。');
        if (file.size > 20 * 1024 * 1024) throw new Error('文件超过 20 MB，请拆分词表后导入。');
        const buffer = await file.arrayBuffer(); if (token !== importToken) return;
        const extension = file.name.split('.').pop().toLowerCase();
        fallback.filename = file.name; fallback.origin = file.name;
        if (fallback.source === '自建词库') fallback.source = file.name.replace(/\.[^.]+$/, '');
        if (['xlsx', 'xls'].includes(extension)) {
          if (!window.XLSX) throw new Error('Excel 读取组件未能加载，请重新打开完整的工具文件。');
          result = await readExcel(buffer, fallback);
        } else {
          let content;
          try { content = new TextDecoder('utf-8', { fatal: true }).decode(buffer); }
          catch (_) { content = new TextDecoder('gb18030').decode(buffer); }
          if (extension === 'json') {
            const data = JSON.parse(content.replace(/^\uFEFF/, ''));
            if (!data || data.format !== BACKUP_FORMAT || ![1, BACKUP_VERSION].includes(data.version) || !Array.isArray(data.entries)) throw new Error('这不是本工具支持的词库备份，请选择导出的 JSON 文件。');
            if (data.entries.length > 50000) throw new Error('词库超过 50,000 条，请分开导入。');
            result.entries = C.mergeEntries([], data.entries).entries;
            backupSettings = { settings: C.safeSettings(data.settings), modes: practiceModes(data) };
            if (data.dictation || data.mistakes?.dictation || data.ordered?.dictation) {
              backupPractice = { dictation: data.dictation, mistakes: data.mistakes, ordered: data.ordered };
            }
            result.notes.push('词库备份将与当前词汇合并，相同来源和章节内的重复词会合并。');
          } else if (['txt', 'csv', 'tsv'].includes(extension)) result.entries = C.parseText(content, fallback);
          else throw new Error('请使用 Excel、CSV、TXT、TSV 或本工具的 JSON 备份。');
        }
      }
      if (token !== importToken) return;
      const cleanEntries = C.mergeEntries([], result.entries).entries;
      if (!cleanEntries.length && !backupSettings) throw new Error(result.notes.join(' ') || '没有识别到英文词汇。请检查内容，或使用“英文 | 中文 | 章节”的格式。');
      const merged = C.mergeEntries(entries, cleanEntries);
      preview = { entries: cleanEntries, backupSettings, backupPractice };
      const groupCounts = C.groupsOf(cleanEntries);
      $('previewSummary').textContent = '识别到 ' + cleanEntries.length + ' 个词，' + groupCounts.length + ' 个章节。将新增 ' + merged.added + ' 词，更新 ' + merged.updated + ' 条释义。';
      $('previewRows').replaceChildren();
      cleanEntries.slice(0, 8).forEach(item => { const row = element('tr'); row.append(element('td', '', item.word), element('td', '', item.meaning || '—'), element('td', '', item.chapter)); $('previewRows').append(row); });
      $('previewNotes').textContent = [...(result.notes || []), groupCounts.map(g => g.source + ' / ' + g.chapter + '：' + g.count + ' 词').join('；'), cleanEntries.length > 8 ? '这里只预览前 8 条，加入时会导入全部。' : ''].filter(Boolean).join('\n');
      $('restoreSettingsLabel').hidden = !backupSettings; $('restoreSettings').checked = false;
      $('restorePracticeLabel').hidden = !backupPractice; $('restorePractice').checked = false;
      $('importPreview').hidden = false; $('commitImport').disabled = false; $('importFeedback').textContent = '请核对下方词汇和章节。';
    } catch (error) {
      if (token !== importToken) return;
      $('importFeedback').textContent = error.message || '无法读取文件，请检查格式后重试。'; $('importFeedback').classList.add('error');
    } finally { if (token === importToken) $('previewImport').disabled = false; }
  }
  function commitImport() {
    if (!preview) return;
    player.stop();
    lastImport = { entries: entries.map(item => ({ ...item })), settings: { ...settings }, ...practiceSnapshot() };
    const merged = C.mergeEntries(entries, preview.entries); entries = merged.entries;
    preview.entries.forEach(item => selections.mistakes.add(C.groupKey(item)));
    if (preview.backupSettings && $('restoreSettings').checked) {
      scopeModes = preview.backupSettings.modes;
      settings = { ...preview.backupSettings.settings, mode: scopeModes[scope] };
    }
    if (scope === 'mistakes' && settings.mode !== 'dictation') session.focusId = preview.entries[0]?.id || session.focusId;
    query = ''; $('search').value = '';
    if (preview.backupPractice && $('restorePractice').checked) {
      restorePractice(preview.backupPractice, false);
      $('dictationAnswer').dataset.entry = '';
    }
    renderContent(); renderSidebar(); renderSettings(); persist(); closeDialog('importDialog');
    showToast('已新增 ' + merged.added + ' 词，更新 ' + merged.updated + ' 条释义。', true);
  }
  function undoImport() {
    if (!lastImport) return;
    const previousImport = lastImport;
    player.stop(); ({ entries, settings } = previousImport); lastImport = null;
    scopeModes = practiceModes(previousImport);
    settings.mode = scopeModes[scope];
    restorePractice(previousImport);
    $('dictationAnswer').dataset.entry = '';
    renderContent(); renderSidebar(); renderSettings(); persist(); showToast('已撤销本次添加。');
  }
  function downloadBackup() {
    const data = { format: BACKUP_FORMAT, version: BACKUP_VERSION, exportedAt: new Date().toISOString(), entries, settings, ...practiceSnapshot() };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob), anchor = element('a'); anchor.href = url; anchor.download = '错词词库备份_' + new Date().toISOString().slice(0, 10) + '.json';
    document.body.append(anchor); anchor.click(); anchor.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
    showToast('已导出词库备份。恢复时在“添加词汇 → 导入文件”中选择它。');
  }

  $('reviewMode').addEventListener('click', () => changeMode('review'));
  $('practiceMode').addEventListener('click', () => changeMode('practice'));
  $('dictationMode').addEventListener('click', () => changeMode('dictation'));
  $('orderedScope').addEventListener('click', () => changeScope('ordered'));
  $('mistakesScope').addEventListener('click', () => changeScope('mistakes'));
  $('manageMistakes').addEventListener('click', () => {
    if (query) $('clearSearch').click();
    if (!words.length) $('selectAll').click();
    changeMode('review');
    $('wordList').querySelector('.remove-mistake')?.focus();
  });
  $('dictationForm').addEventListener('submit', submitDictation);
  $('dictationAnswer').addEventListener('input', () => { session.drafts.set(session.focusId, $('dictationAnswer').value); $('dictationInputError').hidden = true; persist(); });
  $('dictationAnswer').addEventListener('compositionstart', () => { composingAnswer = true; });
  $('dictationAnswer').addEventListener('compositionend', () => { composingAnswer = false; });
  $('dictationAnswer').addEventListener('keydown', event => { if (event.key === 'Enter' && (event.repeat || event.isComposing || composingAnswer)) event.preventDefault(); });
  $('resetDictation').addEventListener('click', resetDictation); $('restartDictation').addEventListener('click', resetDictation);
  $('display').addEventListener('change', () => { settings.display = $('display').value; renderContent(); persist(); });
  $('search').addEventListener('input', () => { invalidateList(); query = $('search').value; renderContent(); });
  $('clearSearch').addEventListener('click', () => { invalidateList(); query = ''; $('search').value = ''; renderContent(); });
  $('selectAll').addEventListener('click', () => { invalidateList(); selections[scope] = new Set(groups.map(g => g.key)); renderContent(); renderSidebar(); persist(); });
  $('selectNone').addEventListener('click', () => { invalidateList(); selections[scope].clear(); renderContent(); renderSidebar(); persist(); });
  $('emptyAction').addEventListener('click', () => {
    if (!orderedEntries.length && !entries.length) { window.EnglishPublicData.open(); return; }
    if (scope === 'mistakes' && !entries.length) changeScope('ordered');
    else if (!groups.some(group => selections[scope].has(group.key))) $('selectAll').click();
    else $('clearSearch').click();
  });
  $('british').addEventListener('click', () => setAccent('en-GB')); $('american').addEventListener('click', () => setAccent('en-US'));
  $('dialogAccent').addEventListener('change', () => setAccent($('dialogAccent').value));
  ['rate', 'repeat', 'gap'].forEach(id => $(id).addEventListener('change', () => {
    const raw = { ...settings, [id]: $(id).value }; settings = C.safeSettings(raw);
    if (player.active) { player.stop(); showToast('参数已更新，点击播放即可按新参数练习。'); }
    renderSettings(); persist();
  }));
  $('voiceChoice').addEventListener('change', () => { if (player.active) player.stop(); settings.voiceURI = $('voiceChoice').value; renderVoiceStatus(); persist(); });
  $('playAll').addEventListener('click', togglePlay); $('stop').addEventListener('click', () => player.stop());
  $('replay').addEventListener('click', () => startSingle(session.focusId)); $('previous').addEventListener('click', () => move(-1)); $('next').addEventListener('click', () => move(1));
  ['openVoice', 'voiceDetails'].forEach(id => $(id).addEventListener('click', () => { refreshVoices(); openDialog('voiceDialog'); }));
  $('testVoice').addEventListener('click', () => { const voice = currentVoice(); if (voice) { hasPlayed = true; player.start([{ id: '__voice_sample', word: 'Hello. Welcome to your listening practice.' }], { ...options(), repeat: 1, gap: 0 }, 'sample'); } });
  $('refreshVoices').addEventListener('click', () => { refreshVoices(); showToast('已刷新，英音 ' + C.voicesFor(voices, 'en-GB').length + ' 个，美音 ' + C.voicesFor(voices, 'en-US').length + ' 个。'); });
  ['openImport', 'addFromSidebar'].forEach(id => $(id).addEventListener('click', openImport));
  $('pasteTab').addEventListener('click', () => setImportMode('paste')); $('fileTab').addEventListener('click', () => setImportMode('file'));
  ['pasteText', 'importSource', 'importChapter'].forEach(id => $(id).addEventListener('input', invalidatePreview));
  $('importFile').addEventListener('change', () => { invalidatePreview(); if ($('importFile').files.length) prepareImport(); });
  $('previewImport').addEventListener('click', prepareImport); $('commitImport').addEventListener('click', commitImport);
  $('backup').addEventListener('click', downloadBackup); $('undoImport').addEventListener('click', undoImport); $('closeToast').addEventListener('click', hideToast);
  $('openExcelSync').addEventListener('click', () => { if (document.body.dataset.practiceArea !== 'vocabulary') openDialog('excelSyncDialog'); });
  $('pickWorkbook').addEventListener('click', () => workbookSync.pick());
  $('syncWorkbook').addEventListener('click', () => workbookSync.sync());
  $('useLocalWorkbook').addEventListener('click', () => workbookSync.connectLocal());
  document.querySelectorAll('[data-close]').forEach(button => button.addEventListener('click', () => closeDialog(button.dataset.close)));
  ['importDialog', 'voiceDialog', 'excelSyncDialog'].forEach(id => {
    $(id).addEventListener('click', event => { const rect = $(id).getBoundingClientRect(); if (event.target === $(id) && (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom)) closeDialog(id); });
    $(id).addEventListener('close', () => { if (id === 'importDialog') invalidatePreview(); if (id === 'voiceDialog' && player.mode === 'sample' && player.active) player.stop(); });
  });
  document.addEventListener('keydown', event => {
    if (document.body.dataset.practiceArea === 'vocabulary') return;
    if (event.ctrlKey || event.altKey || event.metaKey || event.isComposing || document.querySelector('dialog[open]')) return;
    if (event.target.closest('input,textarea,select,a,[contenteditable="true"],summary')) return;
    if (event.code === 'Space' && event.target.closest('button')) return;
    if (event.code === 'Space') { event.preventDefault(); togglePlay(); }
    if (event.code === 'ArrowLeft') { event.preventDefault(); move(-1); }
    if (event.code === 'ArrowRight') { event.preventDefault(); move(1); }
  });
  window.addEventListener('beforeunload', () => { persist(); player.stop(false); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') persist(); });

  window.DictationAudio = {
    stop: () => player.stop(),
    read(word) {
      if (!currentVoice()) { refreshVoices(); openDialog('voiceDialog'); return; }
      hasPlayed = true;
      player.start([{ id: '__vocabulary_sample', word }], { ...options(), repeat: 1, gap: 0 }, 'sample');
    }
  };
  load(); recompute(); renderSidebar(); renderContent(); renderSettings();
  workbookSync.start(initialWorkbookRecords());
  if (storageError) { $('storageNotice').textContent = storageError; $('storageNotice').hidden = false; }
  if (window.matchMedia('(max-width: 760px)').matches) $('chapterDetails').open = false;
  if (window.speechSynthesis?.addEventListener) window.speechSynthesis.addEventListener('voiceschanged', refreshVoices);
  refreshVoices(); [150, 700, 1800].forEach(ms => setTimeout(refreshVoices, ms));
  setTimeout(() => { voiceLoaded = true; renderVoiceStatus(); }, 2200);
  setInterval(() => { if (player.phase === 'waiting') renderTransport(); }, 250);
})();
