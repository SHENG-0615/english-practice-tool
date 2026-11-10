(function () {
  'use strict';
  const C = window.VocabularyCore;
  const $ = id => document.getElementById(id);
  const seed = JSON.parse($('vocabularyData').textContent);
  const entries = seed.entries;
  const byId = new Map(entries.map(entry => [entry.id, entry]));
  const lists = [...new Set(entries.map(entry => entry.list))];
  const KEY = window.EnglishPublicData.keys.vocabulary;
  const areaKey = 'english-public-area-v1';
  const modes = ['spelling', 'meaning'];
  const modesText = { study: '词汇学习', spelling: '拼写练习', meaning: '释义练习' };
  const element = (tag, className, text) => {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  };
  const uuid = () => window.crypto?.randomUUID?.() || Date.now() + '-' + Math.random().toString(36).slice(2);
  const history = [];
  const initial = { mode: 'study', scope: 'lists', bankStatus: 'all', bankKind: 'all', bankList: 'all', bankQuery: '', display: 'both', selected: lists.length ? [lists[0]] : [], onlyMistakes: false, search: '', studyIndex: 0, sessions: {}, bankSessions: {}, spellingErrors: [], meaningErrors: [] };
  const spellingErrors = new Set(), meaningErrors = new Set();
  for (const mode of modes) {
    const session = { answers: {}, drafts: {}, rounds: {}, lastId: '' };
    for (const name of lists) {
      const previous = entries.filter(entry => entry.list === name).flatMap(entry => entry.prior[mode]);
      const index = Math.max(0, ...previous.map(item => item.index));
      session.rounds[name] = seed.freshLists?.[mode]?.includes(name)
        ? { id: 'round-' + mode + '-' + uuid(), native: false }
        : { id: 'native-' + mode + '-' + name + '-' + index, index, native: true };
    }
    for (const entry of entries) {
      for (const old of entry.prior[mode]) {
        const answers = mode === 'spelling' ? [old.answer] : C.splitExistingMeaning(entry, old.answer);
        const grade = C.grade(entry, mode, answers);
        const round = { id: 'native-' + mode + '-' + entry.list + '-' + old.index, index: old.index, native: true };
        history.push({ id: 'origin-' + entry.id + '-' + mode + '-' + old.index, origin: true, mode, entryId: entry.id, list: entry.list, row: entry.row, word: entry.word, answers, round, feedback: grade.feedback, status: grade.status, time: '' });
        if (!grade.correct && mode === 'spelling') spellingErrors.add(entry.id);
        if (grade.status !== 'full' && mode === 'meaning') meaningErrors.add(entry.id);
        if (old.index === session.rounds[entry.list].index) session.answers[entry.id] = { answers, round, sequence: entry.row, ...grade };
      }
    }
    initial.sessions[mode] = session;
  }
  let state = initial, savedState = {}, words = [], composing = false, area = 'dictation';
  try {
    const saved = JSON.parse(localStorage.getItem(KEY));
    if (saved?.version === 1) {
      savedState = saved;
      state = { ...initial, ...saved, sessions: initial.sessions };
      state.mode = ['study', ...modes].includes(saved.mode) ? saved.mode : 'study';
      state.display = ['both', 'chinese', 'english'].includes(saved.display) ? saved.display : 'both';
      state.selected = Array.isArray(saved.selected) ? saved.selected.filter(name => lists.includes(name)) : lists.length ? [lists[0]] : [];
      for (const mode of modes) {
        const stored = saved.sessions?.[mode];
        if (!stored) continue;
        const current = state.sessions[mode];
        current.answers = {};
        for (const [id, record] of Object.entries(stored.answers || {})) {
          if (!byId.has(id) || !Array.isArray(record.answers)) continue;
          const grade = C.grade(byId.get(id), mode, record.answers);
          if (!grade.empty) current.answers[id] = { ...record, ...grade };
        }
        current.drafts = stored.drafts || {};
        current.rounds = { ...current.rounds, ...stored.rounds };
        current.lastId = stored.lastId || '';
      }
      for (const id of saved.spellingErrors || []) if (byId.has(id)) spellingErrors.add(id);
      for (const id of saved.meaningErrors || []) if (byId.has(id)) meaningErrors.add(id);
    }
  } catch (_) { /* The workbook snapshot remains available. */ }

  state.scope = state.scope === 'bank' ? 'bank' : 'lists';
  state.bankStatus = ['all', 'wrong', 'partial'].includes(state.bankStatus) ? state.bankStatus : 'all';
  state.bankKind = ['all', ...modes].includes(state.bankKind) ? state.bankKind : 'all';
  state.bankList = lists.includes(state.bankList) ? state.bankList : 'all';
  state.bankSessions = {};
  for (const mode of modes) {
    const session = { answers: {}, drafts: {}, rounds: {}, lastId: '', queueIds: [], queueKey: '' };
    for (const name of lists) session.rounds[name] = { id: 'bank-' + mode + '-' + uuid(), native: false };
    const stored = savedState.bankSessions?.[mode];
    if (stored) {
      for (const [id, record] of Object.entries(stored.answers || {})) if (byId.has(id) && Array.isArray(record.answers)) {
        const grade = C.grade(byId.get(id), mode, record.answers);
        if (!grade.empty) session.answers[id] = { ...record, ...grade };
      }
      session.drafts = stored.drafts || {}; session.rounds = { ...session.rounds, ...stored.rounds };
      session.lastId = stored.lastId || ''; session.queueKey = stored.queueKey || '';
      session.queueIds = (stored.queueIds || []).filter(id => byId.has(id));
    }
    state.bankSessions[mode] = session;
  }
  const bank = window.VocabularyBank.create(entries, savedState, history);
  function refreshErrorSets() {
    spellingErrors.clear(); meaningErrors.clear();
    for (const issue of bank.list()) (issue.mode === 'spelling' ? spellingErrors : meaningErrors).add(issue.entryId);
  }
  refreshErrorSets();
  for (const name of lists) {
    const option = element('option', '', name); option.value = name; $('vocabBankList').append(option);
  }

  const sync = window.DictationXlsxSync.create(renderSync, {
    database: window.EnglishPublicData.keys.vocabularySync, journal: window.EnglishPublicData.keys.vocabularyJournal, version: 2, localKey: 'vocabulary', disableLocal: !seed.syncCompatible,
    append: (bytes, rows, receipt) => window.VocabularyWorkbook.append(bytes, rows, entries, receipt)
  });
  function persist() {
    try {
      localStorage.setItem(KEY, JSON.stringify({ ...state, version: 1, spellingErrors: [...spellingErrors], meaningErrors: [...meaningErrors], ...bank.snapshot() }));
    } catch (_) {
      $('vocabInputError').textContent = '浏览器未能保存进度，请先完成工作簿同步。';
      $('vocabInputError').hidden = false;
    }
  }
  function currentSession() { return (state.scope === 'bank' ? state.bankSessions : state.sessions)[state.mode]; }
  function bankRecords(status = state.bankStatus) {
    return bank.list({ status, mode: state.mode === 'study' ? state.bankKind : state.mode, list: state.bankList, query: state.bankQuery });
  }
  function markImportant(node, entry) {
    node.classList.toggle('important-word', !!entry?.important);
    if (entry?.important) { const mark = element('span', 'priority-mark', '*'); mark.title = '原表加粗重点词'; mark.setAttribute('aria-label', '重点词'); node.append(mark); }
  }
  function addCorrectAnswer(node, entry, mode) {
    if (mode === 'spelling') {
      const spelling = element('div', 'correct-spelling', entry.spellings.join(' / ')); markImportant(spelling, entry); node.append(spelling);
      node.append(element('p', 'answer-meaning', entry.meaning));
    } else {
      for (const group of entry.groups) node.append(element('p', 'answer-meaning', group.meaning));
    }
  }
  function answerText(entry, mode, answers) {
    return mode === 'spelling' ? answers?.[0] || '未保留当时作答'
      : entry.groups.map((group, i) => C.label(group) + ' ' + (answers?.[i] || '未填')).join(' ｜ ');
  }
  function selectedWords() {
    if (state.scope === 'bank') {
      const active = bankRecords();
      if (state.mode === 'study') return [...new Map(active.map(issue => [issue.entryId, byId.get(issue.entryId)])).values()];
      const session = currentSession();
      const queueKey = JSON.stringify([state.mode, state.bankStatus, state.bankKind, state.bankList, state.bankQuery]);
      if (session.queueKey !== queueKey) { session.queueIds = [...new Set(active.map(issue => issue.entryId))]; session.queueKey = queueKey; }
      return session.queueIds.filter(id => bank.has(id, state.mode)).map(id => byId.get(id));
    }
    const selected = new Set(state.selected);
    const query = String(state.search || '').trim().toLowerCase();
    return entries.filter(entry => {
      const mistake = state.mode === 'spelling' ? spellingErrors.has(entry.id)
        : state.mode === 'meaning' ? meaningErrors.has(entry.id)
        : spellingErrors.has(entry.id) || meaningErrors.has(entry.id);
      return selected.has(entry.list) && (!state.onlyMistakes || mistake)
        && (!query || [entry.word, entry.meaning, entry.list].some(value => value.toLowerCase().includes(query)));
    });
  }
  function renderLists() {
    const holder = $('vocabLists');
    holder.replaceChildren();
    for (const name of lists) {
      const chosen = state.selected.includes(name);
      const row = element('label', 'chapter-option' + (chosen ? ' selected' : ''));
      const input = element('input');
      input.type = 'checkbox'; input.checked = chosen;
      const count = entries.filter(entry => entry.list === name).length;
      input.setAttribute('aria-label', name + '，' + count + ' 个词');
      input.addEventListener('change', () => {
        state.selected = input.checked ? [...state.selected, name] : state.selected.filter(value => value !== name);
        render(); persist();
      });
      row.append(input, element('span', 'chapter-name', name), element('span', 'chapter-count', count));
      holder.append(row);
    }
  }
  function studyItem() {
    state.studyIndex = Math.min(Math.max(0, Number(state.studyIndex) || 0), Math.max(0, words.length - 1));
    return words[state.studyIndex];
  }
  function renderStudy() {
    const entry = studyItem();
    $('vocabStudySource').textContent = entry ? entry.list + ' · 原表第 ' + entry.number + ' 词' : '';
    $('vocabStudyPosition').textContent = entry ? (state.studyIndex + 1) + ' / ' + words.length : '0 / 0';
    $('vocabStudyWord').textContent = !entry ? '' : state.display === 'chinese' ? entry.meaning : entry.word;
    markImportant($('vocabStudyWord'), entry);
    $('vocabStudyMeaning').textContent = entry && state.display === 'both' ? entry.meaning : '';
    $('vocabPhonetic').textContent = entry && state.display !== 'chinese' ? entry.phonetic : '';
    $('vocabPronounce').disabled = !entry;
    $('vocabPrevious').disabled = state.studyIndex <= 0;
    $('vocabNext').disabled = !entry || state.studyIndex >= words.length - 1;
  }
  function activeQuestion() {
    return words.find(entry => !currentSession().answers[entry.id]);
  }
  function renderInputs(entry) {
    const holder = $('vocabInputs');
    const marker = state.scope + ':' + state.mode + ':' + (entry?.id || '');
    if (holder.dataset.question === marker) return;
    holder.dataset.question = marker;
    holder.replaceChildren();
    $('vocabInputError').hidden = true;
    if (!entry) return;
    const labels = state.mode === 'spelling' ? ['输入英文拼写'] : entry.groups.map(group => C.label(group) + ' 的中文释义');
    const drafts = currentSession().drafts[entry.id] || [];
    labels.forEach((text, index) => {
      const label = element('label', '', text);
      const input = element('input');
      input.id = 'vocabAnswer' + index;
      input.type = 'text'; input.maxLength = 300; input.autocomplete = 'off'; input.spellcheck = false;
      input.lang = state.mode === 'spelling' ? 'en' : 'zh-CN';
      input.setAttribute('autocorrect', 'off'); input.setAttribute('autocapitalize', 'none');
      input.placeholder = state.mode === 'spelling' ? '看中文，写英文' : '填写这一词性的一个主要意思';
      input.value = drafts[index] || '';
      input.addEventListener('input', () => {
        currentSession().drafts[entry.id] = [...holder.querySelectorAll('input')].map(field => field.value);
        $('vocabInputError').hidden = true; persist();
      });
      label.append(input); holder.append(label);
    });
  }
  function renderFeedback() {
    const session = currentSession();
    const entry = byId.get(session.lastId);
    const record = entry && session.answers[entry.id];
    const feedback = $('vocabFeedback');
    feedback.replaceChildren(); feedback.hidden = !record;
    if (!record) return;
    feedback.classList.toggle('incorrect', !record.correct);
    feedback.classList.toggle('partial', record.partial);
    const details = element('div');
    const verdict = !record.correct ? '答错了' : record.partial ? '少答，算对' : '答对了';
    details.append(element('strong', '', '上一词 · ' + entry.word + ' · ' + verdict));
    if (state.mode === 'spelling') {
      details.append(element('p', '', '你的拼写 ' + record.answers[0]));
      details.append(element('p', 'small muted', '正确拼写'));
      const spelling = element('p', 'correct-spelling', entry.spellings.join(' / ')); markImportant(spelling, entry); details.append(spelling);
      details.append(element('p', '', entry.meaning));
    } else {
      details.append(element('p', 'small muted', '正确释义（按词性）'));
      const groups = element('ul', 'vocab-group-results');
      for (const group of record.groups) {
        groups.append(element('li', group.correct ? '' : 'missing',
          (group.correct ? '✔ ' : '✘ ') + group.label + ' 你的答案 ' + (group.answer || '未填') + ' · ' + group.meaning));
      }
      details.append(groups);
    }
    feedback.append(details);
  }
  function renderPractice() {
    const session = currentSession();
    const completed = words.filter(entry => session.answers[entry.id]);
    const correct = completed.filter(entry => session.answers[entry.id].correct).length;
    const partial = completed.filter(entry => session.answers[entry.id].partial).length;
    const entry = activeQuestion();
    $('vocabPracticeTitle').textContent = entry ? '第 ' + (words.indexOf(entry) + 1) + ' / ' + words.length + ' 词' : '本轮练习结果';
    $('vocabQuestionSource').textContent = entry ? entry.list + ' · 原表第 ' + entry.number + ' 词' : '';
    $('vocabScore').textContent = '已提交 ' + completed.length + ' 词，答对 ' + correct + ' 词' + (partial ? '，其中少答 ' + partial + ' 词' : '');
    $('vocabPrompt').textContent = entry ? state.mode === 'spelling' ? entry.meaning : entry.word : '';
    markImportant($('vocabPrompt'), entry);
    $('vocabPrompt').classList.toggle('chinese', state.mode === 'spelling');
    $('vocabAnswerForm').hidden = !entry;
    $('vocabComplete').hidden = !!entry || !words.length;
    $('vocabCompleteScore').textContent = '共 ' + words.length + ' 词，答对 ' + correct + ' 词，答错 ' + (completed.length - correct) + ' 词。';
    $('vocabInputHint').textContent = state.mode === 'spelling'
      ? '按回车提交并进入下一词。大小写不计错，按表内英美拼写判据检查。'
      : '按词性分别填写。Tab 或回车进入下一栏，最后一栏按回车提交。少答按对计。';
    renderInputs(entry);
    renderFeedback();
  }
  function renderTable() {
    const holder = $('vocabTable'); holder.replaceChildren();
    const study = state.mode === 'study';
    const bankStudy = state.scope === 'bank' && study;
    if (bankStudy) { renderBankTable(holder); return; }
    const rows = study ? words : words.filter(entry => currentSession().answers[entry.id]);
    $('vocabTableTitle').textContent = study ? '本次词汇' : '作答记录';
    $('vocabTableCount').textContent = rows.length + (study ? ' 词' : ' / ' + words.length + ' 词');
    if (!rows.length) return;
    const table = element('table', 'word-table');
    const header = element('tr'), head = element('thead');
    const titles = study
      ? ['序号', ...(state.display !== 'chinese' ? ['英文'] : []), ...(state.display !== 'english' ? ['中文'] : []), 'List']
      : ['序号', '题目', '你的答案', '正确答案', '结果'];
    titles.forEach(title => header.append(element('th', '', title))); head.append(header); table.append(head);
    const body = element('tbody');
    for (const entry of study ? rows : [...rows].reverse()) {
      const row = element('tr');
      row.append(element('td', 'col-number', entry.number));
      if (study) {
        if (state.display !== 'chinese') {
          const cell = element('td', 'col-word'), button = element('button', 'word-button', entry.word);
          markImportant(button, entry);
          button.addEventListener('click', () => { state.studyIndex = words.indexOf(entry); renderStudy(); persist(); });
          cell.append(button); row.append(cell);
        }
        if (state.display !== 'english') row.append(element('td', 'col-meaning', entry.meaning));
        row.append(element('td', 'col-chapter', entry.list));
      } else {
        const record = currentSession().answers[entry.id];
        const prompt = element('td', 'col-word', state.mode === 'spelling' ? entry.meaning : entry.word); markImportant(prompt, entry);
        const correct = element('td', 'vocab-correct-answer'); addCorrectAnswer(correct, entry, state.mode);
        row.append(prompt, element('td', '', answerText(entry, state.mode, record.answers)), correct,
          element('td', '', record.partial ? '少答，算对' : record.correct ? '正确' : '错误'));
      }
      body.append(row);
    }
    table.append(body); holder.append(table);
  }
  function renderBankTable(holder) {
    const issues = bankRecords();
    $('vocabTableTitle').textContent = '词汇错题总库'; $('vocabTableCount').textContent = issues.length + ' 条';
    if (!issues.length) { holder.append(element('p', 'dictation-record-empty', '当前分类没有记录。可以切换“只看错题／只看少答／都看”。')); return; }
    const table = element('table', 'word-table vocab-bank-table'), head = element('thead'), tr = element('tr');
    ['类型', '来源', '英文', '当时作答', '正确答案', '分类', '操作'].forEach(title => tr.append(element('th', '', title))); head.append(tr); table.append(head);
    const body = element('tbody');
    for (const issue of issues) {
      const entry = byId.get(issue.entryId), row = element('tr'); row.dataset.issue = issue.mode + ':' + issue.entryId;
      const word = element('td', 'col-word', entry.word); markImportant(word, entry);
      const correct = element('td', 'vocab-correct-answer'); addCorrectAnswer(correct, entry, issue.mode);
      const actions = element('td'), remove = element('button', 'text-button', '移出');
      remove.setAttribute('aria-label', '移出' + (issue.mode === 'spelling' ? '拼写' : '释义') + '错题 ' + entry.word);
      remove.addEventListener('click', () => { bank.remove(entry.id, issue.mode); refreshErrorSets(); render(); persist(); });
      actions.append(remove);
      row.append(element('td', '', issue.mode === 'spelling' ? '拼写' : '释义'), element('td', 'col-chapter', entry.list), word,
        element('td', '', issue.answers.length ? answerText(entry, issue.mode, issue.answers) : '未保留当时作答'), correct,
        element('td', '', issue.status === 'partial' ? '少答，算对' : issue.status === 'unknown' ? '旧记录，待核对' : '错题'), actions);
      body.append(row);
    }
    table.append(body); holder.append(table);
  }
  function render() {
    words = selectedWords();
    $('publicVocabCount').textContent = lists.length + ' 个词单 · ' + entries.length + ' 词';
    const inBank = state.scope === 'bank';
    $('vocabAreaTitle').textContent = inBank ? '词汇错题总库' : '词汇练习';
    $('vocabScope').textContent = inBank ? '跨全部 List · 当前 ' + bankRecords().length + ' 条记录' : '已选 ' + state.selected.length + ' 个 List，共 ' + words.length + ' 词';
    $('vocabListOptions').hidden = inBank; $('vocabBankHint').hidden = !inBank; $('vocabBankFilters').hidden = !inBank;
    $('vocabListScope').setAttribute('aria-pressed', !inBank); $('vocabBankScope').setAttribute('aria-pressed', inBank);
    $('vocabBankTotal').textContent = bank.list().length;
    for (const [id, status, title] of [['vocabBankAll', 'all', '都看'], ['vocabBankWrong', 'wrong', '只看错题'], ['vocabBankPartial', 'partial', '只看少答']]) {
      $(id).setAttribute('aria-pressed', state.bankStatus === status); $(id).textContent = title + '（' + bankRecords(status).length + '）';
    }
    $('vocabBankKind').value = state.bankKind; $('vocabBankList').value = state.bankList;
    $('vocabOnlyMistakes').checked = !!state.onlyMistakes;
    $('vocabDisplay').value = state.display;
    $('vocabSearch').value = inBank ? state.bankQuery : state.search;
    for (const [mode, id] of [['study', 'vocabStudyMode'], ['spelling', 'vocabSpellingMode'], ['meaning', 'vocabMeaningMode']]) {
      $(id).setAttribute('aria-pressed', mode === state.mode);
    }
    const study = state.mode === 'study';
    $('vocabDisplayLabel').hidden = !study;
    $('vocabStudyPanel').hidden = !study || !words.length;
    $('vocabPracticePanel').hidden = study || !words.length;
    $('vocabEmpty').hidden = !!words.length;
    renderLists();
    if (study) renderStudy(); else renderPractice();
    renderTable();
  }
  function submit(event) {
    event.preventDefault();
    if (state.mode === 'study' || composing) return;
    const entry = activeQuestion();
    if (!entry) return;
    const answers = [...$('vocabInputs').querySelectorAll('input')].map(input => input.value);
    const grade = C.grade(entry, state.mode, answers);
    if (grade.empty) {
      $('vocabInputError').textContent = '请先填写答案，再提交。';
      $('vocabInputError').hidden = false;
      $('vocabInputs').querySelector('input')?.focus(); return;
    }
    const session = currentSession(), now = new Date();
    const round = session.rounds[entry.list];
    session.answers[entry.id] = { ...grade, round, sequence: now.getTime() };
    session.lastId = entry.id;
    delete session.drafts[entry.id];
    const time = new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString().slice(0, 19).replace('T', ' ');
    const record = { id: 'vocab-' + now.getTime() + '-' + uuid(), mode: state.mode, entryId: entry.id, list: entry.list, row: entry.row, word: entry.word, answers: grade.answers, round, feedback: grade.feedback, status: grade.status, time, practice: state.scope === 'bank' ? 'bank' : 'lists' };
    bank.capture({ ...record, at: now.getTime() }, true); refreshErrorSets(); sync.record(record);
    render(); persist();
    if (activeQuestion()) $('vocabInputs').querySelector('input')?.focus({ preventScroll: true });
    else $('vocabComplete').focus({ preventScroll: true });
  }
  function restart() {
    if (state.mode === 'study') return;
    const session = currentSession();
    const targets = state.scope === 'bank' ? words : entries.filter(item => state.selected.includes(item.list));
    for (const name of new Set(targets.map(item => item.list))) session.rounds[name] = { id: 'round-' + state.mode + '-' + uuid(), native: false };
    for (const entry of targets) { delete session.answers[entry.id]; delete session.drafts[entry.id]; }
    session.lastId = '';
    $('vocabInputs').dataset.question = '';
    render(); persist();
    $('vocabInputs').querySelector('input')?.focus({ preventScroll: true });
  }
  function renderSync(status) {
    const receipt = status.writebackReceipt;
    const targetText = receipt?.targets?.slice(0, 3).map(target => target.list + ' · ' + (target.mode === 'spelling' ? '默写' : '释义') + target.round + '（' + target.column + ' 列）· ' + target.ranges.slice(0, 3).join('、')).join('；');
    const receiptText = targetText ? '最近确认写入：' + targetText + (receipt.repairedCells ? '；已补写 ' + receipt.repairedCells + ' 个原表空格' : '') + (receipt.preservedCells?.length ? '；保留 ' + receipt.preservedCells.length + ' 个后来手改的单元格' : '')
      : status.connected ? '已连接这份工作簿；有新作答写入后会显示具体 List、轮次和单元格。' : '';
    $('vocabWritebackReceipt').textContent = receiptText;
    $('vocabDialogWritebackReceipt').textContent = receiptText;
    window.DictationLocalFiles?.showLocation('vocabWorkbookLocation', status);
    $('vocabUseLocalWorkbook').hidden = !status.localAvailable || status.connectionKind === 'local';
    $('vocabUseLocalWorkbook').disabled = status.busy;
    $('vocabConnectionHelp').textContent = status.connectionKind === 'local'
      ? '已自动连接页面显示地址的同文件夹词汇表。请在 WPS 或 Excel 中关闭它，练习记录会自动分批同步。'
      : '在 WPS 或 Excel 中关闭该表，再选择《雅思词汇练习.xlsx》并授权。练习记录会自动分批同步。';
    $('vocabSyncBrief').textContent = !status.loaded ? '正在读取同步进度'
      : !status.connected ? status.error || '词汇表尚未连接 · 点击右上角 Excel 同步'
      : status.error ? '同步未完成 · 打开 Excel 同步查看原因'
      : status.busy ? '正在写入词汇表…'
      : status.pending ? status.connected + ' · 待同步 ' + status.pending + ' 条'
      : status.connected + ' · 已同步';
    $('vocabSyncStatus').textContent = !status.supported ? '请用 Microsoft Edge 或 Chrome 连接工作簿。'
      : !status.loaded ? '正在读取词汇练习存档。'
      : status.busy ? '正在同步词汇工作簿。'
      : status.connected ? '已连接 ' + status.connected + '。已同步 ' + status.synced + ' 条，待同步 ' + status.pending + ' 条。'
      : '尚未连接词汇工作簿，待同步 ' + status.pending + ' 条。';
    $('vocabSyncError').textContent = status.error || (status.connected && !status.writable ? '请点击“立即同步”重新授权。' : '');
    $('vocabSyncError').hidden = !$('vocabSyncError').textContent;
    $('vocabPickWorkbook').disabled = !status.supported || !status.loaded || status.busy;
    $('vocabPickWorkbook').textContent = status.connected ? '更换词汇工作簿' : '选择词汇工作簿';
    $('vocabSyncNow').hidden = !status.connected;
    $('vocabSyncNow').disabled = status.busy || !status.supported;
    if (!seed.syncCompatible) {
      $('vocabPickWorkbook').hidden = true; $('vocabSyncNow').hidden = true; $('vocabUseLocalWorkbook').hidden = true;
      $('vocabSyncStatus').textContent = '当前词单保存在本机。连接原表需先导入兼容的完整 List 工作簿。';
      $('vocabConnectionHelp').textContent = '普通词表可以学习、拼写和释义练习。原列写回只支持包含对应汇总页的完整 List 工作簿。';
      $('vocabSyncBrief').textContent = entries.length ? '当前词单及作答保存在本机浏览器' : '尚未导入词库';
    }
  }
  function showArea(next) {
    area = next;
    window.DictationAudio?.stop();
    document.body.dataset.practiceArea = area;
    $('dictationWorkspace').hidden = area !== 'dictation';
    $('vocabularyWorkspace').hidden = area !== 'vocabulary';
    $('areaDictation').setAttribute('aria-pressed', area === 'dictation');
    $('areaVocabulary').setAttribute('aria-pressed', area === 'vocabulary');
    $('openImport').hidden = area === 'vocabulary';
    $('backup').hidden = area === 'vocabulary';
    try { localStorage.setItem(areaKey, area); } catch (_) {}
  }
  $('areaDictation').addEventListener('click', () => showArea('dictation'));
  $('areaVocabulary').addEventListener('click', () => showArea('vocabulary'));
  $('vocabListScope').addEventListener('click', () => { state.scope = 'lists'; state.studyIndex = 0; render(); persist(); });
  $('vocabBankScope').addEventListener('click', () => { state.scope = 'bank'; state.mode = 'study'; state.studyIndex = 0; render(); persist(); });
  for (const [id, status] of [['vocabBankAll', 'all'], ['vocabBankWrong', 'wrong'], ['vocabBankPartial', 'partial']]) $(id).addEventListener('click', () => { state.bankStatus = status; state.mode = 'study'; render(); persist(); });
  $('vocabBankKind').addEventListener('change', event => { state.bankKind = event.target.value; state.mode = 'study'; render(); persist(); });
  $('vocabBankList').addEventListener('change', event => { state.bankList = event.target.value; state.mode = 'study'; render(); persist(); });
  for (const [mode, id] of [['study', 'vocabStudyMode'], ['spelling', 'vocabSpellingMode'], ['meaning', 'vocabMeaningMode']]) {
    $(id).addEventListener('click', () => { state.mode = mode; if (state.scope === 'bank' && mode !== 'study' && state.bankKind !== 'all') state.bankKind = mode; render(); persist(); });
  }
  $('vocabDisplay').addEventListener('change', event => { state.display = event.target.value; renderStudy(); renderTable(); persist(); });
  $('vocabOnlyMistakes').addEventListener('change', event => { state.onlyMistakes = event.target.checked; render(); persist(); });
  $('vocabSearch').addEventListener('input', event => { if (state.scope === 'bank') state.bankQuery = event.target.value; else state.search = event.target.value; state.studyIndex = 0; render(); persist(); });
  $('vocabSelectAll').addEventListener('click', () => { state.selected = [...lists]; render(); persist(); });
  $('vocabSelectNone').addEventListener('click', () => { state.selected = []; render(); persist(); });
  $('vocabPrevious').addEventListener('click', () => { state.studyIndex--; renderStudy(); persist(); });
  $('vocabNext').addEventListener('click', () => { state.studyIndex++; renderStudy(); persist(); });
  $('vocabPronounce').addEventListener('click', () => { const entry = studyItem(); if (entry) window.DictationAudio?.read(entry.word); });
  $('vocabAnswerForm').addEventListener('submit', submit);
  $('vocabAnswerForm').addEventListener('compositionstart', () => { composing = true; });
  $('vocabAnswerForm').addEventListener('compositionend', () => { composing = false; });
  $('vocabAnswerForm').addEventListener('keydown', event => {
    if (event.key !== 'Enter' || event.target.tagName !== 'INPUT') return;
    if (event.isComposing || composing || event.repeat) { event.preventDefault(); return; }
    event.preventDefault();
    const fields = [...$('vocabInputs').querySelectorAll('input')];
    const index = fields.indexOf(event.target);
    if (index < fields.length - 1) fields[index + 1].focus();
    else $('vocabAnswerForm').requestSubmit();
  });
  $('vocabRestart').addEventListener('click', restart);
  $('vocabNextRound').addEventListener('click', restart);
  $('openExcelSync').addEventListener('click', () => { if (area === 'vocabulary') $('vocabSyncDialog').showModal(); });
  $('vocabPickWorkbook').addEventListener('click', () => sync.pick());
  $('vocabSyncNow').addEventListener('click', () => sync.sync());
  $('vocabUseLocalWorkbook').addEventListener('click', () => sync.connectLocal());
  $('vocabSyncDialog').addEventListener('click', event => {
    const rect = event.currentTarget.getBoundingClientRect();
    if (event.target === event.currentTarget && (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom)) event.currentTarget.close();
  });
  window.addEventListener('beforeunload', persist);
  render();
  sync.start(history);
  sync.getRecords().then(records => { bank.reconcile(records); refreshErrorSets(); render(); persist(); }).catch(() => {});
  try { if (localStorage.getItem(areaKey) === 'vocabulary') showArea('vocabulary'); } catch (_) {}
})();
