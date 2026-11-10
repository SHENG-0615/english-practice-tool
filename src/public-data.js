(function () {
  'use strict';
  const C = window.DictationCore, V = window.VocabularyCore, $ = id => document.getElementById(id);
  const STORAGE = 'english-public-datasets-v1';
  const MAX_WORDS = 50000;
  const pos = /^(?:n|v|vt|vi|adj|adv|prep|conj|pron|num|art|aux|int)\./;
  const text = value => String(value ?? '').trim();
  const value = (sheet, row, column) => text(sheet?.[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })]?.v);
  const emptyDictation = () => ({ version: 1, title: '听写练习', entries: [], ordered: [] });
  const emptyVocabulary = () => ({ version: 1, entries: [], syncCompatible: false });
  let saved = {}, preview = null, ticket = 0;
  try { saved = JSON.parse(localStorage.getItem(STORAGE)) || {}; } catch (_) {}
  function validSeed(item, type) {
    const seed = item?.seed;
    if (!/^[a-f0-9]{24}$/.test(item?.fingerprint || '') || !Array.isArray(seed?.entries) || seed.entries.length > MAX_WORDS) return false;
    if (type === 'dictation') return Array.isArray(seed.ordered) && seed.ordered.length <= MAX_WORDS && [...seed.entries, ...seed.ordered].every(e => C.validWord(e.word));
    return seed.entries.every(e => C.validWord(e.word) && Array.isArray(e.groups) && e.groups.length && e.prior && Array.isArray(e.prior.spelling) && Array.isArray(e.prior.meaning));
  }
  for (const type of ['dictation', 'vocabulary']) if (!validSeed(saved[type], type)) delete saved[type];
  $('seedData').textContent = JSON.stringify(saved.dictation?.seed || emptyDictation());
  $('vocabularyData').textContent = JSON.stringify(saved.vocabulary?.seed || emptyVocabulary());
  const d = saved.dictation?.fingerprint || 'empty', v = saved.vocabulary?.fingerprint || 'empty';
  const keys = {
    dictation: 'english-public-dictation:' + d, dictationSync: 'english-public-dictation-sync:' + d,
    dictationJournal: 'english-public-dictation-pending:' + d + ':',
    vocabulary: 'english-public-vocabulary:' + v, vocabularySync: 'english-public-vocabulary-sync:' + v,
    vocabularyJournal: 'english-public-vocabulary-pending:' + v + ':'
  };
  function canonicalGroups(raw) {
    const merged = [];
    for (const group of raw) {
      if (!pos.test(group.meaning) && merged.length) {
        merged.at(-1).meaning += group.meaning; merged.at(-1).accepted.push(...group.accepted);
      } else merged.push({ ...group, accepted: [...group.accepted] });
    }
    const result = new Map();
    for (const group of merged) {
      const match = /^((?:(?:n|v|vt|vi|adj|adv|prep|conj|pron|num|art|aux|int)\.)[/.a-z]*)\s*(.*)/.exec(group.meaning);
      const labels = match ? match[1].match(/(?:n|v|vt|vi|adj|adv|prep|conj|pron|num|art|aux|int)\./g) : ['释义'];
      const meaning = match ? match[2] : group.meaning;
      const canonical = meaning.replace(/（[^）]*）|\([^)]*\)|\[[^\]]*\]|【[^】]*】/g, '');
      const supplements = canonical.split(/[；，、,;]/).map(text).filter(token => /^[\u3400-\u9fff]+$/.test(token));
      const accepted = [...new Set([...group.accepted, ...supplements].map(text).filter(Boolean))].slice(0, 64);
      if (!accepted.length && meaning) accepted.push(meaning);
      for (const label of labels) {
        const previous = result.get(label);
        result.set(label, {
          label, meaning: previous ? previous.meaning + '；' + meaning : label + ' ' + meaning,
          accepted: [...new Set([...(previous?.accepted || []), ...accepted])]
        });
      }
    }
    return [...result.values()];
  }
  function meaningGroups(meaning) {
    const kind = '(?:n|v|vt|vi|adj|adv|prep|conj|pron|num|art|aux|int)';
    const heads = [...meaning.matchAll(new RegExp('\\b' + kind + '\\.(?:/' + kind + '\\.)*\\s*', 'g'))];
    if (!heads.length) return [{ meaning, accepted: [] }];
    return heads.map((head, index) => ({ meaning: meaning.slice(head.index, heads[index + 1]?.index).replace(/[；;｜|\s]+$/, ''), accepted: [] }));
  }
  function parseVocabulary(book, resume, important) {
    const lists = book.SheetNames.filter(name => /^List\d{2}$/i.test(name));
    if (!lists.length) return null;
    const entries = [], freshLists = { spelling: [], meaning: [] };
    for (const name of lists) {
      const sheet = book.Sheets[name], end = Math.min(XLSX.utils.decode_range(sheet['!ref'] || 'A1').e.r + 1, MAX_WORDS + 2);
      const occupied = { spelling: false, meaning: false };
      for (let row = 3; row <= end; row++) {
        const word = value(sheet, row, 2), meaning = value(sheet, row, 4);
        if (!C.validWord(word) || !meaning) continue;
        const rawGroups = [];
        for (let group = 0; group < 4; group++) {
          const meaning = value(sheet, row, 36 + group);
          if (meaning) rawGroups.push({ meaning, accepted: Array.from({ length: 8 }, (_, i) => value(sheet, row, 40 + group * 8 + i)).filter(Boolean) });
        }
        const groups = canonicalGroups(rawGroups.length ? rawGroups : meaningGroups(meaning));
        const prior = { spelling: [], meaning: [] };
        for (const [mode, start, count] of [['spelling', 5, 10], ['meaning', 25, 5]]) {
          for (let index = 0; index < count; index++) {
            const answer = value(sheet, row, start + index * 2);
            if (answer) { occupied[mode] = true; if (resume) prior[mode].push({ index, answer }); }
          }
        }
        entries.push({ important: important.has(name + ':' + row), id: name + ':' + row, list: name, row, number: row - 2, word, phonetic: value(sheet, row, 3), meaning, spellings: (value(sheet, row, 35) || word).split('|'), groups, prior });
      }
      if (!resume) for (const mode of ['spelling', 'meaning']) if (occupied[mode]) freshLists[mode].push(name);
    }
    if (!entries.length) return null;
    return { version: 1, entries, freshLists, syncCompatible: ['List01', 'List50', '错词总表', '释义不确定'].every(name => book.Sheets[name]) && entries.every(e => e.groups.length <= 4) };
  }
  function parseDictation(book, filename) {
    const ordered = [];
    for (const name of book.SheetNames) {
      const sheet = book.Sheets[name];
      if (!/^听写\s*1$/.test(value(sheet, 2, 5))) continue;
      const end = Math.min(XLSX.utils.decode_range(sheet['!ref'] || 'A1').e.r + 1, MAX_WORDS + 2);
      for (let row = 3; row <= end; row++) {
        const spellings = value(sheet, row, 2).split('=').map(text).filter(C.validWord);
        const number = Number(value(sheet, row, 1));
        if (!spellings.length || !Number.isInteger(number) || number < 1) continue;
        ordered.push({ word: spellings[0], alternatives: spellings.slice(1), meaning: value(sheet, row, 4), chapter: name, number, source: '兼容听写表', origin: '本地导入' });
      }
    }
    if (!ordered.length) return null;
    const mistakes = C.parseWorkbook(book, XLSX, { filename, source: '兼容听写表' }).entries;
    return { version: 1, title: '听写练习', ordered, entries: mistakes };
  }
  function fromWords(words) {
    const counts = new Map();
    const ordered = words.map(item => {
      const key = C.groupKey(item), number = (counts.get(key) || 0) + 1; counts.set(key, number);
      return { ...item, number };
    });
    const vocabulary = words.filter(item => item.meaning).map((item, index) => ({
      id: 'local:' + item.id, list: item.chapter, row: index + 3, number: index + 1, word: item.word, meaning: item.meaning,
      phonetic: '', spellings: [item.word, ...(item.alternatives || [])],
      groups: canonicalGroups(meaningGroups(item.meaning)), prior: { spelling: [], meaning: [] }
    }));
    return { dictation: { version: 1, title: '听写练习', ordered, entries: [] }, vocabulary: { version: 1, entries: vocabulary, syncCompatible: false } };
  }
  async function fingerprint(seed, type) {
    const data = type === 'dictation' ? seed.ordered.map(e => [e.word, e.meaning, e.source, e.chapter, e.number, e.alternatives || []])
      : seed.entries.map(e => [e.word, e.meaning, e.list, e.row, e.spellings, e.groups]);
    const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify(data)));
    return [...new Uint8Array(digest)].map(n => n.toString(16).padStart(2, '0')).join('').slice(0, 24);
  }
  async function prepare() {
    const currentTicket = ++ticket;
    preview = null; $('publicImportCommit').disabled = true; $('publicImportError').hidden = true;
    const file = $('publicImportFile').files[0]; if (!file) return;
    $('publicImportSummary').textContent = '正在本机读取文件…';
    try {
      if (file.size > 64 * 1024 * 1024) throw new Error('文件超过 64 MB，请分成较小词表。');
      const extension = file.name.split('.').pop().toLowerCase();
      let result = {}, automatic;
      if (extension === 'xlsx') {
        const bytes = await file.arrayBuffer();
        const book = XLSX.read(bytes, { type: 'array', cellFormula: false });
        const important = await window.WorkbookImportance.read(bytes);
        const vocabulary = parseVocabulary(book, $('publicResumeImported').checked, important);
        const dictation = parseDictation(book, file.name);
        if (vocabulary) { result = { vocabulary, dictation: fromWords(vocabulary.entries.map(e => C.normalizeEntry({ word: e.word, meaning: e.meaning, chapter: e.list, source: '本地词库' }))).dictation }; automatic = 'vocabulary'; }
        else if (dictation) { result = { dictation, vocabulary: fromWords(dictation.ordered.map(e => C.normalizeEntry(e))).vocabulary }; automatic = 'dictation'; }
        else result = fromWords(C.parseWorkbook(book, XLSX, { filename: file.name, source: '本地词库' }).entries);
      } else {
        let words;
        if (extension === 'json') {
          const raw = JSON.parse(await file.text());
          words = C.mergeEntries([], Array.isArray(raw) ? raw : raw.entries || []).entries;
        } else if (['csv', 'tsv', 'txt'].includes(extension)) words = C.parseText(await file.text(), { source: '本地词库', chapter: '本地词单' });
        else throw new Error('请选择 XLSX、CSV、TSV、TXT 或 JSON 词表。');
        result = fromWords(words);
      }
      let target = $('publicImportTarget').value;
      if (target === 'auto') target = automatic || 'both';
      if (target !== 'both') result = { [target]: result[target] };
      for (const [type, seed] of Object.entries(result)) {
        const count = type === 'dictation' ? seed?.ordered?.length : seed?.entries?.length;
        if (!count) throw new Error(type === 'vocabulary' ? '词汇练习需要英文和中文释义，当前文件未识别到完整词条。' : '当前文件没有识别到英文词汇。');
        if (count > MAX_WORDS) throw new Error('词库超过 50,000 条，请分开导入。');
      }
      const prepared = {};
      for (const [type, seed] of Object.entries(result)) prepared[type] = { seed, fingerprint: await fingerprint(seed, type) };
      if (currentTicket !== ticket) return;
      preview = prepared;
      $('publicImportSummary').textContent = Object.entries(prepared).map(([type, item]) => (type === 'dictation' ? '听写：' + item.seed.ordered.length : '词汇：' + item.seed.entries.length) + ' 词').join('；') + '。原有词库的本机存档会保留在独立存档中。';
      $('publicImportCommit').disabled = false;
    } catch (error) {
      if (currentTicket !== ticket) return;
      $('publicImportSummary').textContent = '未导入，请检查文件后重试。';
      $('publicImportError').textContent = error.message; $('publicImportError').hidden = false;
    }
  }
  function open() { $('publicImportDialog').showModal(); }
  $('publicOpenImport').addEventListener('click', open);
  for (const id of ['publicImportFile', 'publicImportTarget', 'publicResumeImported']) $(id).addEventListener('change', prepare);
  $('publicImportCommit').addEventListener('click', () => {
    if (!preview) return;
    try { localStorage.setItem(STORAGE, JSON.stringify({ ...saved, ...preview })); }
    catch (_) { $('publicImportError').textContent = '浏览器未能保存词库，可能超过本机存储容量。请分成较小词单。'; $('publicImportError').hidden = false; return; }
    location.reload();
  });
  const dCount = saved.dictation?.seed?.ordered?.length || 0, vCount = saved.vocabulary?.seed?.entries?.length || 0;
  $('publicDatasetStatus').textContent = dCount || vCount ? '本机词库 · 听写 ' + dCount + ' 词，词汇 ' + vCount + ' 词' : '空词库 · 词表和练习记录只保存在本机';
  window.EnglishPublicData = { keys, open, canonicalGroups };
})();
