(function () {
  'use strict';
  const X = window.DictationXlsxLog.xml, C = window.VocabularyCore;
  const { add, list, first, readXml, writeXml, filePart } = X;
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const LOG = 'HTML 词汇记录';
  const HEADERS = ['记录ID', '时间', '模式', 'List', '原表行', '英文', '分栏答案', '判定', '轮次', '原表单元格', '状态', '来源'];
  function col(n) { let s = ''; while (n) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); } return s; }
  function num(s) { return [...s].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0); }
  const literal = s => '"' + String(s).replace(/"/g, '""') + '"';
  function table(doc, strings) {
    const data = first(doc, 'sheetData');
    const rows = new Map(list(data, 'row').map(row => [Number(row.getAttribute('r')), row]));
    const cells = new Map(list(data, 'c').map(cell => [cell.getAttribute('r'), cell]));
    let dirty = false;
    function get(address) {
      const c = cells.get(address); if (!c) return '';
      if (c.getAttribute('t') === 's') return strings[Number(first(c, 'v')?.textContent)] || '';
      if (c.getAttribute('t') === 'inlineStr') return list(c, 't').map(run => run.textContent).join('');
      const v = first(c, 'v')?.textContent || '';
      return ['str', 'e'].includes(c.getAttribute('t')) ? v : v === '' ? '' : Number(v);
    }
    function cell(address) {
      if (cells.has(address)) return cells.get(address);
      const r = Number(address.match(/\d+$/)[0]);
      let row = rows.get(r);
      if (!row) {
        row = add(data, 'row', { r }); rows.set(r, row);
        const after = [...data.children].find(node => Number(node.getAttribute('r')) > r);
        if (after) data.insertBefore(row, after);
      }
      const node = add(row, 'c', { r: address });
      const n = num(address.replace(/\d+$/, ''));
      const after = [...row.children].find(c => num(c.getAttribute('r').replace(/\d+$/, '')) > n);
      if (after) row.insertBefore(node, after);
      cells.set(address, node); return node;
    }
    function set(address, value, formula, cacheOnly = false) {
      const c = cell(address); dirty = true;
      const savedFormula = cacheOnly ? first(c, 'f')?.cloneNode(true) : null;
      const f = savedFormula ? savedFormula.textContent : formula;
      c.replaceChildren(); c.removeAttribute('t');
      if (savedFormula) c.append(savedFormula);
      else if (f != null) add(c, 'f').textContent = f;
      if (typeof value === 'number' || typeof value === 'boolean') {
        if (typeof value === 'boolean') c.setAttribute('t', 'b');
        add(c, 'v').textContent = typeof value === 'boolean' ? Number(value) : value;
      } else if (f != null) {
        c.setAttribute('t', 'str'); add(c, 'v').textContent = value ?? '';
      } else if (value != null && value !== '') {
        c.setAttribute('t', 'inlineStr');
        const text = add(add(c, 'is'), 't');
        text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        text.textContent = String(value).toWellFormed().replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '');
      }
    }
    return { doc, get, set, cache: (address, value) => set(address, value, undefined, true), rows, get dirty() { return dirty; } };
  }
  function register(zip, workbook, rels, types) {
    const sheets = list(workbook, 'sheet');
    const id = Math.max(...sheets.map(s => Number(s.getAttribute('sheetId')))) + 1;
    const used = new Set(list(rels, 'Relationship').map(r => r.getAttribute('Id')));
    let n = 1; while (used.has('rId' + n)) n++;
    let part = id; while (zip.file('xl/worksheets/sheet' + part + '.xml')) part++;
    const target = 'worksheets/sheet' + part + '.xml';
    const sheet = add(first(workbook, 'sheets'), 'sheet', { name: LOG, sheetId: id });
    sheet.setAttributeNS(REL, 'r:id', 'rId' + n);
    add(rels.documentElement, 'Relationship', { Id: 'rId' + n, Type: REL + '/worksheet', Target: target });
    add(types.documentElement, 'Override', { PartName: '/xl/' + target, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' });
    const doc = document.implementation.createDocument(NS, 'worksheet');
    add(doc.documentElement, 'dimension', { ref: 'A1:L1' });
    const view = add(add(doc.documentElement, 'sheetViews'), 'sheetView', { workbookViewId: 0 });
    add(view, 'pane', { ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' });
    const columns = add(doc.documentElement, 'cols');
    [24, 22, 12, 12, 10, 24, 46, 44, 24, 16, 10, 14].forEach((width, i) => add(columns, 'col', { min: i + 1, max: i + 1, width, customWidth: 1 }));
    add(doc.documentElement, 'sheetData');
    add(doc.documentElement, 'autoFilter', { ref: 'A1:L1' });
    return { path: 'xl/' + target, doc };
  }
  function meaningFormulas(t, entry) {
    const r = entry.row;
    for (let g = 0; g < 4; g++) {
      const group = entry.groups[g];
      t.set(col(36 + g) + r, group?.meaning || '');
      for (let k = 0; k < 8; k++) t.set(col(40 + g * 8 + k) + r, group?.accepted[k] ? C.normalizeMeaning(group.accepted[k]) : '');
    }
    for (let round = 0; round < 5; round++) {
      const input = '$' + col(25 + round * 2) + r;
      const old = String(t.get(input.slice(1)) || '');
      const answers = C.splitExistingMeaning(entry, old);
      const grade = C.gradeMeaning(entry, answers);
      const flags = [];
      for (let g = 0; g < 4; g++) {
        const group = entry.groups[g];
        const flag = '$' + col(72 + round * 4 + g) + r;
        flags.push(flag);
        if (!group) { t.set(flag.slice(1), true, 'TRUE'); continue; }
        // Three hidden cells retain the per-POS text, its normalized form and its polarity.
        const raw = '$' + col(103 + round * 12 + g * 3) + r;
        const norm = '$' + col(104 + round * 12 + g * 3) + r;
        const neg = '$' + col(105 + round * 12 + g * 3) + r;
        const delimiter = literal(' ｜ ' + group.label + ' ');
        const wrapped = '" ｜ "&' + input + '&" ｜ "';
        const start = 'FIND(' + delimiter + ',' + wrapped + ')+LEN(' + delimiter + ')';
        const extraction = 'IFERROR(MID(' + wrapped + ',' + start + ',FIND(" ｜ ",' + wrapped + ',' + start + ')-(' + start + ')),IF(ISNUMBER(FIND(" ｜ ",' + input + ')),"",' + input + '))';
        let clean = raw;
        for (const char of [' ', '　', '，', '。', '；', '、', '！', '!']) clean = 'SUBSTITUTE(' + clean + ',' + literal(char) + ',"")';
        t.set(raw.slice(1), C.trim(answers[g] || '').replace(/[，。；、！!\s]/g, ''), clean.replaceAll(raw, extraction));
        const degree = 'IF(OR(LEFT(' + raw + ',2)="非常",LEFT(' + raw + ',2)="十分",LEFT(' + raw + ',2)="比较",LEFT(' + raw + ',2)="较为"),MID(' + raw + ',3,999),IF(LEFT(' + raw + ',1)="很",MID(' + raw + ',2,999),' + raw + '))';
        const normalized = 'IF(AND(LEN(' + degree + ')>2,RIGHT(' + degree + ',1)="的"),LEFT(' + degree + ',LEN(' + degree + ')-1),' + degree + ')';
        const normalizedValue = C.normalizeMeaning(answers[g] || '');
        t.set(norm.slice(1), normalizedValue, normalized);
        const polarity = 'OR(ISNUMBER(FIND(LEFT(' + norm + ',1),"不没未无")),AND(LEFT(' + norm + ',1)="并",ISNUMBER(FIND(MID(' + norm + ',2,1),"不没未无"))),AND(OR(LEFT(' + norm + ',2)="完全",LEFT(' + norm + ',2)="根本",LEFT(' + norm + ',2)="从来"),ISNUMBER(FIND(MID(' + norm + ',3,1),"不没未无"))))';
        t.set(neg.slice(1), /^(?:并|完全|根本|从来)?[不没未无]/.test(normalizedValue), 'AND(' + norm + '<>"",' + polarity + ')');
        const tests = [...new Set(group.accepted.map(C.normalizeMeaning))].filter(Boolean).map(token => {
          const expected = literal(token), negative = /^(?:并|完全|根本|从来)?[不没未无]/.test(token);
          return 'AND(' + neg + '=' + (negative ? 'TRUE' : 'FALSE') + ',OR(ISNUMBER(FIND(' + expected + ',' + norm + ')),AND(LEN(' + norm + ')>=2,LEN(' + norm + ')>=LEN(' + expected + ')-1,ISNUMBER(FIND(' + norm + ',' + expected + ')))))';
        });
        const formula = 'AND(' + norm + '<>"",OR(' + tests.join(',') + '))';
        if (formula.length > 8192) throw new Error(entry.word + ' 的释义判据过长，尚未写入。');
        t.set(flag.slice(1), !!grade.groups[g]?.correct, formula);
      }
      const notes = entry.groups.map((group, g) => 'IF(' + flags[g] + ',"✔ ","✘ ")&$' + col(36 + g) + r).join('&" ｜ "&');
      t.set(col(93 + round) + r, grade.empty ? '' : grade.groups.map(g => (g.correct ? '✔ ' : '✘ ') + g.meaning).join(' ｜ '), 'IF(TRIM(' + input + ')="","",' + notes + ')');
      const any = 'OR(' + flags.slice(0, entry.groups.length).join(',') + ')';
      const all = 'AND(' + flags.slice(0, entry.groups.length).join(',') + ')';
      t.set(col(26 + round * 2) + r, grade.empty ? '' : grade.feedback, 'IF(TRIM(' + input + ')="","",IF(' + any + ',$' + col(93 + round) + r + ',"✗ "&$D' + r + '))');
      t.set(col(98 + round) + r, grade.empty || grade.status === 'full' ? 0 : grade.partial ? 1 : 2, 'IF(TRIM(' + input + ')="",0,IF(' + any + ',IF(' + all + ',0,1),2))');
    }
    let columns = first(t.doc, 'cols');
    if (!columns) { columns = add(t.doc.documentElement, 'cols'); t.doc.documentElement.insertBefore(columns, first(t.doc, 'sheetData')); }
    if (!list(columns, 'col').some(c => Number(c.getAttribute('min')) === 103)) add(columns, 'col', { min: 103, max: 162, hidden: 1, width: 14, customWidth: 1 });
    const dimension = first(t.doc, 'dimension');
    if (dimension) dimension.setAttribute('ref', 'A1:FF' + Math.max(...t.rows.keys()));
  }
  function receiptFor(log, repaired, preserved) {
    const groups = new Map(); let total = 0;
    for (const r of log.rows.keys()) {
      if (r === 1 || log.get('L' + r) === '原表已有') continue;
      const address = String(log.get('J' + r)), mode = String(log.get('C' + r)), list = String(log.get('D' + r));
      if (!/^[A-Z]+\d+$/.test(address) || !['spelling', 'meaning'].includes(mode)) continue;
      total++;
      const column = address.replace(/\d+$/, ''), key = list + ':' + column;
      if (!groups.has(key)) groups.set(key, { list, mode, column, cells: [], latest: '' });
      const group = groups.get(key); group.cells.push(address); group.latest = [group.latest, String(log.get('B' + r))].sort().at(-1);
    }
    return { type: 'vocabulary', practiceRecords: total, repairedCells: repaired, preservedCells: preserved,
      targets: [...groups.values()].sort((a, b) => b.latest.localeCompare(a.latest)).map(group => {
        const rows = [...new Set(group.cells.map(cell => Number(cell.match(/\d+$/)[0])))].sort((a, b) => a - b);
        const ranges = []; let start = rows[0], end = start;
        for (const row of rows.slice(1)) { if (row === end + 1) end = row; else { ranges.push(group.column + start + (end === start ? '' : ':' + group.column + end)); start = end = row; } }
        ranges.push(group.column + start + (end === start ? '' : ':' + group.column + end));
        return { list: group.list, mode: group.mode, column: group.column, round: (num(group.column) - (group.mode === 'spelling' ? 5 : 25)) / 2 + 1, ranges, latest: group.latest };
      }) };
  }
  async function append(bytes, records, entries, onReceipt = () => {}) {
    let zip;
    try { zip = await JSZip.loadAsync(bytes); } catch (_) { throw new Error('无法打开词汇工作簿，请选择完整的未加密 XLSX。'); }
    if (zip.file(/^_xmlsignatures\//i).length) throw new Error('请选择一份未签名的工作簿副本。');
    const workbook = await readXml(zip, 'xl/workbook.xml');
    const rels = await readXml(zip, 'xl/_rels/workbook.xml.rels');
    const paths = new Map(list(workbook, 'sheet').map(s => {
      const rel = list(rels, 'Relationship').find(r => r.getAttribute('Id') === s.getAttributeNS(REL, 'id'));
      if (!rel || rel.getAttribute('TargetMode') === 'External') throw new Error('无法定位词汇表。');
      return [s.getAttribute('name'), filePart(rel.getAttribute('Target'))];
    }));
    if (!['List01', 'List50', '释义不确定', '错词总表'].every(name => paths.has(name))) throw new Error('请选择《雅思词汇练习.xlsx》，听写工作簿在另一个入口连接。');
    const shared = list(rels, 'Relationship').find(r => r.getAttribute('Type').endsWith('/sharedStrings'));
    const strings = shared ? list(await readXml(zip, filePart(shared.getAttribute('Target'))), 'si').map(si => list(si, 't').map(t => t.textContent).join('')) : [];
    const loaded = new Map();
    async function load(name) {
      if (!loaded.has(name)) loaded.set(name, table(await readXml(zip, paths.get(name)), strings));
      return loaded.get(name);
    }
    let log;
    if (paths.has(LOG)) log = await load(LOG);
    else {
      const types = await readXml(zip, '[Content_Types].xml');
      const registered = register(zip, workbook, rels, types);
      paths.set(LOG, registered.path); log = table(registered.doc, strings); loaded.set(LOG, log);
      HEADERS.forEach((header, i) => log.set(col(i + 1) + '1', header));
      writeXml(zip, 'xl/_rels/workbook.xml.rels', rels); writeXml(zip, '[Content_Types].xml', types);
    }
    if (log.get('A1') !== HEADERS[0] || log.get('J1') !== HEADERS[9]) throw new Error('HTML 词汇记录表头不同，尚未写入。');
    const known = new Set(), mapping = new Map();
    let next = Math.max(...log.rows.keys(), 1);
    for (const r of log.rows.keys()) {
      if (r === 1) continue;
      known.add(String(log.get('A' + r)));
      const key = [log.get('C' + r), log.get('D' + r), log.get('I' + r)].join('|');
      const address = String(log.get('J' + r));
      if (/^[A-Z]+\d+$/.test(address)) mapping.set(key, num(address.replace(/\d+$/, '')));
    }
    const byId = new Map(entries.map(entry => [entry.id, entry]));
    const changedLists = new Set();
    let repaired = 0; const preserved = [];
    const pending = records.filter(record => !known.has(record.id) || !record.origin).sort((a, b) => Number(!a.origin) - Number(!b.origin) || String(a.time).localeCompare(String(b.time)) || a.id.localeCompare(b.id));
    if (!pending.length) { onReceipt(receiptFor(log, repaired, preserved)); return null; }
    for (const record of pending) {
      const existingRecord = known.has(record.id);
      const entry = byId.get(record.entryId);
      if (!entry || !['spelling', 'meaning'].includes(record.mode) || !record.round?.id) throw new Error('词汇作答记录不完整，尚未写入。');
      const t = await load(entry.list);
      if (String(t.get('B' + entry.row)) !== entry.word) throw new Error(entry.list + ' 词条位置已变化，请保留原表行顺序后再同步。');
      const start = record.mode === 'spelling' ? 5 : 25, limit = record.mode === 'spelling' ? 10 : 5;
      const key = [record.mode, entry.list, record.round.id].join('|');
      let column = mapping.get(key);
      if (!column) {
        if (record.round.native && Number.isInteger(record.round.index) && record.round.index >= 0 && record.round.index < limit) column = start + record.round.index * 2;
        else {
          let occupied = -1;
          for (let i = 0; i < limit; i++) if (entries.some(e => e.list === entry.list && String(t.get(col(start + i * 2) + e.row)) !== '')) occupied = i;
          column = start + (occupied + 1) * 2;
        }
        if (column >= start + limit * 2) throw new Error(entry.list + ' 的' + (record.mode === 'spelling' ? '十轮拼写' : '五轮释义') + '已用完，请先保存记录并换一份练习副本。');
        mapping.set(key, column);
      }
      const address = col(column) + entry.row;
      const grade = C.grade(entry, record.mode, record.answers);
      if (grade.empty) throw new Error('检测到空作答，尚未写入。');
      const answer = record.mode === 'spelling' ? grade.answers[0] : entry.groups.map((g, i) => g.label + ' ' + (grade.answers[i] || '')).join(' ｜ ');
      if (!record.origin) {
        const old = String(t.get(address));
        if (old && old !== answer) {
          if (!existingRecord) throw new Error(entry.list + '!' + address + ' 已有其他答案。记录保留在本机，请先检查表格，避免覆盖。');
          preserved.push(entry.list + '!' + address); continue;
        }
        const expected = record.mode === 'meaning' ? grade.feedback : grade.correct ? 0 : entry.word + '>' + entry.meaning;
        if (!old || t.get(col(column + 1) + entry.row) !== expected) {
          t.set(address, answer);
          if (record.mode === 'meaning') meaningFormulas(t, entry);
          else t.cache(col(column + 1) + entry.row, expected);
          changedLists.add(entry.list);
          if (existingRecord && !old) repaired++;
        }
      }
      if (existingRecord) continue;
      const values = [record.id, record.time || '', record.mode, entry.list, entry.row, entry.word, JSON.stringify(grade.answers), grade.feedback, record.round.id, address, grade.status, record.origin ? '原表已有' : record.practice === 'bank' ? '错题总库作答' : '网页作答'];
      next++; values.forEach((value, i) => log.set(col(i + 1) + next, value)); known.add(record.id);
    }
    onReceipt(receiptFor(log, repaired, preserved));
    if (!log.dirty && !changedLists.size) return null;
    for (const tag of ['dimension', 'autoFilter']) first(log.doc, tag)?.setAttribute('ref', 'A1:L' + next);
    if (changedLists.size) await refreshSummaries(load, entries, changedLists);
    for (const [name, t] of loaded) if (t.dirty) writeXml(zip, paths.get(name), t.doc);
    const calc = first(workbook, 'calcPr') || add(workbook.documentElement, 'calcPr');
    calc.setAttribute('fullCalcOnLoad', '1'); calc.setAttribute('forceFullCalc', '1');
    writeXml(zip, 'xl/workbook.xml', workbook);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
  async function refreshSummaries(load, entries, names) {
    const summary = await load('错词总表'), rates = await load('正确率');
    const stats = entries.map((entry, i) => ({ entry, index: i + 1, row: i + 2, spelling: Number(summary.get('O' + (i + 2))) || 0, wrong: Number(summary.get('R' + (i + 2))) || 0, partial: Number(summary.get('S' + (i + 2))) || 0 }));
    for (const name of names) {
      const t = await load(name), selected = stats.filter(s => s.entry.list === name), count = selected.length;
      const spellErrors = Array(10).fill(0), meaningErrors = Array(5).fill(0);
      let submitted = 0;
      for (const s of selected) {
        const e = s.entry, r = e.row;
        s.spelling = 0; s.wrong = 0; s.partial = 0;
        if (String(summary.get('B' + s.row)) !== e.word) throw new Error('错词总表顺序已变化，尚未写入。');
        for (let k = 0; k < 10; k++) {
          const answer = String(t.get(col(5 + k * 2) + r));
          const incorrect = !!answer && !C.gradeSpelling(e, [answer]).correct;
          if (answer) submitted++;
          if (incorrect) { s.spelling++; spellErrors[k]++; }
          t.cache(col(6 + k * 2) + r, !answer ? '' : incorrect ? e.word + '>' + e.meaning : 0);
          summary.cache(col(5 + k) + s.row, incorrect ? answer : '');
        }
        for (let k = 0; k < 5; k++) {
          const status = Number(t.get(col(98 + k) + r));
          if (status === 2) { s.wrong++; meaningErrors[k]++; }
          if (status === 1) s.partial++;
        }
        summary.cache('O' + s.row, s.spelling); summary.cache('R' + s.row, s.wrong); summary.cache('S' + s.row, s.partial); summary.cache('T' + s.row, s.wrong + s.partial);
      }
      const rateRow = Number(name.slice(4)) + 3;
      spellErrors.forEach((errors, k) => { t.cache(col(6 + k * 2) + '1', 1 - errors / count); rates.cache(col(3 + k) + rateRow, 1 - errors / count); });
      meaningErrors.forEach((errors, k) => { t.cache(col(26 + k * 2) + '1', 1 - errors / count); rates.cache(col(14 + k) + rateRow, 1 - errors / count); });
      t.cache('D1', 1 - spellErrors.reduce((a, b) => a + b, 0) / (count * 10)); rates.cache('M' + rateRow, t.get('D1'));
      t.cache('CN1', 1 - meaningErrors.reduce((a, b) => a + b, 0) / (count * 5)); rates.cache('S' + rateRow, t.get('CN1')); rates.cache('T' + rateRow, submitted);
    }
    const ranking = key => stats.filter(s => key(s) > 0).sort((a, b) => key(b) - key(a) || a.index - b.index);
    const spelling = ranking(s => s.spelling), wrong = ranking(s => s.wrong), uncertain = ranking(s => s.wrong + s.partial);
    for (const [key, rows] of [['P', spelling], ['U', wrong], ['V', uncertain]]) {
      const positions = new Map(rows.map((s, i) => [s.index, i + 1]));
      for (const s of stats) summary.cache(key + s.row, positions.get(s.index) || '');
    }
    const rank = await load('错词排行'), practice = await load('错词练习');
    for (let i = 0; i < stats.length; i++) {
      const s = spelling[i], r = i + 3;
      for (const [key, value] of [['B', s?.entry.word], ['C', s?.entry.meaning], ['D', s?.entry.list], ['E', s?.spelling], ['F', s?.index]]) rank.cache(key + r, value ?? '');
      if (i >= 150) continue;
      for (const [key, value] of [['A', s?.entry.word], ['B', s?.entry.meaning], ['C', s?.spelling], ['G', s?.entry.spellings.join('|')]]) practice.cache(key + r, value ?? '');
      const manual = String(practice.get('D' + r));
      practice.cache('E' + r, !s || !manual ? '' : C.gradeSpelling(s.entry, [manual]).correct ? 0 : '✘ ' + s.entry.word);
    }
    for (const [name, rows] of [['释义错题', wrong], ['释义不确定', uncertain]]) {
      const t = await load(name);
      for (let i = 0; i < 200; i++) {
        const s = rows[i], r = i + 4;
        for (const [key, value] of [['B', s?.entry.word], ['C', s?.entry.meaning], ['D', s?.entry.list], ['E', s?.wrong], ['F', s?.partial], ['G', s?.index]]) t.cache(key + r, value ?? '');
      }
    }
  }
  window.VocabularyWorkbook = { append };
})();
