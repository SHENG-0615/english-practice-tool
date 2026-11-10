(function () {
  'use strict';
  const normalize = value => String(value || '').normalize('NFKC').trim().replace(/\s+/g, ' ').toLowerCase();
  const primary = word => String(word || '').split(' / ')[0];
  const escapeXml = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const column = number => {
    let result = '';
    while (number > 0) { number--; result = String.fromCharCode(65 + number % 26) + result; number = Math.floor(number / 26); }
    return result;
  };
  const columnNumber = name => [...name].reduce((number, letter) => number * 26 + letter.charCodeAt(0) - 64, 0);
  const quoteSheet = name => "'" + name.replace(/'/g, "''") + "'!";

  async function project(zip, paths, logs, strings, io) {
    if (!logs.length) return false;
    const { list, first, add, readXml, writeXml } = io;
    const chapters = new Map(), feedback = new Map();
    const roundColumns = new Map(), latestColumns = new Map(), legacyRounds = new Map();
    const practiceDoc = await readXml(zip, paths.get('错词练习'));

    function text(cell) {
      if (!cell) return '';
      const value = first(cell, 'v')?.textContent || '';
      return cell.getAttribute('t') === 's' ? strings[Number(value)] || '' : list(cell, 't').map(run => run.textContent).join('') || value;
    }
    function table(doc) {
      const data = first(doc, 'sheetData');
      return {
        doc, data, rows: new Map(list(data, 'row').map(row => [Number(row.getAttribute('r')), row])),
        cells: new Map(list(data, 'c').map(cell => [cell.getAttribute('r'), cell])),
        changed: false
      };
    }
    function cellAt(tab, address) {
      if (tab.cells.has(address)) return tab.cells.get(address);
      const rowNumber = Number(address.match(/\d+$/)[0]);
      let row = tab.rows.get(rowNumber);
      if (!row) {
        row = add(tab.data, 'row', { r: rowNumber });
        const next = [...tab.rows].find(([number]) => number > rowNumber)?.[1];
        if (next) tab.data.insertBefore(row, next);
        tab.rows.set(rowNumber, row);
      }
      const cell = add(row, 'c', { r: address });
      const number = columnNumber(address.replace(/\d+$/, ''));
      const next = [...row.children].find(other => other !== cell && columnNumber(other.getAttribute('r').replace(/\d+$/, '')) > number);
      if (next) row.insertBefore(cell, next);
      const sample = tab.cells.get(address.replace(/\d+$/, String(tab.sampleRow || 2)));
      if (sample?.hasAttribute('s')) cell.setAttribute('s', sample.getAttribute('s'));
      tab.cells.set(address, cell);
      return cell;
    }
    function cache(cell, value) {
      for (const child of [...cell.children]) if (child.localName !== 'f') child.remove();
      if (typeof value === 'number') cell.removeAttribute('t');
      else cell.setAttribute('t', 'str');
      add(cell, 'v').textContent = value;
    }
    function input(tab, address, value) {
      const cell = cellAt(tab, address);
      cell.replaceChildren();
      cell.setAttribute('t', 'inlineStr');
      const run = add(add(cell, 'is'), 't');
      run.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      run.textContent = String(value).toWellFormed().replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
      tab.changed = true;
    }
    function shiftFormula(formula, deltaRow) {
      return formula.split(/("(?:[^"]|"")*")/).map((part, index) => index % 2 ? part : part.replace(/(\$?[A-Z]+)(\$?)(\d+)/g, (ref, col, fixed, row) => fixed ? ref : col + (Number(row) + deltaRow))).join('');
    }
    function unshare(tab, cell) {
      const formula = first(cell, 'f');
      if (formula?.getAttribute('t') !== 'shared') return;
      const id = formula.getAttribute('si');
      const group = [...tab.cells.values()].filter(other => first(other, 'f')?.getAttribute('si') === id);
      const anchor = group.find(other => first(other, 'f').textContent);
      if (!anchor) throw new Error('章节里的共享公式不完整，请检查工作簿。');
      const original = first(anchor, 'f').textContent;
      const base = Number(anchor.getAttribute('r').match(/\d+$/)[0]);
      for (const other of group) {
        const node = first(other, 'f');
        node.textContent = shiftFormula(original, Number(other.getAttribute('r').match(/\d+$/)[0]) - base);
        for (const attribute of ['t', 'si', 'ref']) node.removeAttribute(attribute);
      }
    }
    async function chapter(name) {
      if (chapters.has(name)) return chapters.get(name);
      if (!paths.has(name)) return null;
      const tab = table(await readXml(zip, paths.get(name)));
      tab.sampleRow = 3;
      tab.questions = [...tab.rows].filter(([number]) => number >= 3).map(([row]) => ({
        row, number: Number(text(tab.cells.get('A' + row))),
        word: text(tab.cells.get('B' + row)), meaning: text(tab.cells.get('D' + row))
      })).filter(item => item.word && Number.isInteger(item.number) && item.number > 0);
      if (!tab.questions.length || text(tab.cells.get('E2')) !== '听写1') return null;
      chapters.set(name, tab);
      return tab;
    }
    const practice = table(practiceDoc);
    const logTable = table(logs[0].node.ownerDocument);
    const practiceWords = new Map();
    let lastPracticeRow = 1;
    for (const [row] of practice.rows) {
      const word = text(practice.cells.get('A' + row));
      if (row < 2 || !word) continue;
      practiceWords.set(text(practice.cells.get('C' + row)) + '\n' + normalize(word), row);
      lastPracticeRow = Math.max(lastPracticeRow, row);
    }
    function mistakeRow(record) {
      const key = record.chapter + '\n' + normalize(primary(record.word));
      if (practiceWords.has(key)) return practiceWords.get(key);
      const row = ++lastPracticeRow;
      input(practice, 'A' + row, primary(record.word));
      input(practice, 'B' + row, record.meaning || '');
      input(practice, 'C' + row, record.chapter || '');
      practiceWords.set(key, row);
      return row;
    }
    function mark(log, index, value) {
      const address = column(index + 1) + log.number;
      input(logTable, address, value);
      log.fields[index] = value;
    }
    let changed = false;
    for (const log of logs) {
      const f = log.fields;
      const record = { source: f[1], mode: f[2], chapter: f[3], number: Number(f[4]), word: f[5], answer: f[6], result: f[7], meaning: f[8], kind: f[9] };
      const oldMistake = record.kind === '原表错词';
      const manualMistake = record.kind === '手动加入错题';
      const practiceKey = record.chapter + '\n' + record.mode;
      if (!oldMistake && record.answer && !f[13]) {
        let legacy = legacyRounds.get(practiceKey);
        const questionKey = record.mode === '顺序练习' ? record.number : normalize(record.word);
        if (!legacy) legacy = { number: 1, questions: new Set() };
        if (legacy.questions.has(questionKey)) legacy = { number: legacy.number + 1, questions: new Set() };
        legacy.questions.add(questionKey);
        legacyRounds.set(practiceKey, legacy);
        mark(log, 13, 'legacy-' + record.mode + '-' + record.chapter + '-' + legacy.number);
        changed = true;
      }
      const roundKey = practiceKey + '\n' + f[13];
      const previousTarget = f[11]?.match(/^'((?:''|[^'])+)'!([A-Z]+)\d+$/);
      if (previousTarget) {
        const col = columnNumber(previousTarget[2]);
        roundColumns.set(roundKey, col);
        latestColumns.set(practiceKey, col);
      }
      if (!f[12] && (oldMistake || manualMistake || record.result === '答错' || record.mode === '错题练习')) {
        const row = mistakeRow(record);
        let target = 'A' + row;
        if (!oldMistake && !manualMistake && record.mode === '错题练习' && record.answer) {
          let col = 4;
          while (col <= 13 && text(practice.cells.get(column(col) + row))) col++;
          if (col > 13) throw new Error('错词练习中这个词的 10 次听写栏已满，请先在表格中整理记录。');
          target = column(col) + row;
          input(practice, target, record.answer);
        }
        mark(log, 12, quoteSheet('错词练习') + target);
        changed = true;
      }
      if (manualMistake) {
        if (!f[11]) { mark(log, 11, '手动标记不确定'); changed = true; }
        continue;
      }
      if (f[11] || oldMistake || !record.answer) continue;
      if (!['兼容听写表', '剑20'].includes(record.source)) {
        mark(log, 11, '自建词汇，保存在作答明细'); changed = true; continue;
      }
      const tab = await chapter(record.chapter);
      if (!tab) {
        mark(log, 11, '原表没有同名章节'); changed = true; continue;
      }
      const spellings = record.word.split(' / ').map(normalize);
      const matches = tab.questions.filter(item => item.word.split('=').some(word => spellings.includes(normalize(word))));
      let question = record.mode === '顺序练习' ? matches.find(item => item.number === record.number) : matches.find(item => item.meaning === record.meaning);
      if (!question && matches.length === 1 && record.mode !== '顺序练习') question = matches[0];
      if (!question) throw new Error('无法准确匹配 ' + record.chapter + ' 中的“' + record.word + '”，请检查所选表格。');
      const row = question.row;
      let col = roundColumns.get(roundKey);
      if (!col && f[13].startsWith('resume-')) col = latestColumns.get(practiceKey);
      if (!col) {
        let lastUsed = 3;
        for (let candidate = 5; candidate <= 33; candidate += 2) {
          if (tab.questions.some(item => text(tab.cells.get(column(candidate) + item.row)))) lastUsed = candidate;
        }
        col = lastUsed + 2;
      }
      if (col > 33) throw new Error(record.chapter + ' 第 ' + question.number + ' 题的 15 次听写栏已满，请先整理表格。');
      roundColumns.set(roundKey, col);
      latestColumns.set(practiceKey, col);
      const target = column(col) + row, answerCell = column(col + 1) + row;
      if (text(tab.cells.get(target))) throw new Error(record.chapter + ' 的这一轮里，第 ' + question.number + ' 题已有作答。请用“重新练本组”开始下一轮。');
      input(tab, target, record.answer.normalize('NFKC').trim().replace(/\s+/g, ' '));
      const judge = cellAt(tab, answerCell);
      unshare(tab, judge);
      judge.replaceChildren();
      const choices = spellings.map(word => 'LOWER(TRIM(' + target + '))="' + word.replace(/"/g, '""') + '"');
      const expected = primary(record.word).replace(/"/g, '""');
      add(judge, 'f').textContent = 'IF(' + target + '="",0,IF(OR(' + choices.join(',') + '),0,"' + expected + '>"&$D' + row + '))';
      const result = record.result === '答对' ? 0 : primary(record.word) + '>' + question.meaning;
      cache(judge, result);
      feedback.set(quoteSheet(record.chapter) + answerCell, { result, word: primary(record.word), col });
      mark(log, 11, quoteSheet(record.chapter) + target);
      changed = true;
    }
    for (const [name, tab] of chapters) {
      if (!tab.changed) continue;
      for (let col = 5; col <= 33; col += 2) {
        const answers = tab.questions.filter(item => text(tab.cells.get(column(col) + item.row)));
        const correct = tab.questions.filter(item => text(tab.cells.get(column(col + 1) + item.row)) === '0').length;
        const rate = tab.cells.get(column(col + 1) + '1');
        if (rate) cache(rate, answers.length ? correct / tab.questions.length : 0);
      }
      writeXml(zip, paths.get(name), tab.doc);
    }
    if (practice.changed) {
      const dimension = first(practice.doc, 'dimension');
      if (dimension) dimension.setAttribute('ref', 'A1:M' + lastPracticeRow);
      writeXml(zip, paths.get('错词练习'), practice.doc);
    }
    if (feedback.size) {
      const summaryPath = paths.get('错词总表');
      let summary = await zip.file(summaryPath).async('string');
      function cached(cell, value) {
        const headEnd = cell.indexOf('>');
        let head = cell.slice(0, headEnd).replace(/\/$/, '').replace(/\s+t="[^"]*"/, '');
        if (typeof value !== 'number') head += ' t="str"';
        const body = (cell.endsWith('/>') ? '' : cell.slice(headEnd + 1, -4)).replace(/<v\b[^>]*(?:\/>|>[\s\S]*?<\/v>)/g, '').replace(/<is\b[^>]*>[\s\S]*?<\/is>/g, '');
        return head + '>' + body + '<v>' + escapeXml(value) + '</v></c>';
      }
      summary = summary.replace(/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g, (rowXml, number) => {
        const updates = new Map();
        const cells = [...rowXml.matchAll(/<c\b[^>]*\br="([A-Z]+)\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)];
        for (const [cell, col] of cells) {
          const formula = cell.match(/<f\b[^>]*>([\s\S]*?)<\/f>/)?.[1]?.replace(/&apos;/g, "'").replace(/&amp;/g, '&');
          const result = feedback.get(formula);
          if (!result) continue;
          const attempt = (result.col - 5) / 2;
          const wrong = result.result !== 0;
          updates.set(col, result.result);
          updates.set(column(5 + attempt), wrong ? result.word : '');
          updates.set(column(48 + attempt), wrong ? 1 : 0);
        }
        if (!updates.size) return rowXml;
        let total = 0;
        for (const [cell, col] of cells) {
          const index = columnNumber(col);
          if (index >= 48 && index <= 67) total += Number(updates.has(col) ? updates.get(col) : cell.match(/<v>([^<]*)<\/v>/)?.[1]) || 0;
        }
        updates.set('AT', total);
        const updated = rowXml.replace(/<c\b[^>]*\br="([A-Z]+)\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, (cell, col) => updates.has(col) ? cached(cell, updates.get(col)) : cell);
        const parsed = new DOMParser().parseFromString(updated, 'application/xml');
        if (parsed.getElementsByTagName('parsererror').length) throw new Error('无法更新错词总表，请保留本机记录并重新同步。');
        return updated;
      });
      zip.file(summaryPath, summary, { createFolders: false });
    }
    return changed;
  }

  window.DictationXlsxChapters = { project };
})();
