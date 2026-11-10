(function () {
  'use strict';
  const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const SHEET = 'HTML 听写记录';
  const WORKBOOK = 'xl/workbook.xml';
  const RELATIONS = 'xl/_rels/workbook.xml.rels';
  const TYPES = '[Content_Types].xml';
  const HEADERS = ['作答时间', '词汇来源', '练习入口', '章节', '原表题号', '正确拼写', '你的作答', '结果', '中文释义', '记录类别', '记录ID', '章节写入位置', '错题写入位置', '练习轮次'];
  const xml = new XMLSerializer();
  const list = (node, name) => [...node.getElementsByTagNameNS('*', name)];
  const first = (node, name) => node.getElementsByTagNameNS('*', name)[0];
  const filePart = target => decodeURIComponent(new URL(target, 'https://workbook.invalid/' + WORKBOOK).pathname.slice(1));

  function add(parent, name, attributes = {}) {
    const node = parent.ownerDocument.createElementNS(parent.namespaceURI, name);
    for (const [key, value] of Object.entries(attributes)) node.setAttribute(key, value);
    parent.append(node);
    return node;
  }
  async function readXml(zip, path) {
    const file = zip.file(path);
    if (!file) throw new Error('工作簿缺少必要内容，请选择完整的 XLSX。');
    const doc = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
    if (first(doc, 'parsererror')) throw new Error('无法读取工作簿的记录结构。');
    return doc;
  }
  function writeXml(zip, path, doc) {
    zip.file(path, xml.serializeToString(doc), { createFolders: false });
  }
  function registerLog(zip, workbook, relationships, types, sheets) {
    const sheetId = sheets.reduce((max, sheet) => Math.max(max, Number(sheet.getAttribute('sheetId')) || 0), 0) + 1;
    const used = new Set(list(relationships, 'Relationship').map(rel => rel.getAttribute('Id')));
    let next = 1;
    while (used.has('rId' + next)) next++;
    const relId = 'rId' + next;
    let partNumber = sheetId;
    while (zip.file('xl/worksheets/sheet' + partNumber + '.xml')) partNumber++;
    const target = 'worksheets/sheet' + partNumber + '.xml';
    const path = filePart(target);
    const sheet = add(first(workbook, 'sheets'), 'sheet', { name: SHEET, sheetId });
    sheet.setAttributeNS(DOCUMENT_REL, 'r:id', relId);
    add(relationships.documentElement, 'Relationship', { Id: relId, Type: DOCUMENT_REL + '/worksheet', Target: target });
    add(types.documentElement, 'Override', { PartName: '/' + path, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' });
    writeXml(zip, WORKBOOK, workbook);
    writeXml(zip, RELATIONS, relationships);
    writeXml(zip, TYPES, types);
    return path;
  }
  function newLog() {
    const doc = document.implementation.createDocument(MAIN, 'worksheet');
    const sheet = doc.documentElement;
    add(sheet, 'dimension', { ref: 'A1:N1' });
    const view = add(add(sheet, 'sheetViews'), 'sheetView', { workbookViewId: 0 });
    add(view, 'pane', { ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' });
    const cols = add(sheet, 'cols');
    [22, 17, 18, 11, 12, 32, 30, 10, 40, 15, 18, 18, 18, 18].forEach((width, index) => {
      const number = index + 1;
      add(cols, 'col', { min: number, max: number, width, customWidth: 1, ...(number >= 11 ? { hidden: 1 } : {}) });
    });
    addRow(add(sheet, 'sheetData'), 1, HEADERS);
    add(sheet, 'autoFilter', { ref: 'A1:N1' });
    return doc;
  }
  function addRow(data, number, values) {
    const row = add(data, 'row', { r: number });
    values.forEach((value, index) => {
      if (value === '' || value == null) return;
      const address = String.fromCharCode(65 + index) + number;
      if (typeof value === 'number') {
        add(add(row, 'c', { r: address }), 'v').textContent = value;
      } else {
        const cell = add(row, 'c', { r: address, t: 'inlineStr' });
        const run = add(add(cell, 'is'), 't');
        run.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        run.textContent = String(value).toWellFormed().replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '');
      }
    });
    return row;
  }
  async function sharedStrings(zip, relationships) {
    let strings = [];
    const relation = list(relationships, 'Relationship').find(rel => rel.getAttribute('Type').endsWith('/sharedStrings'));
    if (relation) {
      const doc = await readXml(zip, filePart(relation.getAttribute('Target')));
      strings = list(doc, 'si').map(item => list(item, 't').map(run => run.textContent).join(''));
    }
    return strings;
  }
  function readLogRows(worksheet, strings) {
    const read = cell => cell.getAttribute('t') === 's'
      ? strings[Number(first(cell, 'v')?.textContent)]
      : (list(cell, 't').map(run => run.textContent).join('') || first(cell, 'v')?.textContent);
    const cells = list(worksheet, 'c');
    const header = cells.find(cell => cell.getAttribute('r') === 'K1');
    if (!header || read(header) !== HEADERS[10]) throw new Error('同名记录页的表头不同，请检查所选工作簿。');
    return list(worksheet, 'row').filter(row => Number(row.getAttribute('r')) > 1).map(node => {
      const fields = Array(14).fill('');
      for (const cell of [...node.children]) {
        const col = cell.getAttribute('r')?.replace(/\d+$/, '');
        if (/^[A-N]$/.test(col)) fields[col.charCodeAt(0) - 65] = read(cell) || '';
      }
      return { node, number: Number(node.getAttribute('r')), fields };
    }).filter(row => row.fields[10]);
  }
  async function append(bytes, records) {
    let zip;
    try { zip = await JSZip.loadAsync(bytes); }
    catch (_) { throw new Error('无法打开这份 XLSX，请选择未加密的完整工作簿。'); }
    if (zip.file(/^_xmlsignatures\//i).length) throw new Error('这份工作簿含数字签名，请选一份未签名的副本。');
    const [workbook, relationships] = await Promise.all([readXml(zip, WORKBOOK), readXml(zip, RELATIONS)]);
    const sheets = list(workbook, 'sheet');
    if (!['错词总表', '错词练习'].every(name => sheets.some(sheet => sheet.getAttribute('name') === name))) {
      throw new Error('请选择包含“错词总表”和“错词练习”的完整语料库 XLSX。');
    }
    const existing = sheets.find(sheet => sheet.getAttribute('name').toLowerCase() === SHEET.toLowerCase());
    let path, worksheet;
    if (existing) {
      const id = existing.getAttributeNS(DOCUMENT_REL, 'id');
      const relation = list(relationships, 'Relationship').find(rel => rel.getAttribute('Id') === id);
      if (!relation?.getAttribute('Type').endsWith('/worksheet') || relation.getAttribute('TargetMode') === 'External') {
        throw new Error('无法定位这份工作簿里的 HTML 听写记录页。');
      }
      path = filePart(relation.getAttribute('Target'));
      worksheet = await readXml(zip, path);
    } else {
      const types = await readXml(zip, TYPES);
      path = registerLog(zip, workbook, relationships, types, sheets);
      worksheet = newLog();
    }
    const strings = await sharedStrings(zip, relationships);
    const logs = readLogRows(worksheet, strings);
    const known = new Set(logs.map(row => row.fields[10]));
    const data = first(worksheet, 'sheetData');
    let rowNumber = list(data, 'row').reduce((max, row) => Math.max(max, Number(row.getAttribute('r')) || 0), 0);
    let added = 0;
    for (const row of records) {
      if (known.has(row.id)) continue;
      if (rowNumber >= 1048576) throw new Error('记录页已满，请连接一份新的语料库副本。');
      const fields = [row.time || '', row.source || '', row.mode || '', row.chapter || '', row.number || '', row.word || '', row.answer || '', row.result || '', row.meaning || '', row.kind || '', row.id, '', '', row.round || ''];
      const node = addRow(data, ++rowNumber, fields);
      logs.push({ node, number: rowNumber, fields });
      known.add(row.id);
      added++;
    }
    const paths = new Map(sheets.map(sheet => {
      const id = sheet.getAttributeNS(DOCUMENT_REL, 'id');
      const relation = list(relationships, 'Relationship').find(rel => rel.getAttribute('Id') === id);
      return [sheet.getAttribute('name'), filePart(relation.getAttribute('Target'))];
    }));
    const projected = await window.DictationXlsxChapters.project(zip, paths, logs, strings, { list, first, add, readXml, writeXml });
    if (existing && !added && !projected) return null;
    const headerRow = list(data, 'row').find(row => row.getAttribute('r') === '1');
    for (const index of [11, 12, 13]) {
      const address = String.fromCharCode(65 + index) + '1';
      if (!list(headerRow, 'c').some(cell => cell.getAttribute('r') === address)) {
        add(add(add(headerRow, 'c', { r: address, t: 'inlineStr' }), 'is'), 't').textContent = HEADERS[index];
      }
    }
    const cols = first(worksheet, 'cols');
    if (cols && !list(cols, 'col').some(col => Number(col.getAttribute('min')) === 12)) {
      add(cols, 'col', { min: 12, max: 14, hidden: 1, width: 18, customWidth: 1 });
    }
    for (const tag of ['dimension', 'autoFilter']) {
      const range = first(worksheet, tag);
      if (range && /^A1:[KMN]\d+$/.test(range.getAttribute('ref'))) range.setAttribute('ref', 'A1:N' + rowNumber);
    }
    writeXml(zip, path, worksheet);
    if (projected) {
      const calc = first(workbook, 'calcPr') || add(workbook.documentElement, 'calcPr');
      calc.setAttribute('fullCalcOnLoad', '1');
      calc.setAttribute('forceFullCalc', '1');
      writeXml(zip, WORKBOOK, workbook);
    }
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  window.DictationXlsxLog = { append, xml: { add, list, first, readXml, writeXml, filePart } };
})();
