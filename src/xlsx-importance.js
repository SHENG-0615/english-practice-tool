(function () {
  'use strict';
  // Read formatting directly; SheetJS' community cell styles omit font emphasis.
  async function read(bytes) {
    const zip = await JSZip.loadAsync(bytes), X = window.DictationXlsxLog.xml;
    const { readXml, filePart } = X;
    const list = (node, name) => node ? X.list(node, name) : [];
    const first = (node, name) => node ? X.first(node, name) : null;
    const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const workbook = await readXml(zip, 'xl/workbook.xml');
    const relationships = list(await readXml(zip, 'xl/_rels/workbook.xml.rels'), 'Relationship');
    const relationship = ending => relationships.find(r => r.getAttribute('Type').endsWith('/' + ending));
    const on = node => !!node && !['0', 'false', 'off'].includes((node.getAttribute('val') || '1').toLowerCase());
    let fonts = [], formats = [], bases = [];
    const style = relationship('styles');
    if (style && style.getAttribute('TargetMode') !== 'External') {
      const doc = await readXml(zip, filePart(style.getAttribute('Target')));
      fonts = list(first(doc, 'fonts'), 'font').map(font => on(first(font, 'b')));
      formats = list(first(doc, 'cellXfs'), 'xf'); bases = list(first(doc, 'cellStyleXfs'), 'xf');
    }
    const boldStyle = index => {
      const format = formats[Number(index) || 0];
      const base = bases[Number(format?.getAttribute('xfId')) || 0];
      return !!fonts[Number(format?.getAttribute('fontId') || base?.getAttribute('fontId') || 0)];
    };
    const shared = relationship('sharedStrings');
    const strings = shared && shared.getAttribute('TargetMode') !== 'External' ? list(await readXml(zip, filePart(shared.getAttribute('Target'))), 'si') : [];
    const boldText = (node, inherited) => {
      const runs = list(node, 'r');
      if (!runs.length) return inherited;
      return runs.some(run => {
        if (!first(run, 't')?.textContent.trim()) return false;
        const b = first(first(run, 'rPr'), 'b');
        return b ? on(b) : inherited;
      });
    };
    const marked = new Set();
    for (const sheet of list(workbook, 'sheet')) {
      const rel = relationships.find(r => r.getAttribute('Id') === sheet.getAttributeNS(REL, 'id'));
      if (!rel || rel.getAttribute('TargetMode') === 'External') continue;
      const doc = await readXml(zip, filePart(rel.getAttribute('Target')));
      const column = list(doc, 'col').find(c => Number(c.getAttribute('min')) <= 2 && Number(c.getAttribute('max')) >= 2);
      for (const row of list(first(doc, 'sheetData'), 'row')) {
        for (const cell of [...row.children].filter(c => c.localName === 'c' && /^B\d+$/.test(c.getAttribute('r')))) {
          const inherited = boldStyle(cell.getAttribute('s') ?? row.getAttribute('s') ?? column?.getAttribute('style') ?? 0);
          const text = cell.getAttribute('t') === 's' ? strings[Number(first(cell, 'v')?.textContent)] : first(cell, 'is');
          if (text ? boldText(text, inherited) : inherited) marked.add(sheet.getAttribute('name') + ':' + Number(cell.getAttribute('r').slice(1)));
        }
      }
    }
    return marked;
  }
  window.WorkbookImportance = { read };
})();
