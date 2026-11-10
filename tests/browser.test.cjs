const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const assert = require('node:assert/strict');
const { pathToFileURL } = require('node:url');
const { chromium } = require('playwright');
const XLSX = require('../vendor/xlsx.full.min.js');
const JSZip = require('../vendor/jszip.min.js');
const entry = path.resolve(__dirname, '../index.html');
function mockSpeech() {
  Object.defineProperty(window, 'speechSynthesis', { configurable: true, value: {
    getVoices: () => [{ name: 'Test UK', lang: 'en-GB', voiceURI: 'gb', localService: true }],
    addEventListener() {}, cancel() {}, pause() {}, resume() {}, speak(u) { setTimeout(() => u.onend?.(), 5); }
  } });
  // Sentinels prove the public app does not restore another app's browser data.
  localStorage.setItem('dictation-player-library-v1', JSON.stringify({ entries: [{ word: 'private-sentinel', meaning: 'PRIVATE_SENTINEL' }] }));
  localStorage.setItem('vocabulary-practice-v1', JSON.stringify({ version: 1, selected: ['PRIVATE_SENTINEL'] }));
}
async function importLocal(page, filename, target) {
  await page.locator('#publicOpenImport').click();
  await page.locator('#publicImportTarget').selectOption(target);
  await page.locator('#publicImportFile').setInputFiles(filename);
  await page.waitForFunction(() => !document.getElementById('publicImportCommit').disabled);
  await page.locator('#publicImportCommit').click();
  await page.waitForFunction(() => !document.getElementById('publicImportDialog').open && document.getElementById('publicDatasetStatus').textContent.includes('本机词库'));
}
(async () => {
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'english-tool-test-'));
  const words = path.join(temporary, 'synthetic.csv');
  fs.writeFileSync(words, '英文,中文,章节\norbit,n. 轨道,Demo\nbridge,n. 桥；v. 架桥,Demo', 'utf8');
  const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHANNEL ? { channel: process.env.PLAYWRIGHT_CHANNEL } : {}) });
  try {
    const ctx = await browser.newContext({ viewport: { width: 1366, height: 900 } });
    await ctx.addInitScript(mockSpeech);
    const page = await ctx.newPage(), errors = [], requests = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('request', request => { if (/^https?:/.test(request.url())) requests.push(request.url()); });
    await page.goto(pathToFileURL(entry).href);
    assert((await page.locator('#publicDatasetStatus').textContent()).includes('空词库'));
    assert.equal(await page.locator('#chapterGroups input').count(), 0);
    assert(!await page.locator('body').textContent().then(value => value.includes('PRIVATE_SENTINEL')));
    await page.locator('#areaVocabulary').click();
    assert.equal(await page.locator('#vocabLists input').count(), 0);
    await page.locator('#areaDictation').click();
    await importLocal(page, words, 'both');
    assert.equal(await page.locator('#chapterGroups input').count(), 1);
    await page.locator('#dictationMode').click();
    await page.locator('#dictationAnswer').fill('orbit'); await page.locator('#dictationAnswer').press('Enter');
    assert((await page.locator('#dictationScore').textContent()).includes('已提交 1 题，答对 1 题'));
    await page.locator('#areaVocabulary').click();
    assert((await page.locator('#publicVocabCount').textContent()).includes('2 词'));
    await page.locator('#vocabDisplay').selectOption('chinese');
    assert(!await page.locator('#vocabTable').textContent().then(value => value.includes('orbit')));
    await page.locator('#vocabDisplay').selectOption('english');
    assert.equal(await page.locator('#vocabStudyMeaning').textContent(), '');
    await page.locator('#vocabSpellingMode').click();
    await page.locator('#vocabAnswer0').fill('ORBIT'); await page.locator('#vocabAnswer0').press('Enter');
    assert((await page.locator('#vocabFeedback').textContent()).includes('答对了'));
    await page.locator('#vocabAnswer0').fill('draft'); await page.reload();
    assert.equal(await page.locator('#vocabAnswer0').inputValue(), 'draft');
    await page.locator('#vocabMeaningMode').click();
    await page.locator('#vocabSearch').fill('bridge');
    assert.equal(await page.locator('#vocabInputs input').count(), 2);
    await page.locator('#vocabAnswer0').fill('桥'); await page.locator('#vocabAnswer0').press('Enter');
    await page.locator('#vocabAnswer1').press('Enter');
    assert((await page.locator('#vocabFeedback').textContent()).includes('少答，算对'));
    await page.locator('#vocabOnlyMistakes').check();
    assert((await page.locator('#vocabTableCount').textContent()).includes('1 / 1'));
    await page.locator('#vocabBankScope').click(); await page.locator('#vocabBankPartial').click();
    assert.equal(await page.locator('#vocabTable tbody tr').count(), 1);
    assert((await page.locator('#vocabTable tbody').textContent()).includes('架桥'));
    await page.locator('#vocabBankWrong').click(); assert.equal(await page.locator('#vocabTable tbody tr').count(), 0);
    await page.locator('#vocabBankAll').click();
    await page.setViewportSize({ width: 390, height: 844 });
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth > innerWidth), false);
    if (process.env.REVIEW_SCREENSHOTS) {
      fs.mkdirSync(process.env.REVIEW_SCREENSHOTS, { recursive: true });
      await page.screenshot({ path: path.join(process.env.REVIEW_SCREENSHOTS, 'vocab-bank-mobile.png'), fullPage: true });
    }
    await page.setViewportSize({ width: 1366, height: 900 });
    // A synthetic workbook exercises the hidden criterion columns without any real source workbook.
    const book = XLSX.utils.book_new();
    const sheet = { '!ref': 'A1:CX3', A2: { t: 's', v: '序号' }, B3: { t: 's', v: 'synthetic' }, D3: { t: 's', v: 'adj. 合成的' }, AI3: { t: 's', v: 'synthetic' }, AJ3: { t: 's', v: 'adj. 合成的' }, AN3: { t: 's', v: '合成' } };
    XLSX.utils.book_append_sheet(book, sheet, 'List01');
    const xlsx = path.join(temporary, 'synthetic.xlsx');
    const zip = await JSZip.loadAsync(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
    let styles = await zip.file('xl/styles.xml').async('string');
    function addFormat(tag, content) {
      const re = new RegExp('<' + tag + '\\b([^>]*)count="(\\d+)"([^>]*)>([\\s\\S]*?)</' + tag + '>');
      const match = re.exec(styles); assert(match);
      const index = Number(match[2]);
      styles = styles.replace(re, '<' + tag + match[1] + 'count="' + (index + 1) + '"' + match[3] + '>' + match[4] + content + '</' + tag + '>');
      return index;
    }
    const fontId = addFormat('fonts', '<font><b/><sz val="11"/><name val="Calibri"/></font>');
    const styleId = addFormat('cellXfs', '<xf numFmtId="0" fontId="' + fontId + '" fillId="0" borderId="0" xfId="0" applyFont="1"/>');
    zip.file('xl/styles.xml', styles);
    zip.file('xl/worksheets/sheet1.xml', (await zip.file('xl/worksheets/sheet1.xml').async('string')).replace(/<c r="B3"/, '<c s="' + styleId + '" r="B3"'));
    fs.writeFileSync(xlsx, await zip.generateAsync({ type: 'nodebuffer' }));
    await importLocal(page, xlsx, 'vocabulary');
    await page.locator('#vocabListScope').click(); await page.locator('#vocabStudyMode').click();
    assert.equal(await page.locator('#vocabStudyWord .priority-mark').textContent(), '*');
    await page.locator('#vocabMeaningMode').click();
    await page.locator('#vocabAnswer0').fill('合成'); await page.locator('#vocabAnswer0').press('Enter');
    assert((await page.locator('#vocabFeedback').textContent()).includes('答对了'));
    await page.setViewportSize({ width: 390, height: 844 });
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth > innerWidth), false);
    assert.deepEqual(errors, []); assert.deepEqual(requests, []);
    assert(!fs.readFileSync(entry, 'utf8').includes('PRIVATE_SENTINEL'));
    if (process.env.REVIEW_SCREENSHOTS) {
      fs.mkdirSync(process.env.REVIEW_SCREENSHOTS, { recursive: true });
      const fresh = await browser.newContext({ viewport: { width: 1366, height: 900 } });
      await fresh.addInitScript(mockSpeech);
      const blank = await fresh.newPage(); await blank.goto(pathToFileURL(entry).href);
      await blank.screenshot({ path: path.join(process.env.REVIEW_SCREENSHOTS, 'public-empty.png') });
      await blank.locator('#areaVocabulary').click();
      await blank.screenshot({ path: path.join(process.env.REVIEW_SCREENSHOTS, 'public-vocabulary-empty.png') });
      await fresh.close();
    }
    console.log(JSON.stringify({ status: 'passed', checks: ['empty first launch', 'isolated browser storage', 'local CSV and XLSX import', 'original bold priority import', 'dictation Enter feedback', 'study display options', 'strict spelling', 'per-POS meaning inputs', 'partial credit', 'independent vocabulary bank and status filters', 'correct answers', 'mobile bank layout', 'draft resume', 'no network requests', 'no file changes from local imports'] }));
    await ctx.close();
  } finally {
    await browser.close();
    for (const filename of fs.readdirSync(temporary)) fs.unlinkSync(path.join(temporary, filename));
    fs.rmdirSync(temporary);
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
