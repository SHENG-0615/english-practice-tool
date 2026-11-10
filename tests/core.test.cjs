const assert = require('node:assert/strict');
const C = require('../src/core.js');
const V = require('../src/vocab-core.js');
const words = C.parseText('英文|中文|章节\norbit|n. 轨道|Demo\nbridge|n. 桥；v. 架桥|Demo', { source: 'test' });
assert.equal(words.length, 2);
assert.equal(words[0].word, 'orbit');
const spelling = { word: 'orbit', spellings: ['orbit'] };
assert(V.gradeSpelling(spelling, ['ORBIT']).correct);
assert(!V.gradeSpelling(spelling, ['orbits']).correct);
const entry = { meaning: 'n. 桥；v. 架桥', groups: [
  { label: 'n.', meaning: 'n. 桥', accepted: ['桥'] },
  { label: 'v.', meaning: 'v. 架桥', accepted: ['架桥'] }
] };
assert(V.gradeMeaning(entry, ['桥', '']).partial);
assert(V.gradeMeaning(entry, ['桥', '架桥']).correct);
assert(V.gradeMeaning(entry, ['', '']).empty);
assert(!V.matchesMeaning('不可靠', '可靠'));
assert(V.matchesMeaning('很可靠', '可靠的'));
assert.equal(C.safeSettings({}).repeat, 1);
assert.equal(C.safeSettings({}).gap, 3);
console.log('Core checks passed; fixtures are authored inside the test and never bundled as a word list.');
const B = require('../src/vocab-bank.js');
const bankEntry = { ...entry, id: 'authored:1', list: 'Demo', row: 3, word: 'bridge', spellings: ['bridge'] };
const bank = B.create([bankEntry]);
bank.capture({ entryId: bankEntry.id, mode: 'meaning', answers: ['桥', ''] }, true);
assert.equal(bank.list({ status: 'partial' }).length, 1);
assert.equal(bank.list({ status: 'wrong' }).length, 0);
bank.remove(bankEntry.id, 'meaning');
const restored = B.create([bankEntry], bank.snapshot(), [{ entryId: bankEntry.id, mode: 'meaning', answers: ['桥', ''] }]);
assert.equal(restored.list().length, 0);
restored.capture({ entryId: bankEntry.id, mode: 'meaning', answers: ['未知', '未知'] }, true);
assert.equal(restored.list({ status: 'wrong' }).length, 1);
