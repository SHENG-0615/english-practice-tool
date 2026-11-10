(function (root) {
  'use strict';
  const trim = value => String(value ?? '').normalize('NFKC').trim().replace(/\s+/g, ' ');
  function normalizeMeaning(value) {
    let text = trim(value).replace(/[，。；、！!\s]/g, '');
    text = text.replace(/^(非常|十分|比较|较为|很)/, '');
    if (text.length > 2) text = text.replace(/的$/, '');
    return text;
  }
  function matchesMeaning(answer, keyword) {
    const input = normalizeMeaning(answer), expected = normalizeMeaning(keyword);
    if (!input || !expected) return false;
    const negative = value => /^(?:并|完全|根本|从来)?[不没未无]/.test(value);
    if (negative(input) !== negative(expected)) return false;
    return input.includes(expected) || (input.length >= 2 && input.length >= expected.length - 1 && expected.includes(input));
  }
  function label(group) {
    return group.label || group.meaning.match(/^[^\u3400-\u9fff]+/)?.[0].trim() || '释义';
  }
  function gradeSpelling(entry, values) {
    const answer = trim(Array.isArray(values) ? values[0] : values);
    const correct = entry.spellings.some(word => trim(word).toLowerCase() === answer.toLowerCase());
    return { empty: !answer, correct: !!answer && correct, partial: false, answers: [answer], status: correct ? 'full' : 'wrong', feedback: correct ? '正确' : entry.spellings.join(' / ') };
  }
  function gradeMeaning(entry, values) {
    const answers = entry.groups.map((_, index) => trim(values?.[index]));
    const groups = entry.groups.map((group, index) => ({
      label: label(group), meaning: group.meaning, answer: answers[index],
      correct: group.accepted.some(keyword => matchesMeaning(answers[index], keyword))
    }));
    const count = groups.filter(group => group.correct).length;
    const empty = answers.every(answer => !answer);
    const partial = count > 0 && count < groups.length;
    const correct = !empty && count > 0;
    const status = !correct ? 'wrong' : partial ? 'partial' : 'full';
    return {
      empty, correct, partial, answers, groups, status,
      feedback: !correct ? '✗ ' + entry.meaning : groups.map(group => (group.correct ? '✔ ' : '✘ ') + group.meaning).join(' ｜ ')
    };
  }
  function splitExistingMeaning(entry, answer) {
    const parts = String(answer || '').split(/\s+[｜|]\s+/);
    const tagged = parts.some(part => entry.groups.some(group => part.startsWith(group.label + ' ')));
    return entry.groups.map(group => tagged
      ? (parts.find(part => part.startsWith(group.label + ' ')) || '').slice(group.label.length).trim()
      : answer);
  }
  function grade(entry, mode, values) {
    return mode === 'spelling' ? gradeSpelling(entry, values) : gradeMeaning(entry, values);
  }
  const api = { trim, label, matchesMeaning, normalizeMeaning, gradeSpelling, gradeMeaning, grade, splitExistingMeaning };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.VocabularyCore = api;
})(typeof window !== 'undefined' ? window : globalThis);
