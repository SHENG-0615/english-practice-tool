(function (root) {
  'use strict';
  function create(entries, saved = {}, records = []) {
    const C = root.VocabularyCore || require('./vocab-core.js');
    const byId = new Map(entries.map(entry => [entry.id, entry]));
    const issues = new Map(), removed = new Set(saved.removedVocabularyMistakes || []);
    const key = (id, mode) => mode + ':' + id;
    for (const issue of saved.vocabularyMistakeBank || []) {
      if (!byId.has(issue.entryId) || !['spelling', 'meaning'].includes(issue.mode) || !['wrong', 'partial', 'unknown'].includes(issue.status)) continue;
      if (!removed.has(key(issue.entryId, issue.mode))) issues.set(key(issue.entryId, issue.mode), { ...issue, answers: Array.isArray(issue.answers) ? issue.answers : [] });
    }
    function capture(record, live = false) {
      const entry = byId.get(record.entryId);
      if (!entry || !['spelling', 'meaning'].includes(record.mode)) return;
      const issueKey = key(entry.id, record.mode);
      if (!live && removed.has(issueKey)) return;
      const grade = C.grade(entry, record.mode, record.answers);
      if (grade.empty || grade.status === 'full') return;
      const at = Number(record.at) || (record.time ? Date.parse(record.time.replace(' ', 'T')) : 0) || 0;
      const previous = issues.get(issueKey);
      if (!live && previous && previous.status !== 'unknown' && previous.at > at) return;
      if (live) removed.delete(issueKey);
      issues.set(issueKey, { entryId: entry.id, mode: record.mode, status: grade.partial ? 'partial' : 'wrong', answers: grade.answers, at, eventId: record.id || '', origin: !!record.origin });
    }
    function reconcile(history) {
      [...history].sort((a, b) => String(a.time || '').localeCompare(String(b.time || '')) || Number(a.round?.index || 0) - Number(b.round?.index || 0)).forEach(record => capture(record));
    }
    reconcile(records);
    if (!saved.vocabularyMistakeBank) {
      for (const mode of ['spelling', 'meaning']) {
        const legacy = mode === 'spelling' ? saved.spellingErrors : saved.meaningErrors;
        for (const id of legacy || []) {
          if (!byId.has(id) || issues.has(key(id, mode)) || removed.has(key(id, mode))) continue;
          const answer = saved.sessions?.[mode]?.answers?.[id];
          if (answer && C.grade(byId.get(id), mode, answer.answers).status !== 'full') capture({ entryId: id, mode, answers: answer.answers });
          else issues.set(key(id, mode), { entryId: id, mode, status: 'unknown', answers: [], at: 0 });
        }
      }
    }
    function list(filters = {}) {
      const query = String(filters.query || '').trim().toLowerCase();
      return [...issues.values()].filter(issue => {
        const entry = byId.get(issue.entryId);
        return (!filters.status || filters.status === 'all' || issue.status === filters.status)
          && (!filters.mode || filters.mode === 'all' || issue.mode === filters.mode)
          && (!filters.list || filters.list === 'all' || entry.list === filters.list)
          && (!query || [entry.word, entry.meaning, entry.list].some(text => text.toLowerCase().includes(query)));
      }).sort((a, b) => b.at - a.at || byId.get(a.entryId).list.localeCompare(byId.get(b.entryId).list, 'zh-CN', { numeric: true }) || byId.get(a.entryId).row - byId.get(b.entryId).row || a.mode.localeCompare(b.mode));
    }
    return {
      capture, reconcile, list,
      has: (id, mode) => issues.has(key(id, mode)),
      remove(id, mode) { issues.delete(key(id, mode)); removed.add(key(id, mode)); },
      snapshot: () => ({ vocabularyMistakeBank: [...issues.values()], removedVocabularyMistakes: [...removed] })
    };
  }
  const api = { create };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.VocabularyBank = api;
})(typeof window !== 'undefined' ? window : globalThis);
