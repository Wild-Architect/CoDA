(function exposePdfSearchIndex(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.PdfSearchIndex = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, () => {
  const whitespace = /\s/u;

  function itemHeight(item) {
    const transform = Array.isArray(item?.transform) ? item.transform : [];
    return Math.max(Math.abs(Number(item?.height) || 0), Math.hypot(Number(transform[2]) || 0, Number(transform[3]) || 0), Math.hypot(Number(transform[0]) || 0, Number(transform[1]) || 0), 1);
  }

  function positionedItems(content) {
    let rawOffset = 0;
    return (Array.isArray(content?.items) ? content.items : []).map((item, order) => {
      const str = String(item?.str || ''), transform = Array.isArray(item?.transform) ? item.transform : [], height = itemHeight(item), width = Math.abs(Number(item?.width) || 0);
      const result = { str, order, rawStart: rawOffset, rawEnd: rawOffset + str.length, x: Number(transform[4]) || 0, y: Number(transform[5]) || 0, width, height, hasEOL: Boolean(item?.hasEOL) };
      rawOffset += str.length;
      return result;
    }).filter(item => item.str.length);
  }

  function groupLines(items) {
    const lines = [];
    items.forEach(item => {
      const current = lines.at(-1), tolerance = current ? Math.max(2, Math.min(current.height, item.height) * .55) : 0;
      if (!current || Math.abs(current.y - item.y) > tolerance || current.hasEOL) {
        lines.push({ items: [item], y: item.y, height: item.height, hasEOL: item.hasEOL });
        return;
      }
      current.items.push(item);
      current.height = Math.max(current.height, item.height);
      current.y = current.items.reduce((sum, entry) => sum + entry.y, 0) / current.items.length;
      current.hasEOL ||= item.hasEOL;
    });
    return lines.map(line => {
      line.items.sort((left, right) => left.x - right.x || left.order - right.order);
      line.xStart = Math.min(...line.items.map(item => item.x));
      line.xEnd = Math.max(...line.items.map(item => item.x + item.width));
      return line;
    });
  }

  function needsParagraphBreak(previous, current) {
    if (!previous) return false;
    const referenceHeight = Math.max(previous.height, current.height, 1);
    const verticalGap = Math.abs(previous.y - current.y);
    const positiveIndent = current.xStart - previous.xStart;
    return verticalGap > referenceHeight * 1.55 || positiveIndent > Math.max(8, referenceHeight * .8);
  }

  function buildPageIndex(content) {
    const rawText = (Array.isArray(content?.items) ? content.items : []).map(item => String(item?.str || '')).join('');
    const grouped = groupLines(positionedItems(content)), characters = [], rawOffsets = [], lines = [];
    const appendSeparator = value => { for (const char of value) { characters.push(char); rawOffsets.push(-1); } };
    grouped.forEach((line, lineIndex) => {
      if (lineIndex) appendSeparator(needsParagraphBreak(grouped[lineIndex - 1], line) ? '\n\n' : '\n');
      const start = characters.length;
      line.items.forEach((item, itemIndex) => {
        const previous = line.items[itemIndex - 1];
        if (previous && !whitespace.test(previous.str.at(-1) || '') && !whitespace.test(item.str[0] || '')) {
          const averageCharacter = previous.width / Math.max(1, previous.str.trim().length);
          if (item.x - (previous.x + previous.width) > Math.max(.8, averageCharacter * .18)) appendSeparator(' ');
        }
        for (let index = 0; index < item.str.length; index++) { characters.push(item.str[index]); rawOffsets.push(item.rawStart + index); }
      });
      lines.push({ start, end: characters.length, xStart: line.xStart, y: line.y, height: line.height });
    });
    const structuredText = characters.join(''), normalizedCharacters = [], structuredOffsets = [], normalizedRawOffsets = [];
    let pendingWhitespace = null;
    for (let index = 0; index < structuredText.length; index++) {
      if (whitespace.test(structuredText[index])) { if (normalizedCharacters.length && pendingWhitespace === null) pendingWhitespace = index; continue; }
      if (pendingWhitespace !== null) { normalizedCharacters.push(' '); structuredOffsets.push(pendingWhitespace); normalizedRawOffsets.push(-1); pendingWhitespace = null; }
      normalizedCharacters.push(structuredText[index]); structuredOffsets.push(index); normalizedRawOffsets.push(rawOffsets[index]);
    }
    return { rawText, structuredText, lines, normalizedText: normalizedCharacters.join(''), structuredOffsets, rawOffsets: normalizedRawOffsets };
  }

  function normalizeQuery(value) { return String(value || '').replace(/\s+/gu, ' ').trim(); }

  function matchContext(pageIndex, normalizedStart, normalizedEnd, surroundingLines = 2) {
    const structuredStart = pageIndex.structuredOffsets[normalizedStart], structuredEnd = pageIndex.structuredOffsets[normalizedEnd - 1] + 1;
    const raw = pageIndex.rawOffsets.slice(normalizedStart, normalizedEnd).filter(offset => offset >= 0);
    if (!raw.length) return null;
    let firstLine = pageIndex.lines.findIndex(line => line.end > structuredStart);
    if (firstLine < 0) firstLine = 0;
    let lastLine = firstLine;
    for (let index = firstLine; index < pageIndex.lines.length && pageIndex.lines[index].start < structuredEnd; index++) lastLine = index;
    const fromLine = Math.max(0, firstLine - surroundingLines), toLine = Math.min(pageIndex.lines.length - 1, lastLine + surroundingLines);
    const contextStart = pageIndex.lines[fromLine]?.start ?? 0, contextEnd = pageIndex.lines[toLine]?.end ?? pageIndex.structuredText.length;
    return {
      start: Math.min(...raw), end: Math.max(...raw) + 1,
      before: `${fromLine ? '…\n' : ''}${pageIndex.structuredText.slice(contextStart, structuredStart)}`,
      term: pageIndex.structuredText.slice(structuredStart, structuredEnd),
      after: `${pageIndex.structuredText.slice(structuredEnd, contextEnd)}${toLine < pageIndex.lines.length - 1 ? '\n…' : ''}`
    };
  }

  function findMatches(pageIndex, query, limit = 5000) {
    const needle = normalizeQuery(query).toLocaleLowerCase('uk');
    if (!needle) return [];
    const haystack = pageIndex.normalizedText.toLocaleLowerCase('uk'), matches = [];
    let offset = 0, found;
    while ((found = haystack.indexOf(needle, offset)) !== -1 && matches.length < limit) {
      const context = matchContext(pageIndex, found, found + needle.length);
      if (context) matches.push(context);
      offset = found + Math.max(1, needle.length);
    }
    return matches;
  }

  return { buildPageIndex, findMatches, normalizeQuery };
});
