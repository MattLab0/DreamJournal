function generateScoreTable(skipZeroScores) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paras = body.getParagraphs();

  fillInScores(); // Ensure scores are updated

  const dateRe = /^Dreams?\s+(\d{2}\/\d{2}\/\d{4})/i;
  const uncertainDateRe = /^Dreams?\s+(\?{1,2}\/\d{2}\/\d{4})/i;
  const scoreRe = /D:(\d+(?:[.,]\d+)?)\s+LD:(\d+(?:[.,]\d+)?)(?:\s+Score:(\d+(?:[.,]\d+)?))?/i;

  // Build headingLinks TOC 
  const headingLinks = {};
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.TABLE_OF_CONTENTS) {
      const toc = el;
      for (let j = 0; j < toc.getNumChildren(); j++) {
        const p = toc.getChild(j);
        if (p.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
        const para = p.asParagraph();
        const textChild = para.getNumChildren() > 0 ? para.getChild(0).asText() : null;
        const link = textChild ? textChild.getLinkUrl() : null;
        if (!link) continue;
        const text = para.getText();
        const m = text.match(dateRe);
        const um = text.match(uncertainDateRe);
        if (m) headingLinks[m[1]] = link;
        else if (um) headingLinks[um[1]] = link;
      }
      break;
    }
  }

  // Build countsMap
  const countsMap = {};
  let currentH1Key = null;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (!el || el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const para = el.asParagraph();
    const heading = para.getHeading();
    const text = para.getText().trim();

    if (heading === DocumentApp.ParagraphHeading.HEADING1) {
      const m = text.match(dateRe);
      const um = text.match(uncertainDateRe);
      if (m) currentH1Key = m[1];
      else if (um) currentH1Key = um[1];
      else currentH1Key = null;
      if (currentH1Key && countsMap[currentH1Key] == null) countsMap[currentH1Key] = 0;
      continue;
    }

    if (currentH1Key && heading === DocumentApp.ParagraphHeading.HEADING2) {
      const lower = text.toLowerCase();
      if (!lower) continue;

      if (lower.startsWith('fragment')) {
        countsMap[currentH1Key] = (countsMap[currentH1Key] || 0) + 0.5;
        continue;
      }

      if (!lower.startsWith('top words') && !lower.startsWith('word cloud') && lower !== 'average and median') {
        countsMap[currentH1Key] = (countsMap[currentH1Key] || 0) + 1;
      }
    }
  }

  // Extract entries
  const entries = [];
  const uncertainEntries = [];
  paras.forEach(p => {
    if (p.getHeading() !== DocumentApp.ParagraphHeading.HEADING1) return;
    const text = p.getText();
    const scoreMatch = text.match(scoreRe);
    if (!scoreMatch) return;

    const dateMatch = text.match(dateRe);
    if (dateMatch) {
      const date = dateMatch[1];
      const d = parseFloat(scoreMatch[1].replace(',', '.'));
      const ld = parseFloat(scoreMatch[2].replace(',', '.'));
      const score = scoreMatch[3] ? parseFloat(scoreMatch[3].replace(',', '.')) : d + ld;
      if (!(skipZeroScores && score <= 0.25)) entries.push({ date, d, ld, score });
      return;
    }

    const uncertainMatch = text.match(uncertainDateRe);
    if (uncertainMatch) {
      const date = uncertainMatch[1];
      const d = parseFloat(scoreMatch[1].replace(',', '.'));
      const ld = parseFloat(scoreMatch[2].replace(',', '.'));
      const score = scoreMatch[3] ? parseFloat(scoreMatch[3].replace(',', '.')) : d + ld;
      uncertainEntries.push({ date, d, ld, score });
    }
  });

  if (entries.length === 0 && uncertainEntries.length === 0) {
    DocumentApp.getUi().alert("No valid scores found.");
    return;
  }

  // Remove old tables
  const headingsToRemove = ['Score Table', 'Uncertain Date Table'];
  for (let i = body.getNumChildren() - 1; i >= 0; i--) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH &&
      el.asParagraph().getHeading() === DocumentApp.ParagraphHeading.HEADING1 &&
      headingsToRemove.includes(el.asParagraph().getText().trim())) {
      body.removeChild(el);
      if (i < body.getNumChildren() && body.getChild(i).getType() === DocumentApp.ElementType.TABLE) {
        body.removeChild(body.getChild(i));
      }
    }
  }

  // Find insert position 
  let insertAt = body.getNumChildren();
  for (let i = 0; i < insertAt; i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH &&
      el.asParagraph().getHeading() === DocumentApp.ParagraphHeading.HEADING1 &&
      el.asParagraph().getText().trim() === 'Index') {
      for (let j = i + 1; j < body.getNumChildren(); j++) {
        const el2 = body.getChild(j);
        if (el2.getType() === DocumentApp.ElementType.PARAGRAPH &&
          el2.asParagraph().getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
          insertAt = j;
          break;
        }
      }
      break;
    }
  }

  if (insertAt > 0) {
    const prevEl = body.getChild(insertAt - 1);
    let hasPageBreak = false;
    if (prevEl.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const prevP = prevEl.asParagraph();
      for (let k = 0; k < prevP.getNumChildren(); k++) {
        if (prevP.getChild(k).getType() === DocumentApp.ElementType.PAGE_BREAK) {
          hasPageBreak = true;
          break;
        }
      }
    }
    if (!hasPageBreak) {
      body.insertPageBreak(insertAt);
      insertAt++;
    }
  } else {
    body.insertPageBreak(insertAt);
    insertAt++;
  }

  // Insert Uncertain Date Score Table
  if (uncertainEntries.length > 0) {
    body.insertParagraph(insertAt, 'Uncertain Date Table').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const tableData = [['📅 Date', 'Month', '# Dreams', 'D Score', 'LD Score', 'Total Score']];

    uncertainEntries.forEach(entry => {
      const parts = entry.date.split('/');
      const monthStr = parts.length > 2 ? new Date(Number(parts[2]), Number(parts[1]) - 1, 1).toLocaleString('default', { month: 'long' }) : '?';
      const dreamCount = countsMap[entry.date] || 0;

      tableData.push([
        entry.date,
        monthStr,
        Number.isFinite(dreamCount) ? (dreamCount % 1 === 0 ? dreamCount.toString() : dreamCount.toFixed(1)) : '0',
        entry.d.toString(),
        entry.ld.toString(),
        entry.score.toString()
      ]);
    });

    const table = body.insertTable(insertAt + 1, tableData);
    table.getRow(0).editAsText().setBold(true);

    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      const date = row.getCell(0).getText();
      const dateText = row.getCell(0).editAsText();
      if (headingLinks[date]) dateText.setLinkUrl(headingLinks[date]);

      for (let c = 0; c < row.getNumCells(); c++) {
        row.getCell(c).editAsText().setFontSize(10);
        row.getCell(c).getChild(0).asParagraph().setLineSpacing(1);
      }

      const ldVal = parseFloat(row.getCell(4).getText().replace(',', '.'));
      if (!isNaN(ldVal) && ldVal >= 1) row.getCell(4).editAsText().setForegroundColor('#008000');
      const totalVal = parseFloat(row.getCell(5).getText().replace(',', '.'));
      if (!isNaN(totalVal) && totalVal === 0) row.getCell(5).editAsText().setForegroundColor('#FF0000');
    }
    insertAt += 2;
  }

  // Insert main Score Table
  if (entries.length > 0) {
    // Ensure continuous date coverage if skipZeroScores is false
    let finalEntries = [...entries];
    if (!skipZeroScores) {
      // Parse dates
      const sorted = entries
        .map(e => {
          const [d, m, y] = e.date.split('/').map(Number);
          return { jsDate: new Date(y, m - 1, d), ...e };
        })
        .sort((a, b) => a.jsDate - b.jsDate);

      const startDate = sorted[0].jsDate;
      const endDate = sorted[sorted.length - 1].jsDate;

      // Create a map for lookup
      const entryMap = {};
      sorted.forEach(e => {
        entryMap[e.date] = e;
      });

      // Fill missing days
      finalEntries = [];
      let current = new Date(startDate);
      while (current <= endDate) {
        const dd = String(current.getDate()).padStart(2, '0');
        const mm = String(current.getMonth() + 1).padStart(2, '0');
        const yyyy = current.getFullYear();
        const dateStr = `${dd}/${mm}/${yyyy}`;

        if (entryMap[dateStr]) {
          finalEntries.push(entryMap[dateStr]);
        } else {
          finalEntries.push({
            date: dateStr,
            d: 0,
            ld: 0,
            score: 0
          });
        }
        current.setDate(current.getDate() + 1);
      }
    }

    body.insertParagraph(insertAt, 'Score Table').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const tableData = [['📅 Date', 'Week', 'Month', '# Dreams', 'D Score', 'LD Score', 'Total Score']];

    finalEntries.forEach(entry => {
      const [day, month, year] = entry.date.split('/').map(Number);
      const jsDate = new Date(year, month - 1, day);
      const monthStr = jsDate.toLocaleString('default', { month: 'long' });
      const weekStr = `${jsDate.getFullYear()}-W${getISOWeek(jsDate).toString().padStart(2, '0')}`;
      const dreamCount = countsMap[entry.date] || 0;

      // Force comma decimal separator
      const fmt = (num) => Number.isFinite(num) ? num.toString().replace('.', ',') : '0';

      tableData.push([
        entry.date,
        weekStr,
        monthStr,
        Number.isFinite(dreamCount) ? (dreamCount % 1 === 0 ? dreamCount.toString() : dreamCount.toFixed(1).replace('.', ',')) : '0',
        fmt(entry.d),
        fmt(entry.ld),
        fmt(entry.score)
      ]);
    });

    const table = body.insertTable(insertAt + 1, tableData);
    table.getRow(0).editAsText().setBold(true);

    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      const date = row.getCell(0).getText();
      const dateText = row.getCell(0).editAsText();
      if (headingLinks[date]) dateText.setLinkUrl(headingLinks[date]);

      for (let c = 0; c < row.getNumCells(); c++) {
        row.getCell(c).editAsText().setFontSize(10);
        row.getCell(c).getChild(0).asParagraph().setLineSpacing(1);
      }

      const ldVal = parseFloat(row.getCell(5).getText().replace(',', '.'));
      if (!isNaN(ldVal) && ldVal >= 1) row.getCell(5).editAsText().setForegroundColor('#008000');
      const totalVal = parseFloat(row.getCell(6).getText().replace(',', '.'));
      if (!isNaN(totalVal) && totalVal === 0) row.getCell(6).editAsText().setForegroundColor('#FF0000');
    }
  }

}

// ISO week helper
function getISOWeek(date) {
  const tmpDate = new Date(date.getTime());
  tmpDate.setHours(0, 0, 0, 0);
  tmpDate.setDate(tmpDate.getDate() + 4 - (tmpDate.getDay() || 7));
  const yearStart = new Date(tmpDate.getFullYear(), 0, 1);
  return Math.ceil((((tmpDate - yearStart) / 86400000) + 1) / 7);
}
