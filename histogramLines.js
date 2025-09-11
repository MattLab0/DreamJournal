//------------------------------------------------------------------------------------

function insertHistogram(body, startIndex, data) {
  const binSize = 1;
  const titleText = "Dream Length Distribution (lines)";

  // --- Build histogram text ---
  const maxVal = data.length ? Math.max(...data) : 0;

  // special handling for binSize === 1: single-number labels and cap at 50 (50+)
  const capSingleBin = 50;
  let binsCount = Math.max(1, Math.ceil((maxVal + 1) / binSize));
  if (binSize === 1 && binsCount > capSingleBin + 1) {
    binsCount = capSingleBin + 1; // last bin = 50+
  }

  const bins = Array(binsCount).fill(0);
  data.forEach(val => {
    let idx = Math.floor(val / binSize);
    if (binSize === 1 && idx > capSingleBin) idx = capSingleBin; // fold into 50+
    if (idx >= 0 && idx < bins.length) bins[idx]++;
  });

  const maxCount = Math.max(...bins);
  const maxBarLength = 30;

  // label width: narrow for binSize===1 so the bar is near the number;
  // otherwise use the range width (e.g. " 0- 1") to keep alignment.
  const sampleRange = `${String(0).padStart(2)}-${String(binSize === 1 ? 0 : (binSize - 1)).padEnd(2)}`;
  const labelWidth = (binSize === 1) ? 3 : sampleRange.length;

  const lines = bins.map((count, i) => {
    const start = i * binSize;
    const end = start + binSize - 1;

    let label;
    if (binSize === 1) {
      // single-number label, overflow folded into "50+"
      label = (i === capSingleBin) ? `${capSingleBin}+` : String(start).padStart(2);
      // right-pad to narrow labelWidth (usually 3) so bar is close
      if (label.length < labelWidth) label = label + ' '.repeat(labelWidth - label.length);
    } else {
      // normal range label like " 0- 1"
      label = `${String(start).padStart(2)}-${String(end).padEnd(2)}`;
    }

    const barLength = maxCount > 0 ? Math.round((count / maxCount) * maxBarLength) : 0;
    const bar = "█".repeat(barLength);
    const countText = count === 0 ? "" : ` (${count})`;
    return `${label} | ${bar}${countText}`;
  });

  const histoText = lines.join("\n");




  // --- Helper: ensure there's a paragraph at/near index containing given text.
  // Tries to insert, then finds the nearest paragraph and returns it. Falls back to appendParagraph.
  function ensureParagraphAt(body, index, text) {
    index = Math.max(0, Math.min(index, body.getNumChildren()));
    try {
      // try to insert; in some doc states insertParagraph might return null or throw — that's ok
      body.insertParagraph(index, text);
    } catch (e) {
      // ignore and fall through to scanning
    }

    // scan forward a little bit for the inserted paragraph (and also check the exact index)
    const maxScan = Math.min(body.getNumChildren(), index + 6);
    for (let k = index; k < maxScan; k++) {
      const child = body.getChild(k);
      if (!child) continue;
      if (child.getType && child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const p = child.asParagraph();
        try {
          if (p.getText() !== text) p.setText(text);
        } catch (e) {
          // if setText fails, ignore — we'll still return the paragraph
        }
        return p;
      }
    }

    // scan backward a bit (in case insertion point landed before)
    const minScan = Math.max(0, index - 3);
    for (let k = index - 1; k >= minScan; k--) {
      const child = body.getChild(k);
      if (!child) continue;
      if (child.getType && child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const p = child.asParagraph();
        try {
          if (p.getText() !== text) p.setText(text);
        } catch (e) { }
        return p;
      }
    }

    // final fallback: append at end
    const appended = body.appendParagraph(text);
    return appended;
  }

  // --- Find existing Heading2 title and its histogram paragraph safely ---
  let titlePara = null;
  let titleIndex = -1;
  let histoPara = null;

  const total = body.getNumChildren();
  for (let i = startIndex; i < total; i++) {
    const el = body.getChild(i);
    if (!el) continue;
    if (el.getType && el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      const heading = p.getHeading && p.getHeading();
      const text = (p.getText && p.getText().trim()) || "";
      if (heading === DocumentApp.ParagraphHeading.HEADING2 && text === titleText) {
        titlePara = p;
        titleIndex = i;
        // check next child for histogram paragraph
        if (i + 1 < total) {
          const nextEl = body.getChild(i + 1);
          if (nextEl && nextEl.getType && nextEl.getType() === DocumentApp.ElementType.PARAGRAPH) {
            histoPara = nextEl.asParagraph();
          }
        }
        break;
      }
      // if we hit the next Heading1 after startIndex, stop searching (we shouldn't pass that)
      if (heading === DocumentApp.ParagraphHeading.HEADING1 && i > startIndex) break;
    }
  }
  // --- If title missing: insert (robust) and set heading2 ---
  if (!titlePara) {
    // ensure page break before inserting the title (but don't duplicate)
    let needBreak = true;
    let insertAt = startIndex + 1;
    if (insertAt > 0) {
      const prev = body.getChild(insertAt - 1);
      if (prev && prev.getType && prev.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const pp = prev.asParagraph();
        for (let k = 0; k < pp.getNumChildren(); k++) {
          const child = pp.getChild(k);
          if (child && child.getType && child.getType() === DocumentApp.ElementType.PAGE_BREAK) {
            needBreak = false;
            break;
          }
        }
      } else if (prev && prev.getType && prev.getType() === DocumentApp.ElementType.PAGE_BREAK) {
        needBreak = false;
      }
    }
    if (needBreak) {
      body.insertPageBreak(insertAt);
      insertAt++;
    }

    const p = ensureParagraphAt(body, insertAt, titleText); // ensure paragraph exists and has titleText
    try {
      p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    } catch (e) {
      // if setHeading fails for some reason, ignore (we still have the paragraph)
    }
    titlePara = p;
    titleIndex = body.getChildIndex(titlePara);
    // ensure histogram paragraph below
    const histInsIndex = Math.max(0, titleIndex + 1);
    const h = ensureParagraphAt(body, histInsIndex, histoText);
    try { h.setFontFamily("Roboto Mono"); h.setFontSize(9); } catch (e) { }
    return; // done
  }

  // --- Title found: update it and update/insert histogram below ---
  // ensure there's a page break immediately before the title paragraph
  let needBreak2 = true;
  let titlePos = body.getChildIndex(titlePara);
  if (titlePos > 0) {
    const prev2 = body.getChild(titlePos - 1);
    if (prev2 && prev2.getType && prev2.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const pp2 = prev2.asParagraph();
      for (let k = 0; k < pp2.getNumChildren(); k++) {
        const child = pp2.getChild(k);
        if (child && child.getType && child.getType() === DocumentApp.ElementType.PAGE_BREAK) {
          needBreak2 = false;
          break;
        }
      }
    } else if (prev2 && prev2.getType && prev2.getType() === DocumentApp.ElementType.PAGE_BREAK) {
      needBreak2 = false;
    }
  }
  if (needBreak2) {
    body.insertPageBreak(titlePos);
    // titlePara index shifted; recompute
    titleIndex = body.getChildIndex(titlePara);
  } else {
    titleIndex = titlePos;
  }

  try {
    titlePara.setText(titleText);
    titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  } catch (e) {
    // ignore, continue to try updating histogram
  }

  if (histoPara) {
    try {
      histoPara.setText(histoText);
      histoPara.setFontFamily("Roboto Mono");
      histoPara.setFontSize(9);
    } catch (e) {
      // fallback: replace by inserting new paragraph after title
      const newH = ensureParagraphAt(body, body.getChildIndex(titlePara) + 1, histoText);
      try { newH.setFontFamily("Roboto Mono"); newH.setFontSize(9); } catch (e) { }
    }
  } else {
    // insert a new histogram paragraph right after the title (robust)
    const insertIndex = Math.max(0, body.getChildIndex(titlePara) + 1);
    const newH = ensureParagraphAt(body, insertIndex, histoText);
    try { newH.setFontFamily("Roboto Mono"); newH.setFontSize(9); } catch (e) { }
  }

}

//---------------------------------------------------

function updateHistogramOnly() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // reuse your existing counter logic but stop before writing the table
  const avgCharsPerLine = 95;
  let collecting = false;
  let accumulatedText = '';
  let dreamLineCounts = [];

  const totalElements = body.getNumChildren();
  for (let i = 0; i < totalElements; i++) {
    const element = body.getChild(i);
    if (element.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const para = element.asParagraph();
    const heading = para.getHeading();
    const text = para.getText().trim();

    if (heading === DocumentApp.ParagraphHeading.HEADING1 || heading === DocumentApp.ParagraphHeading.HEADING2) {
      if (collecting) {
        dreamLineCounts.push(Math.ceil(accumulatedText.length / avgCharsPerLine));
        collecting = false;
        accumulatedText = '';
      }
      if (
        heading === DocumentApp.ParagraphHeading.HEADING2 &&
        text !== '' &&
        !text.toLowerCase().startsWith("fragment") &&
        !text.toLowerCase().startsWith("top words") &&
        !text.toLowerCase().startsWith("word cloud") &&
        !text.toLowerCase().startsWith("average and median")
      ) {
        collecting = true;
      }
      continue;
    }
    if (collecting) accumulatedText += text;
  }
  if (collecting) {
    dreamLineCounts.push(Math.ceil(accumulatedText.length / avgCharsPerLine));
  }

  if (dreamLineCounts.length === 0) {
    DocumentApp.getUi().alert("No valid dreams found to build histogram.");
    return;
  }

  // find "Statistics" heading for placement
  let statsHeadingIndex = -1;
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const para = el.asParagraph();
    if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING1 &&
      para.getText().trim().toLowerCase() === "statistics") {
      statsHeadingIndex = i;
      break;
    }
  }
  if (statsHeadingIndex === -1) {
    DocumentApp.getUi().alert("No 'Statistics' heading found. Please add one.");
    return;
  }

  // insert histogram right after statistics section (independent of table)
  insertHistogram(body, statsHeadingIndex + 1, dreamLineCounts);
}

