// Wrappers
function findTopWordsList() {
  findTopWords("list");
}

function findTopWordsBar() {
  findTopWords("bar");
}

/**
Find top words raw frequency, display them as histogram (bar) or as list (list)
**/
function findTopWords(display = "bar") {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // Remove old Frequency section
  let inFreq = false, toRemove = [];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      const txt = p.getText().trim();
      const h = p.getHeading();
      if (!inFreq && h === DocumentApp.ParagraphHeading.HEADING1 && txt === "Frequency") {
        inFreq = true;
        toRemove.push(i);
        continue;
      }
      if (inFreq) {
        // stop when next major section or Word cloud
        if ((h === DocumentApp.ParagraphHeading.HEADING1 && txt !== "Frequency")
          || (h === DocumentApp.ParagraphHeading.HEADING2 && txt === "Word cloud")) {
          break;
        }
        toRemove.push(i);
      }
    } else if (inFreq) {
      toRemove.push(i);
    }
  }
  toRemove.reverse().forEach(idx => body.removeChild(body.getChild(idx)));

  // Extract all paragraphs (excluding “Top words”)
  let text = "";
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      const t = p.getText();
      if (
        p.getHeading() === DocumentApp.ParagraphHeading.NORMAL &&
        !t.startsWith("Top words")
      ) {
        text += t + " ";
      }
    }
  }

  // Tokenize and filter out stop words
  const tokens = tokenizeAndFilter(text);

  // Build rawFreq and aliasMap (normalized -> original forms)
  const rawFreq = {};
  const aliasMap = {};
  tokens.forEach(w => {
    const norm = normalizeWordEquivalents(w);
    rawFreq[norm] = (rawFreq[norm] || 0) + 1;
    aliasMap[norm] = aliasMap[norm] || new Set();
    aliasMap[norm].add(w);
  });

  // Bidirectional groping to avoid inexhistent pairs
  const baseMap = {};
  for (const word of Object.keys(rawFreq)) {
    const bases = getBaseForms(word, rawFreq).filter(b => b !== word && rawFreq[b]);
    baseMap[word] = new Set(bases);
  }

  // Find mutual pairs: A <-> B
  const groupMap = {};
  for (const word of Object.keys(baseMap)) {
    baseMap[word].forEach(base => {
      if (baseMap[base] && baseMap[base].has(word)) {
        const key = orderSingularPlural(word, base);
        groupMap[word] = key;
        groupMap[base] = key;
      }
    });
  }


  // Add standalone words
  for (const word of Object.keys(rawFreq)) {
    if (!groupMap[word]) groupMap[word] = word;
  }

  // Collect display variants
  const groupForms = {};
  const groupCounts = {};
  for (const word of Object.keys(rawFreq)) {
    const key = groupMap[word];
    if (!groupForms[key]) groupForms[key] = new Set();
    if (!groupCounts[key]) groupCounts[key] = 0;
    groupForms[key].add(...(aliasMap[word] || [word]));
    groupCounts[key] += rawFreq[word];
  }

  // Create final grouped object
  const grouped = {};
  for (const key of Object.keys(groupCounts)) {
    const forms = [...groupForms[key]];
    if (forms.length === 2 && key.includes("/")) {
      // Use the group key directly, because it was ordered by orderSingularPlural
      grouped[key] = { display: key, count: groupCounts[key] };
    } else {
      // Otherwise, sort all forms alphabetically (for single words or synonyms)
      const disp = forms.sort().join("/");
      grouped[key] = { display: disp, count: groupCounts[key] };
    }

  }


  // Sort and take top 30
  const topWords = Object.entries(grouped)
    .sort((a, b) => b[1].count - a[1].count)
    .slice(0, 30);

  // Find insertion point after the Score Table
  let insertAt = body.getNumChildren();
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      if (p.getHeading() === DocumentApp.ParagraphHeading.HEADING1
        && p.getText().trim() === "Score Table") {
        // skip over the heading plus its table (anything until next paragraph)
        let j = i + 1;
        while (j < body.getNumChildren() &&
          body.getChild(j).getType() !== DocumentApp.ElementType.PARAGRAPH) {
          j++;
        }
        insertAt = j;
        break;
      }
    }
  }

  // Insert a single page break if not present
  let needBreak = true;
  if (insertAt > 0) {
    const prev = body.getChild(insertAt - 1);
    if (prev.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const pp = prev.asParagraph();
      for (let k = 0; k < pp.getNumChildren(); k++) {
        if (pp.getChild(k).getType() === DocumentApp.ElementType.PAGE_BREAK) {
          needBreak = false;
          break;
        }
      }
    } else if (prev.getType() === DocumentApp.ElementType.PAGE_BREAK) {
      needBreak = false;
    }
  }
  if (needBreak) {
    body.insertPageBreak(insertAt);
    insertAt++;
  }

  // Insert “Frequency” (H1) and “Top words” (H2)
  body.insertParagraph(insertAt++, "Frequency")
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.insertParagraph(insertAt++, "Top words")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  // Render as list or bar
  const maxCount = topWords[0][1].count;


  if (display === "list") {
    topWords.forEach(([_, data], idx) => {
      body.insertParagraph(insertAt++, `${idx + 1}. ${data.display} (${data.count})`);
    });
  } else {
    topWords.forEach(([_, data]) => {
      const barLength = Math.round((data.count / maxCount) * 30);
      const bar = "█".repeat(barLength);
      body.insertParagraph(
        insertAt++,
        `${data.display.padEnd(25)} ${bar} (${data.count})`
      ).setFontFamily("Roboto Mono");
    });
  }

  return topWords;
}

