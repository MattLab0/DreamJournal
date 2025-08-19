//Wrappers
function findTopWordsListDocFreq() {
  findTopItaWordsDocFreq("list");
}

function findTopWordsBarDocFreq() {
  findTopWordsDocFreq("bar");
}

/**
Find top words frequency, capping frequency at 1 per dream, display them as histogram (bar) or as list (list)
**/
function findTopWordsDistinct(display = "bar") {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // Remove old Frequency section
  let inFreq = false, rem = [];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      if (!inFreq && p.getHeading() === DocumentApp.ParagraphHeading.HEADING1 && p.getText().trim() === "Frequency") {
        inFreq = true;
      } else if (inFreq && p.getHeading() === DocumentApp.ParagraphHeading.HEADING2 && p.getText().trim() === "Word cloud") {
        break;
      } else if (inFreq && p.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
        break;
      }
    }
    if (inFreq) rem.push(i);
  }
  rem.reverse().forEach(i => body.removeChild(body.getChild(i)));

  const blocksWords = []; // Stores unique words per text block
  const rawFreq = {};
  const aliasMap = {};
  let currentBlock = new Set();

  // Process document by blocks (consecutive NORMAL paragraphs between any two headings, at any level)
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      const heading = p.getHeading();
      const text = p.getText().trim();

      // When hitting a heading, close current block
      if (heading !== DocumentApp.ParagraphHeading.NORMAL) {
        if (currentBlock.size > 0) {
          blocksWords.push(currentBlock);
          currentBlock = new Set();
        }
        continue;
      }

      if (!text.startsWith("Top words")) {
        // Tokenize and normalize
        const tokens = tokenizeAndFilter(text);
        const normalized = tokens.map(w => normalizeWordEquivalents(w));

        // Count into rawFreq
        normalized.forEach(norm => {
          rawFreq[norm] = (rawFreq[norm] || 0) + 1;
        });

        // Record raw -> canonical in aliasMap
        tokens.forEach(orig => {
          const canon = normalizeWordEquivalents(orig);
          if (!aliasMap[canon]) aliasMap[canon] = new Set();
          aliasMap[canon].add(orig);
        });

        // Add the canonical form into the block set
        normalized.forEach(norm => currentBlock.add(norm));
      }
    }
  }

  // Push what is left in currentBlock
  if (currentBlock.size > 0) {
    blocksWords.push(currentBlock);
  }


  // === STRICT BIDIRECTIONAL BASE GROUPING ===

  // Build base form map
  const baseMap = {};
  Object.keys(rawFreq).forEach(word => {
    const bases = getBaseForms(word, rawFreq).filter(b => b !== word && rawFreq[b]);
    baseMap[word] = new Set(bases);
  });

  // Identify mutual base pairs
  const mutualPairsSet = new Set();
  for (const word in baseMap) {
    baseMap[word].forEach(base => {
      if (baseMap[base] && baseMap[base].has(word)) {
        const key = orderSingularPlural(word, base);
        mutualPairsSet.add(key);
      }
    });
  }

  // Create lookup: word -> groupKey, and groupKey -> display forms
  const variantToGroup = {};
  const groupForms = {};

  mutualPairsSet.forEach(pairKey => {
    const [w1, w2] = pairKey.split("/");
    const displaySet = new Set();

    if (aliasMap[w1]) aliasMap[w1].forEach(f => displaySet.add(f)); else displaySet.add(w1);
    if (aliasMap[w2]) aliasMap[w2].forEach(f => displaySet.add(f)); else displaySet.add(w2);

    groupForms[pairKey] = displaySet;
    variantToGroup[w1] = pairKey;
    variantToGroup[w2] = pairKey;
  });

  // For words not in a mutual pair, map them to themselves
  Object.keys(rawFreq).forEach(word => {
    if (!variantToGroup[word]) {
      const displaySet = aliasMap[word] ? new Set(aliasMap[word]) : new Set([word]);
      const key = [...displaySet].sort().join("/");
      variantToGroup[word] = key;
      groupForms[key] = displaySet;
    }
  });


  // Compute frequency per group
  const groupDocFreq = {};
  blocksWords.forEach(blockSet => {
    const seenGroups = new Set();
    blockSet.forEach(word => {
      const grp = variantToGroup[word] || word;
      seenGroups.add(grp);
    });
    seenGroups.forEach(grp => {
      groupDocFreq[grp] = (groupDocFreq[grp] || 0) + 1;
    });
  });

  // Sort and get top 30
  const sorted = Object.entries(groupDocFreq)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 30);

  // Find insertion point
  let insertAt = body.getNumChildren();
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      if (p.getHeading() === DocumentApp.ParagraphHeading.HEADING1 && p.getText().trim() === "Score Table") {
        insertAt = i + 2;
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
  body.insertParagraph(insertAt++, "Frequency").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.insertParagraph(insertAt++, "Top words distinct").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const maxCount = sorted[0][1];

  // Render as list or bar
  // Render with proper alias grouping
  if (display === "list") {
    sorted.forEach(([groupKey, count], idx) => {
      const forms = [...(groupForms[groupKey] || new Set([groupKey]))];
      let variants;
      if (forms.length === 2 && groupKey.includes("/")) {
        variants = groupKey;  // keep custom order from groupKey
      } else {
        variants = forms.sort().join("/");
      }

      body.insertParagraph(insertAt++, `${idx + 1}. ${variants} (${count})`);
    });
  } else {
    sorted.forEach(([groupKey, count]) => {
      const forms = [...(groupForms[groupKey] || new Set([groupKey]))];
      let variants;
      if (forms.length === 2 && groupKey.includes("/")) {
        variants = groupKey;  // keep custom order from groupKey
      } else {
        variants = forms.sort().join("/");
      }

      const barLength = Math.round((count / maxCount) * 30);
      const bar = "█".repeat(barLength);
      body.insertParagraph(insertAt++, `${variants.padEnd(25)} ${bar} (${count})`)
        .setFontFamily("Roboto Mono");
    });
  }


  return sorted;
}
