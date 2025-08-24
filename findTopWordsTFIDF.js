// Wrappers
function findTopWordsListTFIDF() {
  findTopItaWordsTFIDF("list");
}
function findTopWordsBarTFIDF() {
  findTopItaWordsTFIDF("bar");
}

/**
 * Find top words frequency, capping frequency at 1 per dream,
 * display them as histogram (bar) or as list (list)
 **/
function findTopWordsTFIDF(display = "bar") {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // Remove old Frequency section
  let inFreq = false, rem = [];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      const txt = p.getText().trim();
      const hd = p.getHeading();
      if (!inFreq && hd === DocumentApp.ParagraphHeading.HEADING1 && txt === "Frequency") {
        inFreq = true;
      } else if (inFreq && hd === DocumentApp.ParagraphHeading.HEADING2 && txt === "Word cloud") {
        break;
      } else if (inFreq && hd === DocumentApp.ParagraphHeading.HEADING1) {
        break;
      }
    }
    if (inFreq) rem.push(i);
  }
  rem.reverse().forEach(i => body.removeChild(body.getChild(i)));

  // Normalize each token and record its raw form
  const documents = [];
  const rawFreq = {};
  const aliasMap = {};

  let currentBlockTokens = [];

  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      const hd = p.getHeading();
      const text = p.getText();

      if (hd !== DocumentApp.ParagraphHeading.NORMAL) {
        // Heading encountered → close previous block if any
        if (currentBlockTokens.length > 0) {
          documents.push(currentBlockTokens);
          currentBlockTokens = [];
        }
        continue;
      }

      // Normal paragraph → accumulate tokens
      const tokens = tokenizeAndFilter(text);
      const normalizedTokens = tokens.map(w => normalizeWordEquivalents(w));

      // Count into rawFreq using the canonical form
      normalizedTokens.forEach(norm => {
        rawFreq[norm] = (rawFreq[norm] || 0) + 1;
      });

      // Record which tokens fed into each bucket
      tokens.forEach(orig => {
        const canon = normalizeWordEquivalents(orig);
        if (!aliasMap[canon]) aliasMap[canon] = new Set();
        aliasMap[canon].add(orig);
      });

      // Append normalized tokens to current block
      currentBlockTokens = currentBlockTokens.concat(normalizedTokens);
    }
  }

  // Push last block if any left
  if (currentBlockTokens.length > 0) {
    documents.push(currentBlockTokens);
  }

  const docCount = documents.length;
  const termFreq = {};
  const docFreq = {};

  documents.forEach((words, idx) => {
    const localFreq = {};
    words.forEach(w => {
      let base = w;
      const bases = getBaseForms(w, rawFreq);
      if (bases.length) base = bases[0];
      localFreq[base] = (localFreq[base] || 0) + 1;
    });

    for (let word in localFreq) {
      if (!termFreq[word]) termFreq[word] = Array(docCount).fill(0);
      termFreq[word][idx] = localFreq[word];
    }
  });

  for (let word in termFreq) {
    docFreq[word] = termFreq[word].filter(c => c > 0).length;
  }

  const tfidf = {};
  for (let word in termFreq) {
    const tfValues = termFreq[word];
    const idf = Math.log((1 + docCount) / (1 + docFreq[word])) + 1;
    // Max TF-IDF across all documents
    let maxTfidf = 0;
    tfValues.forEach(tf => {
      const val = tf * idf;
      if (val > maxTfidf) maxTfidf = val;
    });
    tfidf[word] = maxTfidf;
  }



  console.log("=== STRICT BIDIRECTIONAL GROUPING ===");

  // Build a map from word -> Set of base forms that exist in termFreq
  const baseMap = {};
  for (const word of Object.keys(termFreq)) {
    const bases = getBaseForms(word, rawFreq).filter(b => b !== word && termFreq[b]);
    baseMap[word] = new Set(bases);
  }

  // Find mutual pairs: word A and word B are mutual bases if A in baseMap[B] and B in baseMap[A]
  const mutualPairsSet = new Set();
  for (const word of Object.keys(baseMap)) {
    baseMap[word].forEach(base => {
      if (baseMap[base] && baseMap[base].has(word)) {
        const pairKey = orderSingularPlural(word, base);
        mutualPairsSet.add(pairKey);
      }
    });
  }

  // Calculate TF-IDF per pair
  const pairScores = {};
  const pairForms = {};
  mutualPairsSet.forEach(pairKey => {
    const [w1, w2] = pairKey.split("/");
    let pairTF = 0;
    let pairDF = 0;

    for (let docIdx = 0; docIdx < docCount; docIdx++) {
      let docHasWord = false;
      [w1, w2].forEach(word => {
        if (termFreq[word] && termFreq[word][docIdx] > 0) {
          pairTF += termFreq[word][docIdx];
          docHasWord = true;
        }
      });
      if (docHasWord) pairDF++;
    }

    const pairIDF = Math.log((1 + docCount) / (1 + pairDF)) + 1;
    pairScores[pairKey] = Math.round(pairTF * pairIDF * 100) / 100;

    const forms = new Set();
    [w1, w2].forEach(w => {
      if (aliasMap[w]) aliasMap[w].forEach(f => forms.add(f));
      else forms.add(w);
    });
    pairForms[pairKey] = forms;
  });


  const sorted = Object.entries(pairScores)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 30);


  // Find insertion point after "Score Table" if present
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

  // Insert page break if needed
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

  body.insertParagraph(insertAt++, "Frequency").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.insertParagraph(insertAt++, "Top words (TF-IDF)").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const maxScore = sorted.length > 0 ? sorted[0][1] : 1;

  if (display === "list") {
    sorted.forEach(([key, score], idx) => {
      const forms = [...pairForms[key]];
      let variants;
      if (forms.length === 2 && key.includes("/")) {
        variants = key;  // preserve order from pairKey
      } else {
        variants = forms.sort().join("/");
      }

      body.insertParagraph(insertAt++, `${idx + 1}. ${variants} (${score.toFixed(2)})`);
    });
  } else {
    sorted.forEach(([key, score]) => {
      const barLen = Math.round((score / maxScore) * 30);
      const bar = "█".repeat(barLen);
      const forms = [...pairForms[key]];
      let variants;
      if (forms.length === 2 && key.includes("/")) {
        variants = key;  // preserve order from pairKey
      } else {
        variants = forms.sort().join("/");
      }

      body.insertParagraph(insertAt++, `${variants.padEnd(25)} ${bar} (${score.toFixed(2)})`)
        .setFontFamily("Roboto Mono");
    });
  }


  return sorted;
}
