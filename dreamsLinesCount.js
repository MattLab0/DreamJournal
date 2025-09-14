
//------------------------------------------------------------------------------------


function analyzeDreamsForSelectedDay() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const ui = DocumentApp.getUi();
  const avgCharsPerLine = 95;

  const cursor = doc.getCursor();
  if (!cursor) {
    ui.alert("Place your cursor on a Heading 1 paragraph to analyze that day's dreams.");
    return;
  }

  // Get the paragraph containing the cursor by walking up parents if needed
  let element = cursor.getElement();
  while (element && element.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    element = element.getParent();
  }
  if (!element) {
    ui.alert("Cursor is not inside a paragraph. Place the cursor on a Heading 1 paragraph.");
    return;
  }
  const para = element.asParagraph();

  if (para.getHeading() !== DocumentApp.ParagraphHeading.HEADING1) {
    ui.alert("Please place the cursor exactly on a Heading 1 paragraph.");
    return;
  }

  // Find index of this Heading 1 paragraph in the body by matching text and heading
  const totalChildren = body.getNumChildren();
  let startIndex = -1;
  const paraText = para.getText();
  const paraHeading = para.getHeading();

  for (let i = 0; i < totalChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      if (p.getHeading() === paraHeading && p.getText() === paraText) {
        startIndex = i;
        break;
      }
    }
  }
  if (startIndex === -1) {
    ui.alert("Could not locate the selected Heading 1 paragraph in the document.");
    return;
  }

  // Collect dreams between this Heading 1 and the next Heading 1
  let results = [];
  let collecting = false;
  let currentTitle = '';
  let accumulatedText = '';

  for (let i = startIndex + 1; i < totalChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const p = child.asParagraph();
    const heading = p.getHeading();
    const text = p.getText().trim();

    if (heading === DocumentApp.ParagraphHeading.HEADING1) {
      // Reached next Heading 1 (next day), stop collecting
      break;
    }

    if (heading === DocumentApp.ParagraphHeading.HEADING2) {
      // If previously collecting, save that dream’s line count
      if (collecting) {
        const estLines = Math.ceil(accumulatedText.length / avgCharsPerLine);
        results.push(`${currentTitle}: ~${estLines} lines`);
        collecting = false;
        currentTitle = '';
        accumulatedText = '';
      }

      // Skip excluded titles but INCLUDE fragments as special entries (fixed score 0.5)
      const lowerText = text.toLowerCase();
      if (
        text !== '' &&
        !lowerText.startsWith("top words") &&
        !lowerText.startsWith("word cloud") &&
        lowerText !== "average and median"
      ) {
        if (lowerText.startsWith("fragment")) {
          // Mark fragment specially so later we know it's fixed 0.5 (not analyzed)
          results.push(`${text}: ~FRAGMENT`);
        } else {
          collecting = true;
          currentTitle = text;
        }
      }
      continue;
    }

    if (collecting) {
      accumulatedText += text + " ";
    }
  }

  // Save last dream if any
  if (collecting && currentTitle !== '') {
    const estLines = Math.ceil(accumulatedText.length / avgCharsPerLine);
    results.push(`${currentTitle}: ~${estLines} lines`);
  }

  if (results.length === 0) {
    ui.alert(`No valid dreams found under "${para.getText()}".`);
    return;
  }

  let avg = null, median = null, sd = null;
  let mad = null, scaledMad = null, iqr = null;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.TABLE) {
      const table = el.asTable();
      for (let r = 0; r < table.getNumRows(); r++) {
        const key = table.getRow(r).getCell(0).getText().toLowerCase();
        const val = parseFloat(table.getRow(r).getCell(1).getText());
        if (key === "average") avg = val;
        else if (key === "median") median = val;
        else if (key.includes("standard deviation")) sd = val;
        else if (key.includes("median absolute deviation (mad)")) mad = val;
        else if (key.includes("scaled mad")) scaledMad = val;
        else if (key.includes("interquartile range") || key === "iqr") iqr = val;
      }
      if (avg !== null && median !== null && sd !== null) break;  // could also check mad, scaledMad, iqr if needed
    }
  }

  if (avg === null || median === null || sd === null) {
    ui.alert("Could not read average, median or standard deviation from the statistics table.");
    return;
  }

  // Choose upper bound method here:
  const method = "median_scaledMad"; // options: "mean_sd", "median_scaledMad", "median_iqr"

  // Calculate upper bound based on chosen method
  let upperBound;
  if (method === "mean_sd") {
    upperBound = avg + sd;
  } else if (method === "median_scaledMad") {
    if (median !== null && scaledMad !== null) {
      upperBound = median + scaledMad;
    } else {
      ui.alert("Missing median or scaled MAD for chosen method.");
      return;
    }
  } else if (method === "median_iqr") {
    if (median !== null && iqr !== null) {
      upperBound = median + iqr / 1.35;  // approx conversion to std dev scale
    } else {
      ui.alert("Missing median or IQR for chosen method.");
      return;
    }
  } else {
    ui.alert("Unknown method selected for upper bound calculation.");
    return;
  }

  // Now you can use `upperBound` in your scoring or logic

  const lower = median;
  const upper = median + iqr;

  const clamp = (val, min, max) => Math.max(min, Math.min(max, val));
  const roundToHalf = (num) => Math.round(num * 2) / 2;

  const shortTitle = (title) => title.length > 40 ? title.slice(0, 37) + "…" : title;

  // --- compute per-dream continuous score for display; round each dream's score to nearest 0.5 BEFORE summing ---
  let totalScore = 0;
  const output = results.map(result => {
    // detect fragment entries first
    const fragMatch = result.match(/(.+?): ~FRAGMENT/);
    if (fragMatch) {
      const title = shortTitle(fragMatch[1]);
      const rawScore = 0.5; // fixed for fragments
      const rounded = roundToHalf(rawScore); // will be 0.5
      totalScore += rounded;
      return `${title}: → Score fixed: ${rawScore.toFixed(2)}`;
    }

    const match = result.match(/(.+?): ~(\d+) lines/);
    if (!match) return result;
    const title = shortTitle(match[1]);
    const lines = parseInt(match[2], 10);

    // compute continuous per-dream score (keep original logic)
    const denomCandidate = (upper - lower);
    const denom = (denomCandidate === 0 || !isFinite(denomCandidate)) ? 1 : denomCandidate;
    const rawScore = 1 + (lines - lower) / denom;

    // ROUND each dream's score to nearest 0.5 BEFORE adding to suggested total
    const rounded = roundToHalf(rawScore);
    totalScore += rounded;

    return `${title}: ~${lines} lines → Score: ${rawScore.toFixed(2)} (rounded ${rounded.toFixed(1)})`;
  });

  ui.alert(`Dreams for "${para.getText()}":\n\n${output.join("\n")}
Total Score (rounded): ${totalScore.toFixed(1)}
Scoring thresholds:
Short if < ${lower.toFixed(1)}
Long if > ${upper.toFixed(1)}
`);

  // Ask user to confirm updating D:
  const response = ui.alert(
    "Update D value?",
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    const headingText = para.getText();
    const newHeading = headingText.replace(/D:\s*\d+(\.\d+)?/, `D:${totalScore.toFixed(1)}`);
    para.setText(newHeading);
  }

  //---------------------------------------------
}

function updateAllDaysD() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const avgCharsPerLine = 95;

  const totalChildren = body.getNumChildren();

  for (let idx = 0; idx < totalChildren; idx++) {
    const element = body.getChild(idx);
    if (!element || element.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const para = element.asParagraph();
    if (para.getHeading() !== DocumentApp.ParagraphHeading.HEADING1) continue;

    // Skip statistics headings
    const headingText = para.getText().trim();
    if (/statistics/i.test(headingText)) continue;

    // --- Collect dreams under this heading ---
    let results = [];
    let collecting = false;
    let currentTitle = '';
    let accumulatedText = '';

    for (let i = idx + 1; i < totalChildren; i++) {
      const child = body.getChild(i);
      if (!child || child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

      const p = child.asParagraph();
      const heading = p.getHeading();
      const text = p.getText().trim();

      if (heading === DocumentApp.ParagraphHeading.HEADING1) break; // next day

      if (heading === DocumentApp.ParagraphHeading.HEADING2) {
        if (collecting) {
          const estLines = Math.ceil(accumulatedText.length / avgCharsPerLine);
          results.push(`${currentTitle}: ~${estLines} lines`);
          collecting = false;
          currentTitle = '';
          accumulatedText = '';
        }

        const lowerText = text.toLowerCase();
        if (
          text !== '' &&
          !lowerText.startsWith("top words") &&
          !lowerText.startsWith("word cloud") &&
          lowerText !== "average and median"
        ) {
          if (lowerText.startsWith("fragment")) {
            results.push(`${text}: ~FRAGMENT`);
          } else {
            collecting = true;
            currentTitle = text;
          }
        }
        continue;
      }

      if (collecting) accumulatedText += text + " ";
    }

    if (collecting && currentTitle !== '') {
      const estLines = Math.ceil(accumulatedText.length / avgCharsPerLine);
      results.push(`${currentTitle}: ~${estLines} lines`);
    }

    if (results.length === 0) continue;

    // --- Read statistics from table ---
    let avg = null, median = null, sd = null, mad = null, scaledMad = null, iqr = null;
    for (let i = 0; i < totalChildren; i++) {
      const el = body.getChild(i);
      if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
      const table = el.asTable();
      for (let r = 0; r < table.getNumRows(); r++) {
        const key = table.getRow(r).getCell(0).getText().toLowerCase();
        const val = parseFloat(table.getRow(r).getCell(1).getText());
        if (key === "average") avg = val;
        else if (key === "median") median = val;
        else if (key.includes("standard deviation")) sd = val;
        else if (key.includes("median absolute deviation (mad)")) mad = val;
        else if (key.includes("scaled mad")) scaledMad = val;
        else if (key.includes("interquartile range") || key === "iqr") iqr = val;
      }
      if (avg !== null && median !== null && sd !== null) break;
    }

    if (avg === null || median === null || sd === null) continue; // skip if stats missing

    const lower = median;
    const upper = median + iqr;
    const roundToHalf = (num) => Math.round(num * 2) / 2;
    const shortTitle = (title) => title.length > 40 ? title.slice(0, 37) + "…" : title;

    // --- Compute total score for this heading ---
    let totalScore = 0;
    results.forEach(result => {
      const fragMatch = result.match(/(.+?): ~FRAGMENT/);
      if (fragMatch) {
        totalScore += 0.5;
        return;
      }

      const match = result.match(/(.+?): ~(\d+) lines/);
      if (!match) return;
      const lines = parseInt(match[2], 10);

      const denomCandidate = (upper - lower);
      const denom = (denomCandidate === 0 || !isFinite(denomCandidate)) ? 1 : denomCandidate;
      const rawScore = 1 + (lines - lower) / denom;
      const rounded = roundToHalf(rawScore);
      totalScore += rounded;
    });

    // --- Update D: in the heading ---
    const newHeading = para.getText().replace(/D:\s*\d+(\.\d+)?/, `D:${totalScore.toFixed(1)}`);
    para.setText(newHeading);
  }
}
