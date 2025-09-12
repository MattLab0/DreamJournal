function generateStatisticsDreamLines() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const avgCharsPerLine = 95;

  let collecting = false;
  let currentTitle = '';
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
        const estLines = Math.ceil(accumulatedText.length / avgCharsPerLine);
        dreamLineCounts.push(estLines);
        collecting = false;
        currentTitle = '';
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
        currentTitle = text;
      }
      continue;
    }

    if (collecting) {
      accumulatedText += text;
    }
  }

  if (collecting) {
    const estLines = Math.ceil(accumulatedText.length / avgCharsPerLine);
    dreamLineCounts.push(estLines);
  }

  if (dreamLineCounts.length === 0) {
    DocumentApp.getUi().alert("No valid dreams found to calculate statistics.");
    return;
  }

  const totalDreams = dreamLineCounts.length;
  const sumLines = dreamLineCounts.reduce((a, b) => a + b, 0);
  const average = sumLines / totalDreams;

  dreamLineCounts.sort((a, b) => a - b);

  // Median
  let median;
  if (totalDreams % 2 === 1) {
    median = dreamLineCounts[(totalDreams - 1) / 2];
  } else {
    median = (dreamLineCounts[totalDreams / 2 - 1] + dreamLineCounts[totalDreams / 2]) / 2;
  }

  // Standard deviation
  const variance = dreamLineCounts.reduce((acc, val) => acc + Math.pow(val - average, 2), 0) / totalDreams;
  const sd = Math.sqrt(variance);

  // --- Calculate MAD (Median Absolute Deviation) ---
  const absoluteDeviations = dreamLineCounts.map(val => Math.abs(val - median));
  absoluteDeviations.sort((a, b) => a - b);

  let mad;
  if (totalDreams % 2 === 1) {
    mad = absoluteDeviations[(totalDreams - 1) / 2];
  } else {
    mad = (absoluteDeviations[totalDreams / 2 - 1] + absoluteDeviations[totalDreams / 2]) / 2;
  }
  // Scale MAD to be comparable to SD under normal distribution
  const scaledMAD = mad * 1.4826;

  // --- Calculate IQR (Interquartile Range) ---
  // Q1 = 25th percentile, Q3 = 75th percentile
  const q1 = percentile(dreamLineCounts, 25);
  const q3 = percentile(dreamLineCounts, 75);
  const iqr = q3 - q1;

  const ui = DocumentApp.getUi();

  let statsHeadingIndex = -1;
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const para = el.asParagraph();
    if (
      para.getHeading() === DocumentApp.ParagraphHeading.HEADING1 &&
      para.getText().trim().toLowerCase() === "statistics"
    ) {
      statsHeadingIndex = i;
      break;
    }
  }

  if (statsHeadingIndex === -1) {
    ui.alert("No Heading 1 titled 'statistics' found. Please add one to your document.");
    return;
  }

  let avgMedHeadingIndex = -1;
  for (let i = statsHeadingIndex + 1; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const para = el.asParagraph();
    if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING1) break;
    if (
      para.getHeading() === DocumentApp.ParagraphHeading.HEADING2 &&
      para.getText().trim().toLowerCase() === "average and median"
    ) {
      avgMedHeadingIndex = i;
      break;
    }
  }

  if (avgMedHeadingIndex === -1) {
    avgMedHeadingIndex = statsHeadingIndex + 1;
    body.insertParagraph(avgMedHeadingIndex, "average and median").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  }

  let tableIndex = -1;
  for (let i = avgMedHeadingIndex + 1; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.TABLE) {
      tableIndex = i;
      break;
    } else if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = el.asParagraph();
      if (para.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) break;
    }
  }

  // Helper rounding function (if not already declared)
  const round = (num) => num.toFixed(1);

  // Show stats summary alert first (OK only)
  ui.alert(
    "Statistics calculated",
    `Total dreams: ${totalDreams}
Average lines: ${round(average)}
Median lines: ${round(median)}
Standard Deviation: ${round(sd)}
Median Absolute Deviation (MAD): ${round(mad)}
Scaled MAD (approx. SD equivalent): ${round(scaledMAD)}
Interquartile Range (IQR): ${round(iqr)}
Average + SD: ${round(average + sd)}
Average - SD: ${round(Math.max(average - sd, 0))}`,
    ui.ButtonSet.OK
  );

  const updateResponse = ui.alert(
    "Update statistics table?",
    "This will change score system for new dreams.\n" +
    'Recommended to run "Update Dream Score on All Days".',
    ui.ButtonSet.YES_NO
  );


  if (updateResponse === ui.Button.YES) {
    if (tableIndex === -1) {
      const table = body.insertTable(avgMedHeadingIndex + 1, [
        ["Average", round(average)],
        ["Median", round(median)],
        ["Standard Deviation", round(sd)],
        ["Median Absolute Deviation (MAD)", round(mad)],
        ["Scaled MAD", round(scaledMAD)],
        ["Interquartile Range (IQR)", round(iqr)],
        ["Average + SD", round(average + sd)],
        ["Average - SD", round(Math.max(average - sd, 0))],
      ]);
    } else {
      const table = body.getChild(tableIndex).asTable();
      while (table.getNumRows() < 8) table.appendTableRow();
      for (let r = 0; r < 8; r++) {
        const row = table.getRow(r);
        while (row.getNumCells() < 2) row.appendTableCell("");
      }

      table.getRow(0).getCell(0).setText("Average");
      table.getRow(0).getCell(1).setText(round(average));

      table.getRow(1).getCell(0).setText("Median");
      table.getRow(1).getCell(1).setText(round(median));

      table.getRow(2).getCell(0).setText("Standard Deviation");
      table.getRow(2).getCell(1).setText(round(sd));

      table.getRow(3).getCell(0).setText("Median Absolute Deviation (MAD)");
      table.getRow(3).getCell(1).setText(round(mad));

      table.getRow(4).getCell(0).setText("Scaled MAD");
      table.getRow(4).getCell(1).setText(round(scaledMAD));

      table.getRow(5).getCell(0).setText("Interquartile Range (IQR)");
      table.getRow(5).getCell(1).setText(round(iqr));

      table.getRow(6).getCell(0).setText("Average + SD");
      table.getRow(6).getCell(1).setText(round(average + sd));

      table.getRow(7).getCell(0).setText("Average - SD");
      table.getRow(7).getCell(1).setText(round(Math.max(average - sd, 0)));
    }

    insertHistogram(body, tableIndex === -1 ? avgMedHeadingIndex + 1 : tableIndex, dreamLineCounts);

    ui.alert("Statistics updated successfully.");
  } else {
    ui.alert("Statistics update canceled.");
  }
}