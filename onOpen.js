function onOpen() {
  DocumentApp.getUi()
    .createMenu('☰ My Scripts')
    .addItem('☁ Word Cloud (Overall)', 'createWordCloud')
    .addItem('☁ Word Cloud (Distinct Dreams)', 'createWordCloudDocFreq')
    .addItem('☁ Word Cloud (TF-IDF)', 'createWordCloudTFIDF')
    .addSeparator()
    .addItem('📊 Top Words (Overall)', 'promptTopWords')         
    .addItem('📊 Top Words (Distinct Dreams)', 'promptTopWordsDistinct')
    .addItem('📊 top Words (TF-IDF)', 'promptTopWordsTFIDF')
    .addSeparator()
    .addItem('✏️ Fill In Scores', 'fillInScores')
    .addItem('⬜ Generate Score Table', 'promptGenerateScoreTable')
    .addSeparator()
    .addItem('⬜ Update Statistic Dream Lines', 'generateStatisticsDreamLines')
    .addItem('📊 Update Histogram Dream Lines', 'updateHistogramOnly')
    .addItem('🗓️ Get Dream Score on Selected Date', 'analyzeDreamsForSelectedDay')
    .addSeparator()
    .addItem('🗓️🗓️🗓️ Update Dream Score on All Days', 'updateAllDaysD')
    .addToUi();
}

// Top Words
function promptTopWords() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Display in Bar Chart?',
    '✔ YES → Bar\n✖ NO → List',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response === ui.Button.YES) {
    findTopWords("bar");
  } else if (response === ui.Button.NO) {
    findTopWords("list");
  }
}

// Distinct Words
function promptTopWordsDistinct() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Display in Bar Chart?',
    '✔ YES → Bar\n✖ NO → List',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response === ui.Button.YES) {
    findTopWordsDistinct("bar");
  } else if (response === ui.Button.NO) {
    findTopWordsDistinct("list");
  }
}

// TF-IDF Words
function promptTopWordsTFIDF() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Display in Bar Chart?',
    '✔ YES → Bar\n✖ NO → List',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response === ui.Button.YES) {
    findTopWordsTFIDF("bar");
  } else if (response === ui.Button.NO) {
    findTopWordsTFIDF("list");
  }
}

// Score table
function promptGenerateScoreTable() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Skip rows with score less than 0.25?',
    'Update the index with new dreams before!',
    ui.ButtonSet.YES_NO_CANCEL
  );

  const skipZeroScores = response === ui.Button.YES;
  generateScoreTable(skipZeroScores);
}
