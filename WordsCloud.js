// Wrappers
function createWordCloudDocFreq() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Display in Bar Chart?',
    '✔ YES → Bar\n✖ NO → List',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response === ui.Button.YES || response === ui.Button.NO) {
    const original = findTopWords;
    const mode = response === ui.Button.YES ? "bar" : "list";

    findTopWords = function() {
      return findTopWordsDistinct(mode).map(([key, count]) => {
        return [key, { display: key, count: count }];
      });
    };

    createWordCloud(mode);
    findTopWords = original;
  }
}

function createWordCloudTFIDF() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Display in Bar Chart?',
    '✔ YES → Bar\n✖ NO → List',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response === ui.Button.YES || response === ui.Button.NO) {
    const original = findTopWords;
    const mode = response === ui.Button.YES ? "bar" : "list";

    findTopWords = function() {
      return findTopWordsTFIDF(mode).map(([key, score]) => {
        return [key, { display: key, count: score }];
      });
    };

    createWordCloud(mode);
    findTopWords = original;
  }
}

function createWordCloud(mode = null) {
  // If createWordCloud is called directly, show the prompt (mode = null)
  if (mode === null) {
    const ui = DocumentApp.getUi();
    const response = ui.alert(
      'Display in Bar Chart?',
      '✔ YES → Bar\n✖ NO → List',
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (response !== ui.Button.YES && response !== ui.Button.NO) {
      return;
    }
    mode = response === ui.Button.YES ? "bar" : "list";
  }

  // Grab the raw 
  const raw = findTopWords(mode);

  // Copy & shorten every display
  const shortened = raw.map(([key, data]) => {
    return [key, {
      display: shortenSingPlu(data.display),
      count:   data.count
    }];
  });

  // Assign colors in sequence
  const colors = ['#4285F4', '#EA4335', '#FBBC05', '#34A853', '#673AB7'];
  let ci = 0;
  const withColors = shortened.map(([key, d]) => {
    return [key, {
      display: d.display,
      count:   d.count,
      color:   colors[(ci++) % colors.length]
    }];
  });

  // Fire up HTML dialog
  const tpl = HtmlService.createTemplateFromFile('WordCloud');
  tpl.topWords = withColors;
  const html = tpl.evaluate().setWidth(760).setHeight(600);
  DocumentApp.getUi().showModalDialog(html, 'Generating Word Cloud…');
}

function insertWordCloudImage(dataUrl) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // Remove any existing Word Cloud (H2) and its image
  let inCloud = false, toRemove = [];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      if (!inCloud
          && p.getHeading() === DocumentApp.ParagraphHeading.HEADING2
          && p.getText().trim() === 'Word cloud') {
        inCloud = true;
        toRemove.push(i);
      }
      else if (inCloud
         && (p.getHeading() === DocumentApp.ParagraphHeading.HEADING2
             || p.getHeading() === DocumentApp.ParagraphHeading.HEADING1)) {
        break;
      }
      else if (inCloud) {
        toRemove.push(i);
      }
    }
  }
  toRemove.reverse().forEach(idx => body.removeChild(body.getChild(idx)));

  // Find insertion point
  let insertAt = body.getNumChildren();
  let foundTopWords = false;
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      const text = p.getText().trim();
      if (!foundTopWords 
          && p.getHeading() === DocumentApp.ParagraphHeading.HEADING2
          && text.startsWith('Top words')) {
        foundTopWords = true;
        continue;
      }
      if (foundTopWords) {
        if (p.getHeading() === DocumentApp.ParagraphHeading.HEADING2 || 
            p.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
          insertAt = i;
          break;
        }
      }
    }
  }

  // Insert Word Cloud heading
  body.insertParagraph(insertAt, 'Word cloud')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  // Insert the image
  const b64 = dataUrl.split(',')[1];
  const blob = Utilities.newBlob(
    Utilities.base64Decode(b64),
    'image/png',
    'wordcloud.png'
  );
  const imgPara = body.insertParagraph(insertAt + 1, '');
  imgPara.appendInlineImage(blob);

  // Center the paragraph
  imgPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

