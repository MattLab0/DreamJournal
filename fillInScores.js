function fillInScores() {
  const body = DocumentApp.getActiveDocument().getBody();
  const paras = body.getParagraphs();

  const dateRe = /^Dreams?\s+(\d{2}\/\d{2}\/\d{4})/i;
  const uncertainDateRe = /^Dreams?\s+(\?{1,2}\/\d{2}\/\d{4})/i;
  const scoreRe = /D:(\d+(?:[.,]\d+)?)\s+LD:(\d+(?:[.,]\d+)?)(?:\s+Score:(\d+(?:[.,]\d+)?))?/i;

  paras.forEach(p => {
    if (p.getHeading() !== DocumentApp.ParagraphHeading.HEADING1) return;

    const text = p.getText();
    const dateMatch = text.match(dateRe);
    const uncertainMatch = text.match(uncertainDateRe);
    const scoreMatch = text.match(scoreRe);

    if ((dateMatch || uncertainMatch) && scoreMatch) {
      const date = dateMatch ? dateMatch[1] : uncertainMatch[1];
      const d = parseFloat(scoreMatch[1].replace(',', '.'));
      const ld = parseFloat(scoreMatch[2].replace(',', '.'));
      const score = d + ld;

      p.setText(`Dreams ${date} - D:${d} LD:${ld} Score:${score}`);
    }
  });
}
