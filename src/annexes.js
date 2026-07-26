const { Paragraph, TextRun, HeadingLevel } = require("docx");

function buildAnnexesChapter() {
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun("Annexes")],
    }),
    new Paragraph({
      style: "Normal",
      children: [
        new TextRun(
          "EXEMPLE — Insérer ici les annexes du mémoire (documents complets, tableaux détaillés, etc.)."
        ),
      ],
    }),
  ];
}

module.exports = { buildAnnexesChapter };
