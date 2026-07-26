const {
  Header,
  Footer,
  Paragraph,
  TextRun,
  AlignmentType,
  PageNumber,
  BorderStyle,
} = require("docx");
const { FONT } = require("./styles");

// Texte modifiable en un double-clic dans Word : c'est un simple TextRun,
// pas un champ — l'utilisateur le remplace par le nom réel du dossier pour
// chaque nouveau mémoire créé à partir du modèle.
function buildHeader() {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: {
          bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000", space: 4 },
        },
        children: [
          new TextRun({
            text: "[Nom du dossier]",
            bold: true,
            font: FONT,
            size: 20,
          }),
        ],
      }),
    ],
  });
}

// "Page X sur Y" via les champs OOXML natifs PAGE / NUMPAGES (PageNumber.CURRENT
// / PageNumber.TOTAL_PAGES) : ce sont de vrais champs Word, mis à jour
// automatiquement à chaque impression/aperçu, pas du texte figé.
function buildFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Page ", font: FONT, size: 20 }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 }),
          new TextRun({ text: " sur ", font: FONT, size: 20 }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 20 }),
        ],
      }),
    ],
  });
}

module.exports = { buildHeader, buildFooter };
