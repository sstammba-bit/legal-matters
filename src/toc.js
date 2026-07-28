const { TableOfContents, Paragraph, PageBreak, HeadingLevel, TextRun } = require("docx");

// Vraie TOC Word (champ { TOC \o "1-3" \h \z \u }), pas du texte pré-calculé :
// Word la remplit/actualise lui-même à l'ouverture ou via F9. headingStyleRange
// "1-3" ne couvre que Heading1/2/3, donc les Allégations n'y apparaissent
// jamais, conformément au cahier des charges, sans configuration additionnelle.
//
// Le titre "Table des matières" est un Heading1 explicite (numbering: false) :
// l'alias passé à TableOfContents n'est qu'un contenu de repli avant la
// première mise à jour du champ — une fois Word recalculé, il disparaît et
// est remplacé par les entrées réelles, donc il ne peut pas servir de titre
// permanent.
function buildTableOfContents() {
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      numbering: false,
      children: [new TextRun("Table des matières")],
    }),
    new TableOfContents("Table des matières", {
      hyperlink: true,
      headingStyleRange: "1-3",
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

module.exports = { buildTableOfContents };
