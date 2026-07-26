const { TableOfContents, Paragraph, PageBreak } = require("docx");

// Vraie TOC Word (champ { TOC \o "1-3" \h \z \u }), pas du texte pré-calculé :
// Word la remplit/actualise lui-même à l'ouverture ou via F9. headingStyleRange
// "1-3" ne couvre que Heading1/2/3, donc les Allégations n'y apparaissent
// jamais, conformément au cahier des charges, sans configuration additionnelle.
function buildTableOfContents() {
  return [
    // L'alias "Table des matières" sert de titre affiché par le champ TOC
    // lui-même — pas besoin d'un Heading1 séparé, qui serait à tort compté
    // et numéroté comme un chapitre.
    new TableOfContents("Table des matières", {
      hyperlink: true,
      headingStyleRange: "1-3",
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

module.exports = { buildTableOfContents };
