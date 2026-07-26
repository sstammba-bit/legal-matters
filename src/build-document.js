const { Document } = require("docx");
const { paragraphStyles } = require("./styles");
const { numberingConfig } = require("./numbering");
const { buildHeader, buildFooter } = require("./header-footer");
const { buildTableOfContents } = require("./toc");
const { buildExampleChapter, buildConclusionsChapter } = require("./example-chapter");
const { buildPiecesChapter } = require("./pieces-table");
const { buildAnnexesChapter } = require("./annexes");
const { cmToTwip } = require("./units");

function buildDocument() {
  return new Document({
    creator: "Samuel",
    title: "Modèle de mémoire juridique",
    description: "Modèle Word pour mémoires juridiques suisses",
    styles: { paragraphStyles },
    numbering: numberingConfig,
    features: { updateFields: true },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: cmToTwip(21), // A4
              height: cmToTwip(29.7),
            },
            margin: {
              top: cmToTwip(2.5),
              bottom: cmToTwip(2.5),
              left: cmToTwip(2.5),
              right: cmToTwip(2.5),
            },
          },
        },
        headers: { default: buildHeader() },
        footers: { default: buildFooter() },
        children: [
          ...buildTableOfContents(),
          ...buildExampleChapter(),
          ...buildConclusionsChapter(),
          ...buildPiecesChapter(),
          ...buildAnnexesChapter(),
        ],
      },
    ],
  });
}

module.exports = { buildDocument };
