const {
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  HeadingLevel,
  ShadingType,
} = require("docx");

const HEADER_FILL = "D9D9D9";

function headerCell(text, widthPercent) {
  return new TableCell({
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: HEADER_FILL },
    children: [
      new Paragraph({ children: [new TextRun({ text, bold: true })] }),
    ],
  });
}

function dataCell(text, widthPercent) {
  return new TableCell({
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    children: [new Paragraph({ children: [new TextRun({ text })] })],
  });
}

function buildPiecesChapter() {
  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          headerCell("Pièce", 15),
          headerCell("Description", 55),
          headerCell("Observations", 30),
        ],
      }),
      new TableRow({
        children: [dataCell("1", 15), dataCell("EXEMPLE — Courrier recommandé du 1er janvier 2026", 55), dataCell("", 30)],
      }),
      new TableRow({
        children: [dataCell("2", 15), dataCell("EXEMPLE — Accusé de réception postal", 55), dataCell("", 30)],
      }),
      new TableRow({
        children: [dataCell("3", 15), dataCell("EXEMPLE — Relevé des délais", 55), dataCell("", 30)],
      }),
    ],
  });

  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun("Liste des pièces")],
    }),
    table,
  ];
}

// Version générique réutilisable pour un vrai document (pas le modèle de
// démonstration) : rows = [{ piece, description, observations }, ...].
function buildPiecesTable(rows = []) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          headerCell("Pièce", 15),
          headerCell("Description", 55),
          headerCell("Observations", 30),
        ],
      }),
      ...rows.map(
        (row) =>
          new TableRow({
            children: [
              dataCell(row.piece ?? "", 15),
              dataCell(row.description ?? "", 55),
              dataCell(row.observations ?? "", 30),
            ],
          })
      ),
    ],
  });
}

module.exports = { buildPiecesChapter, buildPiecesTable };
