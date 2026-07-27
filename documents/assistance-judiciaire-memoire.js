const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ShadingType,
  PageBreak,
} = require("docx");
const fs = require("fs");
const path = require("path");

const { paragraphStyles, FONT } = require("../src/styles");
const { numberingConfig } = require("../src/numbering");
const { buildHeader, buildFooter } = require("../src/header-footer");
const { buildTableOfContents } = require("../src/toc");
const { buildPiecesTable } = require("../src/pieces-table");
const { cmToTwip } = require("../src/units");

const N = (text) => new Paragraph({ style: "Normal", children: [new TextRun(text)] });

const H1 = (text) =>
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });

const H2 = (text) =>
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });

// Titre + sous-titre + numéro de dossier, hors table des matières (pas un Heading).
function buildTitleBlock() {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [new TextRun({ text: "MÉMOIRE", bold: true, font: FONT, size: 40 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [
        new TextRun({
          text: "relatif à la demande de remboursement de l'assistance judiciaire",
          bold: true,
          font: FONT,
          size: 28,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 480 },
      children: [
        new TextRun({ text: "Dossier AJ19002691", font: FONT, size: 24 }),
      ],
    }),
  ];
}

// Titre Heading1 mais sans numérotation automatique (le Préambule précède le
// chapitre I et n'est traditionnellement pas numéroté dans un mémoire).
function buildPreambule() {
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      numbering: false,
      children: [new TextRun("Préambule")],
    }),
    N(
      "La présente démarche n'a pas pour objet de remettre en cause le principe du remboursement de l'assistance judiciaire, ni les décisions rendues au cours de la procédure judiciaire."
    ),
    N(
      "Son unique objectif est de présenter, de manière complète, transparente et documentée, ma situation personnelle, familiale et financière actuelle, afin de permettre à la Direction générale des affaires institutionnelles et des communes (DGAIC) d'apprécier ma capacité réelle de remboursement et, le cas échéant, les modalités les plus appropriées pour celui-ci."
    ),
    N(
      "Cette analyse ne peut toutefois être dissociée du contexte dans lequel l'assistance judiciaire a été accordée. La procédure de divorce trouve son origine dans la séparation intervenue en 2013 et ne s'est définitivement achevée qu'à l'issue de la procédure d'appel, soit après plus de dix années de procédure judiciaire. Pendant toute cette période, ma situation personnelle, familiale et financière ont été profondément affectées."
    ),
    N(
      "Le présent mémoire expose les faits de manière chronologique, objective et documentée. Mon intention est de collaborer pleinement avec la DGAIC afin de parvenir à une solution réaliste, proportionnée et durable, compatible avec ma situation financière effective ainsi qu'avec les obligations qui demeurent à ma charge."
    ),
  ];
}

function buildChapitreObjet() {
  return [
    H1("Objet de la demande"),
    H2("Contexte de la démarche"),
    N(
      "Le présent mémoire est établi en réponse au courrier de la Direction générale des affaires institutionnelles et des communes (DGAIC), Direction du recouvrement, du 30 juin 2026, relatif au remboursement des prestations accordées au titre de l'assistance judiciaire sous le numéro AJ19002691."
    ),
    N(
      "Selon ce courrier, le montant total des prestations avancées par l'État s'élève à CHF 72'023.00. Après déduction des acomptes déjà versés entre le 8 juillet 2019 et le 29 mai 2026, le solde réclamé est de CHF 67'923.00."
    ),
    N("Le présent mémoire constitue ma première réponse à ce courrier."),

    H2("Objet du présent mémoire"),
    N(
      "Par son courrier du 30 juin 2026, la DGAIC m'invite soit à m'acquitter du montant réclamé dans un délai de trente jours, soit, si je souhaite solliciter un plan de paiement, à présenter une proposition écrite accompagnée des justificatifs relatifs à ma situation financière."
    ),
    N("Le présent mémoire répond à cette invitation."),
    N(
      "Il présente de manière complète, structurée et documentée ma situation personnelle, familiale et financière actuelle afin de permettre une appréciation objective de ma capacité réelle de remboursement."
    ),
    N(
      "Sur cette base, il formule une proposition de remboursement qui se veut réaliste, durable et compatible avec mes obligations financières actuelles."
    ),

    H2("Esprit de la démarche"),
    N(
      "Le présent mémoire ne remet pas en cause le principe du remboursement prévu par la législation applicable ni les décisions rendues au cours de la procédure judiciaire."
    ),
    N(
      "Il s'inscrit dans une démarche de pleine collaboration avec la DGAIC. Son objectif est de fournir tous les éléments utiles à une appréciation complète de ma situation actuelle, afin que les modalités de remboursement puissent être déterminées sur une base objective, documentée et conforme à ma réalité économique."
    ),
    N(
      "L'ensemble des informations présentées dans ce mémoire est appuyé par des pièces justificatives et organisé de manière chronologique afin de faciliter l'examen du dossier."
    ),
  ];
}

function chronologyRow(date, event, widths) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: widths[0], type: WidthType.PERCENTAGE },
        children: [new Paragraph({ children: [new TextRun(date)] })],
      }),
      new TableCell({
        width: { size: widths[1], type: WidthType.PERCENTAGE },
        children: [new Paragraph({ children: [new TextRun(event)] })],
      }),
    ],
  });
}

function buildChapitreChronologie() {
  const widths = [22, 78];
  const rows = [
    ["30 mai 2013", "Séparation des époux."],
    ["4 décembre 2015", "Dépôt de la demande unilatérale en divorce."],
    [
      "3 octobre 2024",
      "Jugement de divorce rendu par le Tribunal d'arrondissement de Lausanne.",
    ],
    [
      "4 novembre 2024",
      "Appel déposé par Madame Susanna Stammbach, limité aux contributions d'entretien.",
    ],
    [
      "27 juin 2025",
      "Arrêt du Tribunal cantonal réformant partiellement le jugement de divorce.",
    ],
    ["10 juillet 2025", "Le jugement de divorce devient entièrement définitif et exécutoire."],
    [
      "30 juin 2026",
      "Courrier de la DGAIC sollicitant le remboursement des prestations d'assistance judiciaire.",
    ],
  ];

  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({
            width: { size: widths[0], type: WidthType.PERCENTAGE },
            shading: { type: ShadingType.CLEAR, color: "auto", fill: "D9D9D9" },
            children: [new Paragraph({ children: [new TextRun({ text: "Date", bold: true })] })],
          }),
          new TableCell({
            width: { size: widths[1], type: WidthType.PERCENTAGE },
            shading: { type: ShadingType.CLEAR, color: "auto", fill: "D9D9D9" },
            children: [
              new Paragraph({ children: [new TextRun({ text: "Événement", bold: true })] }),
            ],
          }),
        ],
      }),
      ...rows.map(([date, event]) => chronologyRow(date, event, widths)),
    ],
  });

  return [
    H1("Chronologie de la procédure"),
    N(
      "Le tableau ci-dessous résume les principales étapes de la procédure ayant conduit à la présente demande de remboursement de l'assistance judiciaire."
    ),
    table,
  ];
}

// Sections dont le contenu reste à documenter (chiffres et pièces à
// rassembler) : un chapitre avec ses sous-titres, sans contenu fabriqué.
function placeholder() {
  return new Paragraph({
    style: "Observation",
    children: [new TextRun("[Section à compléter]")],
  });
}

// sousSections: [{ titre, paragraphes: [string, ...] }] ou [{ titre, elements: [Paragraph|Table, ...] }]
// — rien de fourni => la sous-section reste un repère "[Section à compléter]".
function buildChapitreAvecSousSections(titre, sousSections) {
  return [
    H1(titre),
    ...sousSections.flatMap(({ titre: sousTitre, paragraphes, elements }) => {
      const body =
        elements && elements.length > 0
          ? elements
          : paragraphes && paragraphes.length > 0
          ? paragraphes.map(N)
          : [placeholder()];
      return [H2(sousTitre), ...body];
    }),
  ];
}

function simpleTable(headers, rows, widths) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map(
      (h, i) =>
        new TableCell({
          width: { size: widths[i], type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, color: "auto", fill: "D9D9D9" },
          children: [new Paragraph({ children: [new TextRun({ text: h, bold: true })] })],
        })
    ),
  });
  const dataRows = rows.map(
    (cells) =>
      new TableRow({
        children: cells.map(
          (c, i) =>
            new TableCell({
              width: { size: widths[i], type: WidthType.PERCENTAGE },
              children: [new Paragraph({ children: [new TextRun(c)] })],
            })
        ),
      })
  );
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });
}

function buildChapitreSimple(titre) {
  return [H1(titre), placeholder()];
}

function buildChapitrePieces() {
  return [
    H1("Liste des pièces"),
    N(
      "Afin de limiter les impressions inutiles, les pièces justificatives numérotées vous sont transmises sous forme électronique, en annexe du courriel de ce jour adressé à recouvrement.dgaic@vd.ch. Je me tiens à disposition pour vous faire parvenir des copies papier de tout ou partie de ces pièces si cela s'avérait nécessaire au traitement du dossier."
    ),
    buildPiecesTable([]),
  ];
}

function buildChapitreAnnexes() {
  return [H1("Annexes"), placeholder()];
}

function buildMemoireContent() {
  return [
    ...buildTitleBlock(),
    new Paragraph({ children: [new PageBreak()] }),
    ...buildPreambule(),
    ...buildTableOfContents(),
    ...buildChapitreObjet(),
    ...buildChapitreChronologie(),
    ...buildChapitreAvecSousSections("Situation financière initiale", [
      {
        titre: "Situation personnelle et familiale",
        paragraphes: [
          "Je suis divorcé et père de deux enfants.",
          "Mon fils Elmo Stammbach, né le 9 septembre 2010, vit avec moi. J'en assume la garde de fait et son entretien courant. Il est âgé de 15 ans et vient d'entrer au gymnase. Il est vraisemblable qu'il poursuive une formation jusqu'à son terme. Son entretien restera donc à ma charge pendant plusieurs années encore.",
          "Ma fille Éline Stammbach, née le 3 novembre 2008, vit auprès de sa mère. Conformément au jugement de divorce, je verse une contribution d'entretien de CHF 980.– par mois, allocations familiales en sus, jusqu'à sa majorité ou, au-delà, jusqu'à la fin de sa formation.",
          "Éline effectue actuellement un apprentissage.",
          "Je suis domicilié à l'Avenue Floréal 12, 1006 Lausanne, où je réside avec mon fils Elmo.",
        ],
      },
      {
        titre: "Revenus",
        elements: [
          simpleTable(
            ["Revenus", "Montant", "Pièce", "Observation"],
            [
              ["Salaire brut annuel 2025", "CHF 195'041.00", "1", "Certificat de salaire"],
              ["Salaire brut mensuel moyen", "CHF 16'253.42", "1", ""],
              ["Salaire net annuel 2025", "CHF 155'503.00", "1", ""],
              ["Salaire net mensuel moyen", "CHF 12'958.58", "1", "Revenu disponible"],
            ],
            [30, 22, 10, 38]
          ),
        ],
      },
      { titre: "Charges" },
      {
        titre: "Fortune",
        paragraphes: [
          "Selon les relevés bancaires produits, mon compte UBS présentait un solde en ma faveur de CHF 279.04 au 31 décembre 2025 (contre CHF 10'142.46 au 31 décembre 2024) ; des intérêts débiteurs de CHF 343.77 ont été facturés pour l'année 2025.",
          "Mon compte Swisscard présentait un solde en ma faveur de CHF 12'369.28 au 31 décembre 2025 (contre CHF 21'825.69 au 31 décembre 2024), avec des intérêts en ma faveur de CHF 720.60 pour l'année 2025.",
          "Mes dettes comprennent notamment la garantie de loyer déposée auprès de Firstcaution, d'un montant de CHF 363.55 par an.",
          "Les autres éléments de fortune (biens immobiliers, titres ou placements, dettes non encore listées) seront précisés dans une version ultérieure du présent mémoire.",
        ],
      },
      {
        titre: "Enfants à charge",
        paragraphes: [
          "Mon fils Elmo Stammbach, né le 9 septembre 2010, vit avec moi ; j'en assume l'entretien courant (cf. également ch. 3.1). Le détail des charges afférentes fait l'objet d'une catégorisation en cours dans ma comptabilité personnelle et sera communiqué dès sa finalisation.",
          "Ma fille Éline Stammbach, née le 3 novembre 2008, vit auprès de sa mère. Je verse à ce titre une contribution d'entretien contractuelle de CHF 980.– par mois, allocations familiales en sus, conformément au jugement de divorce. Les frais complémentaires éventuels feront l'objet du même exercice de catégorisation.",
        ],
      },
      { titre: "Motifs ayant conduit à l'octroi de l'assistance judiciaire" },
    ]),
    ...buildChapitreAvecSousSections("Évolution de la situation depuis l'octroi de l'assistance judiciaire", [
      { titre: "Évolution familiale" },
      { titre: "Évolution de la procédure" },
      { titre: "Évolution professionnelle" },
    ]),
    ...buildChapitreAvecSousSections("Situation financière actuelle", [
      { titre: "Revenus" },
      { titre: "Charges" },
      { titre: "Fortune" },
      { titre: "Liquidités" },
      { titre: "Deuxième pilier" },
      { titre: "AVS" },
      { titre: "Obligations financières existantes" },
    ]),
    ...buildChapitreAvecSousSections("Analyse de la capacité réelle de remboursement", [
      { titre: "Capacité théorique de remboursement" },
      { titre: "Capacité réelle de remboursement" },
      { titre: "Équilibre budgétaire" },
      { titre: "Réserve financière nécessaire" },
      { titre: "Proposition financière soutenable" },
    ]),
    ...buildChapitreAvecSousSections("Éléments particuliers à prendre en considération", [
      { titre: "Durée exceptionnelle de la procédure" },
      { titre: "Conséquences financières du divorce" },
      { titre: "Âge et proximité de la retraite" },
      { titre: "Incertitudes professionnelles" },
      { titre: "Autres circonstances pertinentes" },
    ]),
    ...buildChapitreAvecSousSections("Proposition de remboursement", [
      { titre: "Principes retenus" },
      { titre: "Proposition de plan de remboursement" },
      { titre: "Engagement de collaboration" },
    ]),
    ...buildChapitreSimple("Conclusions"),
    ...buildChapitrePieces(),
    ...buildChapitreAnnexes(),
  ];
}

function buildDocument() {
  return new Document({
    creator: "Samuel",
    title: "Mémoire relatif à la demande de remboursement de l'assistance judiciaire — AJ19002691",
    description: "Mémoire DGAIC — dossier AJ19002691",
    styles: { paragraphStyles },
    numbering: numberingConfig,
    features: { updateFields: true },
    sections: [
      {
        properties: {
          page: {
            size: { width: cmToTwip(21), height: cmToTwip(29.7) },
            margin: {
              top: cmToTwip(2.5),
              bottom: cmToTwip(2.5),
              left: cmToTwip(2.5),
              right: cmToTwip(2.5),
            },
          },
        },
        headers: { default: buildHeader("Dossier AJ19002691") },
        footers: { default: buildFooter() },
        children: buildMemoireContent(),
      },
    ],
  });
}

const DEFAULT_OUTPUT_PATH = path.join(
  "G:",
  "My Drive",
  "05. - Admin",
  "09. Legal matters",
  "assistance-judiciaire",
  "Memoire_Remboursement_AJ19002691.docx"
);

async function main() {
  const outputPath = process.argv[2] || DEFAULT_OUTPUT_PATH;
  const buffer = await Packer.toBuffer(buildDocument());
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, buffer);
  console.log("Mémoire généré :", outputPath);
}

main().catch((err) => {
  console.error("Échec de la génération :", err);
  process.exit(1);
});
