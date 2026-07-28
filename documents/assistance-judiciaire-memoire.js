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
  CommentRangeStart,
  CommentRangeEnd,
  CommentReference,
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

// Paragraphe dont le texte reste à vérifier/compléter par Samuel : le texte
// proposé est affiché normalement, mais signalé par un vrai commentaire Word
// en marge (pas de balise visible dans le corps du texte).
let commentIdCounter = 0;
const genericComments = [];
const NG = (text, commentText) => {
  const id = ++commentIdCounter;
  genericComments.push({
    id,
    author: "Claude",
    initials: "IA",
    date: new Date(),
    children: [new Paragraph({ children: [new TextRun(commentText)] })],
  });
  return new Paragraph({
    style: "Normal",
    children: [
      new CommentRangeStart(id),
      new TextRun(text),
      new CommentRangeEnd(id),
      new CommentReference(id),
    ],
  });
};

const H1 = (text) =>
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });

const H2 = (text) =>
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });

const titleLine = (text, opts = {}, spacingBefore = 0) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: spacingBefore, after: 120 },
    children: [new TextRun({ text, font: FONT, size: 24, ...opts })],
  });

// Titre + sous-titre, expéditeur, destinataire, numéro de dossier — hors
// table des matières (pas un Heading).
function buildTitleBlock() {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [new TextRun({ text: "MÉMOIRE", bold: true, font: FONT, size: 40 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 480 },
      children: [
        new TextRun({
          text: "Relatif à la demande de remboursement de l'assistance judiciaire",
          bold: true,
          font: FONT,
          size: 28,
        }),
      ],
    }),
    titleLine("Samuel Stammbach", { bold: true }),
    titleLine("Avenue Floréal 12"),
    titleLine("1006 Lausanne"),
    titleLine("À l'attention de la", {}, 360),
    titleLine("Direction générale des affaires institutionnelles et des communes (DGAIC)"),
    titleLine("Direction du recouvrement"),
    titleLine("Dossier AJ19002691", { bold: true }, 360),
    titleLine("Lausanne, le [date]"),
  ];
}

// Titre Heading1 mais sans numérotation automatique (le Préambule précède le
// chapitre I et n'est traditionnellement pas numéroté dans un mémoire).
// Reste sur la page de garde (pas de saut de page avant), mais s'en démarque
// visuellement par un espacement généreux avant le titre.
function buildPreambule() {
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      numbering: false,
      spacing: { before: 960 },
      children: [new TextRun("Préambule")],
    }),
    N(
      "La présente démarche n'a pas pour objet de remettre en cause le principe du remboursement de l'assistance judiciaire, ni les décisions rendues au cours de la procédure judiciaire. Son objectif est de présenter, de manière complète et documentée, ma situation personnelle, familiale et financière actuelle, afin de permettre à la DGAIC d'apprécier ma capacité réelle de remboursement."
    ),
    N(
      "Cette situation ne peut être dissociée du contexte dans lequel l'assistance judiciaire a été accordée : une procédure de divorce ayant duré plus de dix ans, de la séparation en 2013 à l'issue définitive de la procédure d'appel en 2025."
    ),
    new Paragraph({ children: [new PageBreak()] }),
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

// Table financière (3.3) : colonne "Poste" alignée à gauche et NON justifiée
// (le style Normal hérité est justifié par défaut, ce qui étire les espaces
// dans une cellule étroite multi-lignes — d'où l'override explicite ici),
// colonnes de montants alignées à droite. mainText + noteText (optionnel, en
// italique) permettent d'ajouter une remarque à la suite d'un poste.
function chargesRow(mainText, noteText, montant2025, montantFutur, bold) {
  const posteChildren = [new TextRun({ text: mainText, bold: !!bold })];
  if (noteText) {
    posteChildren.push(new TextRun({ text: " — ", bold: !!bold }));
    posteChildren.push(new TextRun({ text: noteText, italics: true }));
  }
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 55, type: WidthType.PERCENTAGE },
        children: [new Paragraph({ alignment: AlignmentType.LEFT, children: posteChildren })],
      }),
      new TableCell({
        width: { size: 22.5, type: WidthType.PERCENTAGE },
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: montant2025, bold: !!bold })],
          }),
        ],
      }),
      new TableCell({
        width: { size: 22.5, type: WidthType.PERCENTAGE },
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: montantFutur, bold: !!bold })],
          }),
        ],
      }),
    ],
  });
}

function buildChargesTable() {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      new TableCell({
        width: { size: 55, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, color: "auto", fill: "D9D9D9" },
        children: [
          new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [new TextRun({ text: "Poste", bold: true })],
          }),
        ],
      }),
      new TableCell({
        width: { size: 22.5, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, color: "auto", fill: "D9D9D9" },
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: "Budget 2025 (réel)", bold: true })],
          }),
        ],
      }),
      new TableCell({
        width: { size: 22.5, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, color: "auto", fill: "D9D9D9" },
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: "Budget futur proposé", bold: true })],
          }),
        ],
      }),
    ],
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      headerRow,
      chargesRow(
        "Logement (location, électricité/Services industriels, Radio/TV-Serafe, assurance ménage et RC, internet)",
        null,
        "2'638 fr.",
        "2'638 fr."
      ),
      chargesRow(
        "Alimentation (épicerie, fournitures ménage, hors restaurant)",
        null,
        "1'756 fr.",
        "1'720 fr."
      ),
      chargesRow("Impôts", null, "1'900 fr.", "1'900 fr."),
      chargesRow(
        "Enfant Éline (contribution, habits, argent de poche, voyages, …)",
        null,
        "1'437 fr.",
        "1'437 fr."
      ),
      chargesRow(
        "Enfant Elmo (habits, éducation, équipements, santé, téléphone, argent de poche)",
        "le budget futur augmente car Elmo entre au gymnase à 16 ans : écolage ~700 fr., frais de transport CFF plus élevés",
        "1'351 fr.",
        "1'600 fr."
      ),
      chargesRow(
        "Santé (primes assurance maladie/complémentaire, médecine alternative, compléments, dentiste)",
        null,
        "1'036 fr.",
        "1'000 fr."
      ),
      chargesRow(
        "Dettes (crédit Migros + intérêts)",
        "budget futur : réduction des intérêts de cartes, arrêt de leur utilisation",
        "713 fr.",
        "650 fr."
      ),
      chargesRow("Transports", "réduction future ; pas de voiture", "530 fr.", "400 fr."),
      chargesRow(
        "Divers/utilities : services streaming, services smartphone, services logiciels (Microsoft)",
        "budget futur réduit : arrêt Netflix et autres services payants, utilisation des services gratuits ou déjà couverts par la redevance",
        "491 fr.",
        "400 fr."
      ),
      chargesRow("Thora (chien)", null, "240 fr.", "240 fr."),
      chargesRow(
        "Sous-total charges fixes/quasi-fixes",
        null,
        "12'092 fr.",
        "11'985 fr.",
        true
      ),
      chargesRow(
        "Loisirs/vacances (activités, vacances, restaurant)",
        null,
        "2'372 fr.",
        "600 fr."
      ),
      chargesRow("Dépenses personnelles Sam", null, "442 fr.", "100 fr."),
      chargesRow("Dépenses exceptionnelles", null, "733 fr.", "65 fr."),
      chargesRow("TOTAL CHARGES", null, "15'639 fr.", "12'750 fr.", true),
      chargesRow("Revenu net", null, "12'959 fr.", "12'959 fr."),
      chargesRow(
        "Solde mensuel avant remboursement",
        null,
        "-2'680 fr.",
        "+209 fr.",
        true
      ),
      chargesRow("Remboursement proposé", null, "—", "100 fr."),
      chargesRow("Solde final", null, "", "109 fr.", true),
    ],
  });
}

// rows: [{ cells: [string, ...], bold: boolean }]
function boldableTable(headers, rows, widths) {
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
    ({ cells, bold }) =>
      new TableRow({
        children: cells.map(
          (c, i) =>
            new TableCell({
              width: { size: widths[i], type: WidthType.PERCENTAGE },
              children: [new Paragraph({ children: [new TextRun({ text: c, bold: !!bold })] })],
            })
        ),
      })
  );
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });
}

function buildChapitrePieces() {
  const rows = [
    { piece: "1", description: "Certificat de salaire 2025 (Skyguide)" },
    { piece: "2", description: "Police d'assurance LAMal Samuel" },
    { piece: "3", description: "Police d'assurance complémentaire LCA Samuel" },
    { piece: "4", description: "Police d'assurance Elmo (LAMal + LCA)" },
    { piece: "5", description: "Bail à loyer" },
    { piece: "6", description: "Avenant n°1 au bail" },
    { piece: "7a", description: "Relevé UBS 2024", observations: "Dette carte : CHF 10'142.46" },
    { piece: "7b", description: "Relevé UBS 2025", observations: "Dette carte : CHF 279.04" },
    { piece: "8a", description: "Relevé Swisscard 2024", observations: "Dette carte : CHF 21'825.69" },
    { piece: "8b", description: "Relevé Swisscard 2025", observations: "Dette carte : CHF 12'369.28" },
    { piece: "9", description: "Facture électricité SIL, période 14.11.2024 – 30.11.2025" },
    { piece: "10", description: "Facture électricité SIL, période 01.03.2026 – 31.05.2026" },
    { piece: "11", description: "Assurance ménage/incendie ECA 2026" },
    { piece: "12", description: "Redevance radio/TV SERAFE" },
    { piece: "13", description: "Assurance animal Thora (Vaudoise Animalia)" },
    { piece: "14", description: "Facture internet Wingo (juin 2026)" },
    { piece: "15", description: "Garantie de loyer Firstcaution" },
    {
      piece: "16",
      description: "Certificat de prévoyance 2e pilier (Fondation Skycare, situation au 01.06.2026)",
    },
    { piece: "17", description: "Contrat et/ou relevés du crédit personnel Banque Migros", observations: "À fournir" },
    { piece: "18", description: "Jugement de divorce et convention d'entretien pour Éline", observations: "À fournir" },
    { piece: "19", description: "Documents fiscaux / avis de taxation 2025", observations: "À fournir" },
    {
      piece: "20",
      description: "Extrait de comptabilité personnelle 2025",
      observations: "À fournir",
    },
  ];

  return [
    H1("Liste des pièces"),
    N(
      "Afin de limiter les impressions inutiles, les pièces justificatives numérotées vous sont transmises sous forme électronique, en annexe du courriel de ce jour adressé à recouvrement.dgaic@vd.ch. Je me tiens à disposition pour vous faire parvenir des copies papier de tout ou partie de ces pièces si cela s'avérait nécessaire au traitement du dossier."
    ),
    buildPiecesTable(rows),
    N(
      "Le plan social de Skyguide (« DR0539E Social plan for Managers ») mentionné au chapitre 5 n'est pas joint en tant que pièce ; les faits y relatifs sont exposés en texte libre et recoupés avec des sources de presse publiques."
    ),
  ];
}

function buildMemoireContent() {
  return [
    ...buildTitleBlock(),
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
          "Au-delà de la contribution d'entretien contractuelle, je prends en charge, selon mes moyens et au fil de ses besoins, un ensemble de frais complémentaires pour Éline : argent de poche (CHF 2'169.90 sur l'année 2025), vêtements (CHF 183.78), cadeaux (CHF 516.00), frais de transport (CFF) pour ses visites (CHF 601.00), sa part de vacances communes (CHF 421.68), ainsi qu'un soutien ponctuel complémentaire à sa mère lorsque la situation le justifie (CHF 1'492.80 en 2025). Je m'efforce de répondre à ses besoins au cas par cas, dans la mesure de mes moyens.",
          "Je suis domicilié à l'Avenue Floréal 12, 1006 Lausanne, où je réside en qualité de locataire avec mon fils Elmo.",
        ],
      },
      {
        titre: "Revenus",
        elements: [
          N("Je suis employé auprès de Skyguide. Mon salaire constitue ma seule source de revenus."),
          simpleTable(
            ["Revenus", "Montant", "Pièce"],
            [
              ["Salaire brut annuel 2025", "CHF 195'041.00", "1"],
              ["Salaire brut mensuel moyen", "CHF 16'253.42", "1"],
              ["Salaire net annuel 2025", "CHF 155'503.00", "1"],
              ["Salaire net mensuel moyen", "CHF 12'958.58", "1"],
            ],
            [45, 35, 20]
          ),
        ],
      },
      {
        titre: "Charges",
        elements: [
          N(
            "Les montants retenus sont extraits de ma comptabilité personnelle relative à l'année civile 2025 — une année complète, offrant ainsi une moyenne mensuelle représentative de ma situation. Certains postes de charges (notamment l'alimentation, les vacances et les transports) sont constitués d'un très grand nombre de transactions individuelles au cours de l'année, rendant peu praticable la production d'un justificatif distinct pour chacune d'entre elles. Pour ces postes, le montant retenu est extrait de ma comptabilité personnelle, tenue tout au long de l'année à partir des relevés bancaires et de carte de crédit (Pièce 20), qui peut être mise à disposition dans son détail sur demande."
          ),
          N(
            "Le montant retenu pour l'alimentation exclut le restaurant (reclassé en loisirs) et intègre la présence effective de mes enfants à mon domicile — à plein temps pour Elmo, et au prorata de la garde légale pour Éline (1 week-end par mois et la moitié des vacances scolaires, soit environ 4.5 jours par mois en moyenne). L'objectif retenu est une alimentation biologique et saine, avec de la viande une fois par semaine."
          ),
          buildChargesTable(),
          N(
            "À titre d'illustration, le mois de septembre 2025 a représenté le pic annuel de mes dépenses d'alimentation (épicerie et restaurant confondus), avec un total de CHF 2'878.69, contre une moyenne mensuelle de CHF 2'442.54 sur l'année."
          ),
        ],
      },
      {
        titre: "Fortune",
        paragraphes: [
          "Selon les relevés bancaires produits, mon compte de carte UBS présentait une dette de CHF 279.04 au 31 décembre 2025 (contre CHF 10'142.46 au 31 décembre 2024) ; des intérêts débiteurs de CHF 343.77 m'ont été facturés pour l'année 2025 sur ce solde dû.",
          "Mon compte Swisscard présentait une dette de CHF 12'369.28 au 31 décembre 2025 (contre CHF 21'825.69 au 31 décembre 2024) ; des intérêts débiteurs de CHF 720.60 m'ont été facturés pour l'année 2025 sur ce solde dû.",
          "Mes dettes comprennent également la garantie de loyer déposée auprès de Firstcaution, d'un montant de CHF 363.55 par an.",
          "Ces dettes trouvent principalement leur origine dans ma volonté d'offrir à mes enfants des vacances et des moments de qualité malgré le contexte difficile de ces dernières années — un choix qui a représenté une charge budgétaire significative (cf. section 3.3, poste loisirs/vacances). J'ai conscience que ce train de vie doit être revu à la baisse, et le budget futur proposé en section 3.3 prévoit une réduction substantielle de ce poste, avec des vacances plus modestes à l'avenir.",
          "Mon capital de prévoyance professionnelle (2e pilier, CHF 1'312'189.55 au 1er juin 2026 selon mon certificat de prévoyance, Pièce 16) n'est pas saisissable avant son échéance (art. 39 al. 2 LPP) et ne constitue donc pas une fortune disponible pour le remboursement de cette dette.",
          "Les autres éléments de fortune (biens immobiliers, titres ou placements, autres dettes non encore listées) seront précisés dans une version ultérieure du présent mémoire si nécessaire.",
        ],
      },
      {
        titre: "Motifs ayant conduit à l'octroi de l'assistance judiciaire",
        elements: [
          NG(
            "L'assistance judiciaire m'a été accordée dans le cadre de la procédure de divorce, à un moment où mes ressources ne permettaient pas d'assumer seul les frais de justice et d'avocat d'une procédure qui s'est étendue sur plus de dix ans, de la séparation en 2013 à l'issue définitive de la procédure d'appel en 2025.",
            "GÉNÉRIQUE — à vérifier/compléter par Samuel, notamment la référence exacte de la décision d'octroi."
          ),
        ],
      },
    ]),
    ...buildChapitreAvecSousSections("Analyse de la capacité réelle de remboursement", [
      {
        titre: "Capacité théorique de remboursement",
        paragraphes: [
          "Sur la base de mon revenu net actuel (12'959 fr./mois) et du budget futur proposé (12'750 fr./mois, section 3.3), ma capacité théorique de remboursement s'élève à 209 fr. par mois avant constitution d'une réserve de sécurité.",
        ],
      },
      {
        titre: "Capacité réelle de remboursement",
        paragraphes: [
          "Compte tenu de l'absence d'épargne liquide (section 3.4) et de l'incertitude professionnelle documentée au chapitre 5, il est nécessaire de conserver une marge de sécurité plutôt que d'engager l'intégralité de ce montant théorique. Ma capacité réelle de remboursement, prudente et soutenable, s'établit à 100 fr. par mois.",
        ],
      },
      {
        titre: "Équilibre budgétaire",
        paragraphes: [
          "Le budget futur proposé en section 3.3 repose sur une réduction volontaire et significative de mes dépenses discrétionnaires (loisirs/vacances, dépenses personnelles, dépenses exceptionnelles), permettant de passer d'un déficit mensuel réel de -2'680 fr. en 2025 à un équilibre positif de +209 fr. avant remboursement.",
        ],
      },
      {
        titre: "Réserve financière nécessaire",
        paragraphes: [
          "En l'absence d'épargne disponible (section 3.4), une réserve mensuelle minimale doit être maintenue pour faire face à des dépenses imprévues (santé, réparations, imprévus liés à mes enfants) sans devoir recourir à nouveau à l'endettement. Le solde résiduel de 109 fr. par mois après remboursement constitue cette réserve, volontairement limitée.",
        ],
      },
      {
        titre: "Proposition financière soutenable",
        paragraphes: [
          "Sur la base de cette analyse, je propose un remboursement mensuel de CHF 100.–, détaillé au chapitre 6.",
        ],
      },
    ]),
    ...buildChapitreAvecSousSections("Éléments particuliers à prendre en considération", [
      {
        titre: "Durée exceptionnelle de la procédure",
        paragraphes: [
          "La procédure de divorce s'est étendue sur plus de dix ans, de la séparation en 2013 à l'entrée en force définitive du jugement le 10 juillet 2025. Cette durée exceptionnelle explique l'ampleur des prestations d'assistance judiciaire accordées au fil de la procédure, ainsi que la difficulté à anticiper et provisionner leur remboursement.",
        ],
      },
      {
        titre: "Conséquences financières du divorce",
        paragraphes: [
          "Le divorce a entraîné une réorganisation complète de ma situation financière : entretien de deux ménages distincts (le mien avec Elmo, et la contribution versée pour Éline), partage des frais de procédure, et nécessité de reconstruire une situation stable pour mes enfants dans ce nouveau contexte familial.",
        ],
      },
      {
        titre: "Âge et proximité de la retraite",
        paragraphes: [
          "Ma situation professionnelle actuelle est marquée par une incertitude significative liée au plan de restructuration engagé par Skyguide et annoncé publiquement le 19 mai 2026 (jusqu'à 220 postes supprimés d'ici fin 2027, ramené à une centaine à l'issue de la consultation selon le communiqué de l'entreprise du 13 juillet 2026), les départs à la retraite anticipée étant explicitement identifiés comme un des leviers pour absorber une partie de ces suppressions.",
          "J'ai 60 ans et plus de 20 ans d'ancienneté au sein de Skyguide. Deux scénarios se présentent à moi.",
          "Dans le pire des cas, je pourrais être amené à prendre une retraite anticipée dès novembre 2026, à 61 ans — une retraite complète mais avec une rente nettement plus basse que mon salaire actuel. Selon mon certificat de prévoyance (Pièce 16, Fondation de prévoyance Skycare, situation au 1er juin 2026), cette rente s'élèverait à CHF 4'829.10 (sans 13e rente) à CHF 5'231.53 (avec 13e rente, sous condition de taux de couverture), contre un revenu net actuel de CHF 12'958.58 — soit une réduction de revenu d'environ 60 à 63%.",
          "Dans le meilleur des cas, je poursuivrais mon activité jusqu'à l'âge ordinaire de la retraite, à 63 ans (novembre 2028), bénéficiant ainsi de deux années supplémentaires de salaire à mon niveau actuel — mais dans un contexte professionnel qui reste incertain, avec un disponible mensuel qui demeure restreint, comme démontré par l'analyse détaillée de mes charges.",
          "Dans l'un comme dans l'autre de ces scénarios, un engagement de remboursement mensuel modeste reste soutenable ; un montant plus élevé, basé sur ma seule situation actuelle, deviendrait rapidement intenable si la retraite anticipée se concrétisait dès 2026.",
          "S'agissant de l'AVS, ma carrière de cotisation devrait me permettre de percevoir une rente proche du maximum légal dès l'âge de référence de 65 ans. Le montant maximal de la rente AVS pour une personne seule s'élève, selon les données 2026, à environ CHF 2'450 à 2'520.– par mois (les sources disponibles varient légèrement sur ce montant ; un extrait de compte individuel AVS permettrait de le confirmer précisément), auquel s'ajoutera une 13e rente dès décembre 2026.",
          "Dans l'hypothèse où je conserverais mon emploi jusqu'à l'âge ordinaire de la retraite, à 63 ans (novembre 2028), je percevrais dès cet âge une rente LPP de CHF 5'777.50 par mois (sans 13e rente ; Pièce 16). Dès 65 ans, cette rente LPP se combinerait avec la rente AVS, proche du maximum comme indiqué ci-dessus (environ CHF 2'450 à 2'520.–/mois), portant mon revenu total à environ CHF 8'200 à 8'300.– par mois — soit encore nettement en dessous de mon revenu net actuel de CHF 12'958.58/mois. Cette situation devra être assumée alors que ma fille Éline pourrait encore être en formation (jusqu'en 2035 au plus tard, selon la durée de ses études), ce qui maintient une charge financière significative bien après le début de ma retraite.",
        ],
      },
      {
        titre: "Incertitudes professionnelles",
        paragraphes: [
          "Au-delà de la question spécifique de la retraite anticipée développée ci-dessus, le climat général d'incertitude lié au plan de restructuration Skyguide affecte l'ensemble du personnel administratif et technique, dont je fais partie. Cette incertitude renforce la nécessité d'une proposition de remboursement prudente et modulable.",
        ],
      },
      { titre: "Autres circonstances pertinentes" },
    ]),
    ...buildChapitreAvecSousSections("Proposition de remboursement", [
      {
        titre: "Principes retenus",
        paragraphes: [
          "Ma proposition repose sur trois principes : premièrement, présenter une situation financière réelle et documentée plutôt qu'une estimation théorique ; deuxièmement, proposer un montant que je suis certain de pouvoir honorer durablement, y compris en cas de baisse de revenu ; troisièmement, privilégier un engagement modeste mais fiable plutôt qu'un montant plus élevé risquant d'être interrompu.",
        ],
      },
      {
        titre: "Proposition de plan de remboursement",
        paragraphes: [
          "Au vu de l'ensemble des éléments exposés ci-dessus — charges réelles 2025, budget réduit proposé pour l'avenir proche, et risque concret d'une baisse de revenu significative dès novembre 2026 en cas de retraite anticipée — je propose un remboursement mensuel de CHF 100.–, montant que je m'engage à honorer de manière régulière et soutenable. Ce montant est proposé comme étant soutenable tant dans le meilleur que dans le pire des scénarios professionnels décrits au chapitre 5.",
        ],
      },
      {
        titre: "Engagement de collaboration",
        paragraphes: [
          "Je m'engage à informer la DGAIC de toute évolution significative de ma situation financière ou professionnelle, notamment en cas de retraite anticipée effective, et à fournir tout document complémentaire utile à l'examen de mon dossier.",
        ],
      },
    ]),
    H1("Conclusions"),
    N(
      "Le présent mémoire expose de manière complète et documentée ma situation personnelle, familiale et financière. Il démontre qu'un remboursement de CHF 100.– par mois constitue le montant maximal que je peux m'engager à honorer de manière durable, compte tenu de mes charges actuelles et des incertitudes professionnelles à venir. Je reste à disposition de la DGAIC pour tout complément d'information."
    ),
    ...buildChapitrePieces(),
    H1("Annexes"),
    N("Sans objet à ce stade."),
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
    comments: { children: genericComments },
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
