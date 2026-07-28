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
      "La présente démarche n'a pas pour objet de remettre en cause le principe du remboursement de l'assistance judiciaire, ni les décisions rendues au cours de la procédure judiciaire. Son objectif est de présenter, de manière complète et documentée, ma situation personnelle, familiale et financière actuelle, afin de permettre à la DGAIC d'apprécier ma capacité réelle de remboursement."
    ),
    N(
      "Cette situation ne peut être dissociée du contexte dans lequel l'assistance judiciaire a été accordée : une procédure de divorce ayant duré plus de dix ans, de la séparation en 2013 à l'issue définitive de la procédure d'appel en 2025."
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
      "Le plan social de Skyguide (« DR0539E Social plan for Managers ») mentionné au chapitre 7 n'est pas joint en tant que pièce ; les faits y relatifs sont exposés en texte libre et recoupés avec des sources de presse publiques."
    ),
  ];
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
          boldableTable(
            ["Poste", "Budget 2025 (réel)", "Budget futur proposé"],
            [
              { cells: ["Logement", "-2'638.10", "-2'638.10"] },
              { cells: ["Alimentation (épicerie, hors restaurant)", "-1'756.44", "-1'720.00"] },
              { cells: ["Impôts", "-1'899.94", "-1'899.94"] },
              { cells: ["Enfant Éline (contribution + frais)", "-1'436.75", "-1'436.75"] },
              { cells: ["Enfant Elmo", "-1'351.47", "-1'351.47"] },
              { cells: ["Santé (hors primes)", "-1'036.17", "-1'036.17"] },
              { cells: ["Dettes (crédit Migros + intérêts)", "-713.06", "-713.06"] },
              { cells: ["Transports", "-530.44", "-530.44"] },
              { cells: ["Divers/utilities", "-490.81", "-490.81"] },
              { cells: ["Thora (chien)", "-239.93", "-239.93"] },
              {
                cells: ["Sous-total charges fixes/quasi-fixes", "-12'093.11", "-12'056.67"],
                bold: true,
              },
              { cells: ["Loisirs/vacances (activités, vacances, restaurant)", "-2'372.14", "-600.00"] },
              { cells: ["Dépenses personnelles Sam", "-441.69", "-100.00"] },
              { cells: ["Dépenses exceptionnelles", "-732.50", "-65.00"] },
              { cells: ["TOTAL CHARGES", "-15'639.44", "-12'821.67"], bold: true },
              { cells: ["Revenu net", "+12'958.58", "+12'958.58"] },
              { cells: ["Solde mensuel avant remboursement", "CHF -2'680.86", "CHF +136.91"], bold: true },
              { cells: ["Remboursement proposé", "—", "CHF 100.00"] },
              { cells: ["Solde final", "", "CHF +36.91"], bold: true },
            ],
            [46, 27, 27]
          ),
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
    ...buildChapitreAvecSousSections("Évolution de la situation depuis l'octroi de l'assistance judiciaire", [
      {
        titre: "Évolution familiale",
        elements: [
          NG(
            "Depuis l'octroi de l'assistance judiciaire, ma situation familiale a évolué : mon fils Elmo, aujourd'hui âgé de 15 ans, vit avec moi et est entré au gymnase ; ma fille Éline, âgée de 17 ans, effectue un apprentissage. La charge financière liée à mes deux enfants demeure significative et continuera pendant plusieurs années.",
            "GÉNÉRIQUE — à vérifier/compléter par Samuel."
          ),
        ],
      },
      {
        titre: "Évolution de la procédure",
        paragraphes: [
          "La procédure judiciaire, résumée au chapitre 2, s'est achevée le 10 juillet 2025 avec l'entrée en force définitive du jugement de divorce, à l'issue d'un arrêt du Tribunal cantonal du 27 juin 2025 réformant partiellement le jugement initial du 3 octobre 2024 sur les seules contributions d'entretien. Cette durée exceptionnelle a prolongé d'autant la période durant laquelle l'assistance judiciaire m'a été nécessaire.",
        ],
      },
      {
        titre: "Évolution professionnelle",
        elements: [
          NG(
            "Mon emploi auprès de Skyguide est resté stable durant toute cette période. Cette stabilité est aujourd'hui remise en question par le plan de restructuration en cours, détaillé au chapitre 7 (Éléments particuliers).",
            "GÉNÉRIQUE — à vérifier/compléter par Samuel."
          ),
        ],
      },
    ]),
    ...buildChapitreAvecSousSections("Situation financière actuelle", [
      {
        titre: "Revenus",
        paragraphes: [
          "Ma situation de revenus actuelle demeure identique à celle exposée en section 3.2 : mon salaire auprès de Skyguide constitue ma seule source de revenus, à hauteur de CHF 12'958.58 nets par mois en moyenne.",
        ],
      },
      {
        titre: "Charges",
        paragraphes: [
          'Ma structure de charges actuelle demeure conforme au tableau exposé en section 3.3, colonne « Budget futur proposé », qui reflète les ajustements que je mets en œuvre pour dégager une capacité de remboursement durable.',
        ],
      },
      {
        titre: "Fortune",
        paragraphes: [
          "Ma situation de fortune demeure celle exposée en section 3.4 : absence de fortune disponible significative, dettes de cartes en diminution, et capital de prévoyance non disponible avant échéance.",
        ],
      },
      {
        titre: "Liquidités",
        paragraphes: [
          "Je ne dispose pas d'épargne liquide significative au-delà des comptes bancaires et de carte de crédit déjà mentionnés en section 3.4, qui présentent des soldes débiteurs plutôt que des avoirs.",
        ],
      },
      {
        titre: "Deuxième pilier",
        paragraphes: [
          "Mon capital de prévoyance professionnelle s'élève à CHF 1'312'189.55 au 1er juin 2026 (Pièce 16). Ce capital n'est pas disponible avant son échéance et les perspectives de rente sont détaillées au chapitre 7 (Âge et proximité de la retraite).",
        ],
      },
      {
        titre: "AVS",
        elements: [
          NG(
            "Je cotise normalement à l'AVS dans le cadre de mon activité salariée auprès de Skyguide. Je ne dispose pas d'éléments particuliers à signaler concernant ma situation AVS à ce stade.",
            "GÉNÉRIQUE — à compléter par Samuel si des éléments spécifiques existent, ex. lacunes de cotisations."
          ),
        ],
      },
      {
        titre: "Obligations financières existantes",
        paragraphes: [
          "Mes obligations financières actuelles comprennent : la contribution d'entretien pour Éline (CHF 980.– par mois, obligation légale), le crédit personnel auprès de la Banque Migros, les soldes dus sur mes comptes de carte UBS et Swisscard (section 3.4), et la garantie de loyer déposée auprès de Firstcaution.",
        ],
      },
    ]),
    ...buildChapitreAvecSousSections("Analyse de la capacité réelle de remboursement", [
      {
        titre: "Capacité théorique de remboursement",
        paragraphes: [
          "Sur la base de mon revenu net actuel (CHF 12'958.58/mois) et du budget futur proposé (CHF 12'821.67/mois, section 3.3), ma capacité théorique de remboursement s'élève à CHF 136.91 par mois avant constitution d'une réserve de sécurité.",
        ],
      },
      {
        titre: "Capacité réelle de remboursement",
        paragraphes: [
          "Compte tenu de l'absence d'épargne liquide (section 5.4) et de l'incertitude professionnelle documentée au chapitre 7, il est nécessaire de conserver une marge de sécurité plutôt que d'engager l'intégralité de ce montant théorique. Ma capacité réelle de remboursement, prudente et soutenable, s'établit à CHF 100.– par mois.",
        ],
      },
      {
        titre: "Équilibre budgétaire",
        paragraphes: [
          "Le budget futur proposé en section 3.3 repose sur une réduction volontaire et significative de mes dépenses discrétionnaires (loisirs/vacances, dépenses personnelles, dépenses exceptionnelles), permettant de passer d'un déficit mensuel réel de CHF -2'680.86 en 2025 à un équilibre positif de CHF +136.91 avant remboursement.",
        ],
      },
      {
        titre: "Réserve financière nécessaire",
        paragraphes: [
          "En l'absence d'épargne disponible (section 5.4), une réserve mensuelle minimale doit être maintenue pour faire face à des dépenses imprévues (santé, réparations, imprévus liés à mes enfants) sans devoir recourir à nouveau à l'endettement. Le solde résiduel de CHF 36.91 par mois après remboursement constitue cette réserve, volontairement limitée.",
        ],
      },
      {
        titre: "Proposition financière soutenable",
        paragraphes: [
          "Sur la base de cette analyse, je propose un remboursement mensuel de CHF 100.–, détaillé au chapitre 8.",
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
          "Au vu de l'ensemble des éléments exposés ci-dessus — charges réelles 2025, budget réduit proposé pour l'avenir proche, et risque concret d'une baisse de revenu significative dès novembre 2026 en cas de retraite anticipée — je propose un remboursement mensuel de CHF 100.–, montant que je m'engage à honorer de manière régulière et soutenable. Ce montant est proposé comme étant soutenable tant dans le meilleur que dans le pire des scénarios professionnels décrits au chapitre 7.",
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
