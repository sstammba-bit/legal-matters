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

function buildChapitreSimple(titre) {
  return [H1(titre), placeholder()];
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
            "Certains postes de charges (notamment l'alimentation, les vacances et les transports) sont constitués d'un très grand nombre de transactions individuelles au cours de l'année, rendant peu praticable la production d'un justificatif distinct pour chacune d'entre elles. Pour ces postes, le montant retenu est extrait de ma comptabilité personnelle, tenue tout au long de l'année à partir des relevés bancaires et de carte de crédit (Pièce 20), qui peut être mise à disposition dans son détail sur demande."
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
          "Mon capital de prévoyance professionnelle (2e pilier, CHF 1'312'189.55 au 1er juin 2026 selon mon certificat de prévoyance, Pièce 16) n'est pas saisissable avant son échéance (art. 39 al. 2 LPP) et ne constitue donc pas une fortune disponible pour le remboursement de cette dette.",
          "Les autres éléments de fortune (biens immobiliers, titres ou placements, autres dettes non encore listées) seront précisés dans une version ultérieure du présent mémoire si nécessaire.",
        ],
      },
      {
        titre: "Enfants à charge",
        paragraphes: [
          "Mon fils Elmo Stammbach, né le 9 septembre 2010, vit avec moi à plein temps (garde de fait). J'assume l'intégralité de son entretien courant.",
          "Ma fille Éline Stammbach, née le 3 novembre 2008, vit chez sa mère. Je verse une contribution d'entretien contractuelle de CHF 980.– par mois, allocations familiales en sus, conformément au jugement de divorce (Pièce 18 — manquante). Je l'accueille un week-end par mois et la moitié des vacances scolaires, soit environ un mois par an.",
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
      { titre: "Incertitudes professionnelles" },
      { titre: "Autres circonstances pertinentes" },
    ]),
    ...buildChapitreAvecSousSections("Proposition de remboursement", [
      { titre: "Principes retenus" },
      {
        titre: "Proposition de plan de remboursement",
        paragraphes: [
          "Au vu de l'ensemble des éléments exposés ci-dessus — charges réelles 2025, budget réduit proposé pour l'avenir proche, et risque concret d'une baisse de revenu significative dès novembre 2026 en cas de retraite anticipée — je propose un remboursement mensuel de CHF 100.–, montant que je m'engage à honorer de manière régulière et soutenable. Ce montant est proposé comme étant soutenable tant dans le meilleur que dans le pire des scénarios professionnels décrits au chapitre 7.",
        ],
      },
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
