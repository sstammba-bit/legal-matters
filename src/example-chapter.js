const {
  Paragraph,
  TextRun,
  Bookmark,
  InternalHyperlink,
  PageReference,
  HeadingLevel,
} = require("docx");
const { CONCLUSIONS_NUMBERING } = require("./numbering");

// Chapitre de démonstration : illustre chaque style du modèle avec un
// contenu fictif. À SUPPRIMER par l'utilisateur avant de rédiger un vrai
// mémoire (cf. README.md à côté du .dotx).
//
// Des Bookmark sont posés sur le premier chapitre et la première allégation
// pour prouver que le mécanisme de renvoi (Insertion → Renvoi → Signet /
// Page du signet) fonctionne. Pour ses futures allégations, l'utilisateur
// posera lui-même un signet (Insertion → Signet) sur les paragraphes qu'il
// veut pouvoir cibler par renvoi de page/texte — voir limitations dans le
// README. Le renvoi "Élément numéroté" fonctionne nativement sur TOUTES les
// allégations sans signet, puisqu'elles appartiennent à allegation-numbering.

const P = (style, text) =>
  new Paragraph({ style, children: [new TextRun({ text })] });

function buildExampleChapter() {
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new Bookmark({
          id: "chap_objet",
          children: [new TextRun("Objet")],
        }),
      ],
    }),
    P(
      "Normal",
      "EXEMPLE — Le présent mémoire a pour objet de démontrer l'application des styles définis dans ce modèle. Ce chapitre entier doit être supprimé avant la rédaction d'un mémoire réel."
    ),

    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun("Historique")],
    }),
    P(
      "Normal",
      "EXEMPLE — Bref rappel chronologique des faits de la cause (à remplacer par le contenu réel)."
    ),

    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new Bookmark({
          id: "chap_analyse",
          children: [new TextRun("Analyse juridique")],
        }),
      ],
    }),

    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("Faits allégués")],
    }),

    new Paragraph({
      style: "Allegation",
      children: [
        new Bookmark({
          id: "alleg_1",
          children: [
            new TextRun(
              "EXEMPLE — Le 1er janvier 2026, la partie demanderesse a adressé un courrier recommandé à la partie défenderesse."
            ),
          ],
        }),
      ],
    }),
    P("Preuve", "Preuve :"),
    P("Piece", "Pièce 1 — Courrier recommandé du 1er janvier 2026"),

    P(
      "Allegation",
      "EXEMPLE — La partie défenderesse n'a pas répondu dans le délai imparti de 30 jours."
    ),
    P("Preuve", "Preuve :"),
    P("Piece", "Pièce 2 — Accusé de réception postal"),
    P("Piece", "Pièce 3 — Relevé des délais"),

    P(
      "Allegation",
      "EXEMPLE — Ce troisième allégué illustre la numérotation continue : il reste numéroté séquentiellement même si l'on insère un allégué au milieu du document."
    ),

    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun("Discussion juridique")],
    }),

    new Paragraph({
      heading: HeadingLevel.HEADING_3,
      children: [new TextRun("Base légale applicable")],
    }),
    P(
      "CitationJuridique",
      "EXEMPLE FICTIF — ATF 000 I 0 consid. 0 : « citation d'illustration, à remplacer par une référence jurisprudentielle réelle. »"
    ),
    new Paragraph({
      style: "Observation",
      children: [
        new TextRun(
          "EXEMPLE — Observation : commentaire de liaison entre deux sections, sans numérotation, avec un fond légèrement grisé pour le distinguer visuellement du texte normal."
        ),
      ],
    }),

    new Paragraph({
      heading: HeadingLevel.HEADING_3,
      children: [new TextRun("Illustration d'un renvoi interne")],
    }),
    new Paragraph({
      style: "Normal",
      children: [
        new TextRun("EXEMPLE — Voir le chapitre "),
        new InternalHyperlink({
          anchor: "chap_objet",
          children: [new TextRun({ text: "« Objet »", underline: {} })],
        }),
        new TextRun(", page "),
        new PageReference("chap_objet"),
        new TextRun("."),
      ],
    }),
  ];
}

// Chapitre "Conclusions" : titre en gras (style Conclusions) suivi d'une
// liste numérotée indépendante (conclusions-numbering, arabe décimal).
function buildConclusionsChapter() {
  return [
    new Paragraph({
      style: "Conclusions",
      children: [new TextRun("Conclusions")],
    }),
    new Paragraph({
      style: "Normal",
      numbering: { reference: CONCLUSIONS_NUMBERING, level: 0 },
      children: [new TextRun("EXEMPLE — Constater que...")],
    }),
    new Paragraph({
      style: "Normal",
      numbering: { reference: CONCLUSIONS_NUMBERING, level: 0 },
      children: [new TextRun("EXEMPLE — Condamner la partie défenderesse à...")],
    }),
    new Paragraph({
      style: "Normal",
      numbering: { reference: CONCLUSIONS_NUMBERING, level: 0 },
      children: [new TextRun("EXEMPLE — Sous suite de frais et dépens.")],
    }),
  ];
}

module.exports = { buildExampleChapter, buildConclusionsChapter };
