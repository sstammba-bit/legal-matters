const { AlignmentType, ShadingType, LineRuleType } = require("docx");
const { cmToTwip, ptToHalfPt, ptToTwips } = require("./units");
const {
  CHAPITRE_NUMBERING,
  ALLEGATION_NUMBERING,
} = require("./numbering");

const FONT = "Times New Roman";
const BLACK = "000000";
const SINGLE_LINE = { line: 240, lineRule: LineRuleType.AUTO };

// Chaque style définit explicitement son "Next Style" (paragraph.next côté
// Word = propriété `next` ici) : c'est ce qui garantit qu'appuyer sur Entrée
// après une Allégation retombe sur Normal, après une Preuve enchaîne sur une
// Pièce, etc. — exigence explicite du cahier des charges.
const paragraphStyles = [
  {
    id: "Normal",
    name: "Normal",
    quickFormat: true,
    next: "Normal",
    run: { font: FONT, size: ptToHalfPt(12), color: BLACK },
    paragraph: {
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: ptToTwips(6), ...SINGLE_LINE },
    },
  },
  {
    id: "Heading1",
    name: "Titre 1",
    basedOn: "Normal",
    next: "Normal",
    quickFormat: true,
    run: { font: FONT, size: ptToHalfPt(16), bold: true, color: BLACK },
    paragraph: {
      keepNext: true,
      outlineLevel: 0,
      spacing: { before: 480, after: 240 },
      numbering: { reference: CHAPITRE_NUMBERING, level: 0 },
    },
  },
  {
    id: "Heading2",
    name: "Titre 2",
    basedOn: "Normal",
    next: "Normal",
    quickFormat: true,
    run: { font: FONT, size: ptToHalfPt(14), bold: true, color: BLACK },
    paragraph: {
      keepNext: true,
      outlineLevel: 1,
      spacing: { before: 360, after: 180 },
      numbering: { reference: CHAPITRE_NUMBERING, level: 1 },
    },
  },
  {
    id: "Heading3",
    name: "Titre 3",
    basedOn: "Normal",
    next: "Normal",
    quickFormat: true,
    run: { font: FONT, size: ptToHalfPt(12), bold: true, color: BLACK },
    paragraph: {
      keepNext: true,
      outlineLevel: 2,
      spacing: { before: 240, after: 120 },
      numbering: { reference: CHAPITRE_NUMBERING, level: 2 },
    },
  },
  {
    // Un fait = une allégation. Numérotation continue indépendante des
    // chapitres (cf. numbering.js). N'est jamais un Heading : n'apparaît
    // donc jamais dans la table des matières.
    id: "Allegation",
    name: "Allégation",
    basedOn: "Normal",
    next: "Normal",
    quickFormat: true,
    run: { font: FONT, size: ptToHalfPt(12), color: BLACK },
    paragraph: {
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: ptToTwips(6), ...SINGLE_LINE },
      indent: { left: cmToTwip(0.6) },
      numbering: { reference: ALLEGATION_NUMBERING, level: 0 },
    },
  },
  {
    id: "Preuve",
    name: "Preuve",
    basedOn: "Normal",
    next: "Piece",
    quickFormat: true,
    run: { font: FONT, size: ptToHalfPt(11), italics: true, color: BLACK },
    paragraph: {
      spacing: { after: ptToTwips(6), ...SINGLE_LINE },
      indent: { left: cmToTwip(1) },
    },
  },
  {
    // "Référence de pièce" — plusieurs pièces doivent pouvoir s'enchaîner
    // les unes après les autres : next reboucle donc sur Piece lui-même.
    id: "Piece",
    name: "Pièce",
    basedOn: "Normal",
    next: "Piece",
    quickFormat: true,
    run: { font: FONT, size: ptToHalfPt(12), color: BLACK },
    paragraph: {
      spacing: { after: ptToTwips(6), ...SINGLE_LINE },
      indent: { left: cmToTwip(1.5) },
    },
  },
  {
    id: "CitationJuridique",
    name: "Citation juridique",
    basedOn: "Normal",
    next: "Normal",
    quickFormat: true,
    run: { font: FONT, size: ptToHalfPt(11), italics: true, color: BLACK },
    paragraph: {
      spacing: { after: ptToTwips(6), ...SINGLE_LINE },
      indent: { left: cmToTwip(1) },
    },
  },
  {
    id: "Observation",
    name: "Observation",
    basedOn: "Normal",
    next: "Normal",
    quickFormat: true,
    run: { font: FONT, size: ptToHalfPt(12), color: BLACK },
    paragraph: {
      spacing: { after: ptToTwips(6), ...SINGLE_LINE },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: "F2F2F2" },
    },
  },
  {
    // Titre des conclusions finales. La liste numérotée qui suit applique
    // conclusions-numbering directement au niveau du paragraphe (pas via un
    // style dédié), car ce sont de simples paragraphes Normal numérotés.
    id: "Conclusions",
    name: "Conclusions",
    basedOn: "Normal",
    next: "Normal",
    quickFormat: true,
    run: { font: FONT, size: ptToHalfPt(14), bold: true, color: BLACK },
    paragraph: {
      keepNext: true,
      spacing: { before: 360, after: 180 },
    },
  },
];

module.exports = { paragraphStyles, FONT, BLACK, SINGLE_LINE };
