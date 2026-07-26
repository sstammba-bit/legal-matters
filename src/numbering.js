const { LevelFormat, AlignmentType } = require("docx");
const { cmToTwip } = require("./units");

// Trois compteurs totalement indépendants (trois "reference" distinctes) :
// un insert au milieu d'une liste ne peut jamais affecter les deux autres.
// Tout est en décimal arabe à tous les niveaux (1. / 1.1 / 1.1.1) pour éviter
// les conversions romain<->arabe entre niveaux liés, source classique de
// corruption de numérotation dans Word.

const CHAPITRE_NUMBERING = "chapitre-numbering";
const ALLEGATION_NUMBERING = "allegation-numbering";
const CONCLUSIONS_NUMBERING = "conclusions-numbering";

const numberingConfig = {
  config: [
    {
      reference: CHAPITRE_NUMBERING,
      levels: [
        {
          level: 0,
          format: LevelFormat.DECIMAL,
          text: "%1.",
          alignment: AlignmentType.START,
          isLegalNumberingStyle: true,
          style: {
            style: "Heading1",
            paragraph: { indent: { left: 0, hanging: cmToTwip(0.6) } },
          },
        },
        {
          level: 1,
          format: LevelFormat.DECIMAL,
          text: "%1.%2.",
          alignment: AlignmentType.START,
          isLegalNumberingStyle: true,
          style: {
            style: "Heading2",
            paragraph: { indent: { left: cmToTwip(0.6), hanging: cmToTwip(0.6) } },
          },
        },
        {
          level: 2,
          format: LevelFormat.DECIMAL,
          text: "%1.%2.%3.",
          alignment: AlignmentType.START,
          isLegalNumberingStyle: true,
          style: {
            style: "Heading3",
            paragraph: { indent: { left: cmToTwip(1.2), hanging: cmToTwip(0.6) } },
          },
        },
      ],
    },
    {
      // Numérotation continue des allégations, indépendante des chapitres.
      reference: ALLEGATION_NUMBERING,
      levels: [
        {
          level: 0,
          format: LevelFormat.DECIMAL,
          text: "%1.",
          alignment: AlignmentType.START,
          style: {
            style: "Allegation",
            paragraph: { indent: { left: cmToTwip(0.6), hanging: cmToTwip(0.6) } },
          },
        },
      ],
    },
    {
      // Liste des conclusions finales, indépendante des deux autres.
      reference: CONCLUSIONS_NUMBERING,
      levels: [
        {
          level: 0,
          format: LevelFormat.DECIMAL,
          text: "%1.",
          alignment: AlignmentType.START,
          style: {
            paragraph: { indent: { left: cmToTwip(0.6), hanging: cmToTwip(0.6) } },
          },
        },
      ],
    },
  ],
};

module.exports = {
  numberingConfig,
  CHAPITRE_NUMBERING,
  ALLEGATION_NUMBERING,
  CONCLUSIONS_NUMBERING,
};
