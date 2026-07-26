const { convertMillimetersToTwip } = require("docx");

const cmToTwip = (cm) => convertMillimetersToTwip(cm * 10);
const ptToHalfPt = (pt) => pt * 2;
const ptToTwips = (pt) => pt * 20;

module.exports = { cmToTwip, ptToHalfPt, ptToTwips };
