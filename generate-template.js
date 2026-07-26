const fs = require("fs");
const path = require("path");
const { Packer } = require("docx");
const { buildDocument } = require("./src/build-document");
const { convertDocxBufferToDotx } = require("./convert-to-dotx");

const DEFAULT_OUTPUT_PATH = path.join(
  "G:",
  "My Drive",
  "05. - Admin",
  "09. Legal matters",
  "Modele_Memoire_Juridique.dotx"
);

async function main() {
  const outputPath = process.argv[2] || DEFAULT_OUTPUT_PATH;

  const doc = buildDocument();
  const docxBuffer = await Packer.toBuffer(doc);
  const dotxBuffer = await convertDocxBufferToDotx(docxBuffer);

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, dotxBuffer);

  console.log("Modèle généré :", outputPath);
}

main().catch((err) => {
  console.error("Échec de la génération :", err);
  process.exit(1);
});
