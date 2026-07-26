const JSZip = require("jszip");

const DOCX_MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const DOTX_MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";

// .docx et .dotx ne diffèrent, au niveau du paquet OOXML, que par le
// content-type déclaré pour /word/document.xml (document vs. template) et
// par l'extension de fichier. On génère donc un .docx normalement via
// Packer, puis on ne modifie que cette seule ligne de [Content_Types].xml.
async function convertDocxBufferToDotx(docxBuffer) {
  const zip = await JSZip.loadAsync(docxBuffer);

  const contentTypesEntry = zip.file("[Content_Types].xml");
  if (!contentTypesEntry) {
    throw new Error(
      "[Content_Types].xml introuvable dans le paquet OOXML généré — génération docx invalide."
    );
  }

  let xml = await contentTypesEntry.async("string");

  if (!xml.includes(DOCX_MAIN_CONTENT_TYPE)) {
    throw new Error(
      "Content-type 'document.main+xml' non trouvé dans [Content_Types].xml — la structure du .docx généré a peut-être changé (vérifier la version du package docx)."
    );
  }

  xml = xml.replace(DOCX_MAIN_CONTENT_TYPE, DOTX_MAIN_CONTENT_TYPE);
  zip.file("[Content_Types].xml", xml);

  return zip.generateAsync({ type: "nodebuffer" });
}

module.exports = { convertDocxBufferToDotx };
