import {
  writeDocxDocumentXml,
  readDocxDocumentXml,
  type DocxDocument,
} from '@lp-labs/docx-xml-codec';
const document: DocxDocument = {
  paragraphs: [{ style: 'Heading1', runs: [{ text: 'Notes', bold: true }] }],
};
const xml = writeDocxDocumentXml(document);
console.log(readDocxDocumentXml(xml).paragraphs[0].runs[0].text);
