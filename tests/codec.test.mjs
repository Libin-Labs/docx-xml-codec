import test from 'node:test';
import assert from 'node:assert/strict';
import { readDocxDocumentXml, writeDocxDocumentXml } from '../dist/esm/index.js';

test('writes and reads OOXML paragraph, style, text, and formatting subset', () => {
  const document = {
    paragraphs: [
      {
        style: 'Heading1',
        runs: [
          { text: 'A & B', bold: true },
          { text: ' plain', italic: true },
        ],
      },
    ],
  };
  const xml = writeDocxDocumentXml(document);
  assert.match(xml, /<w:document/);
  assert.deepEqual(readDocxDocumentXml(xml), document);
});
test('handles empty paragraphs and escaped XML characters', () => {
  const document = { paragraphs: [{ runs: [] }, { runs: [{ text: '<>&"\'' }] }] };
  assert.deepEqual(readDocxDocumentXml(writeDocxDocumentXml(document)), document);
});
test('rejects DTDs, unknown elements, external entities, and malformed nesting', () => {
  assert.throws(() => readDocxDocumentXml('<!DOCTYPE x><x/>'), SyntaxError);
  assert.throws(
    () =>
      readDocxDocumentXml(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:script/></w:body></w:document>',
      ),
    SyntaxError,
  );
  assert.throws(
    () =>
      readDocxDocumentXml(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>&evil;</w:t></w:r></w:p></w:body></w:document>',
      ),
    SyntaxError,
  );
});
