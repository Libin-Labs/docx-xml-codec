# @lp-labs/docx-xml-codec

Safe OOXML document.xml subset serializer/parser without ZIP container handling.

## Install

```sh
npm install @lp-labs/docx-xml-codec
```

## Quick start

```js
import { writeDocxDocumentXml, readDocxDocumentXml } from '@lp-labs/docx-xml-codec';

const document = { paragraphs: [{ style: 'Heading1', runs: [{ text: 'Notes', bold: true }] }] };
const xml = writeDocxDocumentXml(document);
console.log(readDocxDocumentXml(xml).paragraphs[0].runs[0].text);
```

## API and type contract

- `DocxDocument` is `{ paragraphs: DocxParagraph[] }`; each paragraph has optional style id and ordered `DocxRun[]`. Runs contain plain text and optional `bold` / `italic` booleans.
- `writeDocxDocumentXml(document, limits?)` emits a standalone XML string. `readDocxDocumentXml(xml, limits?)` parses this subset into a fresh object.
- `DocxXmlLimits` defaults to 2,000,000 text characters, 100,000 paragraphs, and 200,000 runs. Upper configurable ceilings are 20,000,000 characters, 500,000 paragraphs, and 1,000,000 runs.

## OOXML interoperability limits, bounds, and errors

The serializer uses the standard WordprocessingML namespace and supports `w:document`, `w:body`, paragraphs, paragraph style ids, runs, `w:b`, `w:i`, `w:t`, and `w:sectPr`. It interoperates at the `word/document.xml` part level only. It does **not** create/read DOCX ZIP containers, content types, relationships, styles parts, numbering, tables, headers, footnotes, comments, tracked changes, or embedded media.

The parser rejects DTDs, entity declarations, comments, unknown elements/attributes, invalid nesting, and unsupported XML characters; it does not resolve external entities or use a DOM. XML input size is capped at `maxCharacters × 6 + 1,000,000`. Invalid object types throw `TypeError`, invalid limits/text sizes throw `RangeError`, and malformed/unsupported XML throws `SyntaxError`. Tokenization is linear in input size; no full XML tree is materialized.

## Customization and limitations

Supply a ZIP/container adapter separately if a complete `.docx` is needed; pass only its uncompressed `word/document.xml` part here. Preserve unsupported document parts through that adapter rather than silently dropping them.

## Module formats

ESM, CommonJS, and TypeScript declarations are included. CommonJS: `const api = require('@lp-labs/docx-xml-codec')`. There are no runtime or peer dependencies; TypeScript and Prettier are development-only.

## Support

If this library is useful to you, consider supporting its maintenance with a [donation via PayPal](https://www.paypal.me/LibinPrasanth).

## Author

Created by [Libin Prasanth](https://www.linkedin.com/in/libinprasanth/). Connect on LinkedIn for updates and discussion.
