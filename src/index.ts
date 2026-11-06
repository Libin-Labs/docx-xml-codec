export type DocxRun = { text: string; bold?: boolean; italic?: boolean };
export interface DocxParagraph {
  style?: string;
  runs: DocxRun[];
}
export interface DocxDocument {
  paragraphs: DocxParagraph[];
}
export interface DocxXmlLimits {
  maxCharacters?: number;
  maxParagraphs?: number;
  maxRuns?: number;
}

const limits = (value: DocxXmlLimits = {}) => {
  const result = {
    maxCharacters: value.maxCharacters ?? 2_000_000,
    maxParagraphs: value.maxParagraphs ?? 100_000,
    maxRuns: value.maxRuns ?? 200_000,
  };
  if (
    Object.values(result).some((number) => !Number.isSafeInteger(number) || number < 1) ||
    result.maxCharacters > 20_000_000 ||
    result.maxParagraphs > 500_000 ||
    result.maxRuns > 1_000_000
  )
    throw new RangeError('XML limits must be positive safe integers.');
  return result;
};

function validate(
  document: DocxDocument,
  maxParagraphs = 100_000,
  maxRuns = 200_000,
  maxCharacters = 20_000_000,
): void {
  if (
    !document ||
    !Array.isArray(document.paragraphs) ||
    document.paragraphs.length > maxParagraphs
  )
    throw new RangeError('Invalid or oversized document.');
  let runs = 0;
  let characters = 0;
  for (const paragraph of document.paragraphs) {
    if (!paragraph || !Array.isArray(paragraph.runs))
      throw new TypeError('Paragraph runs must be an array.');
    if (paragraph.style !== undefined && !/^[A-Za-z][A-Za-z0-9_-]{0,63}$/.test(paragraph.style))
      throw new TypeError('Invalid paragraph style id.');
    for (const run of paragraph.runs) {
      if (
        typeof run.text !== 'string' ||
        (run.bold !== undefined && typeof run.bold !== 'boolean') ||
        (run.italic !== undefined && typeof run.italic !== 'boolean')
      )
        throw new TypeError('Invalid run.');
      if (++runs > maxRuns) throw new RangeError('Run limit exceeded.');
      characters += run.text.length;
      if (characters > maxCharacters) throw new RangeError('Text limit exceeded.');
    }
  }
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export function writeDocxDocumentXml(document: DocxDocument, options: DocxXmlLimits = {}): string {
  const bound = limits(options);
  validate(document, bound.maxParagraphs, bound.maxRuns, bound.maxCharacters);
  let characterCount = 0;
  const paragraphs = document.paragraphs
    .map((paragraph) => {
      const properties = paragraph.style
        ? `<w:pPr><w:pStyle w:val="${escapeXml(paragraph.style)}"/></w:pPr>`
        : '';
      const runs = paragraph.runs
        .map((run) => {
          characterCount += run.text.length;
          if (characterCount > bound.maxCharacters) throw new RangeError('Text limit exceeded.');
          const formatting =
            run.bold || run.italic
              ? `<w:rPr>${run.bold ? '<w:b/>' : ''}${run.italic ? '<w:i/>' : ''}</w:rPr>`
              : '';
          return `<w:r>${formatting}<w:t xml:space="preserve">${escapeXml(run.text)}</w:t></w:r>`;
        })
        .join('');
      return `<w:p>${properties}${runs}</w:p>`;
    })
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`;
}

function decodeXml(text: string): string {
  if (/&(?!(?:#x[\da-fA-F]+|#\d+|amp|lt|gt|quot|apos);)/.test(text))
    throw new SyntaxError('Unknown or malformed XML entity.');
  return text.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_match, entity: string) => {
    if (entity === 'amp') return '&';
    if (entity === 'lt') return '<';
    if (entity === 'gt') return '>';
    if (entity === 'quot') return '"';
    if (entity === 'apos') return "'";
    const number =
      entity[1]!.toLowerCase() === 'x'
        ? Number.parseInt(entity.slice(2), 16)
        : Number.parseInt(entity.slice(1), 10);
    if (
      !Number.isInteger(number) ||
      !(
        number === 0x9 ||
        number === 0xa ||
        number === 0xd ||
        (number >= 0x20 && number <= 0xd7ff) ||
        (number >= 0xe000 && number <= 0xfffd) ||
        (number >= 0x10000 && number <= 0x10ffff)
      )
    )
      throw new SyntaxError('Invalid XML character reference.');
    return String.fromCodePoint(number);
  });
}

function parseAttributes(source: string): Record<string, string> {
  const attributes: Record<string, string> = {};
  const pattern = /\s+([A-Za-z_:][\w:.-]*)="([^"]*)"/gy;
  let index = 0;
  while (index < source.length) {
    pattern.lastIndex = index;
    const match = pattern.exec(source);
    if (!match) {
      if (source.slice(index).trim())
        throw new SyntaxError('Malformed or unsupported XML attribute.');
      break;
    }
    if (Object.hasOwn(attributes, match[1]!))
      throw new SyntaxError(`Duplicate XML attribute ${match[1]}.`);
    attributes[match[1]!] = decodeXml(match[2]!);
    index = pattern.lastIndex;
  }
  return attributes;
}

export function readDocxDocumentXml(xml: string, options: DocxXmlLimits = {}): DocxDocument {
  const bound = limits(options);
  if (typeof xml !== 'string') throw new TypeError('XML input must be a string.');
  if (xml.length > bound.maxCharacters * 6 + 1_000_000)
    throw new RangeError('XML input exceeds configured size limit.');
  for (const character of xml) {
    const code = character.codePointAt(0)!;
    if (!(
      code === 0x9 ||
      code === 0xa ||
      code === 0xd ||
      (code >= 0x20 && code <= 0xd7ff) ||
      (code >= 0xe000 && code <= 0xfffd) ||
      (code >= 0x10000 && code <= 0x10ffff)
    ))
      throw new SyntaxError('XML input contains a forbidden character.');
  }
  if (/<!DOCTYPE|<!ENTITY/i.test(xml))
    throw new SyntaxError('DTD and entity declarations are forbidden.');
  const tokens = xml.match(/<\?xml\s+[^?]*\?>|<!--[\s\S]*?-->|<[^>]*>|[^<]+/g) ?? [];
  if (tokens.join('') !== xml) throw new SyntaxError('Malformed XML token stream.');
  const allowed = new Set([
    'w:document',
    'w:body',
    'w:p',
    'w:pPr',
    'w:pStyle',
    'w:r',
    'w:rPr',
    'w:b',
    'w:i',
    'w:t',
    'w:sectPr',
  ]);
  const stack: string[] = [];
  const document: DocxDocument = { paragraphs: [] };
  let paragraph: DocxParagraph | undefined;
  let run: DocxRun | undefined;
  let textCount = 0;
  let runCount = 0;
  let declarationSeen = false;
  let rootSeen = false;
  let rootClosed = false;
  let bodySeen = false;
  for (const token of tokens) {
    if (token.startsWith('<?')) {
      if (
        declarationSeen ||
        stack.length ||
        !/^<\?xml\s+version="1\.0"(?:\s+encoding="UTF-8")?(?:\s+standalone="yes")?\?>$/i.test(token)
      )
        throw new SyntaxError('Unsupported XML declaration.');
      declarationSeen = true;
      continue;
    }
    if (token.startsWith('<!--'))
      throw new SyntaxError('XML comments are outside the supported subset.');
    if (token.startsWith('</')) {
      const match = /^<\/([A-Za-z_:][\w:.-]*)\s*>$/.exec(token);
      if (!match || stack.pop() !== match[1])
        throw new SyntaxError('Mismatched XML closing element.');
      if (match[1] === 'w:document') rootClosed = true;
      if (match[1] === 'w:r') run = undefined;
      if (match[1] === 'w:p') {
        if (!paragraph) throw new SyntaxError('Unexpected paragraph close.');
        if (document.paragraphs.length >= bound.maxParagraphs)
          throw new RangeError('Paragraph limit exceeded.');
        document.paragraphs.push(paragraph);
        paragraph = undefined;
      }
      continue;
    }
    if (token.startsWith('<')) {
      const match = /^<([A-Za-z_:][\w:.-]*)([\s\S]*?)\s*(\/?)>$/.exec(token);
      if (!match || !allowed.has(match[1]!))
        throw new SyntaxError('Unknown or malformed OOXML element.');
      const name = match[1]!;
      const rawAttributes = match[2]!;
      const selfClosing = match[3]!;
      if (rootClosed || (stack.length === 0 && rootSeen))
        throw new SyntaxError('Multiple or trailing XML roots.');
      const attrs = parseAttributes(rawAttributes);
      const parent = stack[stack.length - 1];
      const validParent: Record<string, string[]> = {
        'w:document': [''],
        'w:body': ['w:document'],
        'w:p': ['w:body'],
        'w:pPr': ['w:p'],
        'w:pStyle': ['w:pPr'],
        'w:r': ['w:p'],
        'w:rPr': ['w:r'],
        'w:b': ['w:rPr'],
        'w:i': ['w:rPr'],
        'w:t': ['w:r'],
        'w:sectPr': ['w:body'],
      };
      if (!validParent[name]!.includes(parent ?? ''))
        throw new SyntaxError(`Invalid OOXML nesting for ${name}.`);
      const allowedAttributes: Record<string, string[]> = {
        'w:document': ['xmlns:w'],
        'w:pStyle': ['w:val'],
        'w:b': ['w:val'],
        'w:i': ['w:val'],
        'w:t': ['xml:space'],
      };
      if (Object.keys(attrs).some((key) => !allowedAttributes[name]?.includes(key)))
        throw new SyntaxError(`Unsupported ${name} attributes.`);
      if (
        (name === 'w:b' || name === 'w:i') &&
        attrs['w:val'] !== undefined &&
        !['0', '1', 'true', 'false', 'on', 'off'].includes(attrs['w:val']!)
      )
        throw new SyntaxError(`Unsupported ${name} value.`);
      if (
        name === 'w:t' &&
        attrs['xml:space'] !== undefined &&
        !['preserve', 'default'].includes(attrs['xml:space']!)
      )
        throw new SyntaxError('Unsupported XML space policy.');
      if (name === 'w:document') {
        if (rootSeen) throw new SyntaxError('Duplicate document root.');
        rootSeen = true;
        if (
          attrs['xmlns:w'] !== 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' ||
          Object.keys(attrs).some((key) => key !== 'xmlns:w')
        )
          throw new SyntaxError('Unsupported OOXML namespaces.');
      } else if (name === 'w:body') {
        if (bodySeen) throw new SyntaxError('Duplicate document body.');
        bodySeen = true;
      } else if (name === 'w:p') {
        paragraph = { runs: [] };
      } else if (name === 'w:pStyle') {
        if (
          !paragraph ||
          Object.keys(attrs).length !== 1 ||
          !/^[A-Za-z][A-Za-z0-9_-]{0,63}$/.test(attrs['w:val'] ?? '')
        )
          throw new SyntaxError('Invalid paragraph style.');
        paragraph.style = attrs['w:val']!;
      } else if (name === 'w:r') {
        if (++runCount > bound.maxRuns) throw new RangeError('Run limit exceeded.');
        run = { text: '' };
        paragraph!.runs.push(run);
      } else if (name === 'w:b' && run)
        run.bold = attrs['w:val'] !== '0' && attrs['w:val'] !== 'false' && attrs['w:val'] !== 'off';
      else if (name === 'w:i' && run)
        run.italic =
          attrs['w:val'] !== '0' && attrs['w:val'] !== 'false' && attrs['w:val'] !== 'off';
      if (selfClosing) {
        if (!['w:pStyle', 'w:b', 'w:i', 'w:sectPr'].includes(name))
          throw new SyntaxError(`Unexpected empty ${name}.`);
      } else stack.push(name);
      continue;
    }
    if (!stack.length) {
      if (token.trim()) throw new SyntaxError('Text outside document root.');
      continue;
    }
    if (stack[stack.length - 1] === 'w:t') {
      if (!run) throw new SyntaxError('Text is outside a run.');
      const value = decodeXml(token);
      run.text += value;
      textCount += value.length;
      if (textCount > bound.maxCharacters) throw new RangeError('Text limit exceeded.');
    } else if (token.trim()) throw new SyntaxError('Unexpected text in OOXML structure.');
  }
  if (stack.length || paragraph || run || document.paragraphs.length > bound.maxParagraphs)
    throw new SyntaxError('Incomplete or oversized OOXML document.');
  if (!rootSeen || !rootClosed || !bodySeen || tokens.length === 0)
    throw new SyntaxError('Missing or incomplete OOXML document root/body.');
  return document;
}
