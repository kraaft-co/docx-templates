import path from 'path';
import { zipLoad, zipGetText } from '../zip';
import {
  readContentTypes,
  getMainDoc,
  getMetadata,
  parseTemplate,
} from '../main';
import { createReport } from '../index';
import fs from 'fs';
import { setDebugLogSink } from '../debug';
import { findHighestImgId } from '../processTemplate';

if (process.env.DEBUG) setDebugLogSink(console.log);

describe('[Content_Types].xml parser', () => {
  it('Correctly finds the main document xml file in a regular .docx file', async () => {
    const template = await fs.promises.readFile(
      path.join(__dirname, 'fixtures', 'simpleQuery.docx')
    );
    const zip = await zipLoad(template);
    const content_types = await readContentTypes(zip);
    const main_doc = getMainDoc(content_types);
    expect(main_doc).toStrictEqual('document.xml');
  });
  it('Correctly finds the main document xml file in an Office365 .docx file', async () => {
    const template = await fs.promises.readFile(
      path.join(__dirname, 'fixtures', 'office365.docx')
    );
    const zip = await zipLoad(template);
    const content_types = await readContentTypes(zip);
    const main_doc = getMainDoc(content_types);
    expect(main_doc).toStrictEqual('document2.xml');
  });
});

describe('getMetadata', () => {
  it('finds the number of pages', async () => {
    const template = await fs.promises.readFile(
      path.join(__dirname, 'fixtures', 'simpleQuery.docx')
    );
    expect(await getMetadata(template)).toMatchInlineSnapshot(`
      {
        "category": undefined,
        "characters": 24,
        "company": undefined,
        "created": "2015-08-16T18:55:00Z",
        "creator": "Unga Graorg",
        "description": undefined,
        "lastModifiedBy": "Grau Panea, Guillermo",
        "lastPrinted": undefined,
        "lines": 1,
        "modified": "2016-12-15T11:21:00Z",
        "pages": 1,
        "paragraphs": 1,
        "revision": "32",
        "subject": undefined,
        "template": "Normal.dotm",
        "title": undefined,
        "words": 4,
      }
    `);
  });

  it('smoke test: does not crash on normal docx files', async () => {
    expect.hasAssertions();
    const files = await fs.promises.readdir(
      path.join(__dirname, 'fixtures'),
      'utf-8'
    );
    for (const f of files) {
      if (f.startsWith('~$') || !f.endsWith('.docx')) continue;
      const t = await fs.promises.readFile(path.join(__dirname, 'fixtures', f));
      const metadata = await getMetadata(t);
      expect(typeof metadata.modified).toBe('string');
    }
  });
});

describe('findHighestImgId', () => {
  it('returns 0 when doc contains no images', async () => {
    const template = await fs.promises.readFile(
      path.join(__dirname, 'fixtures', 'imageExistingMultiple.docx')
    );
    const { jsTemplate } = await parseTemplate(template);
    expect(findHighestImgId(jsTemplate)).toBe(3);
  });
});

describe('preprocessTemplate delimiter matching', () => {
  // Regression test: a run of plain text that coincidentally starts matching
  // the command delimiter (e.g. its last character equals the delimiter's
  // first character) must not corrupt the surrounding text when that match
  // is later aborted, even if a paragraph break is crossed while the match
  // is still pending. Previously, `openNode._text += ' '` fired for every
  // `w:p` boundary crossed while a delimiter match was pending, regardless
  // of whether the match ever completed. Since the pending character(s)
  // are only flushed back to the text *after* that space is appended, an
  // aborted match left stray spaces injected just before the held-back
  // character(s).
  it('does not inject spaces when a speculative delimiter match spans paragraph breaks and then fails', async () => {
    const template = await fs.promises.readFile(
      path.join(__dirname, 'fixtures', 'trailingDelimiterPrefixMatch.docx')
    );

    const report = await createReport({
      template,
      cmdDelimiter: 'ZZ',
      data: {},
    });

    const zip = await zipLoad(report);
    const doc = await zipGetText(zip, 'word/document.xml');
    const texts = doc
      ? Array.from(doc.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)).map(m => m[1])
      : [];

    expect(texts).toEqual(['GoodbyeZ', 'The End']);
  });
});
