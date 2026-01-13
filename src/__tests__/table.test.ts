/* eslint-env jest */
import path from 'path';
import fs, { unlinkSync, writeFileSync } from 'fs';
import MockDate from 'mockdate';
import QR from 'qrcode';
import { createReport } from '../index';
import { UserOptions } from '../types';
import { setDebugLogSink } from '../debug';
import JSZip from 'jszip';
import docx4js from 'docx4js';

if (process.env.DEBUG) setDebugLogSink(console.log);

it('Dynamic table columns with dynamic rows', async () => {
  const template = await fs.promises.readFile(
    path.join(__dirname, 'fixtures', 'table_falsy_value.docx')
  );
  const result = await createReport({
    template,
    cmdDelimiter: ['##', '##'],
  });

  writeFileSync('/tmp/out_table_falsy_value_1.docx', result);

  const tables = await extractTablesFromReport(Buffer.from(result));
  console.log('Extracted tables:', tables);

  expect(tables[0]).toEqual([['foo', '', '']]);
  expect(tables[1]).toEqual([['foo', '', '']]);
  expect(tables[2]).toEqual([['', '', '']]);
  expect(tables[3]).toEqual([['', 'bar', '']]);

  expect(tables[4]).toEqual([
    ['bar 1', 'bar 2'],
    ['foo', 'foo'],
  ]);
  expect(tables[5]).toEqual([
    ['bar 1', 'foo'],
    ['bar 2', 'foo'],
  ]);
});

async function extractTablesFromReport(report: Buffer) {
  const tempFilename = '/tmp/table_test_' + Date.now() + '.docx';
  writeFileSync(tempFilename, report);

  const docx = await docx4js.load(tempFilename);

  const $ = docx.officeDocument.content;
  const tables = $('w\\:tbl');
  const tablesAsArray = tables.toArray().map((table: cheerio.CheerioAPI) => {
    const rows = $(table).find('w\\:tr');
    return rows.toArray().map((row: cheerio.CheerioAPI) => {
      const cells = $(row).find('w\\:tc');
      return cells.toArray().map((cell: cheerio.CheerioAPI) => {
        const paragraphs = $(cell).find('w\\:p');
        return paragraphs
          .toArray()
          .map((paragraph: cheerio.CheerioAPI) =>
            extractTextOfParagraphTag($, $(paragraph))
          )
          .join('\n');
      });
    });
  });

  unlinkSync(tempFilename);

  return tablesAsArray;
}

function extractTextOfParagraphTag(
  $: cheerio.CheerioAPI,
  element: cheerio.Cheerio
) {
  if (element.is('w\\:p')) {
    return element
      .children()
      .toArray()
      .reduce((accumulatedText, child) => {
        const childElement = $(child);
        if (childElement.is('w\\:hyperlink')) {
          return accumulatedText + extractTextOfHyperlinkTag($, childElement);
        }
        if (childElement.is('w\\:r')) {
          return accumulatedText + extractTextOfRunTag($, childElement);
        }
        return accumulatedText;
      }, '');
  }
  return '';
}

function extractTextOfHyperlinkTag(
  $: cheerio.CheerioAPI,
  element: cheerio.Cheerio
) {
  if (element.is('w\\:hyperlink')) {
    return element
      .children()
      .toArray()
      .reduce((accumulatedText, child) => {
        const childElement = $(child);
        if (childElement.is('w\\:r')) {
          return accumulatedText + extractTextOfRunTag($, childElement);
        }
        return accumulatedText;
      }, '');
  }
  return '';
}

function extractTextOfRunTag($: cheerio.CheerioAPI, element: cheerio.Cheerio) {
  if (element.is('w\\:r')) {
    return element
      .children()
      .toArray()
      .reduce((accumulatedText, child) => {
        const childElement = $(child);
        if (childElement.is('w\\:t')) {
          return accumulatedText + extractTextOfTextTag($, childElement);
        }
        if (childElement.is('w\\:br')) {
          return accumulatedText + extractTextOfBreakTag($, childElement);
        }

        return accumulatedText;
      }, '');
  }
  return '';
}

function extractTextOfTextTag($: cheerio.CheerioAPI, element: cheerio.Cheerio) {
  if (element.is('w\\:t')) {
    return element.text();
  }
  return '';
}

function extractTextOfBreakTag(
  $: cheerio.CheerioAPI,
  element: cheerio.Cheerio
) {
  if (element.is('w\\:br')) {
    return '\n';
  }
  return '';
}
