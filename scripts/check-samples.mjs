import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { basename } from 'node:path';
import { createHash } from 'node:crypto';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import {
  importDocx,
  openDocxPackage,
  findMainDocumentPart,
  getParagraphs,
  getTextRuns,
} from '../dist/index.js';
import { DocumentRenderer } from '../dist/react/index.js';
import {
  parseXml,
  contentElements,
  isWord,
  childWord,
} from '../dist/docx/parser/xml.js';

mkdirSync(new URL('../.cache/sample-previews/', import.meta.url), {
  recursive: true,
});
for (const path of process.argv.slice(2)) {
  const bytes = readFileSync(path);
  const document = await importDocx(bytes);
  const pkg = await openDocxPackage(bytes);
  const main = parseXml(await pkg.readText(await findMainDocumentPart(pkg)));
  const visibleText = (parent) =>
    contentElements(parent)
      .map((node) => {
        if (isWord(node, 't')) return node.textContent ?? '';
        if (
          ['del', 'moveFrom', 'sdtPr', 'rPr', 'pPr', 'instrText'].some((name) =>
            isWord(node, name),
          )
        )
          return '';
        if (isWord(node, 'sdt')) {
          const content = childWord(node, 'sdtContent');
          return content ? visibleText(content) : '';
        }
        return visibleText(node);
      })
      .join('');
  const sourceText = visibleText(main.documentElement);
  const importedText = getTextRuns(document.blocks)
    .map((run) => run.text)
    .join('');
  const hash = (text) =>
    createHash('sha256').update(text.replace(/\s/g, '')).digest('hex');
  const countTables = (blocks) =>
    blocks.reduce(
      (sum, block) =>
        sum +
        (block.type === 'table'
          ? 1 +
            block.rows.reduce(
              (n, row) =>
                n +
                row.cells.reduce((n, cell) => n + countTables(cell.blocks), 0),
              0,
            )
          : 0),
      0,
    );
  const paragraphs = getParagraphs(document.blocks);
  const html = renderToStaticMarkup(
    createElement(DocumentRenderer, { document, pageLayout: true }),
  );
  const name = basename(path);
  writeFileSync(
    new URL(
      `../.cache/sample-previews/${encodeURIComponent(name)}.html`,
      import.meta.url,
    ),
    `<!doctype html><meta charset="utf-8"><title>Local sample preview</title><body style="margin:24px;background:#eee">${html}</body>`,
  );
  console.log(
    JSON.stringify({
      file: name,
      sourceTextCharacters: sourceText.length,
      importedTextCharacters: importedText.length,
      allBodyTextPreserved: hash(sourceText) === hash(importedText),
      paragraphs: paragraphs.length,
      tables: countTables(document.blocks),
      lists: paragraphs.filter((p) => p.properties.list).length,
      images: (html.match(/<img /g) ?? []).length,
      drawings: (html.match(/data-docx-drawing=/g) ?? []).length,
      diagnostics: document.diagnostics?.map((d) => d.code) ?? [],
    }),
  );
  if (hash(sourceText) !== hash(importedText)) process.exitCode = 1;
}
