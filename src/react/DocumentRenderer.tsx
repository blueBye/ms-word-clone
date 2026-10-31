import React from 'react';
import type { CSSProperties } from 'react';
import type { DocumentIR, TextRun } from '../model/document.js';

const PERSIAN_ARABIC = /[\u0600-\u06ff\u0750-\u077f\u08a0-\u08ff\ufb50-\ufdff\ufe70-\ufeff]/u;

function runStyle(run: TextRun): CSSProperties {
  const p = run.properties;
  const isComplex = PERSIAN_ARABIC.test(run.text);
  const c = isComplex ? p.complexScript : undefined;
  return {
    fontFamily: c?.fontFamily ?? p.fontFamily,
    fontSize: (c?.fontSizePt ?? p.fontSizePt) === undefined ? undefined : `${c?.fontSizePt ?? p.fontSizePt}pt`,
    fontWeight: (c?.bold ?? p.bold) === undefined ? undefined : (c?.bold ?? p.bold) ? 700 : 400,
    fontStyle: (c?.italic ?? p.italic) === undefined ? undefined : (c?.italic ?? p.italic) ? 'italic' : 'normal',
    direction: p.direction,
    unicodeBidi: p.direction ? 'isolate' : undefined,
    whiteSpace: 'pre-wrap',
  };
}
export interface DocumentRendererProps {
  document: DocumentIR;
  className?: string;
  style?: CSSProperties;
}
/** Render approximation only: layout, styles.xml, tables and sections are out of scope. */
export function DocumentRenderer({ document, className, style }: DocumentRendererProps) {
  return <div className={className} style={style} data-docx-engine="milestone-1">
    {document.blocks.map((block, index) => (
      <p key={index} dir={block.properties.direction} style={{ textAlign: block.properties.alignment, margin: '0 0 0.5em', whiteSpace: 'pre-wrap' }}>
        {block.children.map((run, idx) => <span key={idx} dir={run.properties.direction} style={runStyle(run)}>{run.text}</span>)}
      </p>
    ))}
  </div>;
}
