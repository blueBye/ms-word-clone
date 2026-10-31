import React from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type {
  Block,
  Border,
  Borders,
  DocumentIR,
  DrawingInline,
  Inline,
  PageLayout,
  Paragraph,
  RunProperties,
  Table,
  TableCell,
  TextRun,
} from '../model/document.js';

const COMPLEX =
  /\p{Script_Extensions=Arabic}|\p{Script_Extensions=Hebrew}|\p{Script_Extensions=Syriac}|\p{Script_Extensions=Thaana}/u;
const EAST_ASIAN =
  /\p{Script_Extensions=Han}|\p{Script_Extensions=Hiragana}|\p{Script_Extensions=Katakana}|\p{Script_Extensions=Hangul}/u;
const NEUTRAL = /^[\p{P}\p{Z}\p{M}\p{Cf}]+$/u;
const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
type Script = 'ascii' | 'hAnsi' | 'eastAsia' | 'cs';

function fontFamily(
  name: string | undefined,
  complex = false,
): string | undefined {
  return name
    ? `${JSON.stringify(name)}, ${complex ? 'Tahoma, ' : ''}Arial, sans-serif`
    : undefined;
}
function runStyle(p: RunProperties, script: Script): CSSProperties {
  const complex = script === 'cs';
  const cs = complex ? p.complexScript : undefined;
  const size = cs?.fontSizePt ?? p.fontSizePt;
  const bold = complex ? (cs?.bold ?? false) : p.bold;
  const italic = complex ? (cs?.italic ?? false) : p.italic;
  const decoration = [
    p.underline && p.underline !== 'none' ? 'underline' : '',
    p.strike ? 'line-through' : '',
  ]
    .filter(Boolean)
    .join(' ');
  return {
    fontFamily: fontFamily(
      complex
        ? (cs?.fontFamily ?? p.fontFamily)
        : (p.fontSlots?.[script as 'ascii' | 'hAnsi' | 'eastAsia'] ??
            p.fontFamily),
      complex,
    ),
    fontSize: size === undefined ? undefined : `${size}pt`,
    fontWeight: bold === undefined ? undefined : bold ? 700 : 400,
    fontStyle: italic === undefined ? undefined : italic ? 'italic' : 'normal',
    color: p.color,
    backgroundColor: p.backgroundColor,
    textDecorationLine: decoration || undefined,
    textDecorationStyle:
      p.underline === 'double'
        ? 'double'
        : p.underline?.includes('dash')
          ? 'dashed'
          : p.underline === 'dotted'
            ? 'dotted'
            : p.underline === 'wave'
              ? 'wavy'
              : undefined,
    verticalAlign:
      p.verticalAlign === 'superscript'
        ? 'super'
        : p.verticalAlign === 'subscript'
          ? 'sub'
          : p.verticalAlign,
    position: 'relative',
    zIndex: 1,
    top: p.positionPt ? `${-p.positionPt}pt` : undefined,
    letterSpacing:
      p.letterSpacingPt === undefined ? undefined : `${p.letterSpacingPt}pt`,
    textTransform: p.caps ? 'uppercase' : undefined,
    fontVariantCaps: p.smallCaps ? 'small-caps' : undefined,
    display: p.hidden ? 'none' : undefined,
    whiteSpace: 'pre-wrap',
  };
}
function segments(run: TextRun): { text: string; script: Script }[] {
  if (run.properties.forceComplexScript || run.properties.direction === 'rtl')
    return [{ text: run.text, script: 'cs' }];
  const result: { text: string; script: Script }[] = [];
  for (const { segment } of graphemes.segment(run.text)) {
    const script: Script = COMPLEX.test(segment)
      ? 'cs'
      : EAST_ASIAN.test(segment)
        ? 'eastAsia'
        : NEUTRAL.test(segment) && result.length
          ? result[result.length - 1].script
          : segment.codePointAt(0)! < 128
            ? 'ascii'
            : 'hAnsi';
    if (result.at(-1)?.script === script)
      result[result.length - 1].text += segment;
    else result.push({ text: segment, script });
  }
  return result.length ? result : [{ text: '', script: 'ascii' }];
}
function borderValue(border: Border | undefined): string | undefined {
  if (!border) return undefined;
  if (border.style === 'nil' || border.style === 'none') return 'hidden';
  const style =
    border.style === 'double'
      ? 'double'
      : /dash/i.test(border.style)
        ? 'dashed'
        : /dot/i.test(border.style)
          ? 'dotted'
          : 'solid';
  return `${border.widthPt}pt ${style} ${border.color}`;
}
function borderStyle(borders: Borders | undefined): CSSProperties {
  return {
    borderTop: borderValue(borders?.top),
    borderRight: borderValue(borders?.right),
    borderBottom: borderValue(borders?.bottom),
    borderLeft: borderValue(borders?.left),
  };
}
interface RenderContext {
  page?: PageLayout;
}

function Drawing({
  drawing: d,
  context,
}: {
  drawing: DrawingInline;
  context: RenderContext;
}) {
  const Container = d.blocks ? 'div' : 'span';
  const anchor = d.position;
  const horizontal = anchor?.horizontal;
  const vertical = anchor?.vertical;
  const pageLeft = context.page?.margins?.left ?? 72;
  const pageTop = context.page?.margins?.top ?? 72;
  const style: CSSProperties = {
    display: 'inline-block',
    width: `${d.widthPt}pt`,
    height: `${d.heightPt}pt`,
    verticalAlign: 'middle',
    position: anchor ? 'absolute' : 'relative',
    overflow: d.blocks ? 'visible' : 'hidden',
    backgroundColor: d.shape?.fill,
    border: d.shape?.stroke
      ? `${d.shape.strokeWidthPt ?? 1}pt solid ${d.shape.stroke}`
      : undefined,
    borderRadius:
      d.shape?.kind === 'ellipse'
        ? '50%'
        : d.shape?.kind === 'roundRect'
          ? '12%'
          : undefined,
    transform: d.shape?.rotation ? `rotate(${d.shape.rotation}deg)` : undefined,
    zIndex: anchor ? (anchor.behindText ? 0 : 2) : undefined,
  };
  if (anchor) {
    const pageRelative = horizontal?.relativeTo === 'page';
    const left = (horizontal?.offsetPt ?? 0) - (pageRelative ? pageLeft : 0);
    if (horizontal?.alignment === 'right')
      style.right = pageRelative ? `${-pageLeft}pt` : 0;
    else if (horizontal?.alignment === 'center') {
      style.left = '50%';
      style.marginLeft = `${-d.widthPt / 2}pt`;
    } else style.left = `${left}pt`;
    style.top = `${(vertical?.offsetPt ?? 0) - (vertical?.relativeTo === 'page' ? pageTop : 0)}pt`;
  }
  const crop = d.image?.crop;
  const width = Math.max(0.001, 1 - (crop?.left ?? 0) - (crop?.right ?? 0));
  const height = Math.max(0.001, 1 - (crop?.top ?? 0) - (crop?.bottom ?? 0));
  return (
    <Container style={style} data-docx-drawing={anchor ? 'anchor' : 'inline'}>
      {d.image && (
        <img
          src={d.image.dataUrl}
          alt={d.alt}
          draggable={false}
          style={{
            display: 'block',
            position: 'absolute',
            maxWidth: 'none',
            width: `${100 / width}%`,
            height: `${100 / height}%`,
            left: `${(-(crop?.left ?? 0) / width) * 100}%`,
            top: `${(-(crop?.top ?? 0) / height) * 100}%`,
          }}
        />
      )}
      {d.blocks && (
        <div
          style={{
            position: 'relative',
            width: '100%',
            height: '100%',
            whiteSpace: 'normal',
          }}
        >
          {renderBlocks(d.blocks, context)}
        </div>
      )}
      {!d.image && !d.shape && !d.blocks && (
        <span
          role="img"
          aria-label={d.alt || 'Unsupported image'}
          style={{ fontSize: '9pt' }}
        >
          [Image]
        </span>
      )}
    </Container>
  );
}
function renderInline(
  inline: Inline,
  index: number,
  context: RenderContext,
): ReactNode {
  if (inline.type === 'drawing')
    return <Drawing key={index} drawing={inline} context={context} />;
  if (inline.type === 'break')
    return (
      <span
        key={index}
        data-docx-break={inline.kind}
        style={{
          display: 'block',
          height: '18pt',
          breakBefore: inline.kind === 'page' ? 'page' : 'column',
        }}
        aria-label={`${inline.kind} break`}
      />
    );
  const pieces = segments(inline);
  const content =
    pieces.length === 1 ? (
      <span
        dir={inline.properties.direction}
        style={{
          ...runStyle(inline.properties, pieces[0].script),
          direction: inline.properties.direction,
          unicodeBidi: inline.properties.direction ? 'embed' : undefined,
        }}
      >
        {inline.text}
      </span>
    ) : (
      <span
        dir={inline.properties.direction}
        style={{ whiteSpace: 'pre-wrap' }}
      >
        {pieces.map((piece, i) => (
          <span key={i} style={runStyle(inline.properties, piece.script)}>
            {piece.text}
          </span>
        ))}
      </span>
    );
  return inline.href && /^(https?:|mailto:|#)/i.test(inline.href) ? (
    <a
      key={index}
      href={inline.href}
      rel="noopener noreferrer"
      style={{ color: 'inherit' }}
    >
      {content}
    </a>
  ) : (
    React.cloneElement(content, { key: index })
  );
}
function renderParagraph(
  block: Paragraph,
  key: number,
  context: RenderContext,
  previous?: Block,
  next?: Block,
) {
  const Container = block.children.some(
    (inline) => inline.type === 'drawing' && inline.blocks,
  )
    ? 'div'
    : 'p';
  const p = block.properties;
  const line = p.lineSpacing;
  const mark = block.markProperties;
  const sameStyle = (other: Block | undefined) =>
    other?.type === 'paragraph' && other.properties.styleId === p.styleId;
  const style: CSSProperties = {
    position: 'relative',
    textAlign: p.alignment ?? 'start',
    whiteSpace: 'pre-wrap',
    marginTop: `${p.contextualSpacing && sameStyle(previous) ? 0 : (p.spaceBeforePt ?? 0)}pt`,
    marginBottom: `${p.contextualSpacing && sameStyle(next) ? 0 : (p.spaceAfterPt ?? 0)}pt`,
    marginLeft:
      p.indentLeftPt === undefined ? undefined : `${p.indentLeftPt}pt`,
    marginRight:
      p.indentRightPt === undefined ? undefined : `${p.indentRightPt}pt`,
    marginInlineStart:
      p.indentStartPt === undefined ? undefined : `${p.indentStartPt}pt`,
    marginInlineEnd:
      p.indentEndPt === undefined ? undefined : `${p.indentEndPt}pt`,
    textIndent: p.list
      ? 0
      : p.firstLinePt === undefined
        ? undefined
        : `${p.firstLinePt}pt`,
    lineHeight: line
      ? line.rule === 'auto'
        ? line.value
        : line.rule === 'atLeast'
          ? `max(${line.value}pt, 1.2em)`
          : `${line.value}pt`
      : 'normal',
    minHeight: line?.rule === 'atLeast' ? `${line.value}pt` : undefined,
    backgroundColor: p.backgroundColor,
    ...borderStyle(p.borders),
    breakAfter: p.keepNext ? 'avoid' : undefined,
    breakInside: p.keepLines ? 'avoid' : undefined,
    breakBefore: p.pageBreakBefore ? 'page' : undefined,
    tabSize: p.tabs?.[0]?.positionPt ? `${p.tabs[0].positionPt}pt` : 4,
    ...(!block.children.length && mark
      ? runStyle(mark, p.direction === 'rtl' ? 'cs' : 'ascii')
      : {}),
  };
  return (
    <Container
      key={key}
      dir={p.direction}
      style={style}
      data-docx-style={p.styleId}
      data-docx-paragraph=""
    >
      {p.list && (
        <span
          aria-label="List marker"
          style={{
            ...runStyle(
              p.list.properties,
              p.direction === 'rtl' ? 'cs' : 'hAnsi',
            ),
            position: 'absolute',
            insetInlineStart: `${-Math.max(12, Math.abs(p.firstLinePt ?? 18))}pt`,
            width: `${Math.max(12, Math.abs(p.firstLinePt ?? 18))}pt`,
            textAlign: 'start',
            whiteSpace: 'pre',
          }}
        >
          {p.list.marker}
          {p.list.suffix === 'nothing' ? '' : ' '}
        </span>
      )}
      {block.children.length ? (
        block.children.map((inline, i) => renderInline(inline, i, context))
      ) : (
        <br />
      )}
    </Container>
  );
}
function renderTable(table: Table, key: number, context: RenderContext) {
  const p = table.properties;
  const width = p.width
    ? `${p.width.value}${p.width.unit === 'pt' ? 'pt' : '%'}`
    : table.columnsPt.length
      ? `${table.columnsPt.reduce((a, b) => a + b, 0)}pt`
      : '100%';
  const cellBorders = (cell: TableCell, rowIndex: number): Borders => ({
    top: rowIndex === 0 ? p.borders?.top : p.borders?.insideH,
    bottom:
      rowIndex + cell.rowSpan === table.rows.length
        ? p.borders?.bottom
        : p.borders?.insideH,
    left:
      p.direction === 'rtl'
        ? cell.column + cell.columnSpan === table.columnsPt.length
          ? p.borders?.left
          : p.borders?.insideV
        : cell.column === 0
          ? p.borders?.left
          : p.borders?.insideV,
    right:
      p.direction === 'rtl'
        ? cell.column === 0
          ? p.borders?.right
          : p.borders?.insideV
        : cell.column + cell.columnSpan === table.columnsPt.length
          ? p.borders?.right
          : p.borders?.insideV,
    ...cell.borders,
  });
  return (
    <table
      key={key}
      dir={p.direction ?? 'ltr'}
      style={{
        width,
        maxWidth: '100%',
        tableLayout: p.layout === 'fixed' ? 'fixed' : 'auto',
        borderCollapse: 'collapse',
        backgroundColor: p.backgroundColor,
        ...borderStyle(p.borders),
        marginLeft:
          p.alignment === 'center' || p.alignment === 'right'
            ? 'auto'
            : p.indentPt
              ? `${p.indentPt}pt`
              : 0,
        marginRight: p.alignment === 'center' ? 'auto' : 0,
      }}
    >
      <colgroup>
        {table.columnsPt.map((column, i) => (
          <col key={i} style={{ width: `${column}pt` }} />
        ))}
      </colgroup>
      <tbody>
        {table.rows.map((row, rowIndex) => (
          <tr
            key={rowIndex}
            style={{
              height: row.heightPt ? `${row.heightPt}pt` : undefined,
              breakInside: row.cantSplit ? 'avoid' : undefined,
            }}
          >
            {row.cells
              .filter((cell) => !cell.continuation)
              .map((cell, i) => {
                const margins = {
                  top: 0,
                  right: 5.4,
                  bottom: 0,
                  left: 5.4,
                  ...p.cellMargins,
                  ...cell.margins,
                };
                return (
                  <td
                    key={i}
                    colSpan={cell.columnSpan}
                    rowSpan={cell.rowSpan}
                    style={{
                      ...borderStyle(cellBorders(cell, rowIndex)),
                      verticalAlign:
                        cell.verticalAlign === 'center'
                          ? 'middle'
                          : (cell.verticalAlign ?? 'top'),
                      backgroundColor: cell.backgroundColor,
                      padding: `${margins.top}pt ${margins.right}pt ${margins.bottom}pt ${margins.left}pt`,
                      width: cell.width
                        ? `${cell.width.value}${cell.width.unit === 'pt' ? 'pt' : '%'}`
                        : undefined,
                      overflowWrap: 'anywhere',
                    }}
                  >
                    {renderBlocks(cell.blocks, context)}
                  </td>
                );
              })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
function renderBlocks(blocks: Block[], context: RenderContext) {
  return blocks.map((block, index) =>
    block.type === 'paragraph'
      ? renderParagraph(
          block,
          index,
          context,
          blocks[index - 1],
          blocks[index + 1],
        )
      : renderTable(block, index, context),
  );
}
export interface DocumentRendererProps {
  document: DocumentIR;
  className?: string;
  style?: CSSProperties;
  pageLayout?: boolean;
}
/** Flow preview with document page settings; not a Word pagination engine. */
export function DocumentRenderer({
  document,
  className,
  style,
  pageLayout = false,
}: DocumentRendererProps) {
  const page = document.page,
    margins = page?.margins;
  const context = { page: pageLayout ? page : undefined };
  const rgb = document.backgroundColor
    ?.slice(1)
    .match(/../g)
    ?.map((component) => parseInt(component, 16));
  const dark = rgb && rgb[0] * 0.2126 + rgb[1] * 0.7152 + rgb[2] * 0.0722 < 128;
  const pageStyle: CSSProperties = pageLayout
    ? {
        width: `${page?.widthPt ?? 595.3}pt`,
        minHeight: `${page?.heightPt ?? 841.9}pt`,
        boxSizing: 'border-box',
        padding: `${margins?.top ?? 72}pt ${margins?.right ?? 72}pt ${margins?.bottom ?? 72}pt ${margins?.left ?? 72}pt`,
        margin: '0 auto',
        backgroundColor: document.backgroundColor ?? '#fff',
        color: dark ? '#fff' : '#000',
        fontSize: '12pt',
        lineHeight: 'normal',
      }
    : { backgroundColor: document.backgroundColor };
  const variant = page?.differentFirstPage ? 'first' : 'default';
  const header = document.headers?.[variant],
    footer = document.footers?.[variant] ?? document.footers?.default;
  return (
    <div
      className={className}
      dir="ltr"
      style={{
        position: 'relative',
        isolation: 'isolate',
        ...pageStyle,
        ...style,
      }}
      data-docx-engine="preview-2"
    >
      {pageLayout && header && (
        <div
          data-docx-story="header"
          style={{
            position: 'absolute',
            top: `${margins?.header ?? 36}pt`,
            left: `${margins?.left ?? 72}pt`,
            right: `${margins?.right ?? 72}pt`,
          }}
        >
          {renderBlocks(header, context)}
        </div>
      )}
      {renderBlocks(document.blocks, context)}
      {pageLayout && footer && (
        <div
          data-docx-story="footer"
          style={{
            position: 'relative',
            marginTop: `${Math.max(18, (margins?.bottom ?? 72) - (margins?.footer ?? 36))}pt`,
          }}
        >
          {renderBlocks(footer, context)}
        </div>
      )}
    </div>
  );
}
