import type { DocumentIR, Paragraph, ParagraphFormat, RunFormat } from '@local/docx-ir-engine';
import { Group, Menu, MenuItem, Tool } from './Controls';

export type RibbonTab='Home'|'Insert'|'Layout';
export interface RibbonProps {
  tab:RibbonTab; document:DocumentIR|null; selected?:Paragraph; range?:{start:number;end:number}|null;
  disabled:boolean; canFormat:boolean; canPage:boolean; families:string[];
  onRun:(format:RunFormat)=>void; onParagraph:(format:ParagraphFormat)=>void; onAction:(name:string)=>void;
}
const later='Preserved on import; editing is not available yet.';
export function Ribbon({tab,document,selected,range,disabled,canFormat,canPage,families,onRun,onParagraph,onAction}:RibbonProps) {
  const runs=selected?.children.filter(c=>c.type==='text') ?? [];
  let offset=0;
  const chosen=runs.filter(run=>{const start=offset;offset+=run.text.length;return !range || range.start===range.end || (offset>range.start && start<range.end);});
  const marks=chosen[0]?.properties;
  const states=(key:'bold'|'italic')=>{
    const values=chosen.flatMap(r=>[!!r.properties[key],!!(r.properties.complexScript?.[key] ?? false)]);
    return values.length && values.every(Boolean) ? true : values.some(Boolean) ? 'mixed' as const : false;
  };
  const family=(selected?.properties.direction==='rtl' ? marks?.complexScript?.fontFamily : marks?.fontFamily) ?? '';
  const size=(selected?.properties.direction==='rtl' ? marks?.complexScript?.fontSizePt : marks?.fontSizePt) ?? 12;
  const paragraphDisabled=disabled||!selected;
  return <div className="ribbon" id={`ribbon-${tab}`} role="tabpanel" aria-labelledby={`tab-${tab}`}>
    {tab==='Home'&&<>
      <Group label="Clipboard"><Tool label="Paste" icon="clipboard" large disabled={disabled||!canFormat} onClick={()=>onAction('paste')}/><div className="tool-stack"><Tool label="Cut" icon="cut" disabled={disabled||!range||range.start===range.end||!canFormat} onClick={()=>onAction('cut')}/><Tool label="Copy" icon="copy" disabled={!selected} onClick={()=>onAction('copy')}/><Tool label="Edit text" icon="page" disabled={disabled||!canFormat} onClick={()=>onAction('edit')}/></div></Group>
      <Group label="Font" more={()=>onAction('fonts')}><div className="font-tools"><div className="tool-row">
        <input key={`${selected?.id}-${family}`} className="font-input" aria-label="Font family" list="ribbon-fonts" defaultValue={family} placeholder="Font" disabled={disabled||!canFormat} onBlur={e=>{if(e.target.value.trim()&&e.target.value!==family)onRun({fontFamily:e.target.value});}} onKeyDown={e=>{if(e.key==='Enter')e.currentTarget.blur();}}/>
        <datalist id="ribbon-fonts">{[...new Set(['Calibri','Arial','Tahoma','Times New Roman','Cambria',...families])].map(f=><option key={f} value={f}/>)}</datalist>
        <input key={`${selected?.id}-${size}`} className="size-input" aria-label="Font size" type="number" min="1" max="400" step="0.5" defaultValue={size} disabled={disabled||!canFormat} onBlur={e=>{if(Number(e.target.value)!==size)onRun({fontSizePt:Number(e.target.value)});}} onKeyDown={e=>{if(e.key==='Enter')e.currentTarget.blur();}}/>
        <Tool label="Increase font size" disabled={disabled||!canFormat} onClick={()=>onRun({fontSizePt:Math.min(400,size+1)})}><span className="grow-type">A↑</span></Tool>
        <Tool label="Decrease font size" disabled={disabled||!canFormat} onClick={()=>onRun({fontSizePt:Math.max(1,size-1)})}><span>A↓</span></Tool>
      </div><div className="tool-row">
        <Tool label="Bold" active={states('bold')} disabled={disabled||!canFormat} onClick={()=>onRun({bold:states('bold')!==true})}><b>B</b></Tool>
        <Tool label="Italic" active={states('italic')} disabled={disabled||!canFormat} onClick={()=>onRun({italic:states('italic')!==true})}><i>I</i></Tool>
        <Tool label="Underline" active={!!marks?.underline&&marks.underline!=='none'} disabled={disabled||!canFormat} onClick={()=>onRun({underline:marks?.underline&&marks.underline!=='none'?'none':'single'})}><u>U</u></Tool>
        <Tool label="Strikethrough" active={!!marks?.strike} disabled={disabled||!canFormat} onClick={()=>onRun({strike:!marks?.strike})}><s>ab</s></Tool>
        <Tool label="Subscript" active={marks?.verticalAlign==='subscript'} disabled={disabled||!canFormat} onClick={()=>onRun({verticalAlign:marks?.verticalAlign==='subscript'?'baseline':'subscript'})}>x<sub>2</sub></Tool>
        <Tool label="Superscript" active={marks?.verticalAlign==='superscript'} disabled={disabled||!canFormat} onClick={()=>onRun({verticalAlign:marks?.verticalAlign==='superscript'?'baseline':'superscript'})}>x<sup>2</sup></Tool>
        <Menu label="Highlight"><div className="color-swatches">{[['Yellow','yellow','#ffff00'],['Green','green','#00ff00'],['Cyan','cyan','#00ffff'],['Pink','magenta','#ff00ff'],['Gray','lightGray','#c0c0c0']].map(([label,value,color])=><button key={value} title={label} aria-label={`Highlight ${label}`} style={{background:color}} disabled={disabled||!canFormat} onClick={()=>onRun({highlight:value})}/>)}</div><MenuItem disabled={disabled||!canFormat} onClick={()=>onRun({highlight:'none'})}>No highlight</MenuItem></Menu>
        <label className={`color-tool ${disabled||!canFormat?'is-disabled':''}`} title="Font color"><b>A</b><input aria-label="Font color" type="color" value={marks?.color?.match(/^#[a-f0-9]{6}$/i)?marks.color:'#000000'} disabled={disabled||!canFormat} onChange={e=>onRun({color:e.target.value})}/></label>
      </div></div></Group>
      <Group label="Paragraph" more={()=>onAction('paragraph')}><div className="tool-stack"><div className="tool-row">
        <Menu label="Bullets" icon="list"><MenuItem disabled reason={later}>Bullet library</MenuItem><MenuItem disabled reason={later}>Define New Bullet…</MenuItem></Menu>
        <Menu label="Numbering"><MenuItem disabled reason={later}>Decimal, Persian, Roman and letters</MenuItem><MenuItem disabled reason={later}>Define Number Format…</MenuItem><MenuItem disabled reason={later}>Multilevel List…</MenuItem></Menu>
        <Tool label="Decrease indent" disabled={paragraphDisabled} onClick={()=>onParagraph({indentStartPt:Math.max(0,(selected?.properties.indentStartPt??selected?.properties.indentLeftPt??0)-18)})}>⇤</Tool>
        <Tool label="Increase indent" disabled={paragraphDisabled} onClick={()=>onParagraph({indentStartPt:(selected?.properties.indentStartPt??selected?.properties.indentLeftPt??0)+18})}>⇥</Tool>
      </div><div className="tool-row">
        {(['left','center','right','justify'] as const).map(alignment=><Tool key={alignment} label={`Align ${alignment}`} icon={alignment==='justify'?'justify':`align${alignment[0].toUpperCase()}${alignment.slice(1)}`} active={selected?.properties.alignment===alignment} disabled={paragraphDisabled} onClick={()=>onParagraph({alignment})}/>)}
        <Menu label="Spacing"><p className="menu-heading">Line spacing</p>{[1,1.15,1.5,2,2.5,3].map(value=><MenuItem key={value} disabled={paragraphDisabled} onClick={()=>onParagraph({lineSpacing:value})}>{value.toFixed(value%1?2:1)}</MenuItem>)}<MenuItem disabled={paragraphDisabled} onClick={()=>onAction('paragraph')}>Line Spacing Options…</MenuItem></Menu>
        <Tool label="Left-to-right paragraph" active={selected?.properties.direction==='ltr'} disabled={paragraphDisabled} onClick={()=>onParagraph({direction:'ltr'})}>¶→</Tool>
        <Tool label="Right-to-left paragraph" active={selected?.properties.direction==='rtl'} disabled={paragraphDisabled} onClick={()=>onParagraph({direction:'rtl'})}>←¶</Tool>
      </div></div></Group>
      <Group label="Styles"><button className="style-tile" onClick={()=>onAction('styles')}><span className="style-sample">AaBbCc</span><span>{selected?.properties.styleId ?? 'Document styles'}</span></button></Group>
      <Group label="Editing"><div className="tool-stack"><Tool label="Find" icon="search" onClick={()=>onAction('find')}/><Tool label="Replace" onClick={()=>onAction('replace')}/><Tool label="Paragraph editor" onClick={()=>onAction('edit')} disabled={disabled||!canFormat}/></div></Group>
    </>}
    {tab==='Insert'&&<>
      <Group label="Pages"><div className="tool-stack"><Tool label="New paragraph" icon="page" disabled={disabled||!document} onClick={()=>onAction('insertParagraph')}/><Tool label="Blank page" icon="plus" disabled={disabled||!document} onClick={()=>onAction('blankPage')}/><Tool label="Page break before" disabled={paragraphDisabled} onClick={()=>onParagraph({pageBreakBefore:!selected?.properties.pageBreakBefore})}/></div></Group>
      <Group label="Tables"><Menu label="Table" icon="table" large><MenuItem disabled={disabled||!document} onClick={()=>onAction('table')}>Insert Table…</MenuItem><MenuItem disabled reason={later}>Draw Table</MenuItem><MenuItem disabled reason={later}>Quick Tables</MenuItem><MenuItem disabled reason={later}>Convert Text to Table…</MenuItem></Menu></Group>
      <Group label="Illustrations"><Tool label="Pictures" icon="image" large disabled reason="Image insertion/export is not available yet."/><Menu label="Shapes" large><MenuItem disabled reason={later}>Lines and rectangles</MenuItem><MenuItem disabled reason={later}>Basic shapes and arrows</MenuItem></Menu><Tool label="Chart" large disabled reason={later}/></Group>
      <Group label="Links"><Tool label="Link" icon="link" large disabled reason="Hyperlink editing is not available yet."/><div className="tool-stack"><Tool label="Bookmark" disabled reason={later}/><Tool label="Cross-reference" disabled reason={later}/></div></Group>
      <Group label="Header & Footer"><Menu label="Header" icon="page" large><MenuItem onClick={()=>onAction('stories')}>Inspect header text</MenuItem><MenuItem disabled reason={later}>Insert / remove header</MenuItem></Menu><Menu label="Footer" icon="page" large><MenuItem onClick={()=>onAction('stories')}>Inspect footer text</MenuItem><MenuItem disabled reason={later}>Insert / remove footer</MenuItem></Menu><Tool label="Page number" large disabled reason="Existing PAGE/NUMPAGES fields are displayed. Field insertion is not available yet."/></Group>
      <Group label="Text"><div className="tool-stack"><Tool label="Text Box" disabled reason={later}/><Tool label="Date & Time" disabled={disabled||!canFormat} onClick={()=>onAction('date')}/><Tool label="Quick Parts" disabled reason={later}/></div></Group>
      <Group label="Symbols"><Menu label="Symbol" large><div className="symbol-grid">{['☐','☑','✓','©','®','™','±','×','÷','α','β','Ω','→','←','…','—'].map(symbol=><button key={symbol} disabled={disabled||!canFormat} onClick={()=>onAction(`symbol:${symbol}`)}>{symbol}</button>)}</div><MenuItem disabled={disabled||!canFormat} onClick={()=>onAction('symbol:\u200c')}>Zero-width non-joiner (ZWNJ)</MenuItem><MenuItem disabled={disabled||!canFormat} onClick={()=>onAction('symbol:\u00a0')}>Nonbreaking space</MenuItem></Menu><Tool label="Equation" large disabled reason={later}/></Group>
    </>}
    {tab==='Layout'&&<>
      <Group label="Page Setup" more={()=>onAction('page')}><Menu label="Margins" icon="margins" large>{[['Normal','72'],['Narrow','36'],['Moderate','54']].map(([label,value])=><MenuItem key={label} disabled={disabled||!canPage} onClick={()=>onAction(`margin:${value}`)}>{label}<small>{Number(value)/72} inch</small></MenuItem>)}<MenuItem disabled={disabled||!canPage} onClick={()=>onAction('page')}>Custom Margins…</MenuItem></Menu>
        <Menu label="Orientation" icon="orientation" large>{['Portrait','Landscape'].map(value=><MenuItem key={value} disabled={disabled||!canPage} onClick={()=>onAction(`orientation:${value}`)}>{value}</MenuItem>)}</Menu>
        <Menu label="Size" icon="page" large>{[['A4','595.3,841.9'],['Letter','612,792'],['Legal','612,1008'],['A5','419.5,595.3']].map(([label,value])=><MenuItem key={label} disabled={disabled||!canPage} onClick={()=>onAction(`size:${value}`)}>{label}</MenuItem>)}<MenuItem disabled={disabled||!canPage} onClick={()=>onAction('page')}>More Paper Sizes…</MenuItem></Menu>
        <Menu label="Columns" icon="columns" large><MenuItem disabled reason="Multi-column section layout is not implemented yet.">One / Two / Three</MenuItem><MenuItem disabled reason={later}>More Columns…</MenuItem></Menu>
        <div className="tool-stack"><Menu label="Breaks"><MenuItem disabled={paragraphDisabled} onClick={()=>onParagraph({pageBreakBefore:true})}>Page break before paragraph</MenuItem><MenuItem disabled reason={later}>Column break</MenuItem><MenuItem disabled reason="Full section editing is not implemented yet.">Section: Next / Continuous / Even / Odd</MenuItem></Menu><Tool label="Line Numbers" disabled reason={later}/><Tool label="Hyphenation" disabled reason={later}/></div>
      </Group>
      <Group label="Paragraph" more={()=>onAction('paragraph')}><div className="layout-fields">
        {([['Before','indentStartPt'],['After','indentEndPt']] as const).map(([label,key])=><label key={key}>Indent {label}<input key={`${selected?.id}-${selected?.properties[key]}`} type="number" step="6" defaultValue={selected?.properties[key]??0} disabled={paragraphDisabled} onBlur={e=>{if(Number(e.target.value)!==(selected?.properties[key]??0))onParagraph({[key]:Number(e.target.value)});}}/><span>pt</span></label>)}
      </div><div className="layout-fields">{([['Before','spaceBeforePt'],['After','spaceAfterPt']] as const).map(([label,key])=><label key={key}>Spacing {label}<input key={`${selected?.id}-${selected?.properties[key]}`} type="number" min="0" max="1000" step="6" defaultValue={selected?.properties[key]??0} disabled={paragraphDisabled} onBlur={e=>{if(Number(e.target.value)!==(selected?.properties[key]??0))onParagraph({[key]:Number(e.target.value)});}}/><span>pt</span></label>)}</div></Group>
      <Group label="Arrange"><Tool label="Position" icon="page" large disabled reason={later}/><Tool label="Wrap Text" icon="alignLeft" large disabled reason={later}/><div className="tool-stack"><Tool label="Bring Forward" disabled reason={later}/><Tool label="Send Backward" disabled reason={later}/><Tool label="Selection Pane" onClick={()=>onAction('outline')}/></div><Menu label="Align" large><MenuItem disabled reason={later}>Align to page / margin</MenuItem><MenuItem disabled reason={later}>Distribute objects</MenuItem><MenuItem disabled reason={later}>Group / Rotate</MenuItem></Menu></Group>
    </>}
  </div>;
}
