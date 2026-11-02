import { useEffect, useMemo, useRef, useState } from 'react';
import type { DocumentIR } from '@local/docx-ir-engine';
import { getRequiredFonts, type FontRequirement } from '@local/docx-ir-engine/react';

export function useDocumentFonts(document: DocumentIR | null) {
  const [revision,setRevision]=useState(0), [missing,setMissing]=useState<FontRequirement[]>([]);
  const [status,setStatus]=useState(''), [family,setFamily]=useState(''), [weight,setWeight]=useState('400'), [style,setStyle]=useState('normal');
  const faces=useRef<FontFace[]>([]), loaded=useRef(new Set<string>());
  const required=useMemo(()=>document ? getRequiredFonts(document) : [],[document]);
  useEffect(()=>{
    let cancelled=false;
    Promise.all(required.map(async font=>{
      if(loaded.current.has(font.family.toLowerCase()))return null;
      const names=[font.family,`${font.family} Regular`,`${font.family.replace(/\s+/g,'')}-Regular`];
      try {await new FontFace('__docx_probe__',names.map(n=>`local(${JSON.stringify(n)})`).join(', ')).load();return null;}catch{return font;}
    })).then(items=>{if(!cancelled)setMissing(items.filter((font):font is FontRequirement=>!!font));});
    return ()=>{cancelled=true;};
  },[required,revision]);
  useEffect(()=>()=>{for(const face of faces.current)window.document.fonts.delete(face);},[]);
  async function load(file: File) {
    if(!family.trim()) {setStatus('Enter the exact font family name first.');return;}
    setStatus('Loading font…');
    try {
      const face=await new FontFace(family.trim(),await file.arrayBuffer(),{weight,style}).load();
      window.document.fonts.add(face);faces.current.push(face);
      if(weight==='400' && style==='normal')loaded.current.add(family.trim().toLowerCase());
      setRevision(r=>r+1);setStatus(`${family} · ${weight==='700'?'Bold':'Regular'}${style==='italic'?' italic':''} loaded for this browser session.`);
    } catch {setStatus('Could not load this font. Choose a TTF, OTF, WOFF or WOFF2 file.');}
  }
  return {revision,missing,required,status,family,setFamily,weight,setWeight,style,setStyle,load};
}
export function FontPanel({fonts}:{fonts:ReturnType<typeof useDocumentFonts>}) {
  return <><p className="pane-description">Use the exact source font for accurate text wrapping. Load each weight and style separately.</p>
    <div className="font-requirements">{fonts.required.map(font=><button key={font.family} onClick={()=>fonts.setFamily(font.family)}><span>{font.family}</span><small>{fonts.missing.some(f=>f.family===font.family)?'Missing':'Available'} · {font.locations.join(', ')}</small></button>)}</div>
    <label className="field">Font family<input value={fonts.family} onChange={e=>fonts.setFamily(e.target.value)} placeholder="e.g. Pofak"/></label>
    <div className="field-row"><label className="field">Weight<select value={fonts.weight} onChange={e=>fonts.setWeight(e.target.value)}><option value="400">Regular (400)</option><option value="700">Bold (700)</option></select></label><label className="field">Style<select value={fonts.style} onChange={e=>fonts.setStyle(e.target.value)}><option value="normal">Normal</option><option value="italic">Italic</option></select></label></div>
    <label className="field">Font file<input type="file" accept=".ttf,.otf,.woff,.woff2" onChange={e=>{const file=e.target.files?.[0];if(file)void fonts.load(file);e.target.value='';}}/></label>
    {fonts.status&&<p role="status" className="pane-description">{fonts.status}</p>}
    <details className="pane-help"><summary>Finding fonts on Windows</summary><p>The Windows Fonts view is a virtual folder. User-installed files are usually in <code>%LOCALAPPDATA%\Microsoft\Windows\Fonts</code>. Paste the font file’s full path into the picker. Use “Pofak”, without Word’s “(Body)” theme label.</p></details>
  </>;
}
