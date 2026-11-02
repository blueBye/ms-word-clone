import type { CSSProperties } from 'react';
const paths: Record<string,string>={
  document:'M6 3h8l4 4v14H6z M14 3v5h4 M9 12h6 M9 16h6',
  folder:'M3 7h7l2 2h9l-3 11H3z M3 7V4h7l2 3h8v2',
  save:'M4 3h14l3 3v15H3V3z M7 3v6h10V3 M7 21v-8h10v8',
  undo:'M8 5 3 10l5 5 M3 10h10a7 7 0 0 1 7 7',
  redo:'M16 5l5 5-5 5 M21 10H11a7 7 0 0 0-7 7',
  search:'M10 3a7 7 0 1 0 0 14 7 7 0 0 0 0-14 M15 15l6 6',
  print:'M6 8V3h12v5 M6 17H3V9h18v8h-3 M6 14h12v7H6z M17 11h1',
  clipboard:'M9 4H5v17h14V4h-4 M9 2h6v5H9z',
  cut:'M9 8l12 12 M9 16 21 4 M5 3a4 4 0 1 0 0 8 4 4 0 0 0 0-8 M5 13a4 4 0 1 0 0 8 4 4 0 0 0 0-8',
  copy:'M8 8h13v13H8z M16 5V2H2v14h3',
  image:'M3 3h18v18H3z M3 17l6-7 5 6 3-3 4 5 M16 6h1',
  table:'M3 4h18v16H3z M3 9h18 M3 14h18 M9 4v16 M15 4v16',
  link:'M10 14l4-4 M8 16l-1 1a4 4 0 0 1-6-6l5-5a4 4 0 0 1 6 0 M16 8l1-1a4 4 0 0 1 6 6l-5 5a4 4 0 0 1-6 0',
  page:'M5 3h14v18H5z M8 7h8 M8 11h8 M8 15h5',
  margins:'M3 3h18v18H3z M7 3v18 M17 3v18 M3 7h18 M3 17h18',
  orientation:'M3 3h11v17H3z M17 8h4v13h-8 M7 7h4 M7 11h4',
  alignLeft:'M3 5h18 M3 10h12 M3 15h18 M3 20h12',
  alignCenter:'M3 5h18 M6 10h12 M3 15h18 M6 20h12',
  alignRight:'M3 5h18 M9 10h12 M3 15h18 M9 20h12',
  justify:'M3 5h18 M3 10h18 M3 15h18 M3 20h18',
  list:'M8 5h13 M8 12h13 M8 19h13 M3 5h1 M3 12h1 M3 19h1',
  columns:'M3 4h7v16H3z M14 4h7v16h-7z',
  info:'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18 M12 10v7 M12 6v1',
  chevron:'M7 10l5 5 5-5', close:'M5 5l14 14 M5 19 19 5',
  plus:'M12 4v16 M4 12h16', minus:'M4 12h16', check:'M4 12l5 5L20 6',
  settings:'M4 6h16 M4 12h16 M4 18h16 M8 3v6 M16 9v6 M10 15v6',
};
export function Icon({name,size=20,style}:{name:string;size?:number;style?:CSSProperties}) {
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.65" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" style={style}><path d={paths[name] ?? paths.document}/></svg>;
}
