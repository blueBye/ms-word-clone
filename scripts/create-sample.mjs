import { mkdirSync, writeFileSync } from 'node:fs';
import { zipSync, strToU8 } from 'fflate';

const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:b/><w:sz w:val="40"/></w:rPr><w:t>A document in two directions</w:t></w:r></w:p>
  <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr><w:t>Welcome to your document workspace. Click this paragraph and type, or use File to open a DOCX document.</w:t></w:r></w:p>
  <w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rFonts w:cs="Tahoma"/><w:szCs w:val="28"/><w:rtl/></w:rPr><w:t>این یک پاراگراف فارسی است. می‌توانید فایل خود را باز کنید و نتیجه را ببینید.</w:t></w:r></w:p>
  <w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rFonts w:cs="Tahoma"/><w:szCs w:val="28"/><w:rtl/></w:rPr><w:t>هذه فقرة باللغة العربية لاختبار اتجاه النص والخط.</w:t></w:r></w:p>
  <w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:cs="Tahoma"/><w:szCs w:val="24"/><w:rtl/></w:rPr><w:t xml:space="preserve">متن ترکیبی: </w:t></w:r>
    <w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">English 123 </w:t></w:r>
    <w:r><w:rPr><w:rFonts w:cs="Tahoma"/><w:szCs w:val="24"/><w:rtl/></w:rPr><w:t>و عدد فارسی ۱۲۳</w:t></w:r>
  </w:p>
  <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:i/><w:sz w:val="22"/></w:rPr><w:t>Use Home to format text, change lists and apply styles. Download DOCX saves an edited copy to your device.</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`;

const sample = zipSync({
  '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
  '_rels/.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
  'word/document.xml': strToU8(xml),
});

const output = new URL('../apps/web/public/', import.meta.url);
mkdirSync(output, { recursive: true });
writeFileSync(new URL('sample.docx', output), sample);
console.log('Created apps/web/public/sample.docx');
