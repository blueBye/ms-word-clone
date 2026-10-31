export const NS = {
  word: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wordStrict: 'http://purl.oclc.org/ooxml/wordprocessingml/main',
  relationships: 'http://schemas.openxmlformats.org/package/2006/relationships',
  officeRelationships: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  officeRelationshipsStrict: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
} as const;
export const WORD_NS = new Set<string>([NS.word, NS.wordStrict]);
