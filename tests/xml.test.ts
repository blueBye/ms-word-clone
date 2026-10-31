import { describe, expect, it } from 'vitest';
import { childElements, halfPoints, onOff, parseXml, wordAttribute } from '../src/docx/parser/xml.js';
import { W, W_STRICT } from './helpers/docx.js';

function element(xml: string) {
  return parseXml(xml).documentElement!;
}

describe('OPC XML boundary', () => {
  it.each(['', '<a>', '<a></b>', '<a/><b/>', '<a x="1" x="2"/>', '<a>&unknown;</a>', '<a x=unquoted/>'])('rejects malformed XML: %s', xml => {
    expect(() => parseXml(xml)).toThrow('Invalid XML');
  });
  it.each(['<!DOCTYPE a><a/>', '<!DOCTYPE a SYSTEM "https://example.invalid/external.dtd"><a/>'])('rejects prohibited DTDs: %s', xml => {
    expect(() => parseXml(xml)).toThrow('DOCTYPE');
  });
  it('allows comments, XML declarations and predefined entities', () => {
    const root = element('<?xml version="1.0"?><!-- <!DOCTYPE text> --><a>A &amp; B &#x6CC;</a>');
    expect(root.textContent).toBe('A & B ی');
  });
  it('only returns child elements, in source order', () => {
    expect(childElements(element('<a>text<!--comment--><b/><c/></a>')).map(e => e.localName)).toEqual(['b', 'c']);
  });
});

describe.each([W, W_STRICT])('Word attributes in %s', namespace => {
  it('handles a nonstandard prefix and ignores unqualified/foreign attributes', () => {
    const root = element(`<x:r xmlns:x="${namespace}" xmlns:y="urn:other" x:val="0" y:other="foreign" bare="value"/>`);
    expect(wordAttribute(root, 'val')).toBe('0');
    expect(wordAttribute(root, 'bare')).toBeUndefined();
    expect(wordAttribute(root, 'other')).toBeUndefined();
  });
  it.each(['1', 'true', 'on'])('reads enabled ST_OnOff value %s', value => {
    expect(onOff(element(`<x:b xmlns:x="${namespace}" x:val="${value}"/>`))).toBe(true);
  });
  it.each(['0', 'false', 'off'])('reads disabled ST_OnOff value %s', value => {
    expect(onOff(element(`<x:b xmlns:x="${namespace}" x:val="${value}"/>`))).toBe(false);
  });
  it('distinguishes absence, implicit enable and invalid values', () => {
    expect(onOff(undefined)).toBeUndefined();
    expect(onOff(element(`<x:b xmlns:x="${namespace}"/>`))).toBe(true);
    expect(onOff(element(`<x:b xmlns:x="${namespace}" x:val="invalid"/>`))).toBeUndefined();
  });
  it.each([['1', 0.5], ['23', 11.5], ['24', 12]])('converts %s half-points to %s points', (raw, points) => {
    expect(halfPoints(element(`<x:sz xmlns:x="${namespace}" x:val="${raw}"/>`))).toBe(points);
  });
  it.each(['0', '-2', '2.5', 'NaN', 'Infinity', ''])('does not coerce invalid font size %s', value => {
    expect(halfPoints(element(`<x:sz xmlns:x="${namespace}" x:val="${value}"/>`))).toBeUndefined();
  });
});
