const d = require('docx');
const {Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, BorderStyle, HeadingLevel} = d;
const AR = /[؀-ۿ]/;
const FONT = 'Arial';
function runs(text, o = {}) {
  // split mixed text into LTR / RTL runs so Arabic renders correctly inside English paragraphs
  const parts = String(text).split(/([؀-ۿ][؀-ۿ\s،؛؟ـ«»،؛؟\d٠-٩\-–—.:()]*[؀-ۿ»)]?)/).filter(s => s !== '');
  return parts.map(p => new TextRun({text: p, font: FONT, size: o.size || 18, bold: o.bold, italics: o.italics,
    color: o.color, rightToLeft: AR.test(p) && !!o.rtlRuns}));
}
function P(text, o = {}) {
  const rtl = o.rtl !== undefined ? o.rtl : false;
  return new Paragraph({bidirectional: rtl, alignment: o.align || (rtl ? AlignmentType.RIGHT : AlignmentType.LEFT),
    spacing: {after: o.after !== undefined ? o.after : 120, before: o.before || 0, line: o.line},
    heading: o.heading, keepNext: o.keepNext,
    children: rtl ? [new TextRun({text: String(text), font: FONT, size: o.size || 20, bold: o.bold, rightToLeft: true})]
                  : runs(text, {size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color, rtlRuns: true})});
}
function H(text, level) { return new Paragraph({heading: level, spacing: {before: 240, after: 120}, keepNext: true, children: runs(text, {size: level === HeadingLevel.HEADING_1 ? 30 : level === HeadingLevel.HEADING_2 ? 26 : 22, bold: true, rtlRuns: true})}); }
const B = {style: BorderStyle.SINGLE, size: 4, color: '999999'};
const borders = {top: B, bottom: B, left: B, right: B};
function cell(content, w, o = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map(c => c instanceof Paragraph ? c : P(c, {size: o.size || 16, rtl: o.rtl, after: 40, bold: o.bold}));
  return new TableCell({width: {size: w, type: WidthType.DXA}, borders, margins: {top: 50, bottom: 50, left: 80, right: 80},
    shading: o.fill ? {type: ShadingType.CLEAR, color: 'auto', fill: o.fill} : undefined, children: paras});
}
function table(widths, rows, o = {}) {
  return new Table({width: {size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA}, columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({tableHeader: i === 0 && o.header !== false, cantSplit: o.cantSplit,
      children: r.map((c, j) => cell(c.v !== undefined ? c.v : c, widths[j], Object.assign({fill: i === 0 && o.header !== false ? 'D9E2F3' : undefined, bold: i === 0 && o.header !== false}, c.o || {}, (o.rtlCols || []).includes(j) ? {rtl: true} : {})))}))});
}
module.exports = {P, H, runs, cell, table, AR, FONT};
