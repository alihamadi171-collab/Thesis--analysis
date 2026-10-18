const fs = require('fs');
const d = require('docx');
const {Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageOrientation, Footer, Header, PageNumber, PageBreak, TableOfContents, LevelFormat, BorderStyle} = d;
const {P, H, table, runs} = require('./common');
const units = Object.fromEntries(JSON.parse(fs.readFileSync('final_units.json')).map(u => [u.uid, u]));
const allUnits = Object.values(units);
const cb = fs.readFileSync('codebook.tsv', 'utf8').trim().split('\n').slice(1).map(l => l.split('\t'));
const th = JSON.parse(fs.readFileSync('themes.json'));
const tabs = JSON.parse(fs.readFileSync('d2_tables.json'));
const mat = JSON.parse(fs.readFileSync('matrix.json'));
const co = JSON.parse(fs.readFileSync('cooccur.json'));
const demo = JSON.parse(fs.readFileSync('demographics_source.json'));
const Q = {}; fs.readFileSync('quotes.tsv', 'utf8').trim().split('\n').slice(1).forEach(l => { const [k, uid, ar, en] = l.split('\t'); Q[k] = {uid, ar, en}; });
const E = Array.from({length: 10}, (_, i) => 'E' + (i + 1));
const lab = Object.fromEntries(cb.map(r => [r[0], r[2]]));
const c2s = {}; th.themes.forEach(t => t.subs.forEach(s => s.codes.forEach(c => c2s[c] = s.id)));
// code stats
const cnt = {}, parts = {};
allUnits.forEach(u => u.codes.forEach(c => { cnt[c] = (cnt[c] || 0) + 1; (parts[c] = parts[c] || new Set()).add(u.interview); }));
const nsubst = cb.filter(r => r[1] !== '0').length;

function quote(k) {
  const q = Q[k]; if (!q) throw new Error('missing quote ' + k);
  if (!units[q.uid] || !units[q.uid].ar.includes(q.ar)) throw new Error('quote not exact ' + k);
  const bd = {left: {style: BorderStyle.SINGLE, size: 12, color: '8EAADB', space: 8}};
  return [new Paragraph({bidirectional: true, alignment: AlignmentType.RIGHT, indent: {left: 567, right: 567}, border: {right: {style: BorderStyle.SINGLE, size: 12, color: '8EAADB', space: 8}}, spacing: {after: 40, before: 80},
            children: [new TextRun({text: q.ar, font: 'Arial', size: 21, rightToLeft: true})]}),
          new Paragraph({indent: {left: 567, right: 567}, border: bd, spacing: {after: 120},
            children: [new TextRun({text: '"' + q.en + '" ', font: 'Arial', size: 20, italics: true}), new TextRun({text: '(' + q.uid + ')', font: 'Arial', size: 18, bold: true, color: '2F5496'})]})];
}
function tbl(name) {
  if (name === 'participants') {
    const workInt = {E1: 'Not working (NGO contract ended a month earlier)', E2: 'Not working; occasional 3-month guard rotations', E3: 'Not working', E4: 'Not working', E5: 'Occasional home catering', E6: 'Working — garment factory, ~8 years', E7: 'Working — unlicensed school transport', E8: 'Working — bag shop, ~1 month', E9: 'Working — NGO children\'s coach since 2019', E10: 'Irregular day labour'};
    const note = {E2: 'Self-reported age 49 (E2-008–010)', E3: 'Self-reported age 49 (E3-016)', E5: 'Self-reported age 55 (E5-260)', E6: 'Work status contradicts table', E7: 'Work status contradicts table', E9: 'Work status contradicts table; K6 total not reproducible', E10: 'K6 total not reproducible'};
    return table([700, 600, 900, 1500, 1100, 2400, 700, 1738], [['ID', 'Sex', 'Age (table)', 'Setting', 'Work (table)', 'Work (as described in interview)', 'K6 (table)', 'Verification note']].concat(demo.slice(1).map(r => { const id = r[0].match(/^E\d+/)[0]; return [id, r[1], r[2], r[3], r[4], workInt[id], r[5], note[id] || '—']; })));
  }
  if (name === 'codebook') {
    const rows = [['Code', 'English label', 'Arabic label', 'Sub-theme', 'Definition', 'Inclusion / exclusion guidance', 'Units', 'Partic.']];
    cb.forEach(r => rows.push([r[0], r[2], {v: r[3], o: {rtl: true}}, c2s[r[0]] || 'non-thematic', r[4], r[5] || '—', String(cnt[r[0]] || 0), String(parts[r[0]] ? parts[r[0]].size : 0)]));
    return table([1500, 2300, 2200, 900, 3500, 3398, 800, 800], rows);
  }
  if (name === 'structure') {
    const rows = [['Theme', 'Subtheme', 'Codes', 'Units', 'Participants']];
    th.themes.forEach(t => t.subs.forEach((s, i) => {
      const us = allUnits.filter(u => u.codes.some(c => s.codes.includes(c)));
      rows.push([i === 0 ? `${t.id} ${t.en}` : '', `${s.id} ${s.en}`, s.codes.map(c => `${c} (${lab[c]})`).join('; '), String(us.length), `${new Set(us.map(u => u.interview)).size}/10`]);
    }));
    return table([2300, 2600, 3238, 700, 800], rows);
  }
  if (name === 'cooccur') {
    const rows = [['Code A', 'Code B', 'Shared units']];
    co.slice(0, 20).forEach(([a, b, n]) => rows.push([`${a} — ${lab[a]}`, `${b} — ${lab[b]}`, String(n)]));
    return table([4200, 4238, 1200], rows);
  }
  if (name === 'matrix') {
    const rows = [['Subtheme'].concat(E, ['Partic.'])];
    th.themes.forEach(t => t.subs.forEach(s => rows.push([`${s.id} ${s.en}`].concat(E.map(e => String(mat[s.id][e] || 0)), [String(E.filter(e => mat[s.id][e]).length) + '/10']))));
    return table([5000].concat(E.map(() => 900), [1398]), rows);
  }
  if (name === 'matrix2') {
    const rows = [['Theme'].concat(E)];
    th.themes.forEach(t => rows.push([`${t.id} ${t.en}`].concat(E.map(e => t.subs.some(s => mat[s.id][e]) ? '●' : '–'))));
    return table([6398].concat(E.map(() => 900)), rows);
  }
  if (name === 'cross') {
    const C = JSON.parse(fs.readFileSync('d2_cross.json'));
    return table([600, 1300, 2000, 2700, 2900, 2200, 3698], [['ID', 'Sex / setting / age', 'Work before displacement', 'Work in Jordan (interview)', 'Main exclusion experienced', 'Main source of strength', 'Future / rebuilding stance']].concat(C));
  }
  if (tabs[name]) { const n = tabs[name][0].length; const w = n === 3 ? [2600, 2000, 5038] : [2200, 2400, 2400, 2638]; return table(w, tabs[name]); }
  throw new Error('unknown table ' + name);
}
// --- parse content into sections
const lines = fs.readFileSync('d2_content.txt', 'utf8').split('\n');
const sections = []; let cur = {orient: 'portrait', kids: []};
const front = [];
front.push(new Paragraph({spacing: {before: 1600, after: 240}, alignment: AlignmentType.CENTER, children: [new TextRun({text: 'Themes, Codes and Analytical Framework — Final', font: 'Arial', size: 40, bold: true})]}));
front.push(P('An Exploration of the Psychosocial Experiences of Older Syrian Refugees in the context of economic empowerment', {align: AlignmentType.CENTER, size: 26, bold: true}));
front.push(P('استكشاف التجارب النفسية لدى للاجئين السوريين من كبار السن في سياق التمكين الاقتصادي', {rtl: true, align: AlignmentType.CENTER, size: 26, after: 360}));
front.push(P('Master of Clinical Social Work — thesis research · Researcher: Ali Ahmad Al Hamadi', {align: AlignmentType.CENTER, size: 22}));
front.push(P(`Foundation for the Findings chapter: approach, bilingual codebook (${nsubst} substantive codes), five themes and eighteen subthemes, evidence, within- and cross-case analysis, matrices, negative cases, memos and chapter outline`, {align: AlignmentType.CENTER, size: 20, italics: true, after: 480}));
front.push(P('Companion document: Transcript_Table_Analysis_FINAL.docx (all transcript units with codes, translations and back translations). Quotation identifiers (e.g., E4-205) refer to that document.', {align: AlignmentType.CENTER, size: 18}));
front.push(new Paragraph({children: [new PageBreak()]}));
front.push(H('Contents', HeadingLevel.HEADING_1));
front.push(new TableOfContents('Contents', {hyperlink: true, headingStyleRange: '1-2'}));
front.push(P('(If the contents list is empty, right-click it in Word and choose "Update Field".)', {size: 16, italics: true}));
front.push(new Paragraph({children: [new PageBreak()]}));
cur.kids.push(...front);
let para = [];
function flush() { if (para.length) { cur.kids.push(P(para.join(' '), {size: 21, after: 140})); para = []; } }
for (const raw of lines) {
  const l = raw.trimEnd();
  if (!l.trim()) { flush(); continue; }
  let m;
  if ((m = l.match(/^\|section:(\w+)\|$/))) { flush(); sections.push(cur); cur = {orient: m[1], kids: []}; continue; }
  if ((m = l.match(/^\|table:(\w+)\|$/))) { flush(); cur.kids.push(tbl(m[1])); cur.kids.push(P('', {after: 120})); continue; }
  if ((m = l.match(/^> (q\w+)$/))) { flush(); cur.kids.push(...quote(m[1])); continue; }
  if (l.startsWith('### ')) { flush(); cur.kids.push(H(l.slice(4), HeadingLevel.HEADING_3)); continue; }
  if (l.startsWith('## ')) { flush(); cur.kids.push(H(l.slice(3), HeadingLevel.HEADING_2)); continue; }
  if (l.startsWith('# ')) { flush(); cur.kids.push(H(l.slice(2), HeadingLevel.HEADING_1)); continue; }
  if (l.startsWith('- ')) { flush(); cur.kids.push(new Paragraph({numbering: {reference: 'bul', level: 0}, spacing: {after: 80}, children: runs(l.slice(2), {size: 21, rtlRuns: true})})); continue; }
  para.push(l.trim());
}
flush(); sections.push(cur);
// verify all quote keys used exist & count
const used = lines.filter(l => /^> q\w+$/.test(l)).map(l => l.slice(2));
console.log('quotes used', used.length, 'unique', new Set(used).size);
const pg = o => o === 'landscape' ? {size: {width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE}, margin: {top: 1000, bottom: 1000, left: 1000, right: 1000}}
                                  : {size: {width: 11906, height: 16838}, margin: {top: 1134, bottom: 1134, left: 1134, right: 1134}};
const doc = new Document({creator: 'Ali Ahmad Al Hamadi', title: 'Themes, Codes and Analytical Framework — FINAL',
  numbering: {config: [{reference: 'bul', levels: [{level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: {paragraph: {indent: {left: 567, hanging: 283}}}}]}]},
  styles: {default: {document: {run: {font: 'Arial', size: 21}}}},
  sections: sections.filter(s => s.kids.length).map(s => ({properties: {page: pg(s.orient)},
    headers: {default: new Header({children: [P('Themes, Codes and Analytical Framework — FINAL', {size: 16, color: '666666', align: AlignmentType.RIGHT})]})},
    footers: {default: new Footer({children: [new Paragraph({alignment: AlignmentType.CENTER, children: [new TextRun({children: ['Page ', PageNumber.CURRENT], size: 16, font: 'Arial'})]})]})},
    children: s.kids}))});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('written', b.length); });
