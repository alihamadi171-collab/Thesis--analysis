const fs = require('fs');
const d = require('docx');
const {Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageOrientation, Footer, Header, PageNumber, PageBreak, TableOfContents} = d;
const {P, H, table} = require('./common');
const units = JSON.parse(fs.readFileSync('final_units.json'));
const demo = JSON.parse(fs.readFileSync('demographics_source.json'));
const cb = fs.readFileSync('codebook.tsv', 'utf8').trim().split('\n').slice(1).map(l => l.split('\t'));
const th = JSON.parse(fs.readFileSync('themes.json'));
const fixes = fs.readFileSync('translation_fixes.tsv', 'utf8').trim().split('\n').slice(1).map(l => l.split('\t'));
const W = [1000, 3600, 1800, 3600, 3200, 2198]; // 15398 = A4 landscape minus 0.5" margins
const kids = [];
const T = 'An Exploration of the Psychosocial Experiences of Older Syrian Refugees in the context of economic empowerment';
kids.push(new Paragraph({spacing: {before: 1800, after: 240}, alignment: AlignmentType.CENTER, children: [new TextRun({text: 'Transcript Table Analysis — Final Version', font: 'Arial', size: 44, bold: true})]}));
kids.push(P(T, {align: AlignmentType.CENTER, size: 28, bold: true, after: 120}));
kids.push(P('استكشاف التجارب النفسية لدى للاجئين السوريين من كبار السن في سياق التمكين الاقتصادي', {rtl: true, align: AlignmentType.CENTER, size: 28, after: 400}));
kids.push(P('Master of Clinical Social Work — thesis research', {align: AlignmentType.CENTER, size: 22}));
kids.push(P('Researcher: Ali Ahmad Al Hamadi', {align: AlignmentType.CENTER, size: 22}));
kids.push(P('Complete unit-by-unit analysis of interviews E1–E10: original Arabic transcript, codes, English translation, back translation and themes', {align: AlignmentType.CENTER, size: 20, italics: true, after: 600}));
kids.push(P('Companion document: Themes_Codes_Analytical_Framework_FINAL.docx (codebook, theme development, cross-case analysis, memos and Findings outline).', {align: AlignmentType.CENTER, size: 18}));
kids.push(new Paragraph({children: [new PageBreak()]}));
kids.push(H('Contents', HeadingLevel.HEADING_1));
kids.push(new TableOfContents('Contents', {hyperlink: true, headingStyleRange: '1-2'}));
kids.push(P('(If the contents list is empty, right-click it in Word and choose "Update Field".)', {size: 16, italics: true}));
kids.push(new Paragraph({children: [new PageBreak()]}));

// ---- Section 1: notes
const n = units.length, sub = units.filter(u => u.codes.some(c => !['PROC', 'K6', 'REFLEX', 'CONTEXT'].includes(c))).length;
const corr = units.filter(u => u.corrected).length;
kids.push(H('1. About this final version', HeadingLevel.HEADING_1));
const notes = [
 'This document is the completed and revised version of Transcript_Table_Analysis_Completed.docx. It keeps the original structure of Transcript.docx — a demographic table followed by the ten interviews in the columns Code | Transcript | الترميز | Translation | Back Translation | Themes — and completes the coding, translation checking and theme assignment for every analytical unit of all ten interviews (E1–E10).',
 'Sources and hierarchy. The approved thesis proposal was used for the study title, research questions, objectives and the five-step thematic analysis procedure. Transcript.docx was treated as the authoritative record of what participants said; its demographic table as the authoritative record of participant characteristics. The previous coded transcript, research summaries and earlier AI-generated material were treated as provisional working material and were checked against the Arabic transcripts before anything was kept.',
 `Preservation of the Arabic transcripts. Every line of the ten transcripts in Transcript.docx appears exactly once, in its original order, in the Transcript column. No participant or interviewer wording was corrected, shortened, reordered or merged. This was verified automatically: for each interview, the Transcript column concatenated in order is character-for-character identical to the corresponding source transcript (whitespace aside). The transcripts are therefore complete and recoverable from this document. The ${n.toLocaleString()} units follow the line breaks of the original transcription; some units contain both interviewer and participant speech where the transcription did not separate them (notably E9, which was transcribed in long turns).`,
 'Unit identifiers. The Code column gives a stable unit identifier (for example E4-205 = interview E4, unit 205). Unit numbers are identical to those of the previous version (E4-205 = P04-C205), so earlier notes remain traceable.',
 `Coding. All units were re-coded line by line against a consolidated bilingual codebook of ${cb.length} codes (definitions and inclusion/exclusion rules are in the companion document). The previous version used inconsistent and partly mislabelled codes (for example, K6 frequency answers coded as "divergence" and some statements given unrelated codes); these were replaced. Interviewer questions, consent and back-channel talk are coded PROC; K6 screening exchanges are coded K6 and used descriptively only; background narrative not bearing on the research questions is coded CONTEXT; interviewer turns that state the researcher's own assumptions or suggest an answer are flagged REFLEX for reflexive auditing. ${sub.toLocaleString()} units carry at least one substantive code.`,
 `Translations and back translations. The English translations of the previous version were checked against the Arabic during coding. ${corr} units were re-translated because the earlier rendering was wrong, reversed in meaning, incomplete or technically corrupted (for example "كارمله" = "as a widow", previously "burden"; "بكره" = "tomorrow", previously "I hate"; "اعمار" = "reconstruction", previously "ages"; "خلاصه عمر" = "the essence of a lifetime", previously "we are finished"). Each correction is listed with its reason in Appendix A. For corrected units a new back translation was produced. The Back Translation column renders the English back into Arabic so that meaning can be compared with the original; it is a checking device and does not replace the transcript. Square brackets [ ] mark words added for clarity, and [—] marks wording that is unclear in the transcription. Speaker turns are marked [Interviewer:] / [Participant:] in corrected translations where a unit mixes speakers.`,
 'Themes. The Themes column lists the subtheme(s) to which the unit\'s codes belong (5 themes, 18 subthemes; see section 3 and the companion document). Procedural, K6 and contextual units are labelled as such and are not used as thematic evidence.',
 'Software statement. The analysis followed an ATLAS.ti-style workflow (segmentation, open and descriptive coding, constant comparison, codebook consolidation, code–theme networks, participant-by-theme matrices, memos and an audit trail), implemented through line-by-line reading and transparent scripts. No proprietary qualitative-analysis software (ATLAS.ti, NVivo, MAXQDA) was used.',
 'Points for the researcher\'s review. (1) The demographic contradictions listed in section 2.3 should be resolved from the original records before submission. (2) Some passages (E7-463 to E7-487 and E7-553 to E7-567) contain sensitive disclosures; they are retained verbatim as required but should not be quoted, and redaction should be considered for participant protection. (3) The demographic table and some transcript lines contain participants\' names; pseudonymised identifiers (E1–E10) should be used in the thesis. (4) Quotations chosen for the Findings chapter should receive a final native-speaker check of the translation.'
];
notes.forEach(t => kids.push(P(t, {size: 20})));

// ---- Section 2: demographics
kids.push(H('2. Participant demographic characteristics', HeadingLevel.HEADING_1));
kids.push(H('2.1 Demographic table — retained exactly as in the source file', HeadingLevel.HEADING_2));
kids.push(table([2600, 1300, 1300, 2400, 2600, 2600], demo.map(r => r), {rtlCols: []}));
kids.push(P('Column headings in English: Code | Sex | Age | Place of residence | Current work status | K6 psychological screening score. NA in the source = not currently working.', {size: 18, italics: true, before: 80}));
kids.push(H('2.2 Summary of the source table', HeadingLevel.HEADING_2));

kids.push(P('The source table lists ten participants: four women (E1, E5, E6, E9) and six men (E2, E3, E4, E7, E8, E10); six live in a camp setting (E1, E2, E3, E4, E9, E10; the interviews refer to Za\'atari camp) and four in urban Amman (E5, E6, E7, E8). Recorded ages range from 50 to 65 (mean 54.3). Only E8 is recorded as currently working ("YES"). Recorded K6 totals range from 12 to 24.', {size: 20}));
kids.push(H('2.3 Verification against the interviews — contradictions flagged (table not altered)', HeadingLevel.HEADING_2));
const ver = [['Issue', 'What the source table says', 'What the interviews show', 'Status / recommendation'],
 ['"All participants are older than 50"', 'Ages 50, 50, 50 for E2, E3, E9; others 51–65', 'Three participants are exactly 50, not older than 50. E2 says he is 49 and "will reach 50" (E2-008 to E2-010); E3 says "49 and some" (E3-016) and later "50" (E3-383). E5 says 55 (E5-260) vs 54 in the table. E10 says he arrived aged about 38 and has been displaced about 15 years (E10-077 to E10-080; E10-224).', 'Describe the sample as "aged 50–65 (the proposal\'s criterion)", not "older than 50". Check E2\'s and E3\'s dates of birth against the inclusion criterion (50–65).'],
 ['"Only one participant is currently employed"', 'E8 = YES; all others NA', 'E8: confirmed — selling in a friend\'s bag shop for about a month (E8-124, E8-282). But E6 works daily in a garment factory and has done so for about eight years (E6-037, E6-073 to E6-076, E6-225); E7 drives school students in an unlicensed car as current paid work (E7-102, E7-357 to E7-364); E9 has worked continuously for NGOs since 2016, currently with Save the Children (E9-021, E9-047). E10 does irregular day labour (E10-055, E10-195).', 'Contradiction for E6, E7 and E9. Clarify how "current work status" was defined (e.g., formal registered employment at recruitment). Do not report "only one is employed" without this clarification.'],
 ['K6 totals', 'E1 13, E2 21, E3 12, E4 24, E5 14, E6 18, E7 18, E8 13, E9 16, E10 13', 'The scale was explained inconsistently: 0–4 in E1; 1–4 in E3–E8 and E10; four verbal options in E9. Totals for E1–E8 can be reproduced from the answers as recorded. E9\'s answers (sometimes, never, sometimes, never, never, never; E9-006 to E9-016) give 4 (0–4) or 10 (1–4), not 16. E10\'s answers (E10-016 to E10-047) give 12 (0–4) or 18 (1–4), not 13.', 'Totals are not comparable across interviews and two cannot be reproduced. Use K6 only descriptively, report the scoring problem as a limitation, and do not interpret K6 diagnostically.']];
kids.push(table([2200, 2600, 6300, 4298], ver));

// ---- Section 3: legend
kids.push(H('3. How to read the interview tables', HeadingLevel.HEADING_1));
kids.push(P('Code = unit identifier. Transcript = original Arabic, unchanged. الترميز = code(s) assigned to the unit, with the Arabic code label. Translation = English translation. Back Translation = Arabic back-rendering of the English, for checking meaning. Themes = subtheme(s) the codes belong to. Grey rows are procedural (interviewer questions, consent, back-channel); blue rows are K6 screening exchanges.', {size: 20}));
kids.push(H('3.1 Code legend', HeadingLevel.HEADING_2));
const t2 = {}; th.themes.forEach(t => t.subs.forEach(s => s.codes.forEach(c => t2[c] = s.id)));
kids.push(table([2000, 4200, 3600, 1600, 3998], [['Code', 'English label', 'Arabic label', 'Subtheme', 'Definition (short)']].concat(cb.map(r => [r[0], r[2], {v: r[3], o: {rtl: true}}, t2[r[0]] || (r[1] === '0' ? '—' : ''), r[4]]))));
kids.push(H('3.2 Themes and subthemes', HeadingLevel.HEADING_2));
th.themes.forEach(t => { kids.push(P(`${t.id}. ${t.en}`, {bold: true, size: 20, after: 40})); kids.push(P(t.ar, {rtl: true, size: 20, after: 60}));
  t.subs.forEach(s => kids.push(P(`${s.id}  ${s.en} — ${s.ar}  [${s.codes.join(', ')}]`, {size: 18, after: 40}))); });

// ---- Interviews
const dem = {}; demo.slice(1).forEach(r => { const id = r[0].match(/^E\d+/)[0]; dem[id] = r; });
for (let i = 1; i <= 10; i++) {
  const e = 'E' + i, us = units.filter(u => u.interview === e), r = dem[e];
  const s = us.filter(u => u.codes.some(c => !['PROC', 'K6', 'REFLEX', 'CONTEXT'].includes(c))).length;
  kids.push(new Paragraph({children: [new PageBreak()]}));
  kids.push(H(`${i + 3}. Interview ${e}`, HeadingLevel.HEADING_1));
  kids.push(P(`Participant ${e} · ${r[1] === 'F' ? 'Female' : 'Male'} · age ${r[2]} (source table) · ${r[3]} · current work (source table): ${r[4]} · K6 (source table): ${r[5]} · ${us.length} units, of which ${s} substantive`, {size: 18, italics: true}));
  const rows = [['Code', 'Transcript', 'الترميز', 'Translation', 'Back Translation', 'Themes']];
  us.forEach(u => {
    const proc = u.codes.every(c => ['PROC', 'REFLEX'].includes(c)), k6 = u.codes.includes('K6') && u.codes.every(c => ['K6', 'PROC'].includes(c));
    const fill = proc ? 'F2F2F2' : k6 ? 'E8F0F8' : undefined;
    const codeP = [].concat(...u.coding.map(c => { const [a, b] = c.split(' — '); return [P(a, {size: 15, bold: true, after: 0}), P(b, {rtl: true, size: 15, after: 40})]; }));
    const themeP = [].concat(...u.themes.map(t => { const k = t.lastIndexOf(' / '); return [P(t.slice(0, k), {size: 15, after: 0}), P(t.slice(k + 3), {rtl: true, size: 15, after: 40})]; }));
    rows.push([{v: u.uid + (u.corrected ? ' *' : ''), o: {fill}}, {v: u.ar, o: {rtl: true, size: 18, fill}}, {v: codeP, o: {fill}},
      {v: u.en, o: {fill}}, {v: u.back, o: {rtl: true, size: 18, fill}}, {v: themeP, o: {fill}}]);
  });
  kids.push(table(W, rows));
  kids.push(P('* = translation and back translation corrected in this version (see Appendix A).', {size: 16, italics: true, before: 60}));
}
// ---- Appendix A
kids.push(new Paragraph({children: [new PageBreak()]}));
kids.push(H('Appendix A. Translation corrections log', HeadingLevel.HEADING_1));
kids.push(P(`The ${fixes.length} units below received a new English translation and back translation. The corrected text appears in the interview tables; this log records why. Units not listed keep the previous translation, which was checked against the Arabic during coding.`, {size: 20}));
kids.push(table([1600, 13798], [['Unit', 'Reason for correction']].concat(fixes.map(f => [f[0].replace(/^P0?(\d+)-C/, 'E$1-'), f[3] || 'Meaning corrected / clarified against the Arabic; speaker turns marked.']))));

const doc = new Document({creator: 'Ali Ahmad Al Hamadi', title: 'Transcript Table Analysis — FINAL',
  styles: {default: {document: {run: {font: 'Arial', size: 20}}}},
  sections: [{properties: {page: {size: {width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE}, margin: {top: 720, bottom: 720, left: 720, right: 720}}},
    headers: {default: new Header({children: [P('Transcript Table Analysis — FINAL · Older Syrian refugees in Jordan', {size: 16, color: '666666', align: AlignmentType.RIGHT})]})},
    footers: {default: new Footer({children: [new Paragraph({alignment: AlignmentType.CENTER, children: [new TextRun({children: ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES], size: 16, font: 'Arial'})]})]})},
    children: kids}]});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('written', b.length); });
