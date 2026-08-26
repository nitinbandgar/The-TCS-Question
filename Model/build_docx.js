const fs = require('fs');
const d = require('docx');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow,
  TableCell, WidthType, BorderStyle, ShadingType, TableOfContents, PageBreak,
  Header, Footer, PageNumber, LevelFormat, convertInchesToTwip, VerticalAlign
} = d;

const D = JSON.parse(fs.readFileSync('/root/tcs/data.json', 'utf8'));

const INK = '1A2226', ACCENT = '0D4351', MUTE = '5C6B70', RULE = 'C6D0D0';
const SEV = 'B23A34', OK = '10745F';
const SANS = 'Calibri', SERIF = 'Cambria';

const fmt = (n, dg = 0) => Number(n).toLocaleString('en-IN',
  { minimumFractionDigits: dg, maximumFractionDigits: dg });
const pct = (n, dg = 1) => (n * 100).toFixed(dg) + '%';
const cr = n => '₹' + fmt(Math.round(n));

/* ---------- helpers ---------- */
const P = (text, o = {}) => new Paragraph({
  spacing: { after: o.after ?? 140, before: o.before ?? 0, line: o.line ?? 276 },
  alignment: o.align,
  indent: o.indent,
  border: o.border,
  keepNext: o.keepNext,
  children: (Array.isArray(text) ? text : [text]).map(t =>
    typeof t === 'string'
      ? new TextRun({ text: t, font: o.font ?? SANS, size: o.size ?? 21,
                      color: o.color ?? INK, bold: o.bold, italics: o.italics })
      : t)
});
const R = (text, o = {}) => new TextRun({
  text, font: o.font ?? SANS, size: o.size ?? 21, color: o.color ?? INK,
  bold: o.bold, italics: o.italics
});

const H1 = (n, text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 420, after: 200 },
  children: [
    new TextRun({ text: n + '  ', font: SANS, size: 22, color: ACCENT, bold: true }),
    new TextRun({ text, font: SERIF, size: 30, color: INK, bold: true })
  ]
});
const H2 = text => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 120 },
  children: [new TextRun({ text, font: SANS, size: 22, color: ACCENT, bold: true })]
});
const EYE = text => new Paragraph({
  spacing: { before: 200, after: 60 },
  children: [new TextRun({ text: text.toUpperCase(), font: SANS, size: 15,
    color: MUTE, bold: true, characterSpacing: 30 })]
});
const RULE_P = () => new Paragraph({
  spacing: { before: 60, after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 1 } },
  children: [new TextRun({ text: '', size: 2 })]
});
const BULLET = (runs, lvl = 0) => new Paragraph({
  numbering: { reference: 'bul', level: lvl },
  spacing: { after: 90, line: 268 },
  children: (Array.isArray(runs) ? runs : [runs]).map(t =>
    typeof t === 'string' ? R(t) : t)
});

const noBorders = {
  top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
  left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
  insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE }
};

function table(cols, head, rows, opts = {}) {
  const total = cols.reduce((a, b) => a + b, 0);
  const cell = (txt, i, o = {}) => new TableCell({
    width: { size: cols[i], type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
    margins: { top: 70, bottom: 70, left: 90, right: 90 },
    verticalAlign: VerticalAlign.CENTER,
    borders: {
      top: { style: o.topRule ? BorderStyle.SINGLE : BorderStyle.NONE, size: 6, color: RULE },
      bottom: { style: BorderStyle.SINGLE, size: o.headRule ? 8 : 3, color: o.headRule ? ACCENT : RULE },
      left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }
    },
    children: [new Paragraph({
      spacing: { after: 0, line: 240 },
      alignment: (i === 0 || o.leftAll) ? AlignmentType.LEFT : AlignmentType.RIGHT,
      children: [new TextRun({ text: String(txt), font: SANS, size: o.size ?? 18,
        color: o.color ?? INK, bold: o.bold })]
    })]
  });
  const trs = [new TableRow({
    tableHeader: true,
    children: head.map((h, i) => cell(h, i, { bold: true, color: ACCENT, size: 16, headRule: true, leftAll: opts.leftAll }))
  })];
  rows.forEach(r => {
    const isTot = String(r[0]).trim().toLowerCase() === 'total';
    trs.push(new TableRow({
      children: r.map((v, i) => cell(v, i, {
        bold: isTot || opts.boldRows?.includes(r[0]),
        fill: isTot ? 'EFF3F3' : undefined, leftAll: opts.leftAll
      }))
    }));
  });
  return new Table({ columnWidths: cols, width: { size: total, type: WidthType.DXA }, rows: trs });
}

/* ---------- content ---------- */
const S = D.scenarios;
const B = D.buckets;
const totRisk = B.reduce((a, b) => a + b.rev * b.auto, 0);

const body = [];

/* --- title block --- */
body.push(new Paragraph({
  spacing: { before: 1400, after: 80 },
  children: [new TextRun({ text: 'STRATEGIC ASSESSMENT', font: SANS, size: 16,
    color: ACCENT, bold: true, characterSpacing: 60 })]
}));
body.push(new Paragraph({
  spacing: { after: 200 },
  children: [new TextRun({ text: 'The TCS Question', font: SERIF, size: 60, color: INK, bold: true })]
}));
body.push(new Paragraph({
  spacing: { after: 400 },
  children: [new TextRun({
    text: 'An honest reading of TCS’s position under AI disruption, the risks to each part of its book, ' +
          'a quantified five-year scenario model, and what each path means for Tata Sons’ capacity to fund the group.',
    font: SERIF, size: 24, color: MUTE, italics: true })]
}));
body.push(RULE_P());
const META = [['Prepared for', 'The Chairman, Tata Sons'],
 ['Date', '26 August 2026'],
 ['Base year', 'FY26 (year ended 31 March 2026)'],
 ['Horizon', 'FY27-FY31'],
 ['Basis', 'Public filings and disclosure. No internal management information used.'],
 ['Companion files', 'Live Excel scenario model, board presentation, interactive web version']];
const mcell = (t, w, mute) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  margins: { top: 34, bottom: 34, left: 0, right: 140 },
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
             left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
  children: [new Paragraph({ spacing: { after: 0, line: 250 },
    children: [new TextRun({ text: t, font: SANS, size: 18,
      color: mute ? MUTE : INK, bold: !!mute })] })]
});
body.push(new Table({
  columnWidths: [1900, 7100],
  width: { size: 9000, type: WidthType.DXA },
  borders: noBorders,
  rows: META.map(([k, v]) => new TableRow({
    children: [mcell(k, 1900, true), mcell(v, 7100, false)] }))
}));
body.push(new Paragraph({ children: [new PageBreak()] }));

/* --- 1. judgement --- */
body.push(H1('1', 'The judgement'));
body.push(P([R('This is not a cyclical trough that will resolve itself. TCS faces three separate problems at once and the market is pricing them as one. Only the first reverses on its own.', { bold: true })]));
body.push(P('FY26 revenue was $30.0bn, down 0.5% in dollars and 2.4% in constant currency — the first annual dollar decline in TCS’s listed history. Operating margin nonetheless rose to 25.0%, and that is the tell: the margin was bought with 23,460 fewer people, not with better pricing. When hiring resumed in Q1 FY27 (+9,279), margin fell 130 basis points to 24.0% in a single quarter.'));
body.push(P('The order book looks reassuring — $40.7bn of TCV in FY26 and $9.5bn in Q1 FY27. It should not. The same scope of work now signs for fewer dollars than it did two years ago. A healthy book-to-bill on a deflating rate card is not growth; it is running to stand still.'));
body.push(P('The competitive comparison is the number that should not be explained away. Infosys grew 3.1% in constant currency on a $20.2bn base in the same year TCS shrank 2.4%. That five-and-a-half point gap is not the cycle. Both faced the same clients, the same tariffs, the same technology. TCS is, for the first time in its modern history, the slowest-growing of the large Indian caps at scale.'));
body.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  indent: { left: 340 },
  border: { left: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 12 } },
  children: [new TextRun({
    text: 'The threat to TCS is not artificial intelligence. It is that TCS sells effort, and AI is collapsing the amount of effort each unit of client outcome requires.',
    font: SERIF, size: 26, color: INK })]
}));
body.push(P('That distinction determines the remedy. If the problem were AI capability, TCS would fix it by hiring — it already has more than 270,000 people with AI competencies and a $2.6bn annualised AI book growing 13.6% a quarter. If the problem is the pricing model, then every productivity gain TCS achieves is a discount it hands to the client at renewal, and capability alone makes the arithmetic worse, not better.'));

body.push(H2('The five numbers that matter'));
body.push(table([2600, 1500, 5100],
  ['Measure', 'Value', 'Why it matters'],
  [
    ['FY26 revenue growth, CC', '−2.4%', 'First annual dollar revenue decline in TCS’s listed history'],
    ['Revenue at automation risk', '42%', '$12.7bn of the FY26 book sits in work AI can substantially automate'],
    ['Tata Sons dependence on TCS', '87%', 'Share of the parent’s FY26 dividend income from a single holding'],
    ['Expected FY31 revenue', '$31.9bn', 'Probability-weighted — 1.2% a year compound from here'],
    ['Five-year dividend spread', '₹1.26 lakh cr', 'Gap between best and worst cumulative dividend to Tata Sons']
  ], { leftAll: false }));

/* --- 2. the record --- */
body.push(H1('2', 'What the record actually says'));
body.push(P('Growth has decelerated in every year since FY23 and turned negative in FY26. The rupee has been masking it: FY26 rupee revenue rose 4.6% while dollar revenue fell 0.5%, a roughly five-point currency tailwind.'));
body.push(table([1450, 1450, 1450, 1450, 1450, 1450],
  ['Fiscal year', 'Revenue $bn', 'Growth, CC', 'Op. margin', 'Headcount', 'Payout ₹cr'],
  D.history.map(h => [h.FY, (h.RevUSDm / 1000).toFixed(1), pct(h.CCg), pct(h.OPM),
    fmt(h.Headcount), fmt(h.PayoutINRcr)])
    .concat([['Q1 FY27', '7.6 (qtr)', '+3.2% YoY', '24.0%', '593,798', '—']])));
body.push(P(''));
body.push(P('Three figures from the same period belong together. Revenue per employee rose from $49,640 in FY25 to $51,354 in FY26 — AI leverage is real and it is showing up in the accounts. Headcount fell 23,460. And total shareholder payout was ₹39,571 crore against free cash flow of ₹42,983 crore: a 92% payout, consistent with a stated policy of returning 80–100% of free cash flow.'));
body.push(P([R('That last figure should concern this office most. There is no retained buffer. ', { bold: true }),
  R('TCS distributes essentially everything it generates. Any fall in free cash flow passes through to Tata Sons within a single dividend cycle, undamped.')]));
body.push(P('The market has already formed a view. The stock is down roughly 35% in calendar 2026 to a six-year low, the Nifty IT index is about 39% below its December 2024 peak, and TCS now trades on a P/E near 15x against a long-run average roughly twice that. Market capitalisation has fallen from about ₹14.8 lakh crore at the 2024 peak to ₹8.3 lakh crore. Five-year share price CAGR is negative 9%.'));

/* --- 3. three forces --- */
body.push(H1('3', 'Three forces, wrongly treated as one'));
body.push(P('Conflating these is how boards get the response wrong. Each has a different half-life and a different remedy.'));

body.push(H2('Force one · Cyclical — discretionary spend is frozen'));
body.push(P('Clients in banking, retail and communications have stretched decision cycles and deferred programmes. Tariffs, Middle East disruption and US federal contraction have all bitten. Accenture cut its FY26 growth guidance from 5% to 4%, saw bookings fall 2% year on year, and lost 18% of its value in a session.'));
body.push(P([R('Half-life: four to eight quarters. ', { bold: true }), R('This reverses on its own. It is also the smallest of the three.')]));

body.push(H2('Force two · Deflationary — the unit of sale is shrinking'));
body.push(P('AI reduces the hours required per unit of delivered outcome. Under time-and-materials and FTE-priced contracts, that reduction is refunded to the client at renewal. Industry reporting is consistent on this: the same scope now translates into a smaller dollar value than a year ago, even as order books grow.'));
body.push(P([R('Half-life: permanent. ', { bold: true }), R('This does not reverse. It is survivable only by changing what is sold and how it is priced.')]));

body.push(H2('Force three · Competitive — the work is moving elsewhere'));
body.push(P('India’s global capability centres are now a $98–100bn market across more than 2,100 centres employing some 2.3 million people — more than three times TCS’s entire revenue, growing at roughly 20% a year. Separately, OpenAI has put over $4bn behind placing forward-deployed engineers inside enterprises and acquired a consultancy to do it. Clients are building the capability TCS used to rent them.'));
body.push(P([R('Half-life: a decade. ', { bold: true }), R('This is a change in who owns the capability, not a change in demand.')]));

body.push(H2('One genuine tailwind'));
body.push(P('The $100,000 H-1B fee, which reads as a headwind, is largely already absorbed. TCS took on roughly 500 new H-1B holders this year and has shifted decisively to local hiring. Analysts put the incremental cost under the old delivery model at about $50m a year — material but not structural. Competitors more dependent on visa-linked onsite staffing carry more of that cost than TCS does.'));

/* --- 4. risk register --- */
body.push(H1('4', 'Risk register — where the exposure sits'));
body.push(P([R('A caveat first, and it is important. ', { bold: true }),
  R('TCS has not disclosed a service-line revenue split since FY21. The seven-bucket decomposition below is a reconstruction from the last disclosure, deal composition and peer structure. It is the single most challengeable input in this assessment and the first thing that should be replaced with internal management information.')]));
body.push(P('"Gross revenue at risk" is FY26 revenue in each bucket multiplied by the share of its work that current and near-term AI can perform with materially less human effort by FY31. It is a measure of exposure, not of expected loss — what is actually lost depends on how much of the productivity gain is passed to the client in price and how far cheaper delivery expands demand.'));
body.push(P(''));
body.push(table([2900, 1150, 1000, 1150, 1200, 1300],
  ['Exposure bucket', 'FY26 $m', 'Share', 'Automatable', 'At risk $m', 'Severe FY31 $m'],
  B.map((b, i) => [b.name, fmt(Math.round(b.rev)), pct(b.w, 1), pct(b.auto, 0),
    fmt(Math.round(b.rev * b.auto)), fmt(Math.round(S.severe.buckets[i][4]))])
   .concat([['Total', fmt(30017), '100.0%', pct(totRisk / 30017, 0), fmt(Math.round(totRisk)),
     fmt(Math.round(S.severe.buckets.reduce((a, r) => a + r[4], 0)))]])));
body.push(P(''));
body.push(P([R('$12.7bn — 42% of FY26 revenue — sits in work AI can substantially automate. ', { bold: true }),
  R('Whether that becomes lost revenue depends on two things TCS partly controls: how much of the productivity gain is passed through in price, and how far cheaper delivery expands the volume clients buy. That is the whole scenario question.')]));

body.push(H2('Bucket by bucket'));
B.forEach((b, i) => {
  body.push(P([R(b.name + ' — ', { bold: true }),
    R('$' + fmt(Math.round(b.rev)) + 'm, ' + pct(b.auto, 0) + ' automatable. ', { bold: true, color: b.auto >= 0.5 ? SEV : (b.auto <= 0.15 ? OK : INK) }),
    R(b.note)], { after: 110 }));
});
body.push(P([R('Note what is not exposed. ', { bold: true }),
  R('Engineering and R&D services and products and platforms together are about 14% of revenue and carry the lowest automation exposure. Products and platforms — BaNCS, ignio, iON, ADD — are the only part of the book whose value AI raises rather than deflates, because their revenue is licensed rather than billed by the hour. They are 5% of revenue. That number is a strategic failure a decade in the making, and it is now the most valuable optionality in the portfolio.')]));

/* --- 5. scenarios --- */
body.push(H1('5', 'Five-year scenarios, FY27–FY31'));
body.push(P('Each scenario is built bucket by bucket from two drivers: annual realisation deflation (AI productivity gain × automatable share × price pass-through to the client) and annual volume growth (underlying demand plus the elasticity response to a lower unit price). Margin, free cash flow conversion, payout and revenue-per-employee paths are set per scenario. Every assumption is visible and editable in the accompanying workbook.'));
body.push(P(''));
const scenRows = [
  ['Bull — AI Dividend', S.bull], ['Base — Managed Transition', S.base],
  ['Bear — Slow Bleed', S.bear], ['Severe — Structural Break', S.severe]
];
body.push(table([2500, 900, 1150, 1050, 1100, 1200, 800],
  ['Scenario', 'Prob.', 'FY31 rev $bn', 'Rev CAGR', 'FY31 margin', 'FY31 headcount', 'Cum. div ₹k cr'],
  scenRows.map(([n, s]) => [n, pct(s.prob, 0), (s.rev[4] / 1000).toFixed(1),
    (s.cagr >= 0 ? '+' : '−') + Math.abs(s.cagr * 100).toFixed(1) + '%',
    pct(s.opm[4]), fmt(s.hc[4]), fmt(Math.round(s.cumts / 1000))])));
body.push(P(''));
body.push(P([R('Probability-weighted, expected FY31 revenue is $31.9bn — a compound growth rate of 1.2% a year. ', { bold: true }),
  R('On an expected-value basis, TCS spends the next five years roughly where it is today. That, not the bear case, is the planning assumption this board should hold.')]));

body.push(H2('Base case in full — Managed Transition (40% probability)'));
body.push(P('Discretionary spend recovers gradually; AI deflation is real but broadly offset by volume expansion and new AI-native work; margin holds in a narrow band around 25%; the workforce stabilises rather than shrinks.'));
body.push(P(''));
const yrs = ['FY26A', 'FY27', 'FY28', 'FY29', 'FY30', 'FY31'];
const bs = S.base;
body.push(table([2500, 1150, 1150, 1150, 1150, 1150, 1150],
  ['Base case'].concat(yrs),
  [
    ['Revenue ($m)', fmt(D.fy26.rev)].concat(bs.rev.map(v => fmt(Math.round(v)))),
    ['Growth, CC', '−2.4%'].concat(bs.ccg.map(v => pct(v))),
    ['Operating margin', pct(D.fy26.opm)].concat(bs.opm.map(v => pct(v))),
    ['Net income ($m)', fmt(D.fy26.pat)].concat(bs.pat.map(v => fmt(Math.round(v)))),
    ['Free cash flow (₹cr)', fmt(D.fy26.fcf)].concat(bs.fcf.map(v => fmt(Math.round(v)))),
    ['Total dividend (₹cr)', fmt(D.fy26.div)].concat(bs.div.map(v => fmt(Math.round(v)))),
    ['Dividend per share (₹)', fmt(D.fy26.dps)].concat(bs.dps.map(v => fmt(v, 0))),
    ['To Tata Sons (₹cr)', fmt(D.fy26.ts)].concat(bs.ts.map(v => fmt(Math.round(v)))),
    ['Revenue per employee ($)', fmt(51354)].concat(bs.rpe.map(v => fmt(Math.round(v)))),
    ['Closing headcount', fmt(D.fy26.hc)].concat(bs.hc.map(v => fmt(Math.round(v))))
  ]));
body.push(P(''));
body.push(P([R('The employment figure in the severe case deserves separate mention. ', { bold: true }),
  R('Headcount falls from 584,519 to 334,575 — a reduction of roughly 250,000 jobs over five years. That is not primarily a financial event for this group; it is a political and social one, and it would be managed as such. Which in turn constrains how fast the cost base can actually be taken out. The model assumes an orderly adjustment. A disorderly one is worse for margin, not better.')]));

/* --- 6. Tata Sons --- */
body.push(H1('6', 'What this means for Tata Sons'));
body.push(P([R('The concentration is the risk, not TCS’s absolute performance. ', { bold: true }),
  R('TCS could perform respectably and still leave this office short, because the parent has committed to a capital programme sized against a dividend that has already begun to fall.')]));
body.push(P('In FY26 Tata Sons received ₹28,291 crore from TCS, down 12% from ₹32,184 crore the year before. Total dividend income across every holding was ₹32,528 crore — TCS was 87% of it. The rest of a listed portfolio worth ₹11.68 trillion contributed ₹4,237 crore. Standalone profit was ₹31,961 crore, up 21.8%.'));
body.push(P('Against that stand the commitments: the Dholera fabrication plant at ₹91,000 crore and the Assam assembly and test facility at ₹27,000 crore; Air India’s 600-aircraft order and a business into which ₹22,618 crore has already gone; Agratas’s £4bn Somerset plant; Tata Digital, carrying ₹22,903 crore of equity against a ₹4,974 crore FY26 loss. The four newest businesses lost roughly ₹30,000 crore between them last year. Parent net cash at the FY26 close was ₹21,841 crore, with no borrowings.'));
body.push(P(''));
body.push(H2('Parent cash bridge by scenario, FY27–FY31'));
body.push(P('Capital deployment is held constant across all four scenarios by design: the commitments do not shrink because TCS slows. Assumes non-TCS dividend income grows 9% a year, parent costs 6%, and equity deployment rises from ₹20,000 crore in FY27 to ₹28,000 crore in FY31 (FY26 actual: ₹15,956 crore).'));
body.push(P(''));
body.push(table([2500, 1700, 1700, 1700, 1700],
  ['Parent position', 'Bull', 'Base', 'Bear', 'Severe'],
  [
    ['Cumulative TCS dividend (₹cr)', fmt(S.bull.cumts), fmt(S.base.cumts), fmt(S.bear.cumts), fmt(S.severe.cumts)],
    ['FY31 TCS dividend (₹cr)', fmt(S.bull.ts[4]), fmt(S.base.ts[4]), fmt(S.bear.ts[4]), fmt(S.severe.ts[4])],
    ['Cumulative surplus / (deficit)', fmt(S.bull.bridge.cumsur), fmt(S.base.bridge.cumsur),
      '(' + fmt(Math.abs(S.bear.bridge.cumsur)) + ')', '(' + fmt(Math.abs(S.severe.bridge.cumsur)) + ')'],
    ['FY31 net cash (₹cr)', fmt(S.bull.bridge.cum[4]), fmt(S.base.bridge.cum[4]),
      fmt(S.bear.bridge.cum[4]), '(' + fmt(Math.abs(S.severe.bridge.cum[4])) + ')'],
    ['Funds the capital call?', 'Yes', 'Yes', 'No, from FY28', 'No, from FY28']
  ]));
body.push(P(''));
body.push(P([R('The break-even is ₹30,827 crore. ', { bold: true }),
  R('That is what TCS must pay Tata Sons in FY31 to cover the parent’s own costs and its committed capital call, after dividend income from every other listed holding. The base case clears it at ₹40,074 crore. The bear case reaches ₹23,583 crore and the severe case ₹12,619 crore — neither ever gets there.')]));
body.push(P('In the bear case Tata Sons runs a cumulative deficit of ₹19,163 crore across FY27–FY31 and its net cash is all but exhausted. In the severe case the deficit is ₹56,517 crore and the parent is in net borrowing from FY30. The gap between the best and worst cumulative dividend from TCS over five years is ₹1,25,649 crore — roughly the whole cost of the Dholera fab plus the Assam facility.'));
body.push(P([R('There is a regulatory dimension that should not wait. ', { bold: true }),
  R('Tata Sons remains on the RBI’s upper-layer NBFC list and the application to deregister as a core investment company is still unresolved. If the shock absorber for a falling dividend is access to debt at scale, that capacity needs to exist and be priced before it is needed, not during.')]));
body.push(new Paragraph({
  spacing: { before: 240, after: 200 },
  indent: { left: 340 },
  border: { left: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 12 } },
  children: [new TextRun({
    text: 'We are not exposed to whether TCS grows. We are exposed to having assumed it would.',
    font: SERIF, size: 26, color: INK })]
}));

body.push(H2('Sensitivity — growth versus margin'));
body.push(P('Two levers determine almost the whole answer. Over the range actually available to management, margin cannot rescue a growth failure. Moving operating margin from 25% to 27% — an outcome nobody has achieved in a decade — is worth roughly ₹2,400 crore of FY31 dividend to Tata Sons. Getting from flat revenue to 4% growth is worth about ₹6,900 crore, and unlike margin it compounds rather than resetting each year.'));
body.push(P('Worse, the two are not independent. In every downside path modelled here, margin falls because growth fails, not alongside it: idle capacity, pricing concessions to hold volume, and the cost of a workforce that cannot be reshaped as fast as revenue moves. Cost discipline buys time. It does not buy the outcome.'));

/* --- 7. recommendations --- */
body.push(H1('7', 'Recommendations'));
body.push(P('Three concern the group’s own balance sheet and should happen regardless of what TCS does. Five concern TCS. One concerns how this board governs the uncertainty.'));

const RECS = [
  ['A1', 'Break the single-source dependency — target below 70% within three years',
    'Eighty-seven per cent of parent dividend income from one holding is not a concentration we would accept in any portfolio we advised. Two levers exist and both are available now. First, the rest of a ₹11.68 trillion listed portfolio returned ₹4,237 crore last year; the payout policies of Tata Motors, Tata Steel, Titan, Tata Power, Tata Consumer and Indian Hotels have never been set with the parent’s funding needs in view. Second, bring third-party capital in at the asset level, as has already been done with TPG in HyperVault and Tata Motors EV, and BlackRock and Mubadala in Tata Power Renewables.',
    'Set an explicit target for TCS as a share of parent dividend income, and report against it to this board each year.'],
  ['A2', 'Sequence the capital call rather than running it in parallel',
    'Dholera, Assam, the Air India fleet, Agratas Somerset and Tata Digital are all being funded simultaneously against a dividend that fell 12% last year. Rank them on two axes: how irreversible the commitment already is, and whether a partner or a government can carry part of the load. The bear case gives roughly until FY28 before this becomes forced rather than chosen.',
    'A ranked, dated sequencing plan with a named decision point for each programme.'],
  ['A3', 'Settle the RBI question before we need the answer',
    'Tata Sons remains on the upper-layer NBFC list with the core-investment-company deregistration plea unresolved. If the shock absorber for a falling dividend is debt capacity at the parent, that capacity has to exist and be priced before the dividend falls, not during. This is a live regulatory question, not a background one.',
    'A clear read on what the parent can borrow, at what cost, under each outcome of the deregistration decision.'],
  ['B1', 'Change what this board asks TCS for',
    'Revenue growth and headcount are the wrong scoreboard in a deflating market. If AI-led delivery works, revenue per employee rises faster than revenue falls — and that is a good outcome, not a bad one, even with flat revenue. The board should be tracking revenue per employee, the share of revenue on non-effort-linked pricing, products and platforms as a percentage of total, and gross margin on the AI book. None of those four is currently a headline metric.',
    'Replace the quarterly growth-and-headcount review with those four measures.'],
  ['B2', 'Force the pricing transition instead of defending the rate card',
    'This is the central recommendation. TCS’s exposure is not AI capability — it has more AI-competent staff than almost anyone. The exposure is billing for effort while effort collapses. Whoever moves to outcome-based and platform pricing first keeps the productivity gain; whoever moves last refunds it. Every quarter spent protecting time-and-materials rates is a quarter of margin handed to clients in the next renewal round.',
    'What proportion of the FY27 order book is priced on outcomes rather than effort, and what is the FY29 target?'],
  ['B3', 'Products and platforms should be 15% of revenue, not 5%',
    'BaNCS, ignio, iON and ADD are the only assets in the portfolio whose value AI raises rather than deflates, because they are licensed rather than billed by the hour. They have been under-invested for a decade because services cash was easier to earn. This is now the most valuable optionality TCS owns, and it is a build-or-buy decision that will not make itself.',
    'A five-year plan to triple IP-led revenue share, with the acquisitions it would require costed.'],
  ['B4', 'Interrogate HyperVault properly, and ring-fence it',
    'Committing up to ₹18,000 crore of equity to AI data centres may well be right. But it converts TCS from an asset-light dividend machine into a capital-intensive infrastructure story at precisely the moment this parent depends most on the dividend — and free cash flow conversion has already fallen to 81% of net income. Capacity of this kind commoditises. The TPG structure is the right instinct; the question is whether it goes far enough.',
    'Hurdle rate, capital schedule, the degree of ring-fencing from the dividend stream, and the exit.'],
  ['B5', 'Fix the workforce model rather than oscillating',
    'Headcount fell 23,460 across FY26 and then rose 9,279 in a single quarter with margin dropping 130 basis points. That is not a strategy; it is a company that has not yet settled what shape a delivery organisation should be when AI does the first draft. The pyramid is being replaced by something with a thinner base and a thicker middle, and the cost of getting there is real. Better to say so and plan it than to discover it a quarter at a time.',
    'A three-year workforce shape plan, with the reskilling cost stated explicitly rather than absorbed into margin.'],
  ['C1', 'Decide the trigger points now, while conditions are calm',
    'The worst version of this is a board that debates the group’s capital plan in the quarter the dividend disappoints. Define in advance what would force a change: two consecutive quarters of negative constant-currency growth, or a dividend to Tata Sons below ₹25,000 crore, or free cash flow conversion below 70%. Agree the response to each while nobody is under pressure.',
    'A one-page trigger schedule, agreed at this board and reviewed each quarter.']
];
RECS.forEach(([n, h, p, ask]) => {
  body.push(new Paragraph({
    spacing: { before: 280, after: 100 },
    children: [
      new TextRun({ text: n + ' ', font: SANS, size: 20, color: ACCENT, bold: true }),
      new TextRun({ text: h, font: SANS, size: 22, color: INK, bold: true })]
  }));
  body.push(P(p, { after: 110 }));
  body.push(new Paragraph({
    spacing: { after: 60 },
    indent: { left: 340 },
    border: { left: { style: BorderStyle.SINGLE, size: 10, color: RULE, space: 10 } },
    children: [
      new TextRun({ text: 'What I would ask for: ', font: SANS, size: 19, color: ACCENT, bold: true }),
      new TextRun({ text: ask, font: SANS, size: 19, color: INK })]
  }));
});

/* --- 8. monitoring --- */
body.push(H1('8', 'What would change this view'));
body.push(H2('Leading indicators, in order of signal value'));
[['Revenue per employee.', ' The cleanest single test of whether AI-led delivery is working. If it rises faster than revenue falls, the transition is being managed. FY26: +3.5%.'],
 ['TCV-to-revenue conversion by deal vintage.', ' A growing order book converting to less revenue is the deflation thesis proving itself. This is internal data and this board should be seeing it.'],
 ['Renewal pricing on the top 50 accounts.', ' Where the productivity refund actually happens. Watch rate-card movement, not headline wins.'],
 ['BFSI constant-currency growth.', ' 32% of revenue and the first vertical to move in any recovery. FY26: +1.0%.'],
 ['Share of revenue on non-effort-linked pricing.', ' Not currently disclosed. It should be, at least to this board.']
].forEach(([b, t]) => body.push(BULLET([R(b, { bold: true }), R(t)])));

body.push(H2('What internal data would sharpen this materially'));
[['Actual service-line revenue and gross margin.', ' The seven-bucket split here is a reconstruction. Real management information would change the risk quantification, in either direction.'],
 ['Pricing model mix.', ' Time-and-materials versus fixed-price versus outcome-linked, and the trend. This determines how much deflation is contractually unavoidable.'],
 ['The HyperVault business case.', ' Hurdle rate, capital schedule, ring-fencing, exit.'],
 ['Tata Sons’ committed versus discretionary capital call by year.', ' The model treats it as fixed. How much of it genuinely is?'],
 ['Client concentration and renewal calendar.', ' Which of the 66 hundred-million-dollar accounts come up for renewal, and when.']
].forEach(([b, t]) => body.push(BULLET([R(b, { bold: true }), R(t)])));

body.push(H2('On the confidence of these numbers'));
body.push(P('The financial history, the Tata Sons dividend figures and the capital commitments are from public filings and are reliable. The seven-bucket revenue decomposition, the per-bucket deflation and volume drivers, the scenario probabilities and the forward capital-call profile are judgements. They are documented and editable in the accompanying workbook, and they should be argued with. A model whose assumptions cannot be challenged is not an analysis; it is an opinion with decimal places.'));

body.push(H1('9', 'Sources'));
body.push(P('TCS Q4 FY26 results press release and fact sheet; TCS Q1 FY27 results press release; TCS Q4 FY25 and Q4 FY24 press releases (tcs.com/investor-relations). Infosys FY26 results and FY27 guidance. Accenture Q3 FY2026 results. Tata Sons FY26 accounts as reported in Business Today (28 July 2026) and Business Standard (17–18 August 2026). Screener.in and stockanalysis.com for market data as at 24 August 2026. Business of GCC industry data for India GCC market size. TCS–TPG HyperVault announcement, November 2025. RBI upper-layer NBFC list, August 2026.',
  { size: 18, color: MUTE }));

/* ---------- document ---------- */
const doc = new Document({
  creator: 'Strategic Assessment',
  title: 'The TCS Question',
  description: 'Strategic assessment prepared for the Chairman, Tata Sons',
  numbering: {
    config: [{
      reference: 'bul',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 340, hanging: 200 } } } }]
    }]
  },
  styles: {
    default: { document: { run: { font: SANS, size: 21, color: INK } } }
  },
  sections: [{
    properties: {
      page: { margin: { top: 1300, right: 1250, bottom: 1300, left: 1250 } }
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 200 },
          children: [
            new TextRun({ text: 'The TCS Question  ·  Prepared for the Chairman, Tata Sons  ·  ',
              font: SANS, size: 15, color: MUTE }),
            new TextRun({ children: [PageNumber.CURRENT], font: SANS, size: 15, color: MUTE })]
        })]
      })
    },
    children: body
  }]
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('/root/tcs/TCS_Chairman_Assessment.docx', b);
  console.log('written', (b.length / 1024).toFixed(1), 'KB');
});
