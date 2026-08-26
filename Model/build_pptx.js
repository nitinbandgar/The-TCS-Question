const fs = require('fs');
const pptxgen = require('pptxgenjs');
const D = JSON.parse(fs.readFileSync('/root/tcs/data.json', 'utf8'));
const S = D.scenarios, B = D.buckets;

const DARK = '0E2028', INK = '1A2226', MUTE = '5C6B70', PANEL = 'EFF3F4',
      RULE = 'D3DCDD', ACCENT = '0D4351', ACC_LT = '78BACB', PAPER = 'FFFFFF';
const BULL = '159178', BASE = '3A72C0', BEAR = 'C8801A', SEV = 'C23B36';
const SERIF = 'Cambria', SANS = 'Calibri';
const W = 13.333, H = 7.5, M = 0.7;

const fmt = (n, d = 0) => Number(n).toLocaleString('en-IN',
  { minimumFractionDigits: d, maximumFractionDigits: d });
const pct = (n, d = 1) => (n * 100).toFixed(d) + '%';

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'Strategic Assessment';
pres.title = 'The TCS Question';

let n = 0;
function slide(dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? DARK : PAPER };
  n++;
  return s;
}
function chrome(s, section, dark) {
  s.addText(section.toUpperCase(), {
    x: M, y: H - 0.55, w: 8, h: 0.3, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 9, charSpacing: 2,
    color: dark ? '92AEB7' : MUTE, align: 'left', valign: 'middle'
  });
  s.addText(String(n).padStart(2, '0'), {
    x: W - M - 1, y: H - 0.55, w: 1, h: 0.3, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 9, charSpacing: 2,
    color: dark ? '92AEB7' : MUTE, align: 'right', valign: 'middle'
  });
}
function title(s, text, sub, dark) {
  s.addText(text, {
    x: M, y: 0.52, w: W - 2 * M, h: 0.85, isTextBox: true, margin: 0,
    fontFace: SERIF, fontSize: 32, bold: true,
    color: dark ? PAPER : INK, valign: 'middle'
  });
  if (sub) s.addText(sub, {
    x: M, y: 1.36, w: W - 2 * M - 1.2, h: 0.44, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 13, color: dark ? ACC_LT : MUTE, valign: 'top'
  });
}
function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.rect, {
    x, y, w, h, fill: { color: fill || PANEL },
    line: { color: fill ? fill : RULE, width: 0.75 }
  });
}

/* ================= 1 · TITLE ================= */
{
  const s = slide(true);
  s.addText('STRATEGIC ASSESSMENT', {
    x: M, y: 2.0, w: 8, h: 0.3, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 11, bold: true, charSpacing: 3, color: ACC_LT
  });
  s.addText('The TCS Question', {
    x: M, y: 2.42, w: 9.5, h: 1.25, isTextBox: true, margin: 0,
    fontFace: SERIF, fontSize: 54, bold: true, color: PAPER
  });
  s.addText('Eighty-seven per cent of Tata Sons’ dividend income comes from one company, and that company has just reported the first revenue decline of its listed life.', {
    x: M, y: 3.78, w: 8.4, h: 0.95, isTextBox: true, margin: 0,
    fontFace: SERIF, fontSize: 16, italic: true, color: 'B9CBD1', lineSpacing: 24
  });
  s.addShape(pres.ShapeType.line, {
    x: M, y: 5.05, w: W - 2 * M, h: 0, line: { color: '2C4550', width: 1 }
  });
  const meta = [['Prepared for', 'The Chairman, Tata Sons'], ['Date', '26 August 2026'],
    ['Base year', 'FY26'], ['Horizon', 'FY27–FY31']];
  meta.forEach(([k, v], i) => {
    const x = M + i * 3.05;
    s.addText(k.toUpperCase(), { x, y: 5.25, w: 2.9, h: 0.24, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 9, charSpacing: 2, color: '5D7C86' });
    s.addText(v, { x, y: 5.5, w: 2.9, h: 0.3, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 13, color: PAPER });
  });
  s.addNotes('Public filings only. Companion files: live Excel scenario model and interactive web version.');
}

/* ================= 2 · THE VERDICT ================= */
{
  const s = slide(false);
  title(s, 'The judgement in five numbers', 'FY26 actuals and model output. Full basis in the accompanying workbook.');
  const items = [
    ['FY26 revenue growth', '−2.4%', 'Constant currency. First annual dollar decline in TCS’s listed history.', SEV],
    ['Revenue at risk', '42%', '$12.7bn of the FY26 book sits in work AI can substantially automate.', SEV],
    ['Tata Sons dependence', '87%', 'Share of the parent’s FY26 dividend income from a single holding.', INK],
    ['Expected FY31 revenue', '$31.9bn', 'Probability-weighted — 1.2% a year compound from here.', INK],
    ['Dividend spread', '₹1.26 lakh cr', 'Gap between best and worst cumulative dividend to Tata Sons.', ACCENT]
  ];
  const cw = (W - 2 * M - 4 * 0.22) / 5;
  items.forEach(([lab, fig, sub, col], i) => {
    const x = M + i * (cw + 0.22);
    card(s, x, 2.15, cw, 2.75);
    s.addText(lab.toUpperCase(), { x: x + 0.22, y: 2.4, w: cw - 0.4, h: 0.3, isTextBox: true,
      margin: 0, fontFace: SANS, fontSize: 9, bold: true, charSpacing: 1.2, color: MUTE });
    s.addText(fig, { x: x + 0.22, y: 2.78, w: cw - 0.36, h: 0.9, isTextBox: true, margin: 0,
      fontFace: SERIF, fontSize: fig.length > 9 ? 19 : 34, bold: true, color: col, valign: 'middle' });
    s.addText(sub, { x: x + 0.22, y: 3.72, w: cw - 0.44, h: 1.1, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, color: INK, lineSpacing: 15, valign: 'top' });
  });
  s.addText('The market has already formed a view: the stock is down roughly 35% in 2026 to a six-year low, on a P/E near 15x against a long-run average roughly twice that.',
    { x: M, y: 5.35, w: W - 2 * M, h: 0.4, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12, italic: true, color: MUTE });
  chrome(s, 'The judgement');
  s.addNotes('The five numbers to hold. Note the last one: the spread between best and worst case is roughly the whole cost of Dholera plus Assam.');
}

/* ================= 3 · NOT A TROUGH ================= */
{
  const s = slide(false);
  title(s, 'Not a trough. A repricing.', 'Three separate problems, priced by the market as one. Only the first reverses on its own.');
  card(s, M, 2.2, 6.2, 3.9);
  s.addText([
    { text: 'The threat is not artificial intelligence.\n', options: { fontFace: SERIF, fontSize: 20, bold: true, color: INK, breakLine: true } },
    { text: 'It is that TCS sells effort, and AI is collapsing the effort each unit of client outcome requires.', options: { fontFace: SERIF, fontSize: 19, color: ACCENT } }
  ], { x: M + 0.35, y: 2.5, w: 5.5, h: 1.75, isTextBox: true, margin: 0, lineSpacing: 26, valign: 'top' });
  s.addText([
    { text: 'If the problem were AI capability, TCS would fix it by hiring. It already has 270,000+ people with AI competencies and a $2.6bn annualised AI book growing 13.6% a quarter.', options: { breakLine: true, paraSpaceAfter: 10 } },
    { text: 'If the problem is the pricing model, then every productivity gain TCS achieves is a discount handed to the client at renewal — and capability alone makes the arithmetic worse, not better.', options: {} }
  ], { x: M + 0.35, y: 4.42, w: 5.5, h: 1.5, isTextBox: true, margin: 0,
       fontFace: SANS, fontSize: 12, color: INK, lineSpacing: 16 });

  const right = [
    ['Margin was bought with people, not price',
     'FY26 operating margin rose to 25.0% while headcount fell 23,460. When hiring resumed in Q1 FY27 (+9,279), margin fell 130bps to 24.0% in one quarter.'],
    ['The order book is deflating',
     '$40.7bn of FY26 TCV looks healthy. But the same scope now signs for fewer dollars than two years ago. Book-to-bill on a falling rate card is running to stand still.'],
    ['The peer gap is not the cycle',
     'Infosys grew 3.1% CC on $20.2bn in the year TCS shrank 2.4%. Same clients, same tariffs, same technology. TCS is now the slowest-growing large Indian cap at scale.']
  ];
  right.forEach(([h, p], i) => {
    const y = 2.2 + i * 1.35;
    s.addText(h, { x: 7.35, y, w: W - M - 7.35, h: 0.32, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 13.5, bold: true, color: ACCENT });
    s.addText(p, { x: 7.35, y: y + 0.36, w: W - M - 7.35, h: 0.85, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11.5, color: INK, lineSpacing: 15 });
  });
  chrome(s, 'The judgement');
  s.addNotes('The pricing-model framing is the single most important idea in this pack. Everything downstream follows from it.');
}

/* ================= 4 · THE RECORD ================= */
{
  const s = slide(false);
  title(s, 'Five years of deceleration', 'Constant-currency growth by fiscal year. The rupee has been masking it — FY26 rupee revenue rose 4.6% while dollar revenue fell 0.5%.');
  s.addChart(pres.ChartType.bar, [{
    name: 'Constant-currency growth',
    labels: D.history.map(h => h.FY).concat(['Q1 FY27']),
    values: D.history.map(h => h.CCg).concat([0.032])
  }], {
    x: M, y: 2.15, w: 7.5, h: 3.9,
    barDir: 'col', barGapWidthPct: 60,
    chartColors: [BASE, BASE, BEAR, BEAR, SEV, BULL],
    varyColors: true,
    showLegend: false, showTitle: false,
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0%',
    dataLabelFontSize: 11, dataLabelColor: INK, dataLabelFontFace: SANS,
    catAxisLabelColor: INK, catAxisLabelFontSize: 11, catAxisLabelFontFace: SANS,
    valAxisLabelColor: MUTE, valAxisLabelFontSize: 10, valAxisLabelFormatCode: '0%',
    valAxisMinVal: -0.05, valAxisMaxVal: 0.20,
    valGridLine: { color: RULE, size: 0.75 }, catGridLine: { style: 'none' },
    catAxisLineShow: false, valAxisLineShow: false
  });
  const facts = [
    ['Revenue per employee', '$51,354', 'up from $49,640 in FY25 — AI leverage is real and showing'],
    ['Headcount change, FY26', '−23,460', 'the source of the margin, not pricing'],
    ['Payout as % of free cash flow', '92%', '₹39,571 cr paid on ₹42,983 cr generated']
  ];
  facts.forEach(([l, v, sub], i) => {
    const y = 2.3 + i * 1.28;
    card(s, 8.5, y, W - M - 8.5, 1.08);
    s.addText(l.toUpperCase(), { x: 8.72, y: y + 0.13, w: 3.6, h: 0.24, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 8.5, bold: true, charSpacing: 1.5, color: MUTE });
    s.addText(v, { x: 8.72, y: y + 0.36, w: 3.6, h: 0.36, isTextBox: true, margin: 0,
      fontFace: SERIF, fontSize: 21, bold: true, color: i === 2 ? SEV : INK });
    s.addText(sub, { x: 8.72, y: y + 0.72, w: 3.7, h: 0.3, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 9.5, color: MUTE });
  });
  s.addText('There is no retained buffer. TCS distributes essentially everything it generates, so any fall in free cash flow passes through to Tata Sons within one dividend cycle.',
    { x: 8.5, y: 6.15, w: W - M - 8.5, h: 0.7, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, bold: true, color: ACCENT, lineSpacing: 15 });
  chrome(s, 'The record');
  s.addNotes('FY26 is the first negative constant-currency year. Q1 FY27 recovery is flattered by a low base and a 7.6% sequential jump in India.');
}

/* ================= 5 · THREE FORCES ================= */
{
  const s = slide(false);
  title(s, 'Three forces, wrongly treated as one', 'Each has a different half-life and a different remedy. Conflating them is how boards get the response wrong.');
  const forces = [
    [BEAR, 'Force one · Cyclical', 'Discretionary spend is frozen',
     'Banking, retail and communications clients have stretched decision cycles and deferred programmes. Tariffs, Middle East disruption and US federal contraction have all bitten. Accenture cut FY26 guidance from 5% to 4%, saw bookings fall 2%, and lost 18% in a session.',
     'Half-life: 4–8 quarters. Reverses on its own — and it is the smallest of the three.'],
    [SEV, 'Force two · Deflationary', 'The unit of sale is shrinking',
     'AI reduces the hours required per unit of delivered outcome. Under time-and-materials and FTE pricing, that reduction is refunded to the client at renewal. The same scope now translates into a smaller dollar value than a year ago, even as order books grow.',
     'Half-life: permanent. Survivable only by changing what is sold and how it is priced.'],
    [BASE, 'Force three · Competitive', 'The work is moving elsewhere',
     'India’s global capability centres are now a $98–100bn market across 2,100+ centres and some 2.3 million people — more than three times TCS’s entire revenue, growing ~20% a year. OpenAI has put $4bn+ behind placing engineers inside enterprises.',
     'Half-life: a decade. Clients are building the capability TCS used to rent them.']
  ];
  const cw = (W - 2 * M - 2 * 0.3) / 3;
  forces.forEach(([col, tag, h, p, v], i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 2.2, cw, 3.55);
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.28, y: 2.48, w: 0.18, h: 0.18, fill: { color: col } });
    s.addText(tag.toUpperCase(), { x: x + 0.56, y: 2.44, w: cw - 0.85, h: 0.26, isTextBox: true,
      margin: 0, fontFace: SANS, fontSize: 8.5, bold: true, charSpacing: 1.5, color: col, valign: 'middle' });
    s.addText(h, { x: x + 0.28, y: 2.78, w: cw - 0.56, h: 0.62, isTextBox: true, margin: 0,
      fontFace: SERIF, fontSize: 17, bold: true, color: INK, valign: 'top' });
    s.addText(p, { x: x + 0.28, y: 3.44, w: cw - 0.56, h: 1.55, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, color: INK, lineSpacing: 15, valign: 'top' });
    s.addText(v, { x: x + 0.28, y: 5.02, w: cw - 0.56, h: 0.62, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 10.5, bold: true, color: col, lineSpacing: 14, valign: 'top' });
  });
  s.addText('One genuine tailwind: the $100,000 H-1B fee is largely already absorbed. TCS took on roughly 500 new H-1B holders this year and has shifted to local hiring; the incremental cost under the old model is put at about $50m a year. Competitors more dependent on visa-linked onsite staffing carry more of it.',
    { x: M, y: 5.95, w: W - 2 * M, h: 0.6, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, color: MUTE, lineSpacing: 15 });
  chrome(s, 'Decomposition');
  s.addNotes('If the board takes one thing from this slide: only force one fixes itself. Forces two and three require decisions.');
}

/* ================= 6 · RISK REGISTER ================= */
{
  const s = slide(false);
  title(s, 'Where the exposure actually sits', '$12.7bn — 42% of FY26 revenue — is in work AI can substantially automate. Exposure, not expected loss.');
  const SHORT = ['App development & maintenance', 'Package application services',
    'Cloud & infrastructure managed', 'Business process services', 'Engineering & R&D',
    'Consulting, data, AI & cyber', 'Products & platforms'];
  s.addChart(pres.ChartType.bar, [
    { name: 'At risk to automation', labels: SHORT, values: B.map(b => Math.round(b.rev * b.auto)) },
    { name: 'Remainder of the bucket', labels: SHORT, values: B.map(b => Math.round(b.rev * (1 - b.auto))) }
  ], {
    x: M, y: 2.15, w: 8.3, h: 4.0,
    barDir: 'bar', barGrouping: 'stacked', barGapWidthPct: 45,
    chartColors: [SEV, 'CFD9DA'],
    showLegend: true, legendPos: 'b', legendFontSize: 10, legendFontFace: SANS, legendColor: INK,
    showTitle: false,
    showValue: false,
    catAxisLabelColor: INK, catAxisLabelFontSize: 10.5, catAxisLabelFontFace: SANS,
    valAxisLabelColor: MUTE, valAxisLabelFontSize: 9.5, valAxisLabelFormatCode: '#,##0',
    valAxisMinVal: 0, valAxisMaxVal: 10000, valAxisMajorUnit: 2500,
    valAxisTitle: 'FY26 revenue ($m)', showValAxisTitle: true,
    valAxisTitleFontSize: 10, valAxisTitleColor: MUTE,
    valGridLine: { color: RULE, size: 0.75 }, catGridLine: { style: 'none' },
    catAxisLineShow: false, valAxisLineShow: false
  });
  const notes = [
    [SEV, 'Most exposed', 'Business process services (65% automatable) and application development and maintenance (55%). Together $12.6bn of revenue and $7.2bn of exposure.'],
    [BULL, 'Least exposed', 'Engineering & R&D and products and platforms — about 14% of revenue, lowest automation risk, and the only revenue AI makes more valuable rather than less.'],
    [ACCENT, 'The strategic point', 'Products and platforms — BaNCS, ignio, iON, ADD — are licensed, not billed by the hour. They are 5% of revenue. That is a decade-old under-investment and now the most valuable optionality TCS owns.']
  ];
  notes.forEach(([col, h, p], i) => {
    const y = 2.3 + i * 1.32;
    s.addText(h.toUpperCase(), { x: 9.25, y, w: W - M - 9.25, h: 0.25, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 9, bold: true, charSpacing: 1.5, color: col });
    s.addText(p, { x: 9.25, y: y + 0.28, w: W - M - 9.25, h: 0.95, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, color: INK, lineSpacing: 15 });
  });
  s.addText('Caveat: TCS has not disclosed a service-line split since FY21. This decomposition is a reconstruction and is the first thing to replace with internal MIS.',
    { x: 9.25, y: 6.05, w: W - M - 9.25, h: 0.6, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 9.5, italic: true, color: '46545A', lineSpacing: 13 });
  chrome(s, 'Risk register');
  s.addNotes('Exposure is not loss. What converts exposure into loss is price pass-through; what offsets it is volume expansion. That is the scenario question.');
}

/* ================= 7 · SCENARIOS ================= */
{
  const s = slide(false);
  title(s, 'Four futures, FY27–FY31', 'Built bucket by bucket from annual realisation deflation and annual volume growth. Every assumption is editable in the workbook.');
  const rows = [
    [BULL, 'Bull', 'AI Dividend', S.bull], [BASE, 'Base', 'Managed Transition', S.base],
    [BEAR, 'Bear', 'Slow Bleed', S.bear], [SEV, 'Severe', 'Structural Break', S.severe]
  ];
  const cw = (W - 2 * M - 3 * 0.25) / 4;
  rows.forEach(([col, nm, full, sc], i) => {
    const x = M + i * (cw + 0.25);
    card(s, x, 2.15, cw, 3.85);
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.25, y: 2.42, w: 0.16, h: 0.16, fill: { color: col } });
    s.addText(nm, { x: x + 0.5, y: 2.36, w: cw - 0.8, h: 0.3, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 13, bold: true, color: col, valign: 'middle' });
    s.addText(full + '  ·  ' + pct(sc.prob, 0) + ' probability',
      { x: x + 0.25, y: 2.66, w: cw - 0.5, h: 0.26, isTextBox: true, margin: 0,
        fontFace: SANS, fontSize: 9.5, color: MUTE });
    s.addText('$' + (sc.rev[4] / 1000).toFixed(1) + 'bn', {
      x: x + 0.25, y: 2.98, w: cw - 0.5, h: 0.55, isTextBox: true, margin: 0,
      fontFace: SERIF, fontSize: 30, bold: true, color: INK, valign: 'middle' });
    s.addText('FY31 revenue · ' + (sc.cagr >= 0 ? '+' : '−') +
      Math.abs(sc.cagr * 100).toFixed(1) + '% a year', {
      x: x + 0.25, y: 3.55, w: cw - 0.5, h: 0.26, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 9.5, color: MUTE });
    const met = [['FY31 operating margin', pct(sc.opm[4])],
      ['FY31 headcount', fmt(sc.hc[4])],
      ['Cumulative dividend to Tata Sons', '₹' + fmt(Math.round(sc.cumts / 1000)) + 'k cr'],
      ['Parent 5-yr surplus / (deficit)', (sc.bridge.cumsur < 0 ? '(₹' : '₹') +
        fmt(Math.abs(Math.round(sc.bridge.cumsur))) + (sc.bridge.cumsur < 0 ? ' cr)' : ' cr')]];
    met.forEach(([l, v], j) => {
      const y = 3.95 + j * 0.5;
      s.addText(l, { x: x + 0.25, y, w: cw - 0.5, h: 0.22, isTextBox: true, margin: 0,
        fontFace: SANS, fontSize: 8.5, color: MUTE });
      s.addText(v, { x: x + 0.25, y: y + 0.2, w: cw - 0.5, h: 0.26, isTextBox: true, margin: 0,
        fontFace: SANS, fontSize: 12.5, bold: true, color: j === 3 && sc.bridge.cumsur < 0 ? SEV : INK });
    });
  });
  s.addText('Probability-weighted, expected FY31 revenue is $31.9bn — compound growth of 1.2% a year. On an expected-value basis TCS spends the next five years roughly where it is. That, not the bear case, is the planning assumption.',
    { x: M, y: 6.2, w: W - 2 * M, h: 0.5, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12, bold: true, color: ACCENT, lineSpacing: 16 });
  chrome(s, 'Scenario model');
  s.addNotes('Probabilities are a judgement. 45% weight on bear or worse reflects that two of the three forces do not reverse on their own.');
}

/* ================= 8 · REVENUE PATHS ================= */
{
  const s = slide(false);
  title(s, 'Revenue paths', 'FY26 actual, FY27–FY31 modelled. The spread is $28bn of annual revenue by FY31.');
  const yrs = ['FY26', 'FY27', 'FY28', 'FY29', 'FY30', 'FY31'];
  s.addChart(pres.ChartType.line, [
    { name: 'Bull — AI Dividend', labels: yrs, values: [D.fy26.rev].concat(S.bull.rev) },
    { name: 'Base — Managed Transition', labels: yrs, values: [D.fy26.rev].concat(S.base.rev) },
    { name: 'Bear — Slow Bleed', labels: yrs, values: [D.fy26.rev].concat(S.bear.rev) },
    { name: 'Severe — Structural Break', labels: yrs, values: [D.fy26.rev].concat(S.severe.rev) }
  ], {
    x: M, y: 2.15, w: 8.2, h: 4.15,
    chartColors: [BULL, BASE, BEAR, SEV], lineSize: 2.5, lineSmooth: false,
    lineDataSymbol: 'circle', lineDataSymbolSize: 7,
    showLegend: true, legendPos: 'b', legendFontSize: 10, legendFontFace: SANS, legendColor: INK,
    showTitle: false, showValue: false,
    catAxisLabelColor: INK, catAxisLabelFontSize: 11, catAxisLabelFontFace: SANS,
    valAxisLabelColor: MUTE, valAxisLabelFontSize: 10, valAxisLabelFormatCode: '#,##0',
    valAxisTitle: 'Revenue ($m)', showValAxisTitle: true,
    valAxisTitleFontSize: 10, valAxisTitleColor: MUTE,
    valAxisMinVal: 15000, valAxisMaxVal: 50000,
    valGridLine: { color: RULE, size: 0.75 }, catGridLine: { style: 'none' },
    catAxisLineShow: false, valAxisLineShow: false
  });
  const pts = [
    ['The base case is not a recovery', 'It is stabilisation. $36.0bn by FY31 is 3.7% a year — below the rate at which TCS grew in every year to FY25.'],
    ['The severe case is not implausible', '−9.2% a year to $18.5bn implies headcount falling to 334,575 — roughly 250,000 fewer jobs. That is a political event before it is a financial one.'],
    ['The bull case needs a pricing change', 'Getting to $46.6bn requires the consulting, data, AI and cyber bucket to compound at 26% and the pricing model to shift. Capability alone will not do it.']
  ];
  pts.forEach(([h, p], i) => {
    const y = 2.35 + i * 1.4;
    s.addText(h, { x: 9.15, y, w: W - M - 9.15, h: 0.5, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12.5, bold: true, color: ACCENT, lineSpacing: 15 });
    s.addText(p, { x: 9.15, y: y + 0.5, w: W - M - 9.15, h: 0.85, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, color: INK, lineSpacing: 15 });
  });
  chrome(s, 'Scenario model');
}

/* ================= 9 · TATA SONS DEPENDENCE ================= */
{
  const s = slide(true);
  title(s, 'The concentration is the risk', 'Not TCS’s absolute performance. TCS could perform respectably and still leave this office short.', true);
  s.addText('87%', { x: M, y: 2.5, w: 3.3, h: 1.5, isTextBox: true, margin: 0,
    fontFace: SERIF, fontSize: 90, bold: true, color: ACC_LT, valign: 'middle' });
  s.addText('of Tata Sons’ FY26 dividend income came from TCS alone — ₹28,291 crore of ₹32,528 crore.',
    { x: M, y: 4.05, w: 3.4, h: 0.9, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 13, color: 'B9CBD1', lineSpacing: 18 });
  s.addText('The rest of a listed portfolio worth ₹11.68 trillion contributed ₹4,237 crore. The TCS dividend has already fallen 12% — from ₹32,184 crore in FY25.',
    { x: M, y: 5.0, w: 3.4, h: 1.1, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11.5, color: '8FA9B2', lineSpacing: 16 });

  s.addText('AGAINST THAT, THE COMMITMENTS', { x: 4.5, y: 2.35, w: 8, h: 0.28, isTextBox: true,
    margin: 0, fontFace: SANS, fontSize: 9.5, bold: true, charSpacing: 2, color: '5D7C86' });
  const commit = [
    ['Tata Electronics', '₹1,18,000 cr', 'Dholera fab ₹91,000 cr + Assam ATMP ₹27,000 cr'],
    ['Air India', '₹22,618 cr invested', '600-aircraft order; a five- to ten-year transformation'],
    ['Agratas', '£4bn Somerset', 'Battery plant, part-offset by a £380m UK grant'],
    ['Tata Digital', '₹22,903 cr equity', 'FY26 loss of ₹4,974 cr on ₹35,990 cr revenue']
  ];
  commit.forEach(([nm, amt, sub], i) => {
    const y = 2.72 + i * 0.86;
    s.addShape(pres.ShapeType.rect, { x: 4.5, y, w: W - M - 4.5, h: 0.74,
      fill: { color: '16303A' }, line: { color: '22414D', width: 0.75 } });
    s.addText(nm, { x: 4.75, y: y + 0.09, w: 2.4, h: 0.28, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12.5, bold: true, color: PAPER, valign: 'middle' });
    s.addText(amt, { x: 4.75, y: y + 0.37, w: 2.4, h: 0.28, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12, color: ACC_LT, valign: 'middle' });
    s.addText(sub, { x: 7.3, y: y + 0.09, w: W - M - 7.55, h: 0.56, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, color: '9FB6BE', valign: 'middle', lineSpacing: 15 });
  });
  s.addText('The four newest businesses lost roughly ₹30,000 crore between them in FY26. Parent net cash at the FY26 close was ₹21,841 crore, with no borrowings.',
    { x: 4.5, y: 6.22, w: W - M - 4.5, h: 0.5, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, italic: true, color: '8FA9B2', lineSpacing: 15 });
  chrome(s, 'The Tata Sons view', true);
  s.addNotes('The commitments were sized against a dividend that has already started falling. That is the exposure.');
}

/* ================= 10 · DIVIDEND vs BREAK-EVEN ================= */
{
  const s = slide(false);
  title(s, 'Two of four paths never reach break-even', 'Dividend to Tata Sons against what the parent needs in FY31 to fund its own costs and its committed capital call.');
  const yrs = ['FY26', 'FY27', 'FY28', 'FY29', 'FY30', 'FY31'];
  s.addChart(pres.ChartType.line, [
    { name: 'Bull', labels: yrs, values: [D.fy26.ts].concat(S.bull.ts) },
    { name: 'Base', labels: yrs, values: [D.fy26.ts].concat(S.base.ts) },
    { name: 'Bear', labels: yrs, values: [D.fy26.ts].concat(S.bear.ts) },
    { name: 'Severe', labels: yrs, values: [D.fy26.ts].concat(S.severe.ts) },
    { name: 'Break-even requirement', labels: yrs, values: yrs.map(() => D.breakeven_fy31) }
  ], {
    x: M, y: 2.15, w: 8.2, h: 4.15,
    chartColors: [BULL, BASE, BEAR, SEV, '6C7B80'],
    lineSize: 2.5, lineSmooth: false, lineDataSymbol: 'circle', lineDataSymbolSize: 7,
    showLegend: true, legendPos: 'b', legendFontSize: 10, legendFontFace: SANS, legendColor: INK,
    showTitle: false, showValue: false,
    catAxisLabelColor: INK, catAxisLabelFontSize: 11, catAxisLabelFontFace: SANS,
    valAxisLabelColor: MUTE, valAxisLabelFontSize: 10, valAxisLabelFormatCode: '#,##0',
    valAxisTitle: 'Dividend to Tata Sons (₹ crore)', showValAxisTitle: true,
    valAxisTitleFontSize: 10, valAxisTitleColor: MUTE,
    valAxisMinVal: 10000, valAxisMaxVal: 60000,
    valGridLine: { color: RULE, size: 0.75 }, catGridLine: { style: 'none' },
    catAxisLineShow: false, valAxisLineShow: false
  });
  card(s, 9.15, 2.3, W - M - 9.15, 1.5);
  s.addText('BREAK-EVEN, FY31', { x: 9.37, y: 2.48, w: 3.2, h: 0.24, isTextBox: true, margin: 0,
    fontFace: SANS, fontSize: 9, bold: true, charSpacing: 1.5, color: MUTE });
  s.addText('₹30,827 cr', { x: 9.37, y: 2.74, w: 3.2, h: 0.5, isTextBox: true, margin: 0,
    fontFace: SERIF, fontSize: 26, bold: true, color: INK, valign: 'middle' });
  s.addText('FY31 capital call plus parent costs, less dividends from all other holdings',
    { x: 9.37, y: 3.24, w: 3.2, h: 0.5, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 9.5, color: MUTE, lineSpacing: 12 });
  const dv = [[BULL, 'Bull', S.bull.ts[4], true], [BASE, 'Base', S.base.ts[4], true],
    [BEAR, 'Bear', S.bear.ts[4], false], [SEV, 'Severe', S.severe.ts[4], false]];
  dv.forEach(([col, nm, v, ok], i) => {
    const y = 4.0 + i * 0.55;
    s.addShape(pres.ShapeType.ellipse, { x: 9.15, y: y + 0.11, w: 0.16, h: 0.16, fill: { color: col } });
    s.addText(nm, { x: 9.4, y, w: 0.9, h: 0.38, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12, bold: true, color: INK, valign: 'middle' });
    s.addText('₹' + fmt(Math.round(v)) + ' cr', { x: 10.15, y, w: 1.4, h: 0.38, isTextBox: true,
      margin: 0, fontFace: SANS, fontSize: 12, color: INK, valign: 'middle', align: 'right' });
    s.addText(ok ? 'clears' : 'short', { x: 11.65, y, w: 1, h: 0.38, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, bold: true, color: ok ? BULL : SEV, valign: 'middle', align: 'right' });
  });
  s.addText('In the bear case Tata Sons runs a cumulative deficit of ₹19,163 cr over five years and exhausts its net cash. In the severe case the deficit is ₹56,517 cr and the parent is borrowing from FY30.',
    { x: 9.15, y: 6.2, w: W - M - 9.15, h: 0.75, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 10.5, color: INK, lineSpacing: 14 });
  chrome(s, 'The Tata Sons view');
}

/* ================= 11 · GROWTH vs MARGIN ================= */
{
  const s = slide(false);
  title(s, 'Margin cannot rescue a growth failure', 'Sensitivity of the FY31 dividend to Tata Sons. Both levers move it; only one compounds.');
  const cmp = [
    [MUTE, 'Operating margin, 25% → 27%', '+₹2,400 cr',
     'An outcome nobody at TCS has achieved in a decade. It buys one step up, then resets — margin does not compound.'],
    [ACCENT, 'Revenue growth, 0% → 4% a year', '+₹6,900 cr',
     'Nearly three times the effect, and it compounds every year. Growth is the term that decides the answer.']
  ];
  cmp.forEach(([col, h, v, p], i) => {
    const x = M + i * ((W - 2 * M) / 2 + 0.15);
    const cw2 = (W - 2 * M) / 2 - 0.15;
    card(s, x, 2.25, cw2, 2.6, i === 1 ? 'E6EFF1' : PANEL);
    s.addText(h.toUpperCase(), { x: x + 0.32, y: 2.52, w: cw2 - 0.64, h: 0.3, isTextBox: true,
      margin: 0, fontFace: SANS, fontSize: 10, bold: true, charSpacing: 1.5, color: col });
    s.addText(v, { x: x + 0.32, y: 2.9, w: cw2 - 0.64, h: 0.85, isTextBox: true, margin: 0,
      fontFace: SERIF, fontSize: 40, bold: true, color: i === 1 ? ACCENT : INK, valign: 'middle' });
    s.addText(p, { x: x + 0.32, y: 3.82, w: cw2 - 0.64, h: 0.8, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12, color: INK, lineSpacing: 16 });
  });
  s.addText('And the two are not independent', { x: M, y: 5.2, w: 6, h: 0.4, isTextBox: true,
    margin: 0, fontFace: SERIF, fontSize: 20, bold: true, color: INK });
  s.addText('In every downside path modelled here, margin falls because growth fails, not alongside it: idle capacity, pricing concessions to hold volume, and the cost of a workforce that cannot be reshaped as fast as revenue moves. Cost discipline buys time. It does not buy the outcome.',
    { x: M, y: 5.62, w: W - 2 * M, h: 0.9, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 13, color: INK, lineSpacing: 19 });
  chrome(s, 'Sensitivity');
  s.addNotes('Holds FCF at 82% of net income, dividend at 92% of FCF, the rupee at 98.22 in FY31, and the stake at 71.74%.');
}

/* ================= 12 · RECOMMENDATIONS A ================= */
{
  const s = slide(false);
  title(s, 'Three things to do regardless of TCS', 'These are about the group’s own balance sheet. None of them depends on how the TCS question resolves.');
  const recs = [
    ['A1', 'Break the single-source dependency',
     'Target TCS below 70% of parent dividend income within three years. Two levers: the payout policies of a ₹11.68 trillion listed portfolio that returned only ₹4,237 crore last year, and third-party capital at the asset level — as with TPG in HyperVault and BlackRock and Mubadala in Tata Power Renewables.',
     'Set the target explicitly and report against it to this board each year.'],
    ['A2', 'Sequence the capital call',
     'Dholera, Assam, the Air India fleet, Agratas Somerset and Tata Digital are being funded in parallel against a dividend that fell 12% last year. Rank them by how irreversible the commitment already is and whether a partner or a government can carry part of the load.',
     'A ranked, dated sequencing plan with a named decision point for each programme.'],
    ['A3', 'Settle the RBI question first',
     'Tata Sons remains on the upper-layer NBFC list with the core-investment-company deregistration plea unresolved. If the shock absorber for a falling dividend is debt capacity at the parent, that capacity must exist and be priced before the dividend falls, not during.',
     'What the parent can borrow, at what cost, under each outcome of the deregistration decision.']
  ];
  const cw = (W - 2 * M - 2 * 0.28) / 3;
  recs.forEach(([nu, h, p, ask], i) => {
    const x = M + i * (cw + 0.28);
    card(s, x, 2.2, cw, 3.95);
    s.addText(nu, { x: x + 0.3, y: 2.45, w: 0.8, h: 0.3, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12, bold: true, color: ACCENT });
    s.addText(h, { x: x + 0.3, y: 2.75, w: cw - 0.6, h: 0.7, isTextBox: true, margin: 0,
      fontFace: SERIF, fontSize: 17, bold: true, color: INK, valign: 'top' });
    s.addText(p, { x: x + 0.3, y: 3.5, w: cw - 0.6, h: 1.85, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, color: INK, lineSpacing: 15, valign: 'top' });
    s.addText([{ text: 'Ask for: ', options: { bold: true, color: ACCENT } }, { text: ask, options: {} }],
      { x: x + 0.3, y: 5.35, w: cw - 0.6, h: 0.65, isTextBox: true, margin: 0,
        fontFace: SANS, fontSize: 10.5, color: INK, lineSpacing: 14, valign: 'top' });
  });
  chrome(s, 'Recommendations');
}

/* ================= 13 · RECOMMENDATIONS B ================= */
{
  const s = slide(false);
  title(s, 'Five things to ask of TCS', 'B2 is the central one. The rest support it.');
  const recs = [
    ['B1', 'Change what this board asks for',
     'Revenue growth and headcount are the wrong scoreboard in a deflating market. Track revenue per employee, share of revenue on non-effort-linked pricing, IP-led revenue share, and gross margin on the AI book. None is currently a headline metric.'],
    ['B2', 'Force the pricing transition',
     'TCS’s exposure is not AI capability — it is billing for effort while effort collapses. Whoever moves to outcome and platform pricing first keeps the productivity gain; whoever moves last refunds it at every renewal.'],
    ['B3', 'Products and platforms: 5% → 15%',
     'BaNCS, ignio, iON and ADD are the only assets AI makes more valuable, because they are licensed rather than billed by the hour. A build-or-buy decision that will not make itself.'],
    ['B4', 'Interrogate HyperVault, and ring-fence it',
     'Up to ₹18,000 crore of equity into AI data centres converts TCS from an asset-light dividend machine into a capital-intensive story — just as this parent depends most on the dividend. Free cash flow conversion has already fallen to 81% of net income.'],
    ['B5', 'Fix the workforce model',
     'Headcount fell 23,460 in FY26 then rose 9,279 in one quarter with margin dropping 130bps. That is oscillation, not strategy. The delivery pyramid is becoming something with a thinner base and a thicker middle; say so and plan it.']
  ];
  recs.forEach(([nu, h, p], i) => {
    const y = 2.15 + i * 0.855;
    s.addShape(pres.ShapeType.rect, { x: M, y, w: W - 2 * M, h: 0.75,
      fill: { color: nu === 'B2' ? 'E6EFF1' : PANEL }, line: { color: nu === 'B2' ? ACC_LT : RULE, width: 0.75 } });
    s.addText(nu, { x: M + 0.28, y: y + 0.02, w: 0.65, h: 0.72, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12, bold: true, color: ACCENT, valign: 'middle' });
    s.addText(h, { x: M + 0.95, y: y + 0.02, w: 3.35, h: 0.72, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 13, bold: true, color: INK, valign: 'middle' });
    s.addText(p, { x: M + 4.45, y: y + 0.04, w: W - M - 4.45 - M - 0.2, h: 0.68, isTextBox: true,
      margin: 0, fontFace: SANS, fontSize: 10.5, color: INK, lineSpacing: 13.5, valign: 'middle' });
  });
  s.addText('What I would ask for on B2: what proportion of the FY27 order book is priced on outcomes rather than effort — and what is the FY29 target?',
    { x: M, y: 6.48, w: W - 2 * M, h: 0.4, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11.5, bold: true, color: ACCENT });
  chrome(s, 'Recommendations');
}

/* ================= 14 · TRIGGERS (dark close) ================= */
{
  const s = slide(true);
  title(s, 'Decide the trigger points now', 'The worst version of this is a board debating the group’s capital plan in the quarter the dividend disappoints.', true);
  const trig = [
    ['Two consecutive quarters of negative constant-currency growth', 'Re-open the capital sequencing plan'],
    ['Dividend to Tata Sons below ₹25,000 crore', 'Activate non-TCS payout levers and asset-level partners'],
    ['Free cash flow conversion below 70% of net income', 'Review HyperVault capital schedule and ring-fencing'],
    ['Revenue per employee flat or falling for a year', 'The AI transition is not working; escalate to a strategic review']
  ];
  trig.forEach(([t, r], i) => {
    const y = 2.35 + i * 0.95;
    s.addShape(pres.ShapeType.rect, { x: M, y, w: W - 2 * M, h: 0.8,
      fill: { color: '16303A' }, line: { color: '22414D', width: 0.75 } });
    s.addText('IF', { x: M + 0.3, y: y + 0.05, w: 0.45, h: 0.7, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 10, bold: true, charSpacing: 1.5, color: '8FB0BA', valign: 'middle' });
    s.addText(t, { x: M + 0.85, y: y + 0.05, w: 5.5, h: 0.7, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12.5, color: PAPER, valign: 'middle' });
    s.addText('THEN', { x: 7.0, y: y + 0.05, w: 0.7, h: 0.7, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 10, bold: true, charSpacing: 1.5, color: '8FB0BA', valign: 'middle' });
    s.addText(r, { x: 7.75, y: y + 0.05, w: W - M - 7.95, h: 0.7, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 12.5, color: ACC_LT, valign: 'middle' });
  });
  s.addText('We are not exposed to whether TCS grows. We are exposed to having assumed it would.',
    { x: M, y: 6.25, w: W - 2 * M, h: 0.5, isTextBox: true, margin: 0,
      fontFace: SERIF, fontSize: 20, italic: true, color: PAPER });
  chrome(s, 'Governance', true);
  s.addNotes('Agree these while nobody is under pressure. Review each quarter.');
}

/* ================= 15 · BASIS ================= */
{
  const s = slide(false);
  title(s, 'What this rests on, and what would sharpen it', 'A model whose assumptions cannot be challenged is not an analysis. These are the ones to challenge.');
  const cols = [
    [ACCENT, 'Reliable', ['TCS financial history FY22–FY26 and Q1 FY27, from results press releases and fact sheets',
      'Tata Sons FY26 dividend receipts and standalone results as reported',
      'Capital commitments for Tata Electronics, Air India, Agratas and Tata Digital',
      'Peer results: Infosys FY26 and FY27 guidance, Accenture Q3 FY2026',
      'Market data as at 24 August 2026']],
    [BEAR, 'Judgement — challenge these', ['The seven-bucket revenue decomposition (TCS has disclosed no service-line split since FY21)',
      'Per-bucket deflation and volume drivers for each scenario',
      'Scenario probabilities (15 / 40 / 30 / 15)',
      'The forward capital-call profile, held constant across scenarios',
      'Margin, free-cash-flow conversion and payout paths']],
    [BULL, 'What internal data would change', ['Actual service-line revenue and gross margin',
      'Pricing model mix: time-and-materials vs fixed vs outcome-linked, and the trend',
      'TCV-to-revenue conversion by deal vintage',
      'Renewal pricing on the top 50 accounts',
      'HyperVault business case; committed vs discretionary capital call by year']]
  ];
  const cw = (W - 2 * M - 2 * 0.3) / 3;
  cols.forEach(([col, h, items], i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 2.2, cw, 3.35);
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.3, y: 2.5, w: 0.16, h: 0.16, fill: { color: col } });
    s.addText(h.toUpperCase(), { x: x + 0.56, y: 2.45, w: cw - 0.85, h: 0.28, isTextBox: true,
      margin: 0, fontFace: SANS, fontSize: 9.5, bold: true, charSpacing: 1.5, color: col, valign: 'middle' });
    s.addText(items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1, paraSpaceAfter: 8 } })),
      { x: x + 0.3, y: 2.85, w: cw - 0.6, h: 2.6, isTextBox: true, margin: 0,
        fontFace: SANS, fontSize: 10.5, color: INK, lineSpacing: 14, valign: 'top' });
  });
  s.addText('Every assumption is visible and editable in the accompanying workbook. Nothing here uses internal management information.',
    { x: M, y: 5.85, w: W - 2 * M, h: 0.4, isTextBox: true, margin: 0,
      fontFace: SANS, fontSize: 11, italic: true, color: MUTE });
  chrome(s, 'Basis');
}

pres.writeFile({ fileName: '/root/tcs/TCS_Board_Assessment.pptx' })
  .then(f => console.log('written', f));
