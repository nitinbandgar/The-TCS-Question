# The TCS Question

A strategic assessment of Tata Consultancy Services under AI disruption, prepared for the Chairman, Tata Sons — with a driver-based five-year scenario model and an analysis of what each path means for Tata Sons' capacity to fund the group.

**Base year:** FY26 (year ended 31 March 2026) · **Horizon:** FY27–FY31 · **As at:** 26 August 2026

---

## The short version

TCS reported the first annual dollar revenue decline of its listed history in FY26: **$30.0bn, −0.5% reported and −2.4% in constant currency**. Operating margin nonetheless rose to 25.0% — bought with 23,460 fewer people rather than with better pricing. When hiring resumed in Q1 FY27 (+9,279), margin fell 130bps to 24.0% in a single quarter.

Three separate forces are at work and the market prices them as one:

| Force | Nature | Half-life |
|---|---|---|
| Frozen discretionary spend | Cyclical | 4–8 quarters — reverses on its own |
| AI collapsing effort per unit of outcome | Deflationary | Permanent — survivable only by repricing |
| GCC insourcing and AI-native entrants | Competitive | A decade — clients are building the capability |

**42% of FY26 revenue ($12.7bn) sits in work AI can substantially automate.** Whether that becomes lost revenue depends on price pass-through and volume elasticity — which is what the model flexes.

**87% of Tata Sons' FY26 dividend income came from TCS alone** (₹28,291 cr of ₹32,528 cr), already down 12% year on year. Against that sit the Dholera fab (₹91,000 cr), Assam ATMP (₹27,000 cr), the Air India fleet, Agratas Somerset (£4bn) and Tata Digital.

## Scenario results

| Scenario | Prob. | FY31 revenue | Revenue CAGR | Cumulative dividend to Tata Sons | Parent 5-yr cash |
|---|---|---|---|---|---|
| Bull — AI Dividend | 15% | $46.6bn | +9.2% | ₹2,08,225 cr | +₹69,133 cr |
| Base — Managed Transition | 40% | $36.0bn | +3.7% | ₹1,67,137 cr | +₹28,045 cr |
| Bear — Slow Bleed | 30% | $25.8bn | −3.0% | ₹1,19,930 cr | −₹19,163 cr |
| Severe — Structural Break | 15% | $18.5bn | −9.2% | ₹82,576 cr | −₹56,517 cr |

Probability-weighted, expected FY31 revenue is **$31.9bn — 1.2% a year**. On an expected-value basis TCS spends the next five years roughly where it is.

Tata Sons needs **₹30,827 crore** from TCS in FY31 to cover its own costs and its committed capital call. Two of the four paths never reach it.

---

## Contents

```
index.html                           Interactive assessment (GitHub Pages site)
Model/
  TCS_Scenario_Model_FY27-FY31.xlsx  Live scenario model — every assumption editable
  model.py                           The model itself: drivers, engine, scenario logic
  data.json                          Model output, inlined into the web page
  build_xlsx.py                      Workbook generator (openpyxl)
  build_docx.js / build_pptx.js      Memo and deck generators (docx, pptxgenjs)
```

The page is a single self-contained `index.html` — the model output is inlined,
so it has no build step and no runtime dependency beyond Google Fonts. Serve the
repository root with GitHub Pages (Settings → Pages → branch `main`, folder `/ (root)`).

### The Excel model

`TCS_Scenario_Model_FY27-FY31.xlsx` is fully formula-driven — 940 live formulas, no hardcoded results. Change any blue input on the **Assumptions** sheet and every downstream sheet recalculates.

| Sheet | What it holds |
|---|---|
| Read Me | Colour legend and how to use the workbook |
| History | TCS actuals FY22–FY26 and Q1 FY27, with sources |
| Assumptions | Shared parameters, the seven-bucket revenue decomposition, four scenario driver blocks, Tata Sons parameters |
| Revenue Build | Bucket-level revenue build per scenario |
| Scenario Output | Full P&L, cash, dividend and headcount per scenario, plus the probability-weighted case |
| Tata Sons Bridge | Parent cash bridge: TCS dividend against the capital call |
| Sensitivity | FY31 dividend across revenue CAGR × terminal margin, with the break-even calculation |
| Risk Register | Revenue at risk by exposure bucket |

Blue = input · black = formula · green = cross-sheet link · yellow fill = key judgement.

---

## How the model works

Revenue is decomposed into seven buckets by AI exposure. Each year, each bucket moves by:

```
revenue(t) = revenue(t-1) × (1 + volume_growth × decay^n) × (1 + deflation × phasing)
```

where **deflation** = AI productivity gain × automatable share × price pass-through to the client, and **volume growth** = underlying demand plus the elasticity response to a lower unit price. Margin, free-cash-flow conversion, payout ratio and revenue-per-employee growth are set per scenario per year.

Reproduce the results:

```bash
python3 Model/model.py
```

That prints all four scenarios and the Tata Sons bridge. Rebuilding the workbook
needs `openpyxl`; the memo and deck generators need `npm i docx pptxgenjs`.

---

## What is reliable and what is a judgement

**Reliable — from public filings:** TCS financial history FY22–FY26 and Q1 FY27 (results press releases and fact sheets); Tata Sons FY26 dividend receipts and standalone results as reported; capital commitments for Tata Electronics, Air India, Agratas and Tata Digital; peer results for Infosys and Accenture; market data as at 24 August 2026.

**Judgement — challenge these first:**

- **The seven-bucket revenue decomposition.** TCS has disclosed no service-line split since FY21. The weights are reconstructed from the last disclosure, deal composition and peer structure. This is the single most challengeable input in the whole analysis.
- Per-bucket deflation and volume drivers for each scenario.
- Scenario probabilities (15 / 40 / 30 / 15).
- The forward capital-call profile, held constant across scenarios by design — the commitments do not shrink because TCS slows.
- Margin, free-cash-flow conversion and payout paths.

**No internal management information was used.** Actual service-line revenue and gross margin, pricing-model mix, TCV-to-revenue conversion by deal vintage, and renewal pricing on the top 50 accounts would all materially sharpen this — in either direction.

---

## Sources

TCS Q4 FY26 results press release and fact sheet; TCS Q1 FY27 results press release; TCS Q4 FY25 and Q4 FY24 press releases ([tcs.com/investor-relations](https://www.tcs.com/investor-relations)). Infosys FY26 results and FY27 guidance. Accenture Q3 FY2026 results. Tata Sons FY26 accounts as reported in Business Today (28 July 2026) and Business Standard (17–18 August 2026). Screener.in and stockanalysis.com for market data. Business of GCC for India GCC market size. TCS–TPG HyperVault announcement. RBI upper-layer NBFC list, August 2026.

---

*A model whose assumptions cannot be challenged is not an analysis; it is an opinion with decimal places.*
