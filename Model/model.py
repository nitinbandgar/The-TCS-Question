"""
TCS Five-Year Scenario Model (FY27-FY31)
Prepared for the Chairman, Tata Sons.
All history from TCS quarterly fact sheets / press releases (FY24-FY26, Q1 FY27).
"""
import json
import pandas as pd

# ----------------------------------------------------------------------------
# 1. HISTORY (actuals) - source: TCS press releases & fact sheets
# ----------------------------------------------------------------------------
HISTORY = pd.DataFrame([
    # FY, RevUSDm, USDgrowth, CCgrowth, OPM, NetMargin, PATusd, RevINRcr, PATINRcr,
    # TCV, Headcount, DPS, PayoutINRcr
    ["FY22", 25707, 0.157, 0.157, 0.250, 0.204, 5240, 191754, 38449, 34600, 592195, 43, 15671],
    ["FY23", 27927, 0.086, 0.138, 0.244, 0.192, 5364, 225458, 42303, 34100, 614795, 115, 42084],
    ["FY24", 29080, 0.041, 0.034, 0.246, 0.196, 5697, 240893, 46099, 42700, 601546, 73, 26411],
    ["FY25", 30180, 0.038, 0.042, 0.243, 0.190, 5739, 255324, 48797, 39400, 607979, 126, 44962],
    ["FY26", 30017, -0.005, -0.024, 0.250, 0.198, 5938, 267021, 49454, 40700, 584519, 110, 39571],
], columns=["FY", "RevUSDm", "USDg", "CCg", "OPM", "NetM", "PATusdm",
            "RevINRcr", "PATINRcr", "TCVusdbn", "Headcount", "DPS", "PayoutINRcr"])

Q1FY27 = dict(RevUSDm=7624, CCg_yoy=0.032, OPM=0.240, NetM=0.192,
              PATusdm=1460, TCVusdbn=9.5, Headcount=593798, Attrition=0.136,
              AI_annualised_usdbn=2.6)

# ----------------------------------------------------------------------------
# 2. REVENUE DECOMPOSITION BY AI EXPOSURE
#    TCS stopped disclosing service-line splits in FY21. Weights below are an
#    analyst reconstruction from TCS's last disclosed mix, deal composition,
#    delivery-org headcount and peer disclosure. Treat as an ESTIMATE.
# ----------------------------------------------------------------------------
BUCKETS = [
    # name, FY26 weight, AI-automatable share of work, qualitative note
    ("Application Development & Maintenance (custom)", 0.32, 0.55,
     "Code generation, test automation and agentic L1/L2 support attack the core of this book. "
     "Largest single pool and the most exposed."),
    ("Enterprise / Package Application Services", 0.13, 0.40,
     "SAP/Oracle/Salesforce migration and rollout. Config work automates; change management, "
     "data migration and integration hold value."),
    ("Cloud & Infrastructure Managed Services", 0.16, 0.50,
     "AIOps compresses run-cost; offset by AI-infrastructure build, GPU estate operations "
     "and the HyperVault adjacency."),
    ("Business Process Services", 0.10, 0.65,
     "Most directly substitutable. Agentic workflows displace seat-based delivery; the "
     "FTE pricing model is the liability."),
    ("Engineering & R&D Services", 0.09, 0.25,
     "Embedded, semiconductor, automotive software. Least substitutable near-term and "
     "structurally growing."),
    ("Consulting, Data, AI & Cybersecurity", 0.15, 0.15,
     "The growth engine. AI-native demand, data estate rebuild, model governance, "
     "security. TCS AI book already ~8.5% of revenue."),
    ("Products & Platforms (BaNCS, ignio, iON, ADD)", 0.05, 0.10,
     "IP-led, non-linear. AI raises product value rather than deflating it. Strategically "
     "under-scaled at ~5% of revenue."),
]

# ----------------------------------------------------------------------------
# 3. SCENARIO DRIVERS
#    For each bucket and scenario: (annual realisation deflation, annual volume growth)
#    Deflation = AI productivity gain x share automatable x price pass-through to client
#    Volume    = underlying demand + elasticity response to lower unit price
# ----------------------------------------------------------------------------
DRIVERS = {
    # scenario: [(deflation, volume) per bucket, in BUCKETS order]
    "Bull — AI Dividend": [
        (-0.050, 0.090), (-0.030, 0.100), (-0.040, 0.120), (-0.060, 0.100),
        (-0.010, 0.120), (0.000, 0.260), (0.000, 0.160)],
    "Base — Managed Transition": [
        (-0.060, 0.060), (-0.040, 0.070), (-0.050, 0.080), (-0.080, 0.060),
        (-0.020, 0.080), (-0.020, 0.180), (-0.010, 0.110)],
    "Bear — Slow Bleed": [
        (-0.090, 0.020), (-0.060, 0.030), (-0.080, 0.040), (-0.120, 0.020),
        (-0.040, 0.040), (-0.040, 0.110), (-0.030, 0.060)],
    "Severe — Structural Break": [
        (-0.130, 0.000), (-0.090, 0.000), (-0.110, 0.010), (-0.180, -0.020),
        (-0.060, 0.010), (-0.070, 0.050), (-0.060, 0.020)],
}

# Deflation intensity phasing FY27..FY31 (shocks build, then the base resets)
PHASING = {
    "Bull — AI Dividend":        [1.0, 1.0, 0.9, 0.9, 0.8],
    "Base — Managed Transition": [1.0, 1.0, 1.0, 0.9, 0.9],
    "Bear — Slow Bleed":         [0.8, 1.0, 1.1, 1.0, 0.9],
    "Severe — Structural Break": [0.6, 1.0, 1.2, 1.1, 0.9],
}
# High-growth buckets (Consulting/AI, Products) decay toward maturity
GROWTH_DECAY = {5: 0.88, 6: 0.92}

PROBABILITY = {
    "Bull — AI Dividend": 0.15,
    "Base — Managed Transition": 0.40,
    "Bear — Slow Bleed": 0.30,
    "Severe — Structural Break": 0.15,
}

# Operating margin path by scenario (FY27..FY31)
MARGINS = {
    "Bull — AI Dividend":        [0.245, 0.255, 0.262, 0.267, 0.270],
    "Base — Managed Transition": [0.243, 0.246, 0.250, 0.252, 0.255],
    "Bear — Slow Bleed":         [0.240, 0.234, 0.228, 0.224, 0.220],
    "Severe — Structural Break": [0.235, 0.220, 0.202, 0.188, 0.175],
}

# Free-cash-flow conversion (FCF / PAT). Falls as HyperVault / AI-DC capex ramps.
FCF_CONV = {
    "Bull — AI Dividend":        [0.80, 0.75, 0.79, 0.84, 0.88],
    "Base — Managed Transition": [0.80, 0.76, 0.78, 0.82, 0.85],
    "Bear — Slow Bleed":         [0.78, 0.72, 0.74, 0.78, 0.82],
    "Severe — Structural Break": [0.75, 0.66, 0.68, 0.74, 0.80],
}

# Share of FCF paid out as dividend (policy: 80-100% of FCF)
PAYOUT_OF_FCF = {
    "Bull — AI Dividend":        [0.92, 0.92, 0.92, 0.92, 0.92],
    "Base — Managed Transition": [0.93, 0.93, 0.92, 0.92, 0.92],
    "Bear — Slow Bleed":         [0.95, 0.93, 0.91, 0.90, 0.90],
    "Severe — Structural Break": [0.95, 0.90, 0.85, 0.85, 0.85],
}

# Revenue per employee CAGR (AI leverage / de-linking headcount from revenue)
REV_PER_EMP_GROWTH = {
    "Bull — AI Dividend": 0.055,
    "Base — Managed Transition": 0.040,
    "Bear — Slow Bleed": 0.030,
    "Severe — Structural Break": 0.015,
}

# ----------------------------------------------------------------------------
# 4. SHARED ASSUMPTIONS
# ----------------------------------------------------------------------------
YEARS = ["FY27", "FY28", "FY29", "FY30", "FY31"]
FX_FY26 = 267021 * 10 / 30017            # INR crore*10 -> INR mn / USD mn  => ~88.96
FX_DEPRECIATION = 0.020                   # 2.0% p.a. INR depreciation
OTHER_INCOME_MARGIN = 0.016               # other income as % of revenue
TAX_RATE = 0.255
TATA_SONS_STAKE = 0.7174
SHARES_CR = 361.8                         # TCS shares outstanding (crore)
REV_PER_EMP_FY26 = 30017e6 / 584519       # USD 51,354

# Tata Sons side (FY26 actuals, INR crore)
TS_TCS_DIV_FY26 = 28291
TS_TOTAL_DIV_FY26 = 32528
TS_NON_TCS_DIV_FY26 = TS_TOTAL_DIV_FY26 - TS_TCS_DIV_FY26     # 4,237
TS_NON_TCS_GROWTH = 0.09
TS_OPEX_TAX_FY26 = TS_TOTAL_DIV_FY26 - 25544                   # 6,984 (op cash flow gap)
TS_OPEX_GROWTH = 0.06
TS_NET_CASH_FY26 = 21841
# Equity deployment into unlisted/strategic bets (Air India, Tata Electronics,
# Agratas, Tata Digital). FY26 actual = 15,089 (subs) + 867 (assoc) = 15,956.
TS_CAPITAL_CALL = {"FY27": 20000, "FY28": 24000, "FY29": 26000, "FY30": 27000, "FY31": 28000}

FX = {}
_fx = FX_FY26
for y in YEARS:
    _fx *= (1 + FX_DEPRECIATION)
    FX[y] = _fx

# ----------------------------------------------------------------------------
# 5. ENGINE
# ----------------------------------------------------------------------------
def run(scenario):
    drv = DRIVERS[scenario]
    phase = PHASING[scenario]
    rev = [w * HISTORY.iloc[-1]["RevUSDm"] for _, w, _, _ in BUCKETS]
    bucket_rows, rows = [], []
    prev_total = HISTORY.iloc[-1]["RevUSDm"]
    headcount = HISTORY.iloc[-1]["Headcount"]
    rpe = REV_PER_EMP_FY26

    for i, y in enumerate(YEARS):
        new_rev = []
        for b, (defl, vol) in enumerate(drv):
            d = defl * phase[i]
            v = vol * (GROWTH_DECAY.get(b, 1.0) ** i)
            new_rev.append(rev[b] * (1 + v) * (1 + d))
        rev = new_rev
        total = sum(rev)
        ccg = total / prev_total - 1

        opm = MARGINS[scenario][i]
        ebit = total * opm
        pbt = total * (opm + OTHER_INCOME_MARGIN)
        pat = pbt * (1 - TAX_RATE)
        netm = pat / total

        rev_inr = total * FX[y] / 10          # USD mn * INR/USD -> INR mn ; /10 -> INR crore
        pat_inr = pat * FX[y] / 10
        fcf_inr = pat_inr * FCF_CONV[scenario][i]
        div_inr = fcf_inr * PAYOUT_OF_FCF[scenario][i]
        dps = div_inr / SHARES_CR
        div_ts = div_inr * TATA_SONS_STAKE

        rpe *= (1 + REV_PER_EMP_GROWTH[scenario])
        headcount = total * 1e6 / rpe

        rows.append(dict(Scenario=scenario, FY=y, RevUSDm=total, CCgrowth=ccg,
                         OPM=opm, EBITusdm=ebit, NetMargin=netm, PATusdm=pat,
                         FX=FX[y], RevINRcr=rev_inr, PATINRcr=pat_inr,
                         FCFINRcr=fcf_inr, DividendINRcr=div_inr, DPS=dps,
                         DivToTataSonsINRcr=div_ts, Headcount=headcount,
                         RevPerEmpUSD=rpe))
        bucket_rows.append(dict(Scenario=scenario, FY=y,
                                **{BUCKETS[b][0]: rev[b] for b in range(len(BUCKETS))}))
        prev_total = total
    return pd.DataFrame(rows), pd.DataFrame(bucket_rows)


def tata_sons_view(scen_df):
    """Tata Sons parent-level cash bridge for one scenario."""
    out = []
    cash = TS_NET_CASH_FY26
    non_tcs = TS_NON_TCS_DIV_FY26
    opex = TS_OPEX_TAX_FY26
    for i, y in enumerate(YEARS):
        non_tcs *= (1 + TS_NON_TCS_GROWTH)
        opex *= (1 + TS_OPEX_GROWTH)
        tcs_div = scen_df.iloc[i]["DivToTataSonsINRcr"]
        inflow = tcs_div + non_tcs
        capital = TS_CAPITAL_CALL[y]
        surplus = inflow - opex - capital
        cash += surplus
        out.append(dict(FY=y, TCSDividend=tcs_div, OtherDividend=non_tcs,
                        TotalInflow=inflow, ParentOpexTax=opex,
                        CapitalDeployment=capital, Surplus=surplus,
                        CumulativeNetCash=cash,
                        TCSShareOfInflow=tcs_div / inflow))
    return pd.DataFrame(out)


results, buckets, tsviews = {}, {}, {}
for s in DRIVERS:
    r, b = run(s)
    results[s] = r
    buckets[s] = b
    tsviews[s] = tata_sons_view(r)

allres = pd.concat(results.values(), ignore_index=True)

# Probability-weighted expected case
exp = None
for s, p in PROBABILITY.items():
    d = results[s].set_index("FY")[["RevUSDm", "PATusdm", "DividendINRcr",
                                    "DivToTataSonsINRcr", "Headcount"]] * p
    exp = d if exp is None else exp + d
exp = exp.reset_index()

# ---- Revenue-at-risk quantification (FY26 base) ----
base_rev = HISTORY.iloc[-1]["RevUSDm"]
risk_rows = []
for name, w, auto, note in BUCKETS:
    risk_rows.append(dict(Bucket=name, FY26_Weight=w, FY26_RevUSDm=w * base_rev,
                          AutomatableShare=auto,
                          GrossRevAtRiskUSDm=w * base_rev * auto,
                          Note=note))
riskdf = pd.DataFrame(risk_rows)

if __name__ == "__main__":
    pd.set_option("display.width", 220, "display.max_columns", 40)
    for s in DRIVERS:
        print("\n" + "=" * 100 + f"\n{s}\n" + "=" * 100)
        print(results[s][["FY", "RevUSDm", "CCgrowth", "OPM", "PATusdm", "RevINRcr",
                          "PATINRcr", "DividendINRcr", "DPS", "DivToTataSonsINRcr",
                          "Headcount"]].round(1).to_string(index=False))
        print(tsviews[s].round(0).to_string(index=False))
    print("\nREVENUE AT RISK\n", riskdf[["Bucket", "FY26_RevUSDm", "AutomatableShare",
                                        "GrossRevAtRiskUSDm"]].round(0).to_string(index=False))
    print("\nTotal gross revenue in high-automation exposure: $%.1fbn of $%.1fbn"
          % (riskdf.GrossRevAtRiskUSDm.sum() / 1000, base_rev / 1000))
    print("\nPROBABILITY-WEIGHTED\n", exp.round(0).to_string(index=False))
