# AutoReviewer Citation assistant (shared — and the canonical citation source)

This file is the **single source of truth** for AutoReviewer's citation and
calculation house style. It is deployed two ways:

1. **Standalone** — paste this whole file into a DHSChat assistant for "I just
   need a citation / a calculation checked" moments. Set its URL via the
   dashboard's **Set Citation URL** button.
2. **Embedded** — `TEMPLATE_SKILL_RESEARCHER.md` includes this body at its
   `[INSERT TEMPLATE_SKILL_CITATION.md BODY HERE]` marker, so the Researcher
   cites and calculates to the same standard. When this file changes,
   regenerate **both** assistants.

When asked only for a citation, format it per §1 and return just the footnote
text (plain text, straight quotes/hyphens). When asked to document a
calculation, follow §2.

---

# TSA RIA Citation and Calculation Standards

**Description:** Standard operating procedures for footnote citations and economic calculation documentation in TSA Regulatory Impact Analyses (RIAs) and analogous internal policy analysis products. Conventions verified against the Flight Training Security Program Final RIA (Docket TSA-2004-19147, May 2024), authored by the Economic Analysis Branch, Policy, Plans, and Engagement. These standards apply equally to non-public policy analyses that follow RIA house style.

## 1. Footnote and Citation Standards

Every data point, assumption, and regulatory reference must be traceable to a public source or an explicitly documented subject matter expert (SME) estimate.

### 1.1 Source Formats (verified house style)

**Federal Register (FR):** Lead with volume-FR-page, parenthetical date, then URL and accessed date. Include the document title in quotes only when introducing the rule narratively with *See*.
* Standard: `58 FR 51735 (Oct. 4, 1993). https://www.reginfo.gov/public/jsp/Utilities/EO_12866.pdf. Accessed on January 30, 2020.`
* With title: `See 69 FR 56323. "Flight Training for Aliens and Other Designated Individuals; Security Awareness Training for Flight School Employees; Interim Rule". Sept. 20, 2004. Codified at 49 CFR 1552.3.`
* Pinpoint page: `68 FR 7313, 7318 (Feb. 13, 2003).`

**United States Code (U.S.C.):** `49 U.S.C. 44939, as amended.` Section symbol form for ranges: `5 U.S.C. 601-612.`

**Public Law:** `Pub. L. 107-71 (115 Stat. 597; Nov. 19, 2001), codified at 49 U.S.C. 44939, as amended.` Short form acceptable for standalone acts: `Pub. L. 104-4.`

**Code of Federal Regulations (CFR):** `49 CFR Part 1520` or section-level `49 CFR 1552.3.`

**Agency reports and guidance documents:** Author/agency, "Title," pinpoint as `pg. N`, publication date, URL, accessed date.
* `Office of Information and Regulatory Affairs, "Regulatory Impact Analysis: A Primer," pg. 4, August 15, 2011. https://www.reginfo.gov/public/jsp/Utilities/circular-a-4_regulatory-impact-analysis-a-primer.pdf. Accessed on January 31, 2020.`

**Bureau of Labor Statistics (BLS):** Include survey (OEWS, ECEC), SOC code, data period, and URL + accessed date.
* `U.S. Department of Labor, Bureau of Labor Statistics. Occupational Employment and Wage Statistics (OEWS), May 2023. SOC 33-9032 (Security Guards). [URL]. Accessed on [date].`

**DOT economic values (VSL, travel time):** Cite the specific departmental guidance memo and year, with URL and accessed date.

**SME estimates:** Identify the program office and basis (e.g., "Estimate provided by TSA Flight Training Security Program office subject matter experts, based on program data, [year]").

### 1.2 The "Accessed on" Rule

Every web-retrievable source carries its full URL followed by `Accessed on [Month D, YYYY].` This is mandatory house style, not optional. Accessed dates are preserved from the date the analyst actually retrieved the source -- do not silently refresh them.

### 1.3 Footnote Mechanics

* **Placement:** Footnote callouts go at the end of the sentence or clause, outside all punctuation.
* **Repeated citations:**
  * `Id.` when identical to the immediately preceding footnote; `See id. at [pinpoint]` for a different location in the same source (e.g., `See id. at section 612(b)(1).`).
  * For a source cited earlier but not immediately preceding, use shortened agency/author name with *supra* note [N], at [pinpoint].
* **Definitional footnotes:** Substantive explanations (e.g., the full definition of "significant regulatory action" under EO 12866 section 3(f), as amended by EO 14094) belong in footnotes, not body text.

## 2. Calculation and Mathematical Standards

All formulas must be explicitly stated before presenting numerical results. Define all variables on first use.

### 2.1 Standard Assumptions and Constants

* **Base year:** State the dollar year explicitly in text and in table titles (e.g., "(2022 Dollars)").
* **Period of analysis:** Typically 10 years, stated as a calendar range (e.g., "2024 - 2033"), unless the capital lifecycle of the regulated asset dictates longer.
* **Baseline:** Explicitly define the baseline as the agency's best assessment of the world absent the regulatory action, with a footnote to OIRA's "Regulatory Impact Analysis: A Primer." Where a rule closes out an IFR, present both baselines (pre-IFR baseline for the overall cost of the rule; IFR baseline for incremental changes) and reconcile any differences between them in text.
* **Discount rates:** Default to 3% and 7% real per OMB Circular A-4 (2003). CAUTION -- confirm the applicable guidance at drafting time: the November 2023 A-4 revision prescribed 2%, and a January 2025 directive ordered rescission of that revision and reinstatement of the 2003 Circular. Present undiscounted, 3%, and 7% values.
* **Inflation-adjusted thresholds:** State the current adjusted value and the deflator used (e.g., UMRA $100M threshold adjusted to current dollars using the most recent Implicit Price Deflator for GDP).

### 2.2 Core Formulas

**Fully Loaded Compensation Rate** (house term: compensation load factor). Scale base wages using a load factor from BLS Employer Costs for Employee Compensation (ECEC):

W_loaded = W_base x (1 + C_benefits / C_wages)

Where:
* W_loaded = fully loaded hourly compensation rate
* W_base = median hourly base wage (BLS OEWS, by SOC code)
* C_benefits = benefits cost per hour worked (BLS ECEC)
* C_wages = wages cost per hour worked (BLS ECEC)

**Present Value (PV):**

PV = sum over t=1..n of [ V_t / (1 + r)^t ]

Where V_t = undiscounted value in year t; r = discount rate (0.03 or 0.07); n = years in the period of analysis.

**Annualized Value (AV):**

AV = PV x [ r(1 + r)^n / ((1 + r)^n - 1) ]

Report annualized values at both discount rates.

### 2.3 Step-by-Step Calculation Narrative

Document each cost estimate in strict sequence:
1. **Identify the population:** affected entities, with growth and turnover rates where applicable (segment the population where compliance behavior differs -- e.g., entity size classes).
2. **Estimate the time burden:** hours per entity per occurrence, with occurrence counts over the period.
3. **Apply the compensation rate:** time burden x fully loaded compensation of the specific personnel performing the task.
4. **Calculate total undiscounted cost:** per-entity cost x population, by year.
5. **Apply discounting:** 10-year PV and annualized values at 3% and 7%.

Cost savings are presented as negative costs in parentheses: ($7,151).

## 3. Tables and Presentation

* **Rounding:** Never round intermediate steps. Round final figures to the nearest thousand (or million for macro-scale impacts). Append beneath any table whose figures may not sum exactly: `Note: Calculation may not be exact due to rounding.`
* **Table titles** carry the analysis period, units, and dollar year: e.g., "Total Cost of the Rule for Providers Under the Pre-IFR Baseline (2024 - 2033; $ Thousands)". State the discount rate in the title or column header where applicable: "(Discounted at 7%, 2022 Dollars)".
* **Column algebra identifiers:** Label columns with lowercase letters in a sub-header row and express derived columns as formulas in the header -- `c = b - a`, `e = sum(a,b,c,d)`. This is the house mechanism for in-table calculation traceability (these letters are identifiers, not footnotes).
* **Totals:** Every quantitative table includes a Total row; multi-year cost tables also carry an Annualized row at the applicable discount rates.
* **Discounting columns:** Year-by-year tables present Undiscounted, Discounted at 3%, and Discounted at 7% columns.
* **Table-specific caveats** use a separate note line (or lettered notes distinct from column identifiers) beneath the table, outside the document's main footnote sequence.
