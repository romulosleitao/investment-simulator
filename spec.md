# 📊 Specification & Business Rules: Investment Simulator (Excel)

## 1. Objective & Overview
Create a professional Excel tool to simulate Real Estate Investment Funds (FIIs) portfolios, tracking monthly investments, asset accumulation, projected compound returns, and dividend distributions based on specific investor profiles.

---

## 2. Workbook Structure (Sheets)

### Sheet 1: `APP` (Dashboard & Interactive Interface)
* **Purpose:** Main user control panel, input variables, scenario projections, and asset allocation breakdown.
* **Key Sections & Layout:**
  * **Global/Configuration Variables:**
    * Salary / Base Income (`B10`)
    * Portfolio Monthly Yield Rate / Dividend Yield (`B11`, e.g., `0.6%` or `0.006`)
    * Suggested Investment Rule (`B12`, e.g., 30% of salary)
  * **Monthly Investment Simulator (Inputs):**
    * Monthly Contribution Amount (`B15`)
    * Horizon in Years (`B16`)
    * Monthly Interest Rate / Yield (`B17`)
    * *Outputs (Formulas):* 
      * Accumulated Capital (`B18`) -> Compound interest formula on monthly deposits.
      * Monthly Dividends (`B19`) -> Accumulated Capital * Yield Rate.
  * **Scenario Simulator Table (Time Horizons):**
    * Columns: Years (`2`, `5`, `10`, `20`, `30`), Accumulated Value, Projected Monthly Dividends.
  * **Profile Asset Allocation (FII Types):**
    * Profile Selector (`B30`, values: `Conservador`, `Moderado`, `Agressivo`).
    * Dynamic breakdown table pulling weights from the reference data based on the selected profile and calculating individual fund type budget allocations.

### Sheet 2: `Database_Profiles` (Backend Reference Data)
* **Purpose:** Stores the relational distribution rules mapping profiles to FII asset classes.
* **Columns Structure:**
  * `CHAVE` (Concatenation of Profile + FII Type, e.g., `Moderado-PAPEL`)
  * `PERFIL` (Conservador, Moderado, Agressivo)
  * `TIPO DE FII` (PAPEL, TIJOLO, HÍBRIDOS, FOFs, DESENVOLVIMENTO, HOTELARIAS)
  * `%` (Target allocation weight, summing to 1.0 per profile)

---

## 3. Business Rules & Formulas

1. **Compound Interest & Growth Model:**
   * The accumulated capital formula for regular monthly investments must use the future value formula considering monthly compounding.
2. **Dynamic Asset Allocation:**
   * The allocation table on the `APP` sheet must dynamically filter or look up weights from `Database_Profiles` using `XLOOKUP` or `FILTER` based on the active profile selected in cell `B30`.
3. **Visual Uniformity & Design Standards:**
   * Clean financial corporate layout (palette: Dark headers, soft gray gridlines, formatted currency `R$` and percentage `%` cells).