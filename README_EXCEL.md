# Excel Financial Model Generator

This directory contains tools to generate a banker-style Excel financial model for renewable energy projects.

## Files

- **`build_excel.py`**: Python script that generates `excel_template.xlsx` with all formulas
- **`excel_template.xlsx`**: Generated Excel workbook with 8 tabs and all formulas (created by running `build_excel.py`)

## How to Regenerate the Template

1. **Install dependencies** (if not already installed):
   ```bash
   pip install openpyxl
   ```

2. **Run the build script**:
   ```bash
   python build_excel.py
   ```

3. **Open the generated file**:
   - The script creates `excel_template.xlsx` in the current directory
   - Open it in Excel 365 (or Excel 2016+ with XIRR support)

## Model Structure

The Excel model has 8 tabs:

1. **Inputs**: All assumptions (timeline, generation, revenue, CAPEX, OPEX, debt, tax)
   - All values are hardcoded here
   - Named ranges are created for easy reference

2. **Timeline_M**: Monthly timeline from project start to end of operation
   - Month-start dates
   - Year and Phase (Development/Construction/Operation) helpers

3. **CAPEX_M**: Monthly CAPEX schedule
   - Development CAPEX spread over development months
   - Construction CAPEX spread over construction months
   - At COD CAPEX booked once at COD

4. **Revenue_M**: Monthly revenue during operation
   - Energy production with degradation
   - Price with indexation (PPA vs. merchant)
   - Revenue = Energy × Price × 1000

5. **OPEX_M**: Monthly OPEX during operation
   - Fixed OM, Variable OM, Insurance, Grid Fees
   - All indexed by inflation

6. **Debt_M**: Monthly debt schedule
   - Debt draws during construction (pro-rata with CAPEX)
   - Debt service (interest + principal) during operation
   - Debt fees (upfront + commitment)

7. **Cashflow_M**: Monthly cash flows
   - Unlevered CF = Revenue - OPEX - CAPEX
   - Equity CF = Unlevered CF + Debt Draw - Debt Service - Debt Fees

8. **Outputs**: Summary metrics and checks
   - Total CAPEX, Debt, Equity
   - Unlevered IRR (Pre-tax) using XIRR
   - Equity IRR (After-tax) using XIRR
   - Checks: Sources=Uses, Debt draw after COD=0, Debt balance >=0, CAPEX total match

## Key Formulas

- **Debt Draw**: `=IF(AND(C2<>"Operation",DebtEnabled="Yes"),CAPEX_M!G2*DebtPct/100,0)`
- **Equity Injection**: `=CAPEX_M!G2*(1-DebtPct/100)` (implicit in Equity CF calculation)
- **Unlevered CF**: `=Revenue - OPEX - CAPEX`
- **Equity CF**: `=UnleveredCF + DebtDraw - DebtService - DebtFees`
- **IRRs**: `=XIRR(values_range, dates_range)`

## Notes

- All formulas reference named ranges from the Inputs tab
- The model supports up to 500 months (~41 years)
- Formulas use helper columns for readability
- Checks section provides red/green flags for model integrity
- Requires Excel 365 or Excel 2016+ with XIRR function support

## Customization

To modify the model:
1. Edit `build_excel.py` to change formulas or structure
2. Run `python build_excel.py` to regenerate `excel_template.xlsx`
3. Open the new template in Excel

## Dependencies

- Python 3.7+
- openpyxl 3.0+

