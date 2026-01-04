# Delphi Financial Models - Streamlit App

A comprehensive financial modeling application for utility-scale renewable energy projects, built with Streamlit.

## Features

- **Multi-Scenario Analysis**: Create and compare multiple project scenarios
- **Complete Financial Modeling**: 
  - Power generation with degradation
  - Revenue modeling (PPA and custom pricing)
  - CAPEX, OPEX, and SG&A management
  - Depreciation schedules
  - Renewable tax benefits (Colombia)
  - Debt & covenants analysis
  - Unlevered and levered cash flow calculations
  - IRR and payback period calculations
- **Sensitivity Analysis**: Two-variable sensitivity analysis with heatmaps
- **Summary Dashboard**: Comprehensive project summary with key metrics

## Installation

1. Clone this repository:
```bash
git clone <your-repo-url>
cd app
```

2. Create a virtual environment:
```bash
python -m venv .venv
```

3. Activate the virtual environment:
- Windows: `.venv\Scripts\activate`
- Mac/Linux: `source .venv/bin/activate`

4. Install dependencies:
```bash
pip install -r requirements.txt
```

## Running Locally

```bash
streamlit run appv7.py
```

The app will open in your browser at `http://localhost:8501`

## Deployment to Streamlit Cloud

1. Push your code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Sign in with your GitHub account
4. Click "New app"
5. Select your repository and branch
6. Set the main file path to: `appv7.py`
7. Click "Deploy"

## Project Structure

- `appv7.py` - Main application file
- `requirements.txt` - Python dependencies
- `delphi_projects.json` - Project data (created automatically, not in git)

## Usage

1. **Create a Project**: Start by creating a new project in the sidebar
2. **Configure Scenario**: Set up macroeconomic assumptions, timeline, generation, revenues, costs, etc.
3. **Review Results**: Navigate through tabs to see:
   - Power generation and revenues
   - CAPEX breakdown
   - Operating costs
   - Tax benefits
   - Cash flow analysis
   - Debt structure
   - Sensitivity analysis
   - Summary dashboard

## Notes

- All financial values are in Colombian Pesos (COP) by default
- USD conversion uses FX path with CPI indexation
- Data is saved locally in `delphi_projects.json`
- The app supports multiple projects and scenarios

## Support

For questions or issues, please contact the development team.

