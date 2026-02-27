# Alb-Stats

Local Dash app that reads your Excel export and shows Plotly statistics on `localhost`.

## Features
- Uses `Couleurname` as the only shown name field.
- Average age:
  - all members
  - `UP` + `BP` + `EM` (`EM` = Ehrenmitglied, counts with UP/BP)
  - `BU` + `FU` (`BU` = Bursch, `FU` = Fuchs)
  - flexible status selection
- Chargen statistics:
  - total chargen per person
  - chargen per person for a selected semester
  - aktivenchargen vs philisterchargen vs unklare chargen (based on semester vs philistrierung date)
- Additional statistics:
  - median age
  - age distribution bins
  - status counts + percentages
  - chargen intensity per person (total, average per year from `Reception` to today, top percentile cutoff)
  - average chargen per person
- Top 20 Bundesbrueder table by total chargen
- Upload a new `.xlsx` directly in the web UI
- Export filtered table rows to CSV or Excel
- Manual override UI for chargen classification (Aktiven/Philister/Verband/Funktionaere), saved locally in `chargen_class_overrides.json` using date-free keys
- Selectable category graph for: `Aktivenchargen`, `Philisterchargen`, `Verbandschargen (Aktiven)`, `Verbandschargen (Philister)`, `Funktionaere`
- Multi-select chargen-role graph to show top people for selected chargen

## Setup
```powershell
python -m pip install -r requirements.txt
```

## Run
```powershell
python app.py --file Datenexport_20260227_1617.xlsx
```

Then open:
- `http://127.0.0.1:8050`

## Optional data check (without starting server)
```powershell
python app.py --file Datenexport_20260227_1617.xlsx --check-only
```

## Tests
```powershell
pytest -q
```
