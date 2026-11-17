# Production Cost Review & Scenario Model

An anonymised portfolio case study showing how an unstructured production budget can be converted into a decision-ready cost model.

The model supports four business questions:

1. Where is production spend concentrated?
2. Which cost categories differ most from external benchmarks?
3. What are the unit economics per episode, shoot day and finished minute?
4. How would the budget change under different episode counts, shoot days and contingency assumptions?

## Privacy-first public version

This repository uses **synthetic demonstration data**. The original production workbook, project title, client identities, personal information, contracts and real financial values remain private and are not committed.

The public version preserves the analytical workflow, spreadsheet architecture, formulas, controls and scenario logic used in the completed project.

## Deliverable

Download the generated Excel model: [`output/Production_Cost_Review_Demo.xlsx`](output/Production_Cost_Review_Demo.xlsx)

The workbook includes:

- **Summary** - spend by category, mix comparison and unit economics
- **Scenario** - adjustable episode, shoot-day and contingency assumptions
- **People & Paperwork** - workforce-cost and documentation tracking
- **Line Items** - clean analytical source table with formula checks
- **Data Quality** - exceptions requiring review
- **Benchmarks** - external assumptions and source links

## Analytical workflow

```text
Source budget -> clean and classify -> validate totals -> cost model -> scenario analysis -> management review
```

The model separates input cells, formulas and linked cells using consistent spreadsheet formatting. Scenario outputs are formula-driven rather than hard-coded.

## Public sample structure

| File | Purpose |
|---|---|
| `data/sample_line_items.csv` | Synthetic line-item data with the same analytical schema |
| `data/data_quality.csv` | Illustrative validation exceptions |
| `src/build_model.py` | Builds the multi-sheet Excel model with formulas and formatting |
| `output/Production_Cost_Review_Demo.xlsx` | Ready-to-review demonstration workbook |

## Run locally

```bash
pip install -r requirements.txt
python src/build_model.py
```

## Skills demonstrated

- Business requirements translated into a reusable financial model
- Data cleaning, classification and quality-control design
- Cost-driver and unit-economics analysis
- Scenario planning and operational decision support
- Python, pandas, openpyxl and advanced Excel formulas

## Scope note

External benchmark ranges are planning references, not universal production standards. Actual production costs depend on location, labour agreements, creative scope, scheduling, insurance and distribution requirements.

