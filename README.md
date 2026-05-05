# Clinic Scoreboard Converter

Converts `Scoreboard Test.xlsx` into queryable JSON without losing columns or header context.

## Run

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python convert.py
# → output.json
deactivate # To restore your shell's PATH back to normal when done
```

## JSON shape

Each sheet becomes an object with four keys: `metadata`, `rows`, `columns`, `metric_metadata`.

```json
{
  "SCOREBOARD": {
    "metadata": {
      "sheet_name": "SCOREBOARD",
      "header_depth": 7,
      "raw_column_count": 142,
      "integrity_ratio": 1.0,
      "grouped_headers": [
        {
          "header": "PHONE PERFORMANCE",
          "excel_range": "AJ1:AP1",
          "column_count": 7
        }
      ]
    },
    "rows": [
      {
        "_source_row": 8,
        "date": "2026-02-16T00:00:00",
        "metrics": {
          "Total Revenue - All Services": 40454.28,
          "Answer Rate": 0.87
        }
      }
    ],
    "columns": [
      {
        "field_name": "Total_Revenue_All_Services",
        "display_name": "Total Revenue - All Services",
        "header_path": [
          "Total Revenue - All Services",
          "focus=Financial",
          "source=EMR",
          "role=J"
        ],
        "source_column_index": 1,
        "metric_label": "Total Revenue - All Services",
        "category": null,
        "focus": "Financial",
        "source": "EMR",
        "role": "J",
        "qualifiers": []
      }
    ],
    "metric_metadata": {
      "Total Revenue - All Services": {
        "focus": "Financial",
        "source": "EMR",
        "role": "J"
      },
      "Answer Rate": {
        "focus": "Phone",
        "source": "Phone",
        "role": "J",
        "category": "PHONE PERFORMANCE"
      }
    }
  }
}
```

**Why this shape:**

- `rows` are date-centric so dashboards and LLMs can slice by date immediately — no pivot needed.
- `metric_metadata` is hoisted to sheet level (not repeated on every row) to keep row payloads small and avoid redundancy.
- `columns` retains the full structural record (header path, source index) for anything that needs to reconstruct or re-map the original layout.
- `metadata.integrity_ratio` gives a quick sanity check: `converted_non_empty / source_non_empty`. A ratio of 1.0 means no data was silently dropped.

## How the messy bits were handled

| Problem                                              | Decision                                                                                                                                                                                                                  | Trade-off                                                                                                                                        |
| ---------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| **7-row merged header stack**                        | Auto-detect header depth by scanning for the first row where ≥40% of values are numeric — that's where data starts.                                                                                                       | Could misfire on sheets where the first data row is mostly text.                                                                                 |
| **Merged cells (e.g. `AJ1:AP1`)**                    | Use `ws.merged_cells.ranges` to find the exact column span, assign `category` only to columns within that range.                                                                                                          | Categories only captured for horizontally-merged top-row labels; vertical merges in data rows are not specially handled.                         |
| **Keyed header rows (`Focus:`, `Source:`, `Role:`)** | Treat as structured key=value path parts rather than literal column names, then promote to first-class fields.                                                                                                            | Relies on the colon-suffix convention; breaks if a header row uses a different separator.                                                        |
| **Duplicate column names after forward-fill**        | Generate unique labels with `_2`, `_3` suffixes as a fallback. True duplicates are uncommon in this sheet; this keeps things stable without silently dropping columns.                                                    | Labels are stable but not semantically meaningful for true duplicates.                                                                           |
| **Dead null blocks / placeholder columns**           | Drop columns that are empty across all non-spacer data rows, even if they have header metadata. This removes long stretches of useless `null` metrics between real measures like `% Tx Plan Used` and `PT Total Revenue`. | Fully empty columns are no longer represented in `rows.metrics` or `columns`, so layout-level placeholders are not preserved in the main output. |
| **Spacer rows / empty rows**                         | Rows where all cells are empty or contain only decoration characters (`---`, `===`) are dropped and counted in `dropped_spacer_rows`.                                                                                     | A row that is genuinely sparse (a few non-null values) is kept, not dropped.                                                                     |
| **Formula cells**                                    | `openpyxl` is opened with `data_only=True`, which reads the last cached value Excel wrote.                                                                                                                                | If the file was never saved after recalculation, formula cells will read as `null`.                                                              |
| **Newlines in header text**                          | Collapsed to a single space during header normalization.                                                                                                                                                                  | Purely cosmetic; no data impact.                                                                                                                 |

## What I'd do with two more hours

- **Multi-sheet support tested end-to-end** — the converter already iterates all sheets, but the output contract assumes each sheet has the same 7-row header structure. A second sheet with a different layout would need its own detection tuning.
- **Inactive column audit trail** — keep the current behavior of dropping all-empty columns from the main output, but also emit an `inactive_columns` section for debugging and reconciliation with the original sheet layout.
- **CLI flags** — `--sheet`, `--out`, `--no-spacer-drop` so the script is usable without editing source.
- **Round-trip test** — a pytest fixture that reads a known row from the xlsx and asserts the JSON output matches, so future changes to the converter can't silently regress integrity.
- **Warnings appear if the ratio drops below `0.9`** - This helps catch accidental column or row loss during future refactors.
