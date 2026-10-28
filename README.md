# Partlist Validation UI (rule-based; no AI model)

## Files
- `validation_app.py`: Streamlit upload, dashboard, failed-cell editor, full-table editor, and downloads.
- `src/excel_tools.py`: uploaded latest workbook-reading utilities.
- `src/rule_checker.py`: latest rule-checking design with UI-focused fixes (all duplicate occurrences, SILS exemption for blank AU, conditional applicability checks, whole-number checks, and missing `math` import fixed).
- `Stated_Rule_List.xlsx`: bundled 31-rule workbook.
- `tests/test_rule_checker.py`: tests for the current rule-checker API.

## Run
Create/activate a virtual environment, then:

```bash
pip install -r requirements.txt
python -m streamlit run validation_app.py
```

Upload the partlist workbook. The bundled `Stated_Rule_List.xlsx` is used unless another rules workbook is uploaded.

## Notes
- The rules workbook is designed for the `ZaragozaTH,PF,BS` partlist layout. The sheet selector is available, but selecting a different sheet does not make the column-letter rules sheet-specific.
- The UI does not auto-correct data with AI. It lets the user edit values, writes those values into the workbook, then reruns deterministic validation.
- `REVIEW` means the checker cannot safely decide a conditional rule with current information; it is not treated as a pass.
