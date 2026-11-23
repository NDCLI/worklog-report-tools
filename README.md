# Report Tools

Desktop application (PyQt6) for aggregating annotation productivity, inspection/review metrics, error rates (Frame & BBOX), First Inspector Recall, rankings, and cross-file auditing from `Work_log` (including Excel Notes) and `Inspection_And_Rework` workbooks.

## Features
- **Task**: Per-task Worker productivity (`Work Load`, `Work BBOX`, `Work Time`, `Err Frame`, `Err BBOX`, `FRAME/Hour`, `BBOX/Hour`, `Err Rate Frame/BBOX (%)`) with separated Task-level summary columns.
- **Annotator**: Aggregated totals across all tasks per Worker (A–Z).
- **Inspector**: First-level inspection report grouped by `Inspector → Annotator`, with `FRAME/Hour`, `BBOX/Hour`, `Err Rate Frame/BBOX (%)`, and `1st Inspector Recall (%)`. Search and select Inspectors in the sidebar to filter this report and its Excel sheet.
- **Bảng Xếp Hạng**: Side-by-side lowest error rate rankings for Workers and first-level Inspectors.
- **Kiểm tra dữ liệu**: A compact summary for Annotators and Inspectors. Select a person and task in the summary to see the dates, job IDs, frame ranges, and reasons for that group's issues.
- **Excel Export**: Exports formatted multi-sheet `.xlsx` reports with live Excel formulas.

## Run from Source
```bash
pip install -r requirements.txt
python main.py
```

## Build Standalone Executable (Windows)
```bash
python -m PyInstaller --noconfirm --onefile --windowed --name ReportTools --icon app_icon.ico --add-data "app_icon.ico;." main.py
```
