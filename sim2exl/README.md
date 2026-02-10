# Excel Timeline Gantt (VBA)

![Timeline Gantt Screenshot](docs/screenshot.png)

**Overview**
This workbook builds a timeline (Gantt-like) chart in Excel from CSV history files. Each node is rendered on a single row, and state transitions are shown over time on a 1-second grid.

**Features**
- One row per node (sorted by `nodeOrder`)
- 1-second grid with 10-second tick marks
- Cell-based rendering for smooth scrolling and editing
- State color mapping:
  - `process` = green
  - `wait` = orange
  - `down` = blue
  - `idle` = yellow
  - other = gray
- CSV import picker for history files

**Requirements**
- Microsoft Excel (macro-enabled)
- `timeline.xlsm`

**CSV Format**
Expected header columns:
`node,nodeId,nodeOrder,start,end,duration,state,workId,agvId`

**How To Use**
1. Open `timeline.xlsm`.
2. Run `ImportTimelineCsvAndBuild` to select a CSV from `csv_data` and build the chart.
3. Alternatively, run `ImportTimelineCsv` and then `BuildGantt`.

**Key VBA File**
- `timeline/timeline.bas`

**Configuration**
You can adjust these constants in `timeline/timeline.bas`:
- `GRID_SEC` (grid size in seconds)
- `COL_WIDTH` (column width per second)
- `LABEL_STEP` (seconds between grid labels)

**Notes**
- Sheet names are fixed: `Data` and `Gantt`.
- If the Gantt header labels look crowded, increase `LABEL_STEP` or `COL_WIDTH`.
