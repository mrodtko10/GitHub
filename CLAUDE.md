# Project Instructions

## VBA Dashboard Refresh Macro (.bas files)

When the user provides an HTML dashboard file and an Excel workbook and asks to "refresh", "update data", or create a macro — create a `.bas` VBA module that reads data from the Excel sheet and rewrites the JavaScript data constant in the HTML file.

### Pattern to follow

Reference implementations are in this repo:
- `PowerDashboard.bas` — reads "Finalized Projects" sheet → rewrites `const ALL_DATA`
- The `HTMLDashboard_7.bas` reference (uploaded by user) reads "Category Table" → rewrites `const RAW_DATA`

### Critical rules

1. **Row range — always extend well beyond current data.** Never use `UsedRange` or the current last row as the hard limit. Default to scanning **25,000 rows** (i.e., `For i = 2 To 25001`) so new data added in future months is automatically included. If the user specifies a different limit, use that.

2. **Skip truly blank rows** using a check on the key identifier column (e.g., project ID or cost code) AND an adjacent column (e.g., month), not just one field.

3. **Path resolution priority** (always implement all three):
   1. `HTML_PATH_OVERRIDE` Private Const at top of module (user pastes path here)
   2. `Dashboard` sheet cell B4 (or whichever sheet/cell the user specifies)
   3. Workbook folder + filename as fallback

4. **String building** — use `rec = rec & ...` on separate lines, never `& _` line continuations. VBA has a hard limit of 24 continuations per statement; splitting into multiple assignment lines avoids the "Too many line continuations" compile error.

5. **Three macros** to expose:
   - `Refresh*` — reads sheet, writes HTML only (safe, no browser)
   - `Open*` — opens HTML in default browser via `ThisWorkbook.FollowHyperlink`
   - `RefreshAndOpen*` — calls both in sequence

6. **Safety checks** before writing:
   - Path must end in `.html` or `.htm`
   - Path must not equal `ThisWorkbook.FullName`
   - File must exist (`Dir(htmlPath) <> ""`)

7. **JSON number formatting** — use the `J()` helper: `Replace(CStr(CDec(v)), ",", ".")` to avoid locale-specific decimal separators.

8. **Marker replacement** — find `const ALL_DATA = [` (or whatever the JS constant is named), then find the closing `];`, and replace everything between with the new JSON array.

### Deliverable

Produce a `.bas` file saved to the repo. Ask the user which sheet name and column mapping to use if not obvious from the files provided.
