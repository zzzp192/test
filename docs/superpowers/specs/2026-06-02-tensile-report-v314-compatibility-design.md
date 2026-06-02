# Tensile Report V3.14 Compatibility Design

## Goal

Upgrade the source to V3.14 so tensile report extraction and tensile Origin plotting accept both the existing legacy Excel export and the newer Excel export format.

## Scope

- Keep legacy Excel report extraction working.
- Add support for the newer `实验报告` summary sheet layout.
- Keep legacy tensile curve sheet selection working.
- Add support for newer `原始数据` tensile curve sheets.
- Update source version labels and README changelog to V3.14.
- Add focused regression tests.
- Do not package an EXE in this change.
- Do not add `样板数据/` to Git.

## Summary Sheet Detection

The extractor will scan worksheets for a header row containing a recognizable sample ID column. It will build a field-to-column map from normalized header text for:

- sample ID
- thickness
- Rp
- Rm
- Ag
- A
- At

The extractor will use the detected header row and column positions when the required columns are present. For backward compatibility, if dynamic detection is incomplete but a legacy `Sheet1` sheet exists, it will use the existing fixed legacy column positions.

## Curve Sheet Detection

The plotting path will choose a tensile curve worksheet by:

1. preferring names containing `曲线`;
2. accepting names containing `原始数据`;
3. validating that the selected sheet contains paired stress/strain columns;
4. falling back to the first paired-column worksheet if naming does not match.

## Error Handling

If no summary table can be recognized, extraction returns no groups and the existing report-generation error path remains active.

If no valid tensile curve sheet can be recognized, plotting raises a clear error instead of silently reading a photo sheet.

## Testing

Regression tests will cover:

- legacy summary extraction;
- newer summary extraction;
- legacy curve sheet selection;
- newer curve sheet selection;
- rejection of worksheets without paired stress/strain columns.

The sample workbooks remain local verification inputs and are excluded from Git.
