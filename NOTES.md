# Notes

## Assumptions

- Blank spreadsheet cells are treated as empty values during normalization.
- Only truly empty cells are blank. Text such as `NA` or `NULL` is kept as data, so a surname like "Na" is not turned into an empty last name.
- Outside `LIST` values are counted exactly as stored; the assignment gives no separate LIST normalization rule.

## Rule Tradeoffs

1. Initial matching plus transitivity can merge different people. Input records `100064` (John Smith), `100141` (J Smith), and `100003` (Jane Smith) share `3 River Rd`, ZIP `37201`. The initial bridges John and Jane into one group. In production, I would send a group containing conflicting full first names for review instead of merging it automatically.

2. Nickname equivalence can merge separate household members. Input records `100125` (Bill Turner) and `100133` (Will Turner) share `70 Ash Pkwy`, ZIP `85001`; the nickname table maps both first names to William. In production, I would flag distinct full names that collapse to one canonical name for review.

The submitted program keeps both assignment rules unchanged.

## Testing

- Unit tests for name, ZIP, address, unit, suffix, nickname, and initial normalization/matching.
- End-to-end checks for rejection, suppression, transitive groups, priority and REC_ID tie-breaking, duplicate reasons, and `MULTI_COUNT`.
- Workbook checks for exact sheet names and columns, sorted unique rows, complete disposition of mail records, and `INPUT = REJECT + SUPPRESS + DUP_INTER + DUP_INTRA + NET` for every Summary row.
