# Code Review Notes

## Overall impression
The app is a strong single-file prototype: it has clear UX stages, useful visual feedback, and a practical local-only processing model for sensitive data.

## Strengths
- Thoughtful UX flow with progressive tabs for segmentation, filtering, prioritization, and assignment.
- Good baseline security posture for static hosting with a CSP and local-only data processing.
- Consistent escaping (`esc`) before injecting user/data values into HTML.
- Metrics and preview are recalculated centrally through `updateAll()`, reducing UI drift.

## Main risks / improvements
1. **Single-file architecture**: HTML/CSS/JS in one file is hard to test and evolve. Split into modules (`state`, `filters`, `charts`, `ui`).
2. **Performance bottlenecks on large datasets**: frequent full-array filtering and full chart rebuilds can lag. Consider memoization, incremental updates, or Web Worker processing.
3. **Date parsing ambiguity**: `parseDate` falls back to `new Date(str)` which is locale/browser dependent.
4. **Magic constants and localized coupling**: filter config relies on exact Russian column names; add mapping/normalization layer.
5. **State machine clarity**: transitions (`filtersApplied`, `candidateRows`, `finalRows`) are implicit; formalize as explicit app phases.

## Suggested next technical step
Create a small `app.js` module boundary first (without behavior changes), then add a tiny test harness for `matchFilter`, `parseDate`, and `sortByCriteria`.
