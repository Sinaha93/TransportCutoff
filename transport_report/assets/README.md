# Report runtime template

`report_template.pptx` is the unmodified, byte-for-byte application runtime
asset copied from the user-provided `26년 8월 운반비 보고.pptx`. It contains
the existing report's business content and is not a test fixture.

SHA-256: `7bb6269b6d3834e9eb4f7b43aa32d17aa627861d2352a29b5a7db9d7166a9cce`

`app.reporting.pptx_report.prepare_template(source, output)` prepares a
separate named copy. `generate_pptx` can also consume this exact asset and
performs the same naming in memory. Only the recorded SHA can enter the
legacy naming path; other templates must already have the validated semantic
roles. Source decks and spreadsheet inputs must never be output targets.

The fixed adapter has 14 monthly slots (12 destinations, a derived group,
miscellaneous), 12 next-month destination slots, and five selected review
slots. Use `None` for vacant display slots. Supply canonical Task6 results,
including independently calculated grand totals with the configured inclusion
flags, and Task10 chart histories. The plan must be for September 2026.
Row history mappings use raw EA and won. Money appears in thousand won,
unit costs in won/EA, and display rounding is half-up. Missing values remain
blank. Review details are explicitly selected from rows meeting the inclusive
15% unit-cost variance criterion; more than five selections fail.

Original picture crops are retained. New chart pixels are fitted inside their
visible crop windows at the original pixel dimensions so the data tables are
not clipped. Picture IDs, placement and z-order are preserved.

Layout verification: from `transport_report`, run
`python -m tools.verify_ppt_layout TEMPLATE.pptx OUTPUT.pptx`.
Automated tests build a fresh synthetic six-slide template with fictional
destinations and never load this runtime asset. Private PowerPoint render
artifacts are not committed.
