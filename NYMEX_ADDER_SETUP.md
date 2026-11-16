# Regular economics: NYMEX reference

Build TEXAS-NYMEX-04 adds **NYMEX + adder** to Contract pricing. It uses product H settlements for the selected trade date and delivery months. Select the model-required adder or enter a manual offer; regional supply costs still use the selected component curve. The H benchmark is not fuel-adjusted and is not added twice to supply costs.

The app also shows the usage-weighted H benchmark and equivalent modeled/offered NYMEX adders for fixed and regional-index quotes. Equivalent adders exclude separately passed-through transportation. Missing, duplicate or invalid H settlements block NYMEX-based quotes rather than substituting another date. Existing fixed/regional quotes can still be calculated without a valid H comparison.

Client and internal proposal packages now include three pricing worksheets: fixed, regional index and NYMEX, plus the summary page. Client worksheets exclude internal margin, NPV and return metrics.

Transportation pass-back is saved explicitly in workstream inputs and restored as either checked or unchecked. Switching pricing structures and refreshing pricing retain that choice. Save existing workstreams once to retain any corrected selection; past choices that were never saved cannot be recovered.

## Current pricing

The November 16, 2026 file was uploaded by the repository owner as `input/icecleared_2026_11_16.xlsx`; the existing archive workflow has processed it. Dated ICE files are ordered newest first. **Use newest available ICE file** retains customer assumptions, usage and manual offers while selecting the newest source. **Load / refresh selected pricing file** clears the cached parser. Review the displayed actual settlement trade date before quoting, then save the workstream to persist the new pricing reference.

This change does not add a scheduled cross-repository transfer or alter the owner's uploaded dated file. No new Streamlit Secrets are required.
