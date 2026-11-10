# DJI Enterprise management price list

Internal Excel draft for Rectron. Open `output/DJI_Enterprise_MPL_DRAFT.xlsx` in Excel and begin with **Read Me**, then **Settings** and **MPL Internal**.

The workbook contains confidential purchasing costs, margins, sales and customer records. Do not distribute it as a dealer price list.

## Draft pricing

All internal selling prices exclude VAT. Rectron gross-margin baselines are S 10%, A 15%, B 20%, OTHER 25%; selling price = selected cost / (1 - margin). Public pricing uses supplier RRP where available, otherwise a provisional fallback. Exchange rates, handling allowance and fallback margin require review.

The workbook embeds the supplied source reports. Automated refresh and dealer-safe export are not implemented. Review missing mappings, source duplicates and combo allocations before operational use.

## Builder

`tools/build_draft.py` records how this draft was generated and validates its initial pricing calculations. It requires Python 3 and `openpyxl`.

The current builder uses the original upload paths under `/workspace/attachments`, including the corrected CIP upload. Those external input files are not committed separately; the workbook includes their imported data. A fresh checkout alone cannot rerun the builder until those input paths are supplied or adapted. Running the builder replaces the output workbook and does not retain workbook edits. Preserve edited copies first.

This initial draft has structural and cached pricing checks; interactive recalculation and presentation have not been verified in Microsoft Excel.
