# Automated Financial Data Pipeline & Multi-Fund Loader Generator (VBA)

## Project Overview
An end-to-end Excel VBA automation suite designed for Hedge Fund Administration operations. This tool automates Return of Capital (ROC) data extraction, multi-layer exception filtering, complex string parsing, and automated GL file generation split by fund databases.

## Key Features
1. **Rule-Based Data Filtering (`FilteroutROC`):** Auto-filters operational checks, applies custom inclusion/exclusion criteria, and extracts unique tracking identifiers.
2. **Data Parsing & Enrichment (`CreateROCDATASheet`):** Reconstructs trade and position metrics, dynamically parsing complex text tokens (e.g., ROC Terms and Payment Dates) into structured numeric and date formats.
3. **Cross-Workbook Sync (`ExportROCDATASheet`):** Seamlessly bridges data across core accounting workbooks while maintaining required structural line gaps for audit/reporting formats.
4. **Dynamic Database File Splitter (`SplitGLDataIntoXLSXFiles`):** Utilizes `Scripting.Dictionary` to map general ledger entries against a master fund lookup table, dynamically spawning and saving parameterized XML-based `.xlsx` loaders for downstream database ingestion.

## Tech Stack & Concepts
- **VBA Core:** Advanced Sheet Object handling, `Scripting.Dictionary` for $O(1)$ lookups, Dynamic Range manipulation, File I/O automation.
- **Operations Focus:** Financial Data Cleansing, Accrual Reporting, Automated GL Loader Generation, Operational Risk Reduction.

## Impact
- Reduced daily manual processing time by **~80%**.
- Eliminated human copy-paste errors across multi-fund ledger splits.
