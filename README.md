# Finance Service - Corporate Action - ROC Data Aggregator & Multi-Fund Loader Automation (VBA)

## 📌 Business Impact & Overview
Handling Return of Capital (ROC) accrual lifecycles from **Ex-Date to Pay-Date** typically generates thousands of granular trade bookings. In reconciliation platforms like **ATM**, these un-settled positions remain as **outstanding breaks**, creating massive system noise and operational overhead across multiple Brokers and Funds.

This Excel VBA pipeline automates the end-to-end data transformation by **aggregating 1,000+ scattered trade entries into clean, single-line Accrual Journals per Broker/Fund**. This eliminates noise in ATM break tracking, streamlines daily break rec carry, and simplifies final cash reconciliation on Pay-Date.

## 🔑 Key Features
1. **Multi-Fund Break Suppression & Aggregation:** Consolidates 1,000+ granular booking records into single-line Accrual Journals organized by Broker and Fund, replacing fragmented ATM breaks with clean tracking records.
2. **Lifecycle Tracking (Ex-Date $\rightarrow$ Pay-Date):** Maintains position carry integrity throughout the accrual waiting period, reducing daily monitoring overhead until cash settlement.
3. **Complex Text Pattern Parsing:** Uses dynamic string algorithms to extract unstructured tokens (e.g., ROC Terms, Payment Dates) from source feeds into normalized schema formats.
4. **Cross-Workbook Synchronization:** Automatically posts consolidated accruals into master accounting workbooks while preserving custom template line gaps for audit readiness.
5. **Dynamic GL Loader Generator (`SplitGLDataIntoXLSXFiles`):** Utilizes `Scripting.Dictionary` for fast $O(1)$ lookups to map general ledger entries and dynamically spawn parameterized `.xlsx` loaders per fund database.

## 🚀 Quantifiable Operational Impact
- **Noise Reduction:** Reduced thousands of messy ATM breaks down to **1 consolidated line per Broker/Fund**.
- **Efficiency Boost:** Cut daily manual reconciliation and GL file splitting time by **~80%**.
- **Risk Mitigation:** Eliminated manual copy-paste errors and lost micro-bookings during long carry periods prior to Pay-Date.
