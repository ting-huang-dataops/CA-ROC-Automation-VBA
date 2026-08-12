Finance Service - Corporate Action - ROC Data Aggregator & Multi-Fund Loader Automation (VBA)

📌 Business Impact & Overview
Handling Return of Capital (ROC) accrual lifecycles from **Ex-Date through Pay-Date** typically generates trade bookings for thousands of lots. In reconciliation platforms, these unsettled positions remain as **outstanding breaks**, creating massive system noise and operational overhead across multiple Brokers and Funds.

This Excel VBA suite solves both challenges by establishing a two-tier operational automated pipeline:

1. **Global Oversight & Month-End Audit Readiness:** Extracts and consolidates all raw ROC entries into a centralized master workbook, providing an easily inspectable source of truth for auditing all unpaid ROC balances at month-end.
2. **Break Reduction & Accrual Journal Generation:** Automatically transforms and posts structured Accrual Journals into downstream accounting systems. By aggregating fragmented bookings per Broker/Fund, it minimizes cash breaks and streamlines ROC booking carry until final pay date settlement.

## 🔑 Key Features
- **Month-End Unpaid ROC Tracking:** Consolidates multi-source ROC data into a master staging workbook for streamlined month-end review and pending cash verification.
- **Accrual Journal Automation:** Auto-generates clean, single-line journal entries to suppress unnecessary system breaks in reconciliation platforms.
- **Multi-database Journal Loader Splitting:** Leverages `Scripting.Dictionary` for $O(1)$ dynamic mapping to generate parameterized XML/XLSX loader files per fund database.
