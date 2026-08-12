# Corporate Actions Automation Suite (VBA)

An end-to-end VBA-driven automation framework designed to streamline **Corporate Actions (CA)** operations

## Core Modules Overview

### Module 1: Return of Capital (ROC) Automation
* **Business Challenge:** Manual collation of daily ROC events across multiple sub-ledgers created significant overhead during month-end closing and introduced frequent daily cash breaks due to timing mismatches.
* **Solution:** 
  * Automates the batch pull and consolidation of daily ROC transactional data into a standardized Master Workbook.
  * Automatically calculates and formats **Accrual & Cash Journal Entries** to suppress daily cash breaks.
  * Streamlines month-end reconciliation, providing a clean audit trail for finance and fund accounting teams.

### Module 2: GTL Trade Loader & Boxed Position Unboxing
* **Business Challenge:** High-volume **Boxed Positions** (simultaneous Long and Short holdings in the same security across Equity & Swap instruments) caused by Cover/Sell and Buy/Short record mismatches, leading to trade failure and custody lockups.
* **Solution:**
  * Generates dual-sided, double-entry **GTL CSV Loaders** (`Buy Cover` & `Sell` pairs) with unique pseudorandom identifiers.
  * Automatically sanitizes incoming numeric data (clearing non-breaking spaces, text-formatting errors, and thousands separators).
  * Supports automated multi-asset class routing (Equities vs. Equity Swaps).



