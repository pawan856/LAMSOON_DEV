# Quick Reference: Stock Return Post Button Prompt

## For Agents/Developers

If you need to understand, modify, or debug the **Post button** functionality on the Stock Return page, use this prompt:

---

**PROMPT FOR AGENT:**

```
Please review the Stock Return Post button functionality documented in:
/INBOUND/SR/POST_BUTTON_FUNCTIONALITY_PROMPT.md

This document contains complete details about:
- How the Post button works (client-side and server-side)
- What validations are performed
- How stock transactions are processed
- Which database tables are updated
- How serial numbers are handled
- Business logic flow
- Error handling
- Testing requirements

Use this documentation to [understand/modify/debug/enhance] the Post button logic.
```

---

## Quick Facts

**Location**: `/INBOUND/SR/SRMain.aspx` (Line 251) and `SRMain.aspx.vb` (Lines 1476-1633)

**Purpose**: Posts a stock return to inventory, adding returned items back to stock balance

**Status Requirement**: Can only post when status is "APPROVED"

**Main Actions**:
1. Validates all data
2. Saves current changes
3. Processes stock transactions (adds inventory)
4. Updates serial numbers (if applicable)
5. Links to original stock issue
6. Changes status to "POSTED"

**Key Tables Updated**:
- WMS_STOCK_RETURN (master)
- WMS_STOCK_RETURN_D (details)
- WMS_STOCK_TRANS (transactions)
- WMS_ITEM_LOC_BAL (inventory balance)
- WMS_STOCK_SERIAL_TRANS (serial transactions)
- WMS_ITEM_LOC_SERIAL_BAL (serial balance)
- WMS_STOCK_ISSUE (related stock issue)
- WMS_DATE_CODE (lot codes)

**Special Features**:
- Serial number auto-generation
- Cable item support with drum tracking
- Batch/Lot validation and auto-creation
- Expiry date management
- Transaction rollback on errors
- Multi-language support

**Access Control**: Security code "IB_SR", ADMIN_GP for un-post

---

## Common Tasks

### To Understand the Logic:
Read sections: Overview, Server-Side Processing, Business Logic Flow

### To Modify Validation:
See: Pre-Validation section (line 37-58 in docs)

### To Change Stock Processing:
See: Stock Transaction Processing section (line 63-97 in docs)

### To Debug Issues:
See: Error Handling, Testing Recommendations sections

### To Add New Features:
See: Customization Points section

---

## Full Documentation
👉 **[POST_BUTTON_FUNCTIONALITY_PROMPT.md](POST_BUTTON_FUNCTIONALITY_PROMPT.md)**

## Module Overview
👉 **[README.md](README.md)**
