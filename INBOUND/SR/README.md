# Stock Return Module

## Overview
This directory contains the Stock Return (SR) module of the warehouse management system. The Stock Return module handles the return of stock items to the warehouse from customers or consignees.

## Files

### Main Application Files
- **SRMain.aspx** - Main stock return page (UI)
- **SRMain.aspx.vb** - Code-behind for stock return page (Business logic)

### Documentation
- **STOCK_RETURN_PAGE_DOCUMENTATION.md** - Comprehensive documentation explaining everything about the Stock Return page

### Supporting Directories
- **SR_LABEL/** - Item label printing functionality

## Quick Links

### Documentation
For complete details about the Stock Return page, see: [STOCK_RETURN_PAGE_DOCUMENTATION.md](./STOCK_RETURN_PAGE_DOCUMENTATION.md)

The documentation covers:
- **Overview and Purpose** - What the page does and when to use it
- **All Fields** - Detailed explanation of every field on the page
- **Action Buttons** - Functions and behaviors of all buttons
- **Returned Items Grid** - Complete grid documentation
- **Business Logic** - Validation rules and workflows
- **Workflow States** - State transitions and diagrams
- **Database Tables** - Related database tables and fields
- **Technical Details** - Implementation details and classes
- **Troubleshooting** - Common issues and solutions
- **Best Practices** - Recommended usage patterns

## Module Purpose

The Stock Return module allows warehouse staff to:
1. Create new stock return records
2. Track returned items with full details (batch, serial, location, etc.)
3. Manage approval workflow (Submit → Approve → Post)
4. Update inventory levels when returns are posted
5. Link returns to original stock issues
6. Print labels for returned items
7. Attach supporting documents

## Key Features

- **Auto-generated return codes** - System generates unique identifiers
- **Approval workflow** - Submit → Pending → Approved → Posted
- **Inventory integration** - Automatic stock balance updates
- **Serial number tracking** - For serialized items
- **Batch/lot tracking** - Full traceability
- **Multiple return types** - SCRAP, SLCABLE, NS, etc.
- **Item lookup** - Easy selection from item master
- **Stock issue integration** - Import items from stock issues
- **Bilingual support** - English and Chinese

## Workflow

```
NEW → [Submit] → PENDING → [Approve] → APPROVED → [Post] → POSTED
  ↓                ↓                        ↓
[Cancel]      [Un-Submit]              [Cancel]
  ↓                ↓                        ↓
CANCELLED        NEW                    CANCELLED

POSTED → [Un-Post] → NEW
```

## Access

- **Menu Code**: IB_SR
- **Module**: INBOUND
- **Path**: `INBOUND/SR/SRMain.aspx`

## For More Information

Please refer to the complete documentation: [STOCK_RETURN_PAGE_DOCUMENTATION.md](./STOCK_RETURN_PAGE_DOCUMENTATION.md)
