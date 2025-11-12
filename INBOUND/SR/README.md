# Stock Return (SR) Module Documentation

## Overview
This directory contains the Stock Return (SR) module for the LAMSOON_DEV WMS system.

## Main Files
- **SRMain.aspx** - Stock Return page UI
- **SRMain.aspx.vb** - Stock Return page code-behind
- **SR_LABEL/** - Stock Return label printing functionality

## Documentation

### Post Button Functionality
For a comprehensive understanding of the Post button logic and functionality, please refer to:
- **[POST_BUTTON_FUNCTIONALITY_PROMPT.md](POST_BUTTON_FUNCTIONALITY_PROMPT.md)**

This document provides:
- Complete overview of the Post button functionality
- Detailed server-side and client-side processing steps
- Database tables affected
- Validation rules
- Business logic flow
- Testing recommendations
- Customization points

This prompt can be provided to agents or developers who need to understand, modify, or debug the Post button implementation.

## Key Features
- Stock return creation and management
- Item selection and tracking
- Serial number management
- Batch/Lot tracking
- Approval workflow (New → Pending → Approved → Posted)
- Integration with Stock Issue (SI) module
- Multi-language support (English/Chinese)

## Status Flow
1. **NEW** - Initial creation
2. **PENDING** - Submitted for approval
3. **APPROVED** - Approved and ready to post
4. **POSTED** - Posted to inventory (stock added back)
5. **CANCELLED** - Cancelled return
6. **CLOSED** - Closed
7. **PREWEIGHT** - Pre-weighed status

## Access Control
Module security code: **IB_SR**

## Related Modules
- Stock Issue (SI) - Original stock issue documents
- Item Master - Item and pack key information
- Location Management - Warehouse and bin locations
- Vendor Management - Vendor information
