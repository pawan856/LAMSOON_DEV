# Stock Return Page - Post Button Functionality Prompt

## Overview
This document provides a comprehensive description of the Post button functionality on the Stock Return (SR) page located at `/INBOUND/SR/SRMain.aspx`. This prompt can be provided to an agent for understanding, modification, or debugging purposes.

## File Locations
- **ASPX File**: `/INBOUND/SR/SRMain.aspx` (Line 251)
- **Code-Behind File**: `/INBOUND/SR/SRMain.aspx.vb` (Lines 1476-1633)

## Button Declaration
```asp
<asp:Button ID="btnPost" Text="Post" CssClass="all_button" runat="server"/>
```

## Current Implementation

### 1. User Interface Behavior

#### Client-Side Confirmation
- **English**: "Are you sure to post this record?"
- **Chinese**: "確定發布資料?"
- Shows a loading modal popup after confirmation
- Prevents default form submission if user cancels

#### Button Visibility Rules
The Post button visibility is controlled by the stock return status:
- **NEW**: Hidden (btnPost.Visible = False)
- **PENDING**: Hidden (btnPost.Visible = False)
- **APPROVED**: Visible (btnPost.Visible = True)
- **CANCELLED**: Hidden (btnPost.Visible = False)
- **POSTED**: Hidden (btnPost.Visible = False)
- **CLOSED**: Hidden (btnPost.Visible = False)
- **PREWEIGHT**: Hidden (btnPost.Visible = False)

### 2. Server-Side Processing (btnPost_Click Event)

#### Pre-Validation
1. **Grid Validation**: Checks if GridView1 has rows
2. **Data Validation**: Calls `validateAll()` function which checks:
   - Storer Code is not empty
   - Return Type is not empty
   - Warehouse is not empty
   - Valid date format for RT_DATE
   - Valid email format for RT_BY_EMAIL
   - Valid numeric values for total pallet
   - Remark length <= 200 characters
   - Item details validation:
     - RCV Qty is not empty and > 0
     - Location is selected
     - Batch number is valid (8-character date format YYYYMMDD)
     - UOM2 is not blank
     - Qty2 is not blank
     - Expiry date is valid
     - Manufacturing date is valid
     - Manufacturing date <= Expiry date
     - Serial number (if required for item)
     - Item weight (KG) is numeric
     - Drum ID validation for CABLE type items

#### Data Saving
3. **Save Operation**: Calls `save("Y")` to persist current changes before posting

#### Stock Transaction Processing
For each item in the grid (DataTable dt):
4. **Stock Transaction Setup**:
   - Creates StockTrans object (st)
   - Sets transaction properties:
     - STORER_CODE
     - ITM_CODE (Item Code)
     - PACK_KEY
     - IO_DOC = "SR" (Stock Return)
     - IO_DOC_ID = RT_CODE
     - IO_QTY = Received Quantity
     - PALLET_NO
     - IO_WH = Warehouse
     - IO_LOC = Location
     - IO_EXPIRY_DATE (from item location balance or RTD_EXPIRY_DATE)
     - IO_MANU_DATE
     - lO_BATCH_NO
     - lO_VND_CODE (Vendor Code)
   
5. **Serial Number Processing** (if item has serial number):
   - IOS_DRUM_ID
   - IOS_DRUM_LEVEL
   - Generates new serial number using `getNewSerialNo()`
   - IOS_QTY2 and IOS_UOM2
   - Determines if item is CABLE type (sets IOS_SL flag)
   - Sets original serial info

6. **System Sequence**: Generates IO_SYS_SEQ using `getDocNo("SYS_SEQ")`

7. **Stock Updates**:
   - Calls `UpdateStockTrans("IN")` - Updates WMS_STOCK_TRANS table
   - Calls `UpdateStockBalTrans("IN")` - Updates WMS_ITEM_LOC_BAL table
   - If serial number exists:
     - Calls `UpdateStockSerialTrans("IN")` - Updates WMS_STOCK_SERIAL_TRANS
     - Calls `UpdateStockBalSerialTrans("IN")` - Updates WMS_ITEM_LOC_SERIAL_BAL

8. **Detail Record Update**:
   - Updates WMS_STOCK_RETURN_D.RTD_SERIAL_NO with the generated serial number

#### Related Stock Issue Update
9. **Stock Issue Link**: Updates related WMS_STOCK_ISSUE record:
   - Sets IS_RETURN_DATE = Current date
   - Sets IS_RETURN_DOC_NO = RT_CODE
   - Based on RT_REF_DOC_NO (RMA Number)

#### Status Update
10. **Master Record Update**: Updates WMS_STOCK_RETURN:
    - RT_STATUS = 'POSTED'
    - POSTED_DATE = Current date
    - Only if current status <> 'POSTED'

#### Post-Processing
11. **Success Handling**:
    - Commits database transaction
    - Sets RT_STATUS.Value = "POSTED"
    - Updates display status
    - Hides Post button
    - Sets form to view mode (ar.sec_write = "N")
    - Displays success message "1007" (Record saved successfully)
    - Uses RemotePost to reload page with alert message

12. **Error Handling**:
    - Rolls back transaction on any exception
    - Hides loading modal popup
    - Displays error messages:
      - "No item can be posted!" (if no items)
      - Database error messages on transaction failure

### 3. Database Tables Affected

#### Direct Updates:
- **WMS_STOCK_RETURN**: Status and posted date
- **WMS_STOCK_RETURN_D**: Serial numbers
- **WMS_STOCK_ISSUE**: Return date and document number
- **WMS_STOCK_TRANS**: Stock transaction history
- **WMS_ITEM_LOC_BAL**: Item location balance
- **WMS_STOCK_SERIAL_TRANS**: Serial number transactions (if applicable)
- **WMS_ITEM_LOC_SERIAL_BAL**: Serial number balance (if applicable)
- **WMS_DATE_CODE**: Batch/Lot codes (auto-insert if not exists)

### 4. Business Logic Flow

```
User clicks Post button
    ↓
Confirm dialog shown
    ↓
If confirmed → Show loading modal
    ↓
Validate all data
    ↓
Save current changes
    ↓
Begin Database Transaction
    ↓
For Each Item:
    → Setup stock transaction object
    → Get/Generate serial number (if required)
    → Update stock transaction tables
    → Update stock balance tables
    → Update serial transaction tables (if applicable)
    → Update detail serial number
    ↓
Update related Stock Issue record
    ↓
Update Stock Return status to POSTED
    ↓
Commit Transaction
    ↓
Update UI (hide button, set view mode)
    ↓
Reload page with success message
```

### 5. Access Control
- Button is controlled by AccessRightUtils ("IB_SR" security code)
- Special permission check for ADMIN_GP group
- View mode enforced after successful post

### 6. Multi-Language Support
- English and Chinese (Traditional) languages supported
- All messages and labels are language-specific

## Key Features

1. **Transaction Safety**: All database operations wrapped in transaction with rollback on error
2. **Inventory Management**: Properly updates stock balances (adds inventory back)
3. **Serial Number Tracking**: Full support for serialized items with drum ID tracking
4. **Cable Item Support**: Special handling for cable-type items with short-length tracking
5. **Traceability**: Links stock returns to original stock issues
6. **Batch/Lot Management**: Auto-creates lot codes if they don't exist
7. **Expiry Date Management**: Retrieves existing expiry dates from location balance or uses new values
8. **Audit Trail**: Updates system fields (last update by, last update date)

## Customization Points

If you need to modify the Post button functionality, consider these areas:

1. **Validation Rules**: `validateAll()` function (lines 490-841)
2. **Stock Transaction Logic**: `btnPost_Click()` event (lines 1476-1633)
3. **Confirmation Messages**: Lines 104 (English), 149 (Chinese) in Page_Load
4. **Button Visibility**: Lines 204-250 in Page_Load
5. **Post-Transaction Actions**: Lines 1614-1621 (RemotePost section)
6. **Error Messages**: Lines 1625-1629

## Dependencies

- **GlobalDBFunc**: Database connection and operations
- **DBfunc**: Database utility functions
- **StockTrans**: Stock transaction management class
- **UIfunc**: UI utility functions
- **AccessRightUtils**: Security and access control
- **RemotePost**: Page navigation with data
- **GeneralUtils**: General utility functions

## Testing Recommendations

1. Test with items that have serial numbers
2. Test with CABLE type items
3. Test with items without serial numbers
4. Test validation failures
5. Test with related stock issue records
6. Test rollback on database errors
7. Test access control restrictions
8. Test in both English and Chinese languages

## Notes for Developers

- The Post operation is **irreversible** through the UI (though Un-Post functionality exists for ADMIN_GP)
- All stock quantities are added back to inventory (IN operation)
- Serial numbers are auto-generated if items require them
- The function uses OLEDB-style string concatenation for SQL (consider parameterization for security)
- Transaction isolation ensures data consistency
- Modal popup provides user feedback during long operations
