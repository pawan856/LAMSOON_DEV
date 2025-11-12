# Stock Return Page Documentation

## Overview
The Stock Return page (SRMain.aspx) is a comprehensive warehouse management system module that handles the return of stock items to the warehouse. This page is part of the Inbound (IB) module and manages the entire lifecycle of stock returns from creation to posting.

## Page Location
- **Path**: `INBOUND/SR/SRMain.aspx`
- **Code-behind**: `INBOUND/SR/SRMain.aspx.vb`
- **Menu Code**: IB_SR
- **Title**: Stock Return / 退貨維護 (bilingual: English/Chinese)

## Purpose
This page allows warehouse users to:
- Create new stock return records
- Edit existing stock return records
- Track returned items from customers or consignees
- Manage the approval workflow for stock returns
- Post stock returns to update inventory levels
- Print item labels for returned goods
- Attach supporting documents

## Page Modes
The page operates in different modes:

### 1. New Mode (N)
- Used when creating a new stock return record
- Internal Return Code is auto-generated
- All fields are editable
- Status is set to "NEW"

### 2. Edit/View Mode
- Used when viewing or editing existing records
- Code fields are locked
- Edit permissions depend on record status and user access rights

## Main Header Section

### Key Fields

#### 1. Internal Return Code WMS (RT_CODE)
- **Type**: Auto-generated
- **Format**: System-generated unique identifier
- **Purpose**: Primary identifier for the stock return record
- **Editable**: No (auto-generated)

#### 2. Status (RT_STATUS)
- **Type**: Dropdown/Display
- **Possible Values**:
  - **NEW**: Initial status when record is created
  - **PENDING**: Submitted for approval
  - **APPROVED**: Approved by authorized personnel
  - **POSTED**: Inventory has been updated
  - **CANCELLED**: Record has been cancelled
  - **CLOSED**: Process completed
  - **PREWEIGHT**: Pre-weight status
- **Purpose**: Tracks the current state of the stock return

#### 3. Organizations (STORER_CODE)
- **Type**: Dropdown
- **Required**: Yes
- **Purpose**: Identifies the organization/customer returning the stock
- **Validation**: Must be selected before proceeding
- **Note**: Cannot be changed after saving in edit mode

#### 4. Type (RT_TYPE)
- **Type**: Dropdown
- **Required**: Yes
- **Possible Values**:
  - **SCRAP**: Items being returned as scrap
  - **SLCABLE**: Short-length cable items
  - **NS**: Non-stock items
  - **Other return types** as configured in system
- **Purpose**: Categorizes the reason/type of return
- **Impact**: Affects default location selection for items

#### 5. Date (RT_DATE)
- **Type**: Date picker
- **Format**: dd/MM/yyyy
- **Purpose**: Date when the stock return occurred
- **Validation**: 
  - Must be valid date format
  - Cannot be earlier than SO Date (if applicable)
- **Features**: Calendar picker available

#### 6. Received By (RT_RCV_BY)
- **Type**: Text field
- **Max Length**: 20 characters
- **Purpose**: Name of person who received the returned goods

#### 7. Total Pallet (RT_TOT_PALLET)
- **Type**: Numeric
- **Purpose**: Total number of pallets in the return
- **Validation**: Must be a valid decimal number

#### 8. C8 Form - Completed (Y/N) (RT_C8_YN)
- **Type**: Display/Status
- **Values**: Y (Yes) / N (No)
- **Purpose**: Indicates if C8 form has been completed
- **Actions**: 
  - COMPLETE button: Marks C8 as completed
  - UNCOMPLETE button: Reverses C8 completion

#### 9. Returned By (RT_BY)
- **Type**: Text field
- **Max Length**: 30 characters
- **Purpose**: Name of person returning the goods

#### 10. Serial No. (RT_BY_TEL)
- **Type**: Text field
- **Max Length**: 80 characters
- **Purpose**: Contact telephone number

#### 11. Email (RT_BY_EMAIL)
- **Type**: Email field
- **Max Length**: 30 characters
- **Purpose**: Contact email address
- **Validation**: Must be valid email format

#### 12. Issue Control Sheet No. (RT_REF_NO)
- **Type**: Text field
- **Max Length**: 30 characters
- **Purpose**: Reference to issue control sheet

#### 13. Subinventory (RT_WH)
- **Type**: Dropdown
- **Required**: Yes
- **Purpose**: Warehouse location where stock will be returned
- **Impact**: Updates available locations in detail lines

#### 14. Lot (RT_BATCH_NO)
- **Type**: Text field
- **Max Length**: 20 characters
- **Purpose**: Batch/lot number for the return

#### 15. MPAC Control No. (RT_CUS_CODE)
- **Type**: Text field
- **Max Length**: 20 characters
- **Purpose**: Customer control number
- **Note**: Editable even in POSTED status for authorized users

#### 16. RMA Num (RT_CONS_CODE)
- **Type**: Text field
- **Max Length**: 20 characters
- **Purpose**: Return Material Authorization number

#### 17. Fault Remark (RT_FAULT_REM)
- **Type**: Multi-line text area
- **Max Length**: 200 characters
- **Purpose**: Details about faults or reasons for return

#### 18. Credit Forms No. (RT_REF_NO2)
- **Type**: Text field
- **Max Length**: 120 characters
- **Purpose**: Reference to credit form number

#### 19. RMA Num (RT_REF_DOC_NO)
- **Type**: Text field with "Get SI Items" button
- **Max Length**: 20 characters
- **Purpose**: RMA number that can be used to auto-populate items from Stock Issue
- **Feature**: Click "Get SI Items" to load items from corresponding Stock Issue record

#### 20. Approval Code (RT_APPROVE_CODE)
- **Type**: Display only
- **Purpose**: Shows the approval code entered when record was approved
- **Visible**: Only after approval

#### 21. Remarks (RT_REM)
- **Type**: Multi-line text area
- **Max Length**: 200 characters
- **Purpose**: General remarks about the stock return

## Action Buttons

### Primary Actions

#### 1. New Button
- **Function**: Creates a new stock return record
- **Behavior**: 
  - Prompts to save current record if unsaved changes exist
  - Redirects to new blank form
- **Availability**: All modes

#### 2. Save Button
- **Function**: Saves the current record
- **Behavior**: 
  - Validates all required fields
  - Creates new record or updates existing
  - Generates RT_CODE for new records
  - Auto-creates batch numbers in master data if new
- **Availability**: NEW, APPROVED, POSTED (limited)
- **Confirmation**: "Save Record?" prompt

#### 3. Back Button
- **Function**: Returns to search/list page
- **Target**: cms_search.aspx?menu_code=IB_SR
- **Availability**: All modes

### Workflow Actions

#### 4. Submit for Approval Button
- **Function**: Changes status from NEW to PENDING
- **Behavior**:
  - Validates all data
  - Saves record
  - Sets RT_STATUS to 'PENDING'
  - Records submission user and timestamp
  - Makes record read-only except for approvers
- **Availability**: NEW status only

#### 5. Un-Submit for Approval Button
- **Function**: Reverts status from PENDING back to NEW
- **Behavior**:
  - Changes RT_STATUS to 'NEW'
  - Clears submission user and timestamp
  - Makes record editable again
- **Availability**: PENDING status only
- **Confirmation**: "Are you sure to UnSubmit this Record?"

#### 6. Approve Button
- **Function**: Approves the stock return
- **Behavior**:
  - Opens modal popup requesting Approval Code
  - Changes RT_STATUS to 'APPROVED'
  - Records approval code, user, and timestamp
  - Enables Post button
- **Availability**: PENDING status only
- **Required**: Approval Code must be entered

#### 7. Post Button
- **Function**: Posts the stock return to inventory
- **Behavior**:
  - Validates all data
  - Creates stock transactions (WMS_STOCK_TRANS)
  - Updates stock balances (WMS_ITEM_LOC_BAL)
  - Handles serial numbers if applicable
  - Updates related Stock Issue record
  - Changes RT_STATUS to 'POSTED'
  - Shows loading indicator during processing
- **Availability**: APPROVED status only
- **Confirmation**: "Are you sure to post this record?"
- **Impact**: Increases inventory quantities

#### 8. Un-Post Button
- **Function**: Reverses a posted stock return
- **Behavior**:
  - Reverses all stock transactions
  - Reverses stock balance updates
  - Reverses serial number updates
  - Changes RT_STATUS back to 'NEW'
  - Validates that stock is still available
- **Availability**: POSTED status only
- **Access**: Restricted to ADMIN_GP group or users with BT_SR_UNPOST right
- **Confirmation**: "Are you sure to un-post this record?"
- **Warning**: Prompts to save work before un-posting

### Additional Actions

#### 9. Cancel Button
- **Function**: Cancels the stock return record
- **Behavior**:
  - Changes RT_STATUS to 'CANCELLED'
  - Makes record read-only
  - Hides most action buttons
- **Availability**: NEW, PENDING, APPROVED statuses
- **Confirmation**: "Are you sure to cancel this record?"

#### 10. Check Stock Balance Button
- **Function**: Opens stock balance inquiry
- **Behavior**: Opens popup window showing current stock levels
- **Target**: cms_search.aspx with menu_code=INQ_001
- **Parameters**: Filters by selected STORER_CODE
- **Availability**: All modes

#### 11. Attachment Button
- **Function**: Manages document attachments
- **Behavior**: Opens attachment management popup
- **Target**: ../../ATTACH/ATTACH_MAIN.ASPX
- **Parameters**: DOC_TYPE='IB_SR', DOC_NO=combination of IMP_CODE, STORER_CODE, RT_CODE
- **Availability**: All modes

#### 12. Print Item Label Button
- **Function**: Prints labels for returned items
- **Behavior**: Opens label printing popup
- **Target**: SR_LABEL/item_label_print.aspx
- **Parameters**: rt_code, so_date, storer_code
- **Availability**: View/Edit mode only (not in New mode)

#### 13. COMPLETE Button (C8 Form)
- **Function**: Marks C8 form as completed
- **Behavior**: Sets RT_C8_YN to 'Y'
- **Availability**: When C8_YN is not 'Y'
- **Confirmation**: "Are you sure to complete c8?"

#### 14. UNCOMPLETE Button (C8 Form)
- **Function**: Marks C8 form as not completed
- **Behavior**: Sets RT_C8_YN to 'N'
- **Availability**: When C8_YN is 'Y'
- **Confirmation**: "Are you sure to Un-complete c8?"

## Returned Items Section

### Header: "Returned Items" / "退貨物件"

### Item Grid Actions

#### 1. Add Button
- **Function**: Adds a new blank row to the items grid
- **Behavior**: 
  - Generates next sequence number
  - Creates empty row for manual data entry
- **Availability**: Edit mode (hidden in most views)

#### 2. Select Item Button
- **Function**: Opens item lookup to add items
- **Behavior**:
  - Opens item lookup popup
  - Allows selection of multiple items
  - Auto-populates item details
  - Sets default location based on return type
- **Requirements**: Storer Code must be selected first
- **Availability**: NEW status or when editable

#### 3. Copy Items Button
- **Function**: Duplicates selected items
- **Behavior**:
  - Copies checked items
  - Creates new sequence numbers
  - Marks as new items
- **Availability**: When items exist

#### 4. Get SI Items Button
- **Function**: Loads items from Stock Issue record
- **Requirements**: RT_REF_DOC_NO (RMA Num) must be filled
- **Behavior**:
  - Retrieves all items from referenced Stock Issue
  - Clears existing items and replaces with SI items
  - Copies quantities, batch numbers, dates, etc.
  - Only works for POSTED Stock Issue records
- **Validation**: Shows error if SI record not found or not posted

### Item Grid Columns

#### 1. Checkbox Column
- **Purpose**: Select items for copy operation
- **Width**: 30px

#### 2. No. (RTD_SEQ)
- **Type**: Text (read-only, auto-generated)
- **Width**: 50px
- **Purpose**: Sequence number for the line item
- **Format**: Integer, sequential

#### 3. Item Code (ITM_SKU_NO)
- **Type**: Label (read-only)
- **Purpose**: Displays item SKU number
- **Source**: From WMS_ITEM master

#### 4. Pack Key (RTD_PACK_KEY)
- **Type**: Label (read-only)
- **Purpose**: Packaging configuration code
- **Source**: From WMS_ITEM master

#### 5. Lot (RTD_BATCH_NO)
- **Type**: Text input
- **Width**: 70px
- **Max Length**: 20 characters
- **Required**: Yes
- **Validation**: 
  - Must not be blank
  - Must be at least 8 characters (YYYYMMDD format)
  - First 8 characters must be a valid date
- **Behavior**: Auto-creates in WMS_DATE_CODE if new batch

#### 6. Pallet No. (RTD_PALLET_NO)
- **Type**: Text input
- **Width**: 70px
- **Max Length**: 20 characters
- **Default**: "000" if left blank
- **Purpose**: Identifies the pallet

#### 7. Carton No. (RTD_CARTON_NO)
- **Type**: Text input
- **Width**: 70px
- **Max Length**: 20 characters
- **Purpose**: Identifies the carton/box

#### 8. Vendor Code (RTD_VND_CODE)
- **Type**: Label (read-only)
- **Purpose**: Vendor/supplier code
- **Default**: "DEF_VEND" if not specified

#### 9. Item Name (RTD_ITM_NAME)
- **Type**: Label (read-only)
- **Purpose**: Description of the item
- **Source**: From WMS_ITEM master

#### 10. Reference (RTD_REF_NO)
- **Type**: Text input
- **Width**: 90px
- **Max Length**: 20 characters
- **Purpose**: Additional reference number

#### 11. Expiry Date (RTD_EXPIRY_DATE)
- **Type**: Date picker
- **Format**: dd/MM/yyyy
- **Required**: Yes
- **Validation**: 
  - Must be valid date
  - Must be after Manufactory Date (if specified)
- **Features**: Calendar picker

#### 12. Manufactory Date (RTD_MANU_DATE)
- **Type**: Date picker
- **Format**: dd/MM/yyyy
- **Validation**: 
  - Must be valid date (if provided)
  - Must be before Expiry Date (if specified)
- **Features**: Calendar picker

#### 13. Ori. Qty (RTD_KG)
- **Type**: Numeric (read-only)
- **Width**: 40px
- **Purpose**: Original quantity
- **Alignment**: Right
- **Validation**: Must be numeric

#### 14. Ori. UOM (RTD_DRUM_LV)
- **Type**: Text (read-only)
- **Width**: 100px
- **Max Length**: 50 characters
- **Purpose**: Original unit of measure

#### 15. Subinventory (RTD_WH)
- **Type**: Text input with AutoPostBack
- **Width**: 40px
- **Purpose**: Warehouse/subinventory code
- **Behavior**: Updates available locations when changed
- **AutoPostBack**: Yes (triggers location dropdown update)

#### 16. Loc (RTD_LOC)
- **Type**: ComboBox (autocomplete dropdown)
- **Max Length**: 20 characters
- **Required**: Yes
- **Purpose**: Specific bin location in warehouse
- **Data Source**: WMS_WH_BIN filtered by RTD_WH
- **Format**: FL_NUM + AR_CODE + RK_CODE + BN_CODE
- **Features**: 
  - Autocomplete/search capability
  - Dropdown selection
  - Case-insensitive

#### 17. Qty (RTD_RCV_QTY)
- **Type**: Numeric input
- **Width**: 40px
- **Required**: Yes
- **CSS Class**: REQUIRED
- **Validation**: 
  - Cannot be empty
  - Must be numeric
  - Must be greater than zero
- **Alignment**: Right
- **Purpose**: Quantity being returned

#### 18. UOM (ITM_UOM)
- **Type**: Text input
- **Width**: 40px
- **Purpose**: Unit of measure for quantity
- **Alignment**: Right

#### 19. Serial No. (RTD_SERIAL_NO)
- **Type**: Text input
- **Width**: 100px
- **Max Length**: 50 characters
- **Visible**: false (controlled by item configuration)
- **Required**: Only if ITM_SERIAL_NO_YN = 'Y'
- **Purpose**: Serial number for serialized items

#### 20. Drum ID (RTD_DRUM_ID)
- **Type**: Text input
- **Width**: 100px
- **Max Length**: 50 characters
- **Visible**: false (for cable items)
- **Validation**: Must exist in WMS_DRUM master for cable items
- **Purpose**: Drum identifier for cable items

#### 21. Qty2 (RTD_QTY2)
- **Type**: Numeric input
- **Width**: 40px
- **Required**: Yes
- **Validation**: 
  - Cannot be empty
  - Must be numeric
  - Must be greater than zero
- **Alignment**: Right
- **Purpose**: Secondary quantity (e.g., for cable length)

#### 22. UOM2 (RTD_UOM2)
- **Type**: Text input
- **Width**: 40px
- **Required**: Yes
- **Purpose**: Unit of measure for secondary quantity
- **Alignment**: Right

#### 23. Delete Button
- **Type**: Button
- **Width**: 50px
- **Function**: Marks the row for deletion
- **Confirmation**: "Are you sure you want to delete this record?" / "你是否確定要刪除這個資料?"
- **Behavior**: Hides row and marks with mFlag='D'

### Hidden Fields (per item)
- **RTD_ITM_CODE**: Item code (primary key)
- **ITM_SERIAL_NO_YN**: Indicates if item requires serial numbers
- **RTD_LOC_WH**: Location warehouse
- **SAP_MAT_DOC_NO**: SAP material document number
- **SAP_MAT_DOC_ITEM**: SAP material document item
- **RTD_STATUS**: Item status

## System Information Section

Located at the bottom of the form:

### 1. CB (Created By / 創建者)
- **Field**: sys_cb
- **Purpose**: Username who created the record
- **Source**: Session("usr_id") at creation time

### 2. CD (Created Date / 創建日期)
- **Field**: sys_cd
- **Purpose**: Date and time when record was created
- **Format**: DateTime from database

### 3. LUB (Last Updated By / 最後更新者)
- **Field**: sys_lub
- **Purpose**: Username who last modified the record
- **Source**: Session("usr_id") at update time

### 4. LUD (Last Updated Date / 最後更新日期)
- **Field**: sys_lud
- **Purpose**: Date and time when record was last updated
- **Format**: DateTime from database

## Business Logic and Validation Rules

### Header Validation

1. **Organization (STORER_CODE)**
   - Must be selected
   - Error message: "Organizations cannot be empty!" / "部門不能空白!"

2. **Type (RT_TYPE)**
   - Must be selected
   - Error message: "Type cannot be empty!" / "類型不能空白!"

3. **Subinventory (RT_WH)**
   - Must be selected
   - Error message: "Subinventory cannot be empty!" / "子庫存不能空白!"

4. **Date (RT_DATE)**
   - Must be valid date format (dd/MM/yyyy)
   - Cannot be earlier than SO_DATE if SO_DATE exists
   - Error: "Invalid date" / "無效的日期"
   - Error: "Return date can not be less than SO Date!"

5. **Email (RT_BY_EMAIL)**
   - If provided, must be valid email format
   - Error: "Invalid email" / "無效的電子郵箱"

6. **Total Pallet (RT_TOT_PALLET)**
   - Must be numeric
   - Error: "Invalid number, Total Pallet!" / "無效的數字, 貨板總量!"

7. **Remarks (RT_REM)**
   - Maximum 200 characters
   - Error: "Too many characters, maximum length of Remarks is 200!" / "字數太多, 備註最多只限200字!"

### Item Line Validation

1. **Quantity (RTD_RCV_QTY)**
   - Cannot be empty
   - Must be numeric
   - Must be greater than zero
   - Errors:
     - "RCV Qty cannot be empty!" / "收貨數量不能空白!"
     - "Invalid number, RCV Qty!" / "無效的數字, 收貨數量!"
     - "RCV Qty need to be greater than zero!" / "收貨數量不能設零/少於零!"

2. **Location (RTD_LOC)**
   - Cannot be empty
   - Error: "Please select Location!" / "位置不能空白!"

3. **Batch Number (RTD_BATCH_NO)**
   - Cannot be empty
   - Must be at least 8 characters
   - First 8 characters must form valid date (YYYYMMDD)
   - Errors:
     - "Batch No is blank!"
     - "Invalid batch number!"

4. **UOM2 (RTD_UOM2)**
   - Cannot be empty
   - Error: "UOM2 is blank!"

5. **Qty2 (RTD_QTY2)**
   - Cannot be empty
   - Must be numeric
   - Must be greater than zero
   - Errors:
     - "Qty2 is blank!" / "收貨數量2不能空白!"
     - "Invalid number, Qty2!" / "無效的數字, 收貨數量2!"
     - "Qty2 cannot be zero for item!" / "收貨數量2不能設零!"

6. **Expiry Date (RTD_EXPIRY_DATE)**
   - Cannot be empty
   - Must be valid date format
   - Must be after Manufactory Date
   - Errors:
     - "Invalid Expiry Date!" / "無效的失效日期!"
     - "Manu. Date cannot later then Expiry Date!"

7. **Manufactory Date (RTD_MANU_DATE)**
   - If provided, must be valid date format
   - Must be before Expiry Date
   - Error: "Invalid Manufactory Date!" / "無效的生產日期!"

8. **Serial Number (RTD_SERIAL_NO)**
   - Required only if ITM_SERIAL_NO_YN = 'Y'
   - Error: "Serial No. is required for the item: [SKU]"

9. **Original Qty (RTD_KG)**
   - Must be numeric
   - Errors:
     - "Item KG must be numeric!" / "收貨重量必需為數目字!"

10. **Drum ID (RTD_DRUM_ID)**
    - For cable items (ITM_TYPE = 'CABLE'), must exist in WMS_DRUM master
    - Error: "Invalid Drum ID: [ID]. Drum ID Not Exists in Drum master record!"

11. **At Least One Item**
    - Grid must contain at least one non-deleted item
    - Error: "Please Select Item!"

## Workflow States and Transitions

### State Diagram

```
NEW → [Submit] → PENDING → [Approve] → APPROVED → [Post] → POSTED
  ↓                ↓                        ↓
[Cancel]      [Un-Submit]              [Cancel]
  ↓                ↓                        ↓
CANCELLED        NEW                    CANCELLED

POSTED → [Un-Post] → NEW
```

### State-Specific Behaviors

#### NEW Status
- **Edit**: Full edit access
- **Available Actions**: Save, Submit, Cancel, Check Stock Balance, Attachment
- **Visible Buttons**: New, Save, Back, Submit, Cancel, Check Stock Balance, Attachment
- **Hidden Buttons**: Approve, Post, Un-Post, Un-Submit, Print Label

#### PENDING Status
- **Edit**: Read-only except for specific users
- **Available Actions**: Approve, Un-Submit
- **Visible Buttons**: New, Back, Approve, Un-Submit, Check Stock Balance, Attachment
- **Hidden Buttons**: Save, Submit, Post, Un-Post, Cancel, Print Label
- **Access**: Approval rights required

#### APPROVED Status
- **Edit**: Mostly read-only
- **Available Actions**: Post
- **Visible Buttons**: New, Back, Post, Check Stock Balance, Attachment
- **Hidden Buttons**: Save, Submit, Approve, Un-Submit, Un-Post, Cancel, Print Label

#### POSTED Status
- **Edit**: Mostly read-only, limited fields editable
- **Available Actions**: Un-Post (restricted), Save (limited)
- **Visible Buttons**: New, Back, Un-Post, Check Stock Balance, Attachment, Print Label
- **Hidden Buttons**: Regular Save, Submit, Approve, Post, Cancel
- **Special**: Save buttons (saveBtn3, saveBtn4) for RT_CUS_CODE only
- **Access**: Un-Post restricted to ADMIN_GP or BT_SR_UNPOST right

#### CANCELLED Status
- **Edit**: Read-only
- **Available Actions**: None (informational only)
- **Visible Buttons**: New, Back, Check Stock Balance, Attachment
- **Hidden Buttons**: All action buttons (Save, Submit, Approve, Post, etc.)

#### CLOSED / PREWEIGHT Status
- **Edit**: Read-only
- **Available Actions**: Limited
- **Behavior**: Similar to POSTED status

## Database Tables

### Primary Tables

#### 1. WMS_STOCK_RETURN (Header)
Main table storing stock return header information.

**Key Fields**:
- IMP_CODE (PK): Implementation code
- STORER_CODE (PK): Customer/organization code
- RT_CODE (PK): Stock return code (auto-generated)
- RT_TYPE: Return type
- RT_STATUS: Current status
- RT_DATE: Return date
- RT_RCV_BY: Received by
- RT_BY: Returned by
- RT_BY_TEL: Telephone
- RT_BY_EMAIL: Email
- RT_BATCH_NO: Batch number
- RT_REF_NO: Reference number
- RT_WH: Warehouse
- RT_TOT_PALLET: Total pallets
- RT_REM: Remarks
- RT_CUS_CODE: Customer code
- RT_CONS_CODE: Consignee code
- RT_FAULT_REM: Fault remarks
- RT_C8_YN: C8 form completion flag
- RT_REF_NO2: Credit form number
- RT_REF_DOC_NO: RMA number
- RT_APPROVE_CODE: Approval code
- RT_SBM_APP_BY: Submitted by
- RT_SBM_APP_DATE: Submission date
- RT_APPROVED_BY: Approved by
- RT_APPROVED_DATE: Approval date
- POSTED_DATE: Posted date
- sys_cb, sys_cd, sys_lub, sys_lud: Audit fields

#### 2. WMS_STOCK_RETURN_D (Detail)
Stores individual line items for each stock return.

**Key Fields**:
- IMP_CODE (PK): Implementation code
- STORER_CODE (PK): Customer/organization code
- RT_CODE (PK): Stock return code
- RTD_SEQ (PK): Line sequence number
- RTD_PALLET_NO: Pallet number
- RTD_CARTON_NO: Carton number
- RTD_BATCH_NO: Batch/lot number
- RTD_REF_NO: Reference number
- RTD_ITM_CODE: Item code
- RTD_LOC_WH: Location warehouse
- SAP_MAT_DOC_NO: SAP material document
- SAP_MAT_DOC_ITEM: SAP document item
- RTD_STATUS: Item status
- RTD_PACK_KEY: Pack key
- RTD_ITM_NAME: Item name
- RTD_LOC: Bin location
- RTD_RCV_QTY: Received quantity
- RTD_SERIAL_NO: Serial number
- RTD_EXPIRY_DATE: Expiry date
- RTD_MANU_DATE: Manufacturing date
- RTD_VND_CODE: Vendor code
- RTD_UOM2: Secondary UOM
- RTD_QTY2: Secondary quantity
- RTD_KG: Weight
- RTD_DRUM_ID: Drum ID
- RTD_WH: Warehouse
- RTD_DRUM_LV: Drum level
- sys_cb, sys_cd, sys_lub, sys_lud: Audit fields

### Related Tables

#### 3. WMS_STOCK_TRANS
Stores all stock movements/transactions.

**Created Records**: When stock return is posted
- IO_TYPE: "IN" (inbound transaction)
- IO_DOC: "SR" (Stock Return)
- IO_DOC_ID: RT_CODE
- IO_QTY: Quantity received
- And other transaction details

#### 4. WMS_ITEM_LOC_BAL
Stores stock balance by item and location.

**Updated When**: Stock return is posted
- **Action**: Increases ILOC_QTY (available quantity)
- **Keys**: IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILOC_WH, ILOC_LOC, ILOC_BATCH_NO

#### 5. WMS_STOCK_SERIAL_TRANS
Stores serial number transactions.

**Created Records**: When posting items with serial numbers
- IOS_IO_TYPE: "IN"
- IOS_DOC: "SR"
- IOS_DOC_ID: RT_CODE
- IOS_SERIAL_NO: Serial number
- And other serial details

#### 6. WMS_ITEM_LOC_BAL_SERIAL
Stores serial number balances.

**Updated When**: Posting serialized items
- Creates or updates serial number records
- Links to specific locations and batches

#### 7. WMS_STOCK_ISSUE
Related stock issue records.

**Updated When**: Stock return is posted
- IS_RETURN_DATE: Set to current date
- IS_RETURN_DOC_NO: Set to RT_CODE
- **Purpose**: Links return back to original issue

#### 8. WMS_DATE_CODE
Stores valid batch/date codes.

**Auto-Created**: When new batch number is used
- DC_DATE_CODE: Batch number
- Created automatically if doesn't exist

#### 9. WMS_ITEM
Item master table.

**Referenced For**:
- Item details (name, UOM, serial number flag)
- ITM_TYPE (to determine if cable, etc.)
- ITM_SERIAL_NO_YN (to determine if serial tracking required)

#### 10. WMS_ITEM_WH
Item-warehouse configuration.

**Referenced For**:
- IW_PREF_LOC1: Default preferred location by warehouse

#### 11. WMS_WH_BIN
Warehouse bin locations.

**Referenced For**:
- Available locations for dropdown
- Filtered by WH_CODE and area types (scrap, short-length cable, etc.)

#### 12. WMS_DRUM
Drum master for cable items.

**Validated Against**: When RTD_DRUM_ID is provided for cable items

## Special Features

### 1. Auto-Generation of Codes
- **RT_CODE**: Auto-generated using sequence "SR"
- **RTD_SEQ**: Auto-incremented for each line item
- **IO_SYS_SEQ**: System sequence for transactions

### 2. Batch Number Auto-Creation
When saving, if a new batch number (RTD_BATCH_NO) is used:
- Automatically creates entry in WMS_DATE_CODE table
- Prevents duplicate entry errors
- Maintains master list of valid batch codes

### 3. Serial Number Generation
For items requiring serial numbers (ITM_SERIAL_NO_YN = 'Y'):
- System can generate new serial numbers
- Links serial to original serial if applicable
- Creates serial number transactions

### 4. Item Lookup Integration
The "Select Item" button opens a lookup with:
- Filtered by STORER_CODE
- Special filtering based on RT_TYPE:
  - SLCABLE: Shows cable items
  - SCRAP: Shows scrap items
  - NS: Shows non-stock items
- Multi-select capability
- Auto-population of item details

### 5. Stock Issue Integration
"Get SI Items" feature:
- Retrieves all items from a Stock Issue record
- Validates SI is in POSTED status
- Replaces current items with SI items
- Maintains quantities and dates
- Links return to original issue

### 6. Location Auto-Selection
When adding items, default location is determined by:
- **If SCRAP type**: Selects location where AR_SCRAP_AREA = 'Y'
- **If SLCABLE type**: Selects location where AR_SHORTLEN_CABLE_AREA = 'Y'
- **Otherwise**: Uses IW_PREF_LOC1 from WMS_ITEM_WH

### 7. Dynamic Location Updates
When RT_WH (header warehouse) changes:
- All item location dropdowns are refreshed
- Populated with bins from selected warehouse
- Previous selections may be cleared

When RTD_WH (line warehouse) changes:
- Location dropdown for that line is updated
- Attempts to maintain similar location if possible

### 8. Approval Code System
- Required for approval workflow
- Entered in modal popup
- Stored with approval record
- Displayed after approval

### 9. Posting Process
Complex process that:
1. Validates all data
2. Saves any pending changes
3. For each item line:
   - Creates stock transaction (IN)
   - Updates stock balance
   - If serialized:
     - Generates or validates serial number
     - Creates serial transaction
     - Updates serial balance
   - Handles cable-specific logic
4. Updates related Stock Issue record
5. Sets status to POSTED
6. Records posting date

### 10. Un-Posting Process
Reversal process that:
1. Validates record is in POSTED status
2. Checks user permissions (ADMIN_GP or BT_SR_UNPOST)
3. For each item line:
   - Reverses stock transaction
   - Decreases stock balance
   - If serialized:
     - Reverses serial transactions
     - Updates serial balances
   - Validates sufficient stock exists
4. Sets status back to NEW
5. Handles errors gracefully

### 11. Expiry Date Logic
When posting:
- Checks if item already exists in WMS_ITEM_LOC_BAL
- If exists: Uses existing expiry date
- If new: Uses expiry date from return line

### 12. Cable Item Handling
For cable items (ITM_TYPE = 'CABLE'):
- Drum ID tracking (RTD_DRUM_ID)
- Drum level tracking (RTD_DRUM_LV)
- Qty2/UOM2 for length measurement
- Special serial number logic
- IOS_SL flag set to 'Y'
- Validates drum exists in master

## Access Rights and Security

### Page-Level Access
- **Menu Code**: IB_SR
- **Access Control**: AccessRightUtils("IB_SR", usr_id)

### Button-Level Rights
Checked through ar.hasBtnRight():
- **BT_SR_UNPOST**: Required to un-post stock returns

### Group-Based Access
- **ADMIN_GP**: Has special privileges including un-post capability

### Field-Level Security
Based on RT_STATUS and user rights:
- Edit vs. view mode
- Field-specific editability
- Exception lists for special cases

### Status-Based Restrictions
- **NEW**: Full edit for authorized users
- **PENDING**: Restricted except for approvers
- **APPROVED**: Mostly locked, Post button visible
- **POSTED**: Highly restricted, Un-Post for authorized only
- **CANCELLED**: Read-only

## Session Variables

### 1. Session("usr_id")
- Current logged-in user ID
- Used for audit fields (sys_cb, sys_lub)
- Used for access rights validation

### 2. Session("IMP_CODE")
- Implementation/company code
- Part of primary key for all records

### 3. Session("gLang")
- Language preference: "E" (English) or "C" (Chinese)
- Controls all text displays and messages

### 4. Session("pagemode")
- Current mode: "N" (New) or edit mode
- Determines field editability and available actions

### 5. Session("usr_pref_storer")
- User's preferred/default storer code
- Pre-selected in New mode

### 6. Session("SR_IS_Cable")
- Flag: "Y" or "N"
- Used by item lookup for cable filtering

### 7. Session("SR_IS_Scrap")
- Flag: "Y" or "N"
- Used by item lookup for scrap filtering

### 8. Session("SR_IS_NS")
- Flag: "Y" or "N"
- Used by item lookup for non-stock filtering

### 9. Session("gSelectLabel")
- Default label for dropdown select options

## ViewState Variables

### 1. ViewState("dt")
- DataTable containing grid data
- Persists grid state across postbacks

### 2. ViewState("n_cur_seq")
- Current/next sequence number for new items
- Increments with each new row

### 3. ViewState("RT_CODE")
- Stock return code
- Persists after save

### 4. ViewState("STORER_CODE")
- Storer code
- Used for various lookups

## JavaScript Functions

### 1. VendorLookUp()
Opens vendor lookup popup.
- **Target**: cms_search.aspx
- **Menu**: LOOKUP_VEND
- **Parameters**: Storer code

### 2. LocLookUp(lb_id, hd_id)
Opens location lookup popup.
- **Target**: locLookup.aspx
- **Parameters**: Warehouse, field IDs

### 3. DrumLookUp(hd_id, hd_id2)
Opens drum lookup popup.
- **Target**: cms_search.aspx
- **Menu**: LOOKUP_DRUM
- **Parameters**: Storer code, field IDs

### 4. ItemLookUp(STORER_CODE)
Opens item lookup popup.
- **Target**: cms_search.aspx
- **Menu**: LOOKUP_IM
- **Validation**: Requires storer code
- **Parameters**: Storer code, callback function

### 5. selectedItem()
Callback after item selection.
- Sets moduleAction to "SELECTIM"
- Submits form to process selected items

### 6. checkSB(STORER_CODE)
Opens stock balance inquiry.
- **Target**: cms_search.aspx
- **Menu**: INQ_001
- **Parameters**: Storer code

### 7. goToAttach(doc_type, doc_code, p_editmode)
Opens attachment management.
- **Target**: ATTACH_MAIN.ASPX
- **Parameters**: Document type, code, edit mode

### 8. OpenItemLbls()
Opens item label printing.
- **Target**: SR_LABEL/item_label_print.aspx
- **Parameters**: RT code, date, storer code

### 9. getLoad()
Shows loading modal popup.
- Displays "Posting...please wait" message

### 10. NewBtn()
Handles New button click.
- Prompts to save if in edit mode
- Sets new_action flag

### 11. maskKey(objEvent)
Input mask for numeric fields.
- Allows: digits (0-9) and decimal point (.)

### 12. maskDate(objEvent)
Input mask for date fields.
- Allows: digits (0-9) and forward slash (/)

## Error Handling

### Validation Errors
Displayed via `uiFun.displayMsg()`:
- Alert-style popup messages
- Bilingual support (English/Chinese)
- Prevents further processing

### Database Errors
Try-Catch blocks with:
- Transaction rollback
- Error message display
- Connection cleanup

### Common Error Scenarios

1. **Missing Required Fields**
   - Error displayed immediately
   - Focus remains on form

2. **Invalid Data Format**
   - Date format errors
   - Email format errors
   - Numeric format errors

3. **Business Rule Violations**
   - Duplicate batch numbers (auto-handled)
   - Invalid drum IDs
   - Serial number mismatches

4. **Posting Errors**
   - Insufficient stock during un-post
   - Transaction conflicts
   - Database constraint violations

## Internationalization (i18n)

### Supported Languages
1. **English (E)**
2. **Chinese (C)**

### Language Control
- Set via Session("gLang")
- Applies to all labels, buttons, messages

### Key Translations

#### English Labels
- Internal Return Code WMS
- Organizations
- Type
- Date
- Received By
- etc.

#### Chinese Labels
- 退貨單編號 (Return Code)
- 部門 (Department)
- 類型 (Type)
- 日期 (Date)
- 收貨者 (Receiver)
- etc.

### Message Translations
All error and confirmation messages have both English and Chinese versions.

## Technical Implementation Details

### Page Technology
- **Framework**: ASP.NET WebForms
- **Language**: VB.NET
- **Database**: SQL Server
- **AJAX**: UpdatePanels for partial page updates
- **Controls**: AjaxControlToolkit (Calendar, ComboBox, ModalPopup)

### Key Classes Used

#### 1. GlobalDBFunc (gDB)
Database operations:
- getConnection()
- getDataTable()
- amendData()
- getValueFromSQL()

#### 2. DBfunc (DB)
Utility functions:
- getDocNo(): Generates document numbers
- getValueFromSQL(): Executes scalar queries

#### 3. CommonUtils (cU)
Common operations:
- gfBuildDataTableforGridView(): Syncs grid to DataTable
- changeGVLabel(): Multilingual grid headers

#### 4. UIfunc (uiFun)
UI operations:
- load_dropdown(): Populates dropdowns
- load_ComboBox(): Populates combo boxes
- displayMsg(): Shows alert messages
- reOrderDetails(): Resequences items

#### 5. GeneralUtils (gU)
General utilities:
- dbEncode(): SQL injection prevention
- isValidDate(): Date validation
- isValidEmail(): Email validation
- isDecimal(): Numeric validation
- convdbDate(): Date conversion for SQL
- convdbNVCData(): Unicode string handling

#### 6. AccessRightUtils (ar)
Security:
- hideForm(): Controls field visibility
- hasBtnRight(): Checks button permissions
- sec_write, sec_viewMode: Mode flags

#### 7. StockTrans (st)
Stock transaction management:
- UpdateStockTrans(): Creates stock transactions
- UpdateStockBalTrans(): Updates balances
- UpdateStockSerialTrans(): Handles serials
- UnPostStocks(): Reverses transactions
- UnPostSERIAL(): Reverses serial transactions

#### 8. WMSFunc (wFun)
WMS-specific functions:
- getSerialNo(): Serial number generation

### Transaction Management
- SQL Server transactions ensure data consistency
- All posting operations are atomic
- Rollback on any error
- Connection cleanup in Finally blocks

### Grid Management
- DataTable in ViewState for persistence
- mFlag column tracks changes:
  - "N": New row
  - "M": Modified row
  - "D": Deleted row (hidden)
- Row-level operations preserve state

## Performance Considerations

### 1. Database Queries
- Indexed lookups on primary keys
- Filtered queries to minimize data
- Connection pooling

### 2. ViewState
- DataTable stored in ViewState
- Consider pagination for large datasets
- Clear ViewState on page transitions

### 3. Postbacks
- UpdatePanels reduce full postbacks
- Validation on client side where possible
- Async operations for lookups

### 4. Grid Rendering
- Only visible rows rendered
- Deleted rows hidden, not removed
- Lazy loading of combo box data

## Maintenance and Troubleshooting

### Common Issues

#### 1. "Cannot be empty" Errors
**Cause**: Required field validation
**Solution**: Fill all required fields (marked in red/yellow)

#### 2. Date Format Errors
**Cause**: Incorrect date format
**Solution**: Use dd/MM/yyyy format or calendar picker

#### 3. Posting Failures
**Causes**:
- Invalid batch numbers
- Missing locations
- Serial number issues
**Solutions**:
- Validate all items before posting
- Check batch number format (YYYYMMDD)
- Ensure serial numbers for serialized items

#### 4. Un-Post Permission Denied
**Cause**: User lacks required permissions
**Solution**: Must be in ADMIN_GP group or have BT_SR_UNPOST right

#### 5. Location Dropdown Empty
**Cause**: No bins defined for warehouse
**Solution**: Set up bins in WMS_WH_BIN for the warehouse

#### 6. Item Lookup Shows No Items
**Causes**:
- Storer code not selected
- Items not configured for storer
- Wrong RT_TYPE filters
**Solutions**:
- Select storer first
- Check WMS_ITEM configuration
- Verify RT_TYPE matches item types

### Debug Information

#### Enable Debug Mode
Check these session/viewstate values:
- Session("pagemode")
- Session("usr_id")
- ViewState("RT_CODE")
- RT_STATUS.Value

#### Log Locations
- Application logs: LAMSOON_DEV_LOG directory
- Error logs: ErrorLog.txt

#### Common Diagnostic Queries

```sql
-- Check stock return record
SELECT * FROM WMS_STOCK_RETURN 
WHERE RT_CODE = '[code]'

-- Check stock return items
SELECT * FROM WMS_STOCK_RETURN_D 
WHERE RT_CODE = '[code]'

-- Check stock transactions
SELECT * FROM WMS_STOCK_TRANS 
WHERE IO_DOC = 'SR' AND IO_DOC_ID = '[code]'

-- Check stock balances
SELECT * FROM WMS_ITEM_LOC_BAL 
WHERE ITM_CODE = '[item]' AND ILOC_LOC = '[loc]'
```

## Best Practices

### 1. Creating Stock Returns
- Always select Organization first
- Choose appropriate Type for automatic location selection
- Use "Get SI Items" if returning from Stock Issue
- Validate batch numbers follow YYYYMMDD format
- Fill expiry dates carefully

### 2. Approving Stock Returns
- Review all items before approval
- Ensure approval code is properly recorded
- Verify quantities match physical return

### 3. Posting Stock Returns
- Verify all validations pass
- Ensure locations are correct
- Check serial numbers if applicable
- Confirm stock balance updates after posting

### 4. Un-Posting Stock Returns
- Only un-post when absolutely necessary
- Verify sufficient stock exists before un-posting
- Save any pending work before un-posting
- Check for related transactions

### 5. Data Entry
- Use item lookup instead of manual entry
- Copy items feature for duplicate entries
- Regular saves to prevent data loss
- Check stock balance before completing

## Integration Points

### 1. Stock Issue (WMS_STOCK_ISSUE)
- Returns can reference original issue via RT_REF_DOC_NO
- "Get SI Items" populates from issue
- Posting updates issue with return info

### 2. Stock Balance (WMS_ITEM_LOC_BAL)
- Updated when posted
- Queried via "Check Stock Balance" button
- Critical for inventory accuracy

### 3. Attachment System
- Documents can be attached to returns
- Accessed via Attachment button
- Stored in separate attachment system

### 4. Label Printing
- Item labels can be printed
- Uses data from stock return
- Opens in new window

### 5. Item Master (WMS_ITEM)
- Source of item data
- Determines serial number requirements
- Provides UOM and pack key info

## Summary

The Stock Return page is a comprehensive module for managing the return of goods to the warehouse. It provides:

- **Full workflow** from creation to posting
- **Approval system** for controlled processing
- **Inventory integration** with automatic balance updates
- **Serial number tracking** for serialized items
- **Flexible item entry** via lookup or Stock Issue import
- **Audit trail** with complete change history
- **Bilingual support** for international use
- **Security controls** with role-based access

The page handles complex business logic including serial number generation, batch tracking, cable-specific handling, and transaction management, all while maintaining data integrity through proper validation and error handling.

## Document Version
- **Version**: 1.0
- **Last Updated**: 2025-11-12
- **Author**: System Documentation
- **Module**: INBOUND/SR
- **Page**: SRMain.aspx
