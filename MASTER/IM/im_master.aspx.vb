Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WebForms
Imports System.Text
Imports System.Drawing.Imaging
Imports System.Drawing.Printing


Partial Class MASTER_IM_im_master
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As New AccessRightUtils
    Private thumb As ThumbGenerator
    Private dt As New DataTable
    Private imp_code As String = ""
    Private exceptionEditList As List(Of String)
    Private DDFORMAT As String = ""
    Private fileSize As Double = 512000

    Private PHY_PS_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("PHY_PS_DIR")
    Private SYSP_TEMP_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("SYSP_TEMP_DIR")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        DDFORMAT = gU.getConfig("DDFORMATNO")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            Dim nSQL As String = ""

            If Session("gLang") = "E" Then
                nSQL = "select colc_code, colc_eng_value as colc_value from wms_col_code where colc_tabcol = 'WMS_ITEM.ITM_FLAGS'"
            ElseIf Session("gLang") = "C" Then
                nSQL = "select colc_code, colc_chi_value as colc_value from wms_col_code where colc_tabcol = 'WMS_ITEM.ITM_FLAGS'"
            End If

            If Session("pagemode") = "N" Then
                uiFun.load_dropdown(storer_code, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STORER_CODE <> 'PD' AND STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                uiFun.load_checkboxList(ITM_CHE_CLASS, "SELECT COLC_CODE, COLC_ENG_VALUE, COLC_DISPLAY_SEQ FROM WMS_COL_CODE WHERE COLC_TABCOL = 'WMS_ITEM.ITM_CHE_CLASS' ORDER BY COLC_DISPLAY_SEQ", "COLC_CODE", "COLC_ENG_VALUE", "")

                If storer_code.SelectedValue = "" Then
                    storer_code.SelectedValue = Session("usr_pref_storer")
                End If

                itm_type.SelectedValue = "OTHER"
            End If

            uiFun.load_dropdown(itm_pref_wh, "select DISTINCT WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"))
            uiFun.load_dropdown(itm_UOM, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
            uiFun.load_dropdown(ITM_UOM2, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
            'uiFun.load_dropdown(ITM_UOM2, "select colc_code, colc_eng_value as colc_value from wms_col_code where colc_tabcol = 'WMS_ITEM.ITM_UOM2' order by COLC_DISPLAY_SEQ", "colc_code", "colc_value", , Session("gSelectLabel"))
            uiFun.load_dropdown(ITEM_PRICE_CLASS, "select colc_code, colc_eng_value as colc_value from wms_col_code where colc_tabcol = 'WMS_ITEM.ITEM_PRICE_CLASS' order by COLC_DISPLAY_SEQ", "colc_code", "colc_value", , Session("gSelectLabel"))
            uiFun.load_dropdown(ITM_CAP_REV, "select colc_code, colc_eng_value as colc_value from wms_col_code where colc_tabcol = 'WMS_ITEM.ITM_CAP_REV' order by COLC_DISPLAY_SEQ", "colc_code", "colc_value", , Session("gSelectLabel"))
            uiFun.load_dropdown(ITM_CAP_VALUE, "select colc_code, colc_eng_value as colc_value from wms_col_code where colc_tabcol = 'WMS_ITEM.ITM_CAP_VALUE' order by COLC_DISPLAY_SEQ", "colc_code", "colc_value", , Session("gSelectLabel"))
            uiFun.load_dropdown(itm_cat, "Select PO_CATEGORY_CODE,PO_CATEGORY_DESC from PO_CATEGORY_MASTER ORDER BY 2", "PO_CATEGORY_CODE", "PO_CATEGORY_DESC", , Session("gSelectLabel"))
            'WMS_ITEM.ITEM_PRICE_CLASS
            uiFun.load_checkboxList(itm_flags, nSQL, "colc_code", "colc_value")
        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            itm_ref_price.ReadOnly = False
            itm_price.ReadOnly = False
            itm_ref_curr.ReadOnly = False
            btnWh.Visible = False
            'itm_balance.ReadOnly = False
            itm_balance.BackColor = Drawing.Color.Transparent
            itm_balance.BorderStyle = BorderStyle.None
            itm_balance.ReadOnly = True
        Else
            itm_ref_price.BackColor = Drawing.Color.Transparent
            itm_ref_price.BorderStyle = BorderStyle.None
            itm_price.BackColor = Drawing.Color.Transparent
            itm_price.BorderStyle = BorderStyle.None
            itm_ref_curr.BackColor = Drawing.Color.Transparent
            itm_ref_curr.BorderStyle = BorderStyle.None
            itm_balance.BackColor = Drawing.Color.Transparent
            itm_balance.BorderStyle = BorderStyle.None
            itm_ref_price.ReadOnly = True
            itm_price.ReadOnly = True
            itm_ref_curr.ReadOnly = True
            itm_balance.ReadOnly = True

        End If
        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Item Master"
            lbl_itm_code.Text = "Internal Item Code:"
            lbl_pack_key.Text = "Pack Key:"
            lbl_storer_code.Text = "Organizations:"
            lbl_itm_status.Text = "Status:"
            lbl_itm_barcode.Text = "Barcode:"

            'lbl_itm_vend_itm_no.Text = "Vendor Item No."

            lbl_itm_cat.Text = "Category:"
            lbl_itm_type.Text = "Type:"
            lbl_itm_parent.Text = "Parent:"
            lbl_ITM_KIT_QTY.text = "Child Qty:"
            lbl_itm_model.Text = "Style / Model:"
            lbl_itm_series.Text = "Series:"
            lbl_itm_name.Text = "Item Name:"
            lbl_itm_desc.Text = "Description:"
            lbl_itm_remarks.Text = "Remarks:"
            lbl_itm_remarks.Text = "Remarks:"
            lbl_itm_brand.Text = "Product Brand:"
            lbl_itm_pref_wh.Text = "Primary Subinventory:"
            lbl_itm_pref_loc.Text = "Pref. Location:"
            lbl_itm_ref_price.Text = "WholeSale Price:"
            lbl_itm_ref_curr.Text = "Ref. Currency:"
            lbl_itm_price.Text = "Retail Price:"
            lbl_itm_balance.Text = "Balance Qty:"
            lbl_itm_flags.Text = "Flags:"

            lbl_itm_UOM.Text = "Unit of Measure:"
            lbl_itm_pcs_per_uom.Text = "Number per UOM:"

            lbl_itm_qty_of_unit.Text = "Qty of Unit (Report Use):"
            lbl_itm_sku_no.Text = "Item Code:"
            lbl_proj_no.Text = "Project No."
            lbl_ITM_UOM2.Text = "Secondary UOM:"
            lbl_ITM_QTY2.Text = "Qty2:"
            lbl_ITM_GP_CODE.Text = "Group Code:"
            lbl_ITM_KEPT_DIV_CODE.Text = "Kept for Division:"
            lbl_ITM_DIV_CODE.Text = "Division Code:"
            lbl_ITM_EXPENSE_CODE.Text = "Expense Code:"
            lbl_ITM_MFG.Text = "MFG. Part No:"
            lbl_ITM_STORES_CLASS.Text = "Stores Class:"
            lbl_ITM_EMB.Text = "EMB (E/M/B):"
            lbl_ITM_HAZARD_CODE.Text = "Hazard Code:"
            lbl_ITM_DRAWING_NO.Text = "SK Code"
            lbl_ITM_CHE_CLASS.Text = "Chemical Class:"
            lbl_ITM_CRITICAL_YN.Text = "Critical Item (VC/C/NC):"
            lbl_ITM_RESTRICTED_ITEM.Text = "Restriction Item"
            lbl_ITM_WEIGHT_TYPE.Text = "Weight (Heavy/Light)"
            lbl_ITM_INT_ORDER_QTY.Text = "Initial Order Quantity"
            lbl_ITM_MAX_STOCK.Text = "Min Rec Shelf Life"
            lbl_ITM_1ST_INSP.Text = "First Inspection to be Done <br>(For Reference)"
            lbl_ITM_ROP_APL.Text = "ROP (Apleichau)"
            lbl_ITM_ROP_LAM.Text = "ROP (Lamma)"
            lbl_ITM_ORO_YN.Text = "ORO (Y/N)"
            lbl_ITM_ROP_CABLE.Text = "ROP (Cable Warehouse)"
            lbl_ITM_ROP_NP.Text = "ROP (North Point)"
            lbl_ITM_TI_YTD.Text = "Total Issue - YTD"
            lbl_ITM_TI_LYR.Text = "Total Issue - LYR"
            lbl_ITM_TI_PY2.Text = "Total Issue PY2"
            lbl_ITM_TI_PY3.Text = "Total Issue PY3"
            lbl_ITM_SHELF_LIFE.Text = "Shelf Life"
            lbl_ITM_PD_RCV_DATE.Text = "Date of Received"
            lbl_ITM_PD_CONTRACT_NO.Text = "Contract No."
            lbl_ITM_PD_ARR_NOTICE_NO.Text = "Arrival Notice No."
            lbl_ITM_PD_COND_OF_SPARES.Text = "Condition of spares"
            lbl_ITM_PD_ST_1_YEAR.Text = "To be stored for < 1 year or > 1 year"
            lbl_ITM_PD_ST_OVER_1_YEAR.Text = "Justification for Storage over 1 year"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            newrow.Text = "Add"
            CancelBtn.Text = "Cancel"

            lbl_ImageHd.Text = "Images"
            lbl_VendorList.Text = "Alt. Vendors Item"

            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            BtnCopy.OnClientClick = "return confirm(""Are you sure to save this record and clone a new item?"");"


            If Session("pagemode") = "N" Then
                itm_code.Text = "[No. will be auto generated]"
                itm_code.CssClass = "REQUIRED"
                pack_key.CssClass = "REQUIRED"
                storer_code.CssClass = "REQUIRED"
                itm_type.CssClass = "REQUIRED"
                itm_sku_no.CssClass = "REQUIRED"
            Else
                'itm_code.CssClass = ""
                'itm_code.ReadOnly = True
                'itm_code.BorderWidth = 0
                'itm_code.BackColor = Drawing.Color.Transparent

                pack_key.CssClass = ""
                pack_key.ReadOnly = True
                pack_key.BorderWidth = 0
                pack_key.BackColor = Drawing.Color.Transparent

                storer_code.CssClass = ""
            End If

            GridView1.EmptyDataText = "No Item Found."

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "物料主資料庫"
            lbl_itm_code.Text = "物料號碼:"
            lbl_pack_key.Text = "封裝內碼:"
            lbl_storer_code.Text = "部門:"
            lbl_itm_status.Text = "狀態:"

            lbl_itm_barcode.Text = "條碼:"

            'lbl_itm_vend_itm_no.Text = "供應商貨物編號:"

            lbl_itm_cat.Text = "類別:"
            lbl_itm_type.Text = "類型:"
            lbl_itm_parent.Text = "物料Parent:"
            lbl_ITM_KIT_QTY.text = "Qty"
            lbl_itm_model.Text = "Style / 型號:"
            lbl_itm_series.Text = "系列:"
            lbl_itm_name.Text = "貨物名稱:"
            lbl_itm_desc.Text = "說明:"
            lbl_itm_remarks.Text = "備註:"
            lbl_itm_brand.Text = "產品品牌:"
            lbl_itm_pref_wh.Text = "預設子庫存:"
            lbl_itm_pref_loc.Text = "優先地點:"
            lbl_itm_ref_price.Text = "WholeSale 價格:"
            lbl_itm_ref_curr.Text = "貨幣:"
            lbl_itm_price.Text = "Retail 價格:"
            lbl_itm_balance.Text = "結餘數量"
            lbl_itm_flags.Text = "Flags:"
            lbl_ITM_PD_RCV_DATE.Text = "Date of Received"
            lbl_ITM_PD_CONTRACT_NO.Text = "Contract No."
            lbl_ITM_PD_ARR_NOTICE_NO.Text = "Arrival Notice No."
            lbl_ITM_PD_COND_OF_SPARES.Text = "Condition of spares"
            lbl_ITM_PD_ST_1_YEAR.Text = "To be stored for < 1 year or > 1 year"
            lbl_ITM_PD_ST_OVER_1_YEAR.Text = "Justification for Storage over 1 year"

            lbl_itm_UOM.Text = "Unit of Measure:"
            lbl_itm_pcs_per_uom.Text = "Number per UOM:"

            lbl_itm_qty_of_unit.Text = "Qty of Unit (Report Use):"
            lbl_itm_sku_no.Text = "貨物編號"
            lbl_proj_no.Text = "項目編號"
            lbl_itm_remarks.Text = "備註"
            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "儲存"
            saveBtn2.Text = "儲存"
            newrow.Text = "新增"
            CancelBtn.Text = "取消"

            lbl_ImageHd.Text = "相片"
            lbl_VendorList.Text = "次供應商物料表"

            CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"
            BtnCopy.OnClientClick = "return confirm(""Are you sure to save this record and clone a new item?"");"

            If Session("pagemode") = "N" Then
                itm_code.Text = "[号码会自动产生]"

                'itm_code.AutoPostBack = True

            Else
                'itm_code.ReadOnly = True
                'itm_code.BorderWidth = 0
                'itm_code.BackColor = Drawing.Color.Transparent
            End If

            GridView1.EmptyDataText = "沒有資料記錄"
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        REM **********************

        If Not IsPostBack Then
            ViewState("im_dt") = Nothing
            ViewState("image1_name") = ""
            ViewState("image2_name") = ""
            ViewState("wms_itm_code") = ""

            ViewState("remove_image1") = False
            ViewState("remove_image2") = False

            ViewState("pack_key") = ""
            ViewState("storer_code") = ""

            Session.Remove(itm_picture1_upload.ClientID)
            Session.Remove(itm_picture2_upload.ClientID)

            Call BindGV()
        Else
            If itm_picture1_upload.HasFile Then
                Session(itm_picture1_upload.ClientID) = itm_picture1_upload.PostedFile
                itm_picture1_upload_lit.Text = itm_picture1_upload.PostedFile.FileName
                itm_picture1_upload.Visible = False
                itm_picture1_upload_lit.Visible = True
                itm_picture1_edit.Visible = True
            End If

            If itm_picture2_upload.HasFile Then
                Session(itm_picture2_upload.ClientID) = itm_picture2_upload.PostedFile
                itm_picture2_upload_lit.Text = itm_picture2_upload.PostedFile.FileName
                itm_picture2_upload.Visible = False
                itm_picture2_upload_lit.Visible = True
                itm_picture2_edit.Visible = True
            End If

            If Not Session("avi_dt") Is Nothing Then
                ViewState("im_dt") = Session("avi_dt")
                dt = ViewState("im_dt")
                GridView1.DataSource = dt
                GridView1.DataBind()
            Else
                dt = ViewState("im_dt")
            End If

            dsp_itm_pref_loc.Text = itm_pref_loc.Value
            dsp_itm_pref_loc2.Text = itm_pref_loc2.Value

        End If

        btnWh.Attributes.Add("onclick", "javascript:OpenWH(""" & ViewState("wms_itm_code") & """,'" & pack_key.Text.Trim & "');return false;")



        If itm_status.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
            itm_picture1_remove.Visible = False
            itm_picture2_remove.Visible = False
        End If

        'ar.sec_viewMode = "Y"
        setPageCtrlAccess()
        ar.hideForm(Me, editMode, , , exceptionEditList)

        ' ar.setFieldCustomize(Me, "MAST_IM", storer_code.SelectedValue, "WMS_ITEM", "D")

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here
                Call cU.changeGVLabel(oGridViewRow, e, "Vendor Code", "Vendor Code")
                Call cU.changeGVLabel(oGridViewRow, e, "Vendor Name", "Vendor Name")
                Call cU.changeGVLabel(oGridViewRow, e, "Vendor's Item Code", "Vendor's Item Code")
                Call cU.changeGVLabel(oGridViewRow, e, "Date", "Date")
                Call cU.changeGVLabel(oGridViewRow, e, "Status", "Status")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")

                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
            Case DataControlRowType.DataRow
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert)

        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow


                REM **********************
                REM Modify Here
                Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("vnd_code"), HyperLink).Text = DataBinder.Eval(e.Row.DataItem, "vnd_code").ToString.Trim

                CType(e.Row.FindControl("vnd_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "vnd_name").ToString.Trim
                CType(e.Row.FindControl("alv_vnd_itmcode"), Label).Text = DataBinder.Eval(e.Row.DataItem, "alv_vnd_itmcode").ToString.Trim
                'CType(e.Row.FindControl("alv_date_added"), Label).Text = cU.chgToYYYYMMDD(DataBinder.Eval(e.Row.DataItem, "alv_date_added").ToString.Trim)
                CType(e.Row.FindControl("alv_date_added"), Label).Text = DataBinder.Eval(e.Row.DataItem, "alv_date_added").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "vnd_code").ToString.Trim <> "" Then
                    CType(e.Row.FindControl("vnd_code"), HyperLink).NavigateUrl = "Javascript:openDetail('" & (e.Row.RowIndex).ToString.Trim & "','" & DataBinder.Eval(e.Row.DataItem, "vnd_code").ToString.Trim & "')"
                Else
                    CType(e.Row.FindControl("vnd_code"), HyperLink).NavigateUrl = "#"
                End If

                If DataBinder.Eval(e.Row.DataItem, "aitm_status").ToString.Trim <> "" Then
                    CType(e.Row.FindControl("btnChgStatus"), Button).Text = DataBinder.Eval(e.Row.DataItem, "aitm_status").ToString.Trim
                Else
                    CType(e.Row.FindControl("btnChgStatus"), Button).Text = "Active"
                End If

                CType(e.Row.FindControl("btnChgStatus"), Button).ID = "btnChgStatus" & e.Row.RowIndex

                REM **********************

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                nButton.ID = "btnDelete" & e.Row.RowIndex

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    nButton.Text = "删除"
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    Call ar.hideGVRow(GridView1, e.Row)
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                    'nButton.Enabled = False
                End If
        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        Session("im_itm_code") = itm_code.Text
        Session("im_pack_key") = pack_key.Text
        Session("im_storer_code") = storer_code.SelectedValue

        Response.Write("<script>window.open('altVendorItem.aspx?r=N&code=" & itm_code.Text & "', 'altIM', 'menubar=no,scrollbars=yes,resizable=yes,width=1255,height=820,left=5,top=10');</script>")
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        dt.AcceptChanges()
    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        'Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        Dim btnStatus As Button = CType(GridView1.Rows(e.RowIndex).FindControl(("btnChgStatus" & e.RowIndex.ToString)), Button)

        If btnStatus.Text.ToUpper = "INACTIVE" Then
            dt.Rows(e.RowIndex).Item("aitm_status") = "Active"
            btnStatus.Text = "Active"
        ElseIf btnStatus.Text.ToUpper = "ACTIVE" Then
            dt.Rows(e.RowIndex).Item("aitm_status") = "Inactive"
            btnStatus.Text = "Inactive"
        End If

        dt.AcceptChanges()
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""

        If itm_code.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_itm_code.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_itm_code.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If pack_key.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_pack_key.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_pack_key.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If storer_code.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_storer_code.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_storer_code.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If itm_sku_no.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_itm_sku_no.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_itm_sku_no.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If itm_type.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_itm_type.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_itm_type.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If itm_UOM.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_itm_UOM.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_itm_UOM.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If ITM_WEIGHT_TYPE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_ITM_WEIGHT_TYPE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_ITM_WEIGHT_TYPE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If


        If itm_pcs_per_uom.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_itm_pcs_per_uom.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_itm_pcs_per_uom.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If itm_pcs_per_uom.Text <> "" AndAlso Not IsNumeric(itm_pcs_per_uom.Text) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_itm_pcs_per_uom.Text & " is not a valid number!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_itm_pcs_per_uom.Text & " is not a valid number!", Session("gLang"))
            End If
            Return False
        End If

        If ITM_TEMP_FR.Text <> "" AndAlso Not IsNumeric(ITM_TEMP_FR.Text) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Temperature From is not a valid number!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "Temperature From is not a valid number!", Session("gLang"))
            End If
            Return False
        End If

        If ITM_TEMP_TO.Text <> "" AndAlso Not IsNumeric(ITM_TEMP_TO.Text) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Temperature To is not a valid number!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "Temperature To is not a valid number!", Session("gLang"))
            End If
            Return False
        End If


        'If String.IsNullOrWhiteSpace(ITM_EXPENSE_CODE.Text) Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_ITM_EXPENSE_CODE.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_ITM_EXPENSE_CODE.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        'If String.IsNullOrWhiteSpace(ITM_GP_CODE.Text) Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_ITM_GP_CODE.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_ITM_GP_CODE.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        If itm_pref_wh.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Please select " & lbl_itm_pref_wh.Text, Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_itm_pref_wh.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        'If ITM_RESTRICTED_ITEM.SelectedValue = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", "Please select " & lbl_ITM_RESTRICTED_ITEM.Text, Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_ITM_RESTRICTED_ITEM.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        'If ITM_ORO_YN.SelectedValue = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", "Please select " & lbl_ITM_ORO_YN.Text, Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_ITM_ORO_YN.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        'If String.IsNullOrWhiteSpace(ITM_COMMODITY_CODE.Text) Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_ITM_COMMODITY_CODE.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_ITM_COMMODITY_CODE.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If



        If Not String.IsNullOrWhiteSpace(ITM_REPLACEMENT_COST.Text) Then
            If Not gU.isDecimal(ITM_REPLACEMENT_COST.Text) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", lbl_ITM_REPLACEMENT_COST.Text & " is not a valid number!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", lbl_ITM_REPLACEMENT_COST.Text & " is not a valid number!", Session("gLang"))
                End If
                Return False
            End If
        End If

        'If ITM_RESTRICTED_RMK.Text.Length > 1000 Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_ITM_REPLACEMENT_COST.Text & " cannot exceed 1000 character!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_ITM_REPLACEMENT_COST.Text & " cannot exceed 1000 character!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        Dim fileExt As String = ""

        If ITM_PHOTO1.HasFile Then
            fileExt = Path.GetExtension(ITM_PHOTO1.FileName)

            If Not (UCase(fileExt) = ".JPG" OrElse UCase(fileExt) = ".GIF" OrElse UCase(fileExt) = ".PNG") Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ1.SelectedValue & " must be image files!(*.jpg, *.png, *.gif)", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ1.SelectedValue & " must be image files!(*.jpg, *.png, *.gif)", Session("gLang"))
                End If
                Return False
            End If


            If ITM_PHOTO1.PostedFile.ContentLength > fileSize Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ1.SelectedValue & " Size cannot larger than " & fileSize / 1024 & "KB!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ1.SelectedValue & " Size cannot larger than " & fileSize / 1024 & "KB!", Session("gLang"))
                End If
                Return False
            End If
        End If


        If ITM_PHOTO2.HasFile Then
            fileExt = Path.GetExtension(ITM_PHOTO2.FileName)

            If Not (UCase(fileExt) = ".JPG" OrElse UCase(fileExt) = ".GIF" OrElse UCase(fileExt) = ".PNG") Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ2.SelectedValue & " must be image files!(*.jpg, *.png, *.gif)", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ2.SelectedValue & " must be image files!(*.jpg, *.png, *.gif)", Session("gLang"))
                End If
                Return False
            End If

            If ITM_PHOTO2.PostedFile.ContentLength > fileSize Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ2.SelectedValue & " Size cannot larger than " & fileSize / 1024 & "KB!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ2.SelectedValue & " Size cannot larger than " & fileSize / 1024 & "KB!", Session("gLang"))
                End If
                Return False
            End If
        End If

        If ITM_PHOTO3.HasFile Then

            fileExt = Path.GetExtension(ITM_PHOTO3.FileName)

            If Not (UCase(fileExt) = ".JPG" OrElse UCase(fileExt) = ".GIF" OrElse UCase(fileExt) = ".PNG") Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ3.SelectedValue & " must be image files!(*.jpg, *.png, *.gif)", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ3.SelectedValue & " must be image files!(*.jpg, *.png, *.gif)", Session("gLang"))
                End If
                Return False
            End If

            If ITM_PHOTO3.PostedFile.ContentLength > fileSize Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ3.SelectedValue & " Size cannot larger than " & fileSize / 1024 & "KB!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "Photo " & IMG_SEQ3.SelectedValue & " Size cannot larger than " & fileSize / 1024 & "KB!", Session("gLang"))
                End If
                Return False
            End If
        End If
        'If remove_tr.Visible = False Then

        '    Dim nUP1 As HttpPostedFile = Nothing
        '    Dim nUP2 As HttpPostedFile = Nothing

        '    nUP1 = TryCast(Session(itm_picture1_upload.ClientID), HttpPostedFile)
        '    nUP2 = TryCast(Session(itm_picture2_upload.ClientID), HttpPostedFile)

        '    If nUP1 Is Nothing And Not itm_picture1_upload.HasFile Then
        '        If nUP2 Is Nothing And Not itm_picture2_upload.HasFile Then
        '            If Session("gLang") = "E" Then
        '                uiFun.displayMsg(Me, "", "You must upload an image!", Session("gLang"))
        '            Else
        '                uiFun.displayMsg(Me, "", "您必须上传图片!", Session("gLang"))
        '            End If

        '            Return False
        '        End If
        '    End If

        'End If

        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim gConn As SqlConnection
        Dim nextNo As String = ""
        Dim tmp_itm_flags As String = ""
        Dim nFileName1 As String = ""
        Dim nFileName2 As String = ""
        Dim tmpImpCode As String = ""
        Dim tempRemoveFile1 As Boolean = False
        Dim tempRemoveFile2 As Boolean = False

        Dim nUP1 As HttpPostedFile = Nothing
        Dim nUP2 As HttpPostedFile = Nothing

        If Not Session(itm_picture1_upload.ClientID) Is Nothing Then nUP1 = TryCast(Session(itm_picture1_upload.ClientID), HttpPostedFile)
        If Not Session(itm_picture2_upload.ClientID) Is Nothing Then nUP2 = TryCast(Session(itm_picture2_upload.ClientID), HttpPostedFile)

        If imp_code = "" Then tmpImpCode = "W" Else tmpImpCode = imp_code

        Dim old_filename1 As String = ViewState("image1_name")
        Dim old_filename2 As String = ViewState("image2_name")

        Dim tmpFInfo As IO.FileInfo

        Dim n_pack_key As String = ViewState("pack_key")
        Dim n_storer_code As String = ViewState("storer_code")

        Dim hasFile1 As New ArrayList
        Dim hasFile2 As New ArrayList

        Dim removeFile1 As New ArrayList
        Dim removeFile2 As New ArrayList

        Dim flags_data As String = ""

        Dim dupSQL As String
        Dim dupTbl As New DataTable

        Dim lITM_STACKABLE_YN, lITM_INSP_YN, lITM_SCRAP_YN, lITM_SERIAL_NO_YN, lITM_NONSTOCK_YN, lITM_CC, lITM_DG_YN, lITM_REQ_STORE_HUM_YN, lITM_REQ_STORE_AIRCON_YN, lITM_REQ_MSDS_YN, lITM_REQ_STORE_COLD_YN As String
        lITM_NONSTOCK_YN = "NULL"
        lITM_SCRAP_YN = "NULL"
        If itm_flags.SelectedIndex <> -1 Then
            For n As Integer = 0 To itm_flags.Items.Count - 1

                If itm_flags.Items(n).Selected Then
                    If flags_data <> "" Then flags_data = flags_data & ","

                    flags_data = flags_data & itm_flags.Items(n).Value
                End If
            Next
        Else
            flags_data = ""
        End If

        Dim isTemp As String = ""

        If ITM_TEMP_YN.Checked Then
            isTemp = "Y"
        Else
            isTemp = "N"
        End If


        If validateAll() Then
            gConn = gDB.getConnection()
            Dim transaction As SqlTransaction = Nothing

            Try
                gConn = gDB.getConnection()
                transaction = gConn.BeginTransaction()
                ' Start a local transaction

                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("IM", gConn, transaction)

                    'nextNo = itm_code.Text.Trim
                    REM **********************

                    REM **********************
                    REM Modify Here
                    REM char "N" "insert into wms_item " & _

                    dupSQL = "select 1 from wms_item " & _
                                "where itm_code = '" & gU.dbEncode(nextNo.Trim) & "' " & _
                                "and imp_code='" & gU.dbEncode(imp_code.Trim) & "' " & _
                                "and pack_key='" & gU.dbEncode(pack_key.Text.Trim) & "' " & _
                                "and storer_code='" & gU.dbEncode(storer_code.SelectedValue.Trim) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then

                        If nUP1 Is Nothing Then
                            If itm_picture1_upload.HasFile Then
                                tmpFInfo = New FileInfo(itm_picture1_upload.PostedFile.FileName)
                                nFileName1 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_IM_1_" & gU.decodeNullOrEmpty(nextNo.Trim, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.SelectedValue, "0") & tmpFInfo.Extension
                            End If
                        Else
                            tmpFInfo = New FileInfo(nUP1.FileName)
                            nFileName1 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_IM_1_" & gU.decodeNullOrEmpty(nextNo.Trim, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.SelectedValue, "0") & tmpFInfo.Extension
                        End If

                        If nUP2 Is Nothing Then
                            If itm_picture2_upload.HasFile Then
                                tmpFInfo = New FileInfo(itm_picture2_upload.PostedFile.FileName)
                                nFileName2 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_IM_2_" & gU.decodeNullOrEmpty(nextNo.Trim, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.SelectedValue, "0") & tmpFInfo.Extension
                            End If
                        Else
                            tmpFInfo = New FileInfo(nUP2.FileName)
                            nFileName2 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_IM_2_" & gU.decodeNullOrEmpty(nextNo.Trim, "0") & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.SelectedValue, "0") & tmpFInfo.Extension
                        End If

			if ITM_TYPE.text = "CABLE" then
				ITM_SERIAL_NO_YN.Checked = true
			end If

                        If ITM_SERIAL_NO_YN.Checked Then lITM_SERIAL_NO_YN = "'Y'" Else lITM_SERIAL_NO_YN = "NULL"
                        If radLotExist.Checked Then
                            lITM_STACKABLE_YN = "'Y'"
                        ElseIf radLotNotExist.Checked Then
                            lITM_STACKABLE_YN = "'N'"
                        Else
                            lITM_STACKABLE_YN = "NULL"
                        End If
                        'If ITM_STACKABLE_YN.Checked Then lITM_STACKABLE_YN = "'Y'" Else lITM_STACKABLE_YN = "NULL"
                        If ITM_INSP_YN.Checked Then lITM_INSP_YN = "'Y'" Else lITM_INSP_YN = "'N'"
                        'If ITM_SCRAP_YN.Checked Then lITM_SCRAP_YN = "'Y'" Else lITM_SCRAP_YN = "NULL"
                        'If ITM_NONSTOCK_YN.Checked Then lITM_NONSTOCK_YN = "'Y'" Else lITM_NONSTOCK_YN = "'N'"
                        If ITM_CC.Checked Then lITM_CC = "'Y'" Else lITM_CC = "NULL"
                        If ITM_DG_YN.Checked Then lITM_DG_YN = "'Y'" Else lITM_DG_YN = "NULL"
                        If ITM_REQ_STORE_COLD_YN.Checked Then lITM_REQ_STORE_COLD_YN = "'Y'" Else lITM_REQ_STORE_COLD_YN = "NULL"
                        If ITM_REQ_STORE_HUM_YN.Checked Then lITM_REQ_STORE_HUM_YN = "'Y'" Else lITM_REQ_STORE_HUM_YN = "NULL"
                        If ITM_REQ_STORE_AIRCON_YN.Checked Then lITM_REQ_STORE_AIRCON_YN = "'Y'" Else lITM_REQ_STORE_AIRCON_YN = "NULL"
                        If ITM_REQ_MSDS_YN.Checked Then lITM_REQ_MSDS_YN = "'Y'" Else lITM_REQ_MSDS_YN = "NULL"


                        ',ITM_STACKABLE_YN,ITM_INSP_YN,Tex,ITM_SCRAP_YN,

                        sql_string = "insert into wms_item " &
                                      "(itm_code, imp_code, pack_key, storer_code, " &
                                      "itm_vend_itm_no, itm_barcode, itm_name, itm_desc, " &
                                      "itm_cat, itm_type, itm_brand, " &
                                      "itm_pref_wh, itm_pref_loc, itm_ref_curr, " &
                                      "itm_ref_price, itm_price, itm_balance, itm_flags, " &
                                      "itm_remarks, itm_series, itm_series_no, itm_model, " &
                                      "itm_parent, ITM_KIT_QTY, itm_status, " &
                                      "itm_picture1, itm_picture2, " &
                                      "itm_temp_yn, itm_temp_fr, itm_temp_to, itm_uom, " &
                                      "itm_pcs_per_uom, itm_qty_of_unit, itm_sku_no, " &
                                      "ITM_PD_RCV_DATE,ITM_PD_CONTRACT_NO,ITM_PD_ARR_NOTICE_NO,ITM_PD_COND_OF_SPARES,ITM_PD_ST_1_YEAR,ITM_PD_ST_OVER_1_YEAR," &
                                      "ITM_SIZE,ITM_COLOR_CODE,ITM_SERIAL_NO,ITM_PART_NO,ITM_PROD_NO, ITM_PROD_GROUP, itm_pref_loc2," &
                                      "ITM_SERIAL_NO_YN,ITM_STACKABLE_YN,ITM_INSP_YN,ITM_SCRAP_YN,ITEM_PRICE_CLASS,ITM_CAP_REV, ITM_CAP_VALUE, ITM_NONSTOCK_YN,ITM_CC," &
                                      "sys_cb, sys_cd, sys_lub, sys_lud, " &
                                        "proj_no," &
                                        "ITM_SHELF_LIFE," &
                                        "ITM_GP_CODE," &
                                        "ITM_DIV_CODE," &
                                        "ITM_MFG," &
                                        "ITM_STORES_CLASS," &
                                        "ITM_EMB," &
                                        "ITM_HAZARD_CODE," &
                                        "ITM_CRITICAL_YN," &
                                        "ITM_MAX_STOCK," &
                                        "ITM_RESTRICTED_ITEM," &
                                        "ITM_DRAWING_NO," &
                                        "ITM_KEPT_DIV_CODE," &
                                        "ITM_EXPENSE_CODE," &
                                        "ITM_INT_ORDER_QTY," &
                                        "ITM_1ST_INSP," &
                                        "ITM_ORO_YN," &
                                        "ITM_ROP_APL," &
                                        "ITM_ROP_LAM," &
                                        "ITM_ROP_CABLE," &
                                        "ITM_ROP_NP," &
                                        "ITM_TI_YTD," &
                                        "ITM_TI_LYR," &
                                        "ITM_TI_PY2," &
                                        "ITM_TI_PY3," &
                                        "ITM_UOM2," &
                                        "ITM_QTY2," &
                                        "ITM_DG_YN," &
                                        "ITM_DG_CAT," &
                                        "ITM_REQ_STORE_HUM_DESC," &
                                        "ITM_REQ_STORE_COLD_YN," &
                                        "ITM_REQ_STORE_HUM_YN," &
                                        "ITM_REQ_STORE_AIRCON_YN," &
                                        "ITM_CHE_CLASS," &
                                        "ITM_REQ_MSDS_YN," &
                                        "ITM_WEIGHT_TYPE," &
                                        "ITM_MANUFACTURER, ITM_REPLACEMENT_COST, ITM_COMMODITY_CODE, ITM_RESTRICTED_RMK " &
                                      ") " &
                                      "values ( " &
                                      gU.convdbNVCData(gU.dbEncode(nextNo.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(imp_code)) & "," & gU.convdbNVCData(gU.dbEncode(pack_key.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(storer_code.SelectedValue.Trim)) & "," &
                                      gU.convdbNVCData(gU.dbEncode(itm_vend_itm_no.Value)) & "," & gU.convdbNVCData(gU.dbEncode(itm_barcode.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_name.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_desc.Text.Trim)) & ", " &
                                      gU.convdbNVCData(gU.dbEncode(itm_cat.SelectedValue.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_type.SelectedValue.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_brand.Text.Trim)) & "," &
                                      gU.convdbNVCData(gU.dbEncode(itm_pref_wh.SelectedValue.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_pref_loc.Value.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_ref_curr.Text.Trim)) & "," &
                                      gU.dbEncode(gU.decodeNullOrEmpty(itm_ref_price.Text.Trim, "0")) & "," & gU.dbEncode(gU.decodeNullOrEmpty(itm_price.Text.Trim, "0")) & "," & gU.dbEncode(gU.decodeNullOrEmpty(itm_balance.Text.Trim, "0")) & "," & gU.convdbNVCData(gU.dbEncode(flags_data)) & "," &
                                      gU.convdbNVCData(gU.dbEncode(itm_remarks.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_series.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_series_no.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_model.Text.Trim)) & "," &
                                      gU.convdbNVCData(gU.dbEncode(itm_parent.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(ITM_KIT_QTY.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(itm_status.Text.Trim)) & ", " &
                                      gU.convdbNVCData(gU.dbEncode(nFileName1.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(nFileName2.Trim)) & "," &
                                      gU.convdbNVCData(isTemp) & "," & gU.convdbNVCData(gU.dbEncode(ITM_TEMP_FR.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(ITM_TEMP_TO.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_UOM.SelectedValue)) & "," &
                                      gU.convdbNVCData(gU.dbEncode(itm_pcs_per_uom.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_qty_of_unit.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(itm_sku_no.Text.Trim)) & "," &
                                      "Convert(DateTime, " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(ITM_PD_RCV_DATE.Text.Trim, ""))) & ", " & DDFORMAT & ") " & "," & gU.convdbNVCData(gU.dbEncode(ITM_PD_CONTRACT_NO.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(ITM_PD_ARR_NOTICE_NO.Text.Trim)) & "," &
                                      gU.convdbNVCData(gU.dbEncode(ITM_PD_COND_OF_SPARES.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(ITM_PD_ST_1_YEAR.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(ITM_PD_ST_OVER_1_YEAR.Text.Trim)) & "," &
                                      gU.convdbNVCData(gU.dbEncode(ITM_SIZE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(ITM_COLOR_CODE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(ITM_SERIAL_NO.Text.Trim)) & "," &
                                      gU.convdbNVCData(gU.dbEncode(ITM_PART_NO.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(ITM_PROD_NO.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(ITM_PROD_GROUP.Text.Trim)) & "," &
                                      gU.convdbNVCData(gU.dbEncode(itm_pref_loc2.Value)) & "," & lITM_SERIAL_NO_YN & "," & lITM_STACKABLE_YN & "," & lITM_INSP_YN & "," & lITM_SCRAP_YN & "," & gU.convdbNVCData(gU.dbEncode(ITEM_PRICE_CLASS.SelectedValue)) & "," &
                                      gU.convdbNVCData(gU.dbEncode(ITM_CAP_REV.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(ITM_CAP_VALUE.SelectedValue)) & "," &
                                      lITM_NONSTOCK_YN & "," & lITM_CC & "," &
                                      "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate(), " &
                                    gU.convdbNVCData(gU.dbEncode(proj_no.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_SHELF_LIFE.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_GP_CODE.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_DIV_CODE.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_MFG.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_STORES_CLASS.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_EMB.SelectedValue.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_HAZARD_CODE.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_CRITICAL_YN.SelectedValue.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_MAX_STOCK.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_RESTRICTED_ITEM.SelectedValue.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_DRAWING_NO.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_KEPT_DIV_CODE.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_EXPENSE_CODE.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_INT_ORDER_QTY.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_1ST_INSP.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_ORO_YN.SelectedValue.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_ROP_APL.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_ROP_LAM.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_ROP_CABLE.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_ROP_NP.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_TI_YTD.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_TI_LYR.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_TI_PY2.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_TI_PY3.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_UOM2.SelectedValue.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_QTY2.Text.Trim)) & "," &
                                    lITM_DG_YN & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_DG_CAT.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_REQ_STORE_HUM_DESC.Text.Trim)) & "," &
                                    lITM_REQ_STORE_COLD_YN & "," &
                                    lITM_REQ_STORE_HUM_YN & "," &
                                    lITM_REQ_STORE_AIRCON_YN & "," &
                                    gU.convdbNVCData(gU.dbEncode(gU.getChkBoxListValue(ITM_CHE_CLASS))) & "," &
                                    lITM_REQ_MSDS_YN & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_WEIGHT_TYPE.SelectedValue.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_MANUFACTURER.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_REPLACEMENT_COST.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_COMMODITY_CODE.Text.Trim)) & "," &
                                    gU.convdbNVCData(gU.dbEncode(ITM_RESTRICTED_RMK.Text.Trim)) &
                                    ") "

                        '"itm_name_ch, itm_desc_ch, " & _
                        'gU.dbEncode(itm_name_ch.Text) & "'," & gU.dbEncode(itm_name_ch.Text) & "', " & _

                        REM **********************

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
                            For Each rows As DataRow In dt.Rows

                                dupSQL = "select 1 from wms_alt_vend_item " & _
                                   "where itm_code = '" & gU.dbEncode(nextNo.Trim) & "' " & _
                                   "and imp_code='" & gU.dbEncode(imp_code) & "' " & _
                                   "and pack_key='" & gU.dbEncode(pack_key.Text.Trim) & "' " & _
                                   "and storer_code='" & gU.dbEncode(storer_code.SelectedValue.Trim) & "' " & _
                                   "and vnd_code = '" & gU.dbEncode(rows.Item("vnd_code").ToString.Trim) & "'"

                                dupTbl = New DataTable

                                dupTbl = gDB.getDataTable(dupSQL)

                                If dupTbl.Rows.Count = 0 Then
                                    itemSQL = ""
                                    REM **********************
                                    REM Modify Here
                                    Select Case rows.Item("mFlag")
                                        Case "N"
                                            itemSQL = "insert into wms_alt_vend_item " & _
                                                          "(itm_code, imp_code, storer_code, " & _
                                                          "pack_key, vnd_code, vnd_name, " & _
                                                          "alv_vnd_itmcode, alv_date_added, aitm_name, aitm_desc, " & _
                                                          "aitm_spec, aitm_uom, aitm_pcs_per_pack, aitm_qty_per_ctn, " & _
                                                          "aitm_length, aitm_width, aitm_hight, " & _
                                                          "aitm_cbm, aitm_vol, aitm_ref_curr, " & _
                                                          "aitm_ref_price, aitm_remarks, aitm_pref_wh, " & _
                                                          "aitm_pref_loc, aitm_picture1, aitm_picture2, " & _
                                                          "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                          "values ( " & _
                                                             gU.convdbNVCData(gU.dbEncode(nextNo.Trim)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(storer_code.SelectedValue.Trim)) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(pack_key.Text.Trim)) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("vnd_code").ToString.Trim, ""))) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("vnd_name").ToString.Trim, ""))) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("alv_vnd_itmcode").ToString.Trim, ""))) & ", " & _
                                                             "CONVERT(datetime, " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("alv_date_added").ToString.Trim, ""))) & "," & DDFORMAT & ") , " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_name").ToString.Trim, ""))) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_desc").ToString.Trim, ""))) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_spec").ToString.Trim, ""))) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(itm_UOM.SelectedValue, ""))) & ", " & _
                                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_pcs_per_pack").ToString.Trim, "0")) & ", " & _
                                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_qty_per_ctn").ToString.Trim, "1")) & ", " & _
                                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_length").ToString.Trim, "0")) & ", " & _
                                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_width").ToString.Trim, "0")) & ", " & _
                                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_hight").ToString.Trim, "0")) & ", " & _
                                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_cbm").ToString.Trim, "0")) & ", " & _
                                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_vol").ToString.Trim, "0")) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_ref_curr").ToString.Trim, ""))) & ", " & _
                                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_ref_price").ToString.Trim, "0")) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_remarks").ToString.Trim, ""))) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_pref_wh").ToString.Trim, ""))) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_pref_loc").ToString.Trim, ""))) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_picture1").ToString.Trim, ""))) & ", " & _
                                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_picture2").ToString.Trim, ""))) & ", " & _
                                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                            'aitm_name_ch, aitm_desc_ch,
                                            'gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_name_ch").ToString.Trim, ""))) & ", " & _
                                            'gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_desc_ch").ToString.Trim, ""))) & ", " & _

                                    End Select
                                    REM **********************

                                    If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                                Else
                                    If Not transaction Is Nothing Then
                                        transaction.Rollback()
                                        transaction = Nothing
                                    End If

                                    If Session("gLang") = "E" Then
                                        uiFun.displayMsg(Me, "", "Duplicate record has found in " & lbl_VendorList.Text & "!!", Session("gLang"))
                                    Else
                                        uiFun.displayMsg(Me, "", lbl_VendorList.Text & "資料重複!!", Session("gLang"))
                                    End If

                                    Exit Sub
                                End If
                            Next

                        End If
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in " & lheader.Text & "!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", lheader.Text & "資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    If nUP1 Is Nothing Then
                        If itm_picture1_upload.HasFile Then
                            tmpFInfo = New FileInfo(itm_picture1_upload.PostedFile.FileName)
                            nFileName1 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_IM_1_" & itm_code.Text.Trim & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.SelectedValue, "0") & tmpFInfo.Extension
                        End If
                    Else
                        tmpFInfo = New FileInfo(nUP1.FileName)
                        nFileName1 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_IM_1_" & itm_code.Text.Trim & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.SelectedValue, "0") & tmpFInfo.Extension
                    End If

                    If nUP2 Is Nothing Then
                        If itm_picture2_upload.HasFile Then
                            tmpFInfo = New FileInfo(itm_picture2_upload.PostedFile.FileName)
                            nFileName2 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_IM_2_" & itm_code.Text.Trim & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.SelectedValue, "0") & tmpFInfo.Extension
                        End If
                    Else
                        tmpFInfo = New FileInfo(nUP2.FileName)
                        nFileName2 = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_IM_2_" & itm_code.Text.Trim & gU.decodeNullOrEmpty(pack_key.Text, "0") & gU.decodeNullOrEmpty(storer_code.SelectedValue, "0") & tmpFInfo.Extension
                    End If

			if ITM_TYPE.text = "CABLE" then
				ITM_SERIAL_NO_YN.Checked = true
			end If

                    If ITM_SERIAL_NO_YN.Checked Then lITM_SERIAL_NO_YN = "'Y'" Else lITM_SERIAL_NO_YN = "NULL"
                    If radLotExist.Checked Then
                        lITM_STACKABLE_YN = "'Y'"
                    ElseIf radLotNotExist.Checked Then
                        lITM_STACKABLE_YN = "'N'"
                    Else
                        lITM_STACKABLE_YN = "NULL"
                    End If
                    'If ITM_STACKABLE_YN.Checked Then lITM_STACKABLE_YN = "'Y'" Else lITM_STACKABLE_YN = "NULL"
                    If ITM_INSP_YN.Checked Then lITM_INSP_YN = "'Y'" Else lITM_INSP_YN = "'N'"
                    'If ITM_SCRAP_YN.Checked Then lITM_SCRAP_YN = "'Y'" Else lITM_SCRAP_YN = "NULL"
                    'If ITM_NONSTOCK_YN.Checked Then lITM_NONSTOCK_YN = "'Y'" Else lITM_NONSTOCK_YN = "'N'"
                    If ITM_CC.Checked Then lITM_CC = "'Y'" Else lITM_CC = "NULL"
                    If ITM_DG_YN.Checked Then lITM_DG_YN = "'Y'" Else lITM_DG_YN = "NULL"
                    If ITM_REQ_STORE_COLD_YN.Checked Then lITM_REQ_STORE_COLD_YN = "'Y'" Else lITM_REQ_STORE_COLD_YN = "NULL"
                    If ITM_REQ_STORE_HUM_YN.Checked Then lITM_REQ_STORE_HUM_YN = "'Y'" Else lITM_REQ_STORE_HUM_YN = "NULL"
                    If ITM_REQ_STORE_AIRCON_YN.Checked Then lITM_REQ_STORE_AIRCON_YN = "'Y'" Else lITM_REQ_STORE_AIRCON_YN = "NULL"
                    If ITM_REQ_MSDS_YN.Checked Then lITM_REQ_MSDS_YN = "'Y'" Else lITM_REQ_MSDS_YN = "NULL"

                    sql_string = "update wms_item set " &
                                "itm_status = " & gU.convdbNVCData(gU.dbEncode(itm_status.Text)) & ", " &
                                "itm_vend_itm_no = " & gU.convdbNVCData(gU.dbEncode(itm_vend_itm_no.Value)) & ", " &
                                "itm_barcode = " & gU.convdbNVCData(gU.dbEncode(itm_barcode.Text)) & ", " &
                                "itm_name = " & gU.convdbNVCData(gU.dbEncode(itm_name.Text)) & ", " &
                                "itm_desc = " & gU.convdbNVCData(gU.dbEncode(itm_desc.Text)) & ", " &
                                "itm_cat = " & gU.convdbNVCData(gU.dbEncode(itm_cat.SelectedValue)) & ", " &
                                "itm_type = " & gU.convdbNVCData(gU.dbEncode(itm_type.SelectedValue)) & ", " &
                                "itm_brand = " & gU.convdbNVCData(gU.dbEncode(itm_brand.Text)) & ", " &
                                "itm_pref_wh = " & gU.convdbNVCData(gU.dbEncode(itm_pref_wh.SelectedValue)) & ", " &
                                "itm_pref_loc = " & gU.convdbNVCData(gU.dbEncode(itm_pref_loc.Value)) & ", " &
                                "itm_pref_loc2 = " & gU.convdbNVCData(gU.dbEncode(itm_pref_loc2.Value)) & ", " &
                                "itm_ref_curr = " & gU.convdbNVCData(gU.dbEncode(itm_ref_curr.Text)) & ", " &
                                "itm_ref_price =  " & gU.dbEncode(gU.decodeNullOrEmpty(itm_ref_price.Text, "0")) & ", " &
                                "itm_price =  " & gU.dbEncode(gU.decodeNullOrEmpty(itm_price.Text, "0")) & ", " &
                                "itm_balance =  " & gU.dbEncode(gU.decodeNullOrEmpty(itm_balance.Text, "0")) & ", " &
                                "itm_flags = " & gU.convdbNVCData(gU.dbEncode(flags_data)) & ", " &
                                "itm_remarks = " & gU.convdbNVCData(gU.dbEncode(itm_remarks.Text)) & ", " &
                                "itm_series = " & gU.convdbNVCData(gU.dbEncode(itm_series.Text)) & ", " &
                                "itm_series_no = " & gU.convdbNVCData(gU.dbEncode(itm_series_no.Text)) & ", " &
                                "itm_model = " & gU.convdbNVCData(gU.dbEncode(itm_model.Text)) & ", " &
                                "itm_parent = " & gU.convdbNVCData(gU.dbEncode(itm_parent.Text)) & ", " &
                                "ITM_KIT_QTY = " & gU.convdbNVCData(gU.dbEncode(ITM_KIT_QTY.Text)) & ", " &
                                "itm_temp_yn = " & gU.convdbNVCData(gU.dbEncode(isTemp)) & ", " &
                                "itm_temp_fr = " & gU.convdbNVCData(gU.dbEncode(ITM_TEMP_FR.Text)) & ", " &
                                "itm_temp_to = " & gU.convdbNVCData(gU.dbEncode(ITM_TEMP_TO.Text)) & ", " &
                                "itm_uom = " & gU.convdbNVCData(gU.dbEncode(itm_UOM.SelectedValue)) & ", " &
                                "itm_pcs_per_uom = " & gU.convdbNVCData(gU.dbEncode(itm_pcs_per_uom.Text.Trim)) & ", " &
                                "itm_qty_of_unit = " & gU.convdbNVCData(gU.dbEncode(itm_qty_of_unit.Text.Trim)) & ", " &
                                "itm_sku_no = " & gU.convdbNVCData(gU.dbEncode(itm_sku_no.Text.Trim)) & ", " &
                                "ITM_PD_RCV_DATE = Convert(DateTime, " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(ITM_PD_RCV_DATE.Text.Trim, ""))) & ", " & DDFORMAT & "), " &
                                "ITM_PD_CONTRACT_NO = " & gU.convdbNVCData(gU.dbEncode(ITM_PD_CONTRACT_NO.Text.Trim)) & ", " &
                                "ITM_PD_ARR_NOTICE_NO = " & gU.convdbNVCData(gU.dbEncode(ITM_PD_ARR_NOTICE_NO.Text.Trim)) & ", " &
                                "ITM_PD_COND_OF_SPARES = " & gU.convdbNVCData(gU.dbEncode(ITM_PD_COND_OF_SPARES.Text.Trim)) & ", " &
                                "ITM_PD_ST_1_YEAR = " & gU.convdbNVCData(gU.dbEncode(ITM_PD_ST_1_YEAR.Text.Trim)) & ", " &
                                "ITM_PD_ST_OVER_1_YEAR = " & gU.convdbNVCData(gU.dbEncode(ITM_PD_ST_OVER_1_YEAR.Text.Trim)) & ", " &
                    "ITM_SIZE = " & gU.convdbNVCData(gU.dbEncode(ITM_SIZE.Text.Trim)) & ", " &
                    "ITM_COLOR_CODE = " & gU.convdbNVCData(gU.dbEncode(ITM_COLOR_CODE.Text.Trim)) & ", " &
                    "ITM_SERIAL_NO = " & gU.convdbNVCData(gU.dbEncode(ITM_SERIAL_NO.Text.Trim)) & ", " &
                    "ITM_PART_NO = " & gU.convdbNVCData(gU.dbEncode(ITM_PART_NO.Text.Trim)) & ", " &
                    "ITM_PROD_NO = " & gU.convdbNVCData(gU.dbEncode(ITM_PROD_NO.Text.Trim)) & ", " &
                    "ITM_PROD_GROUP=" & gU.convdbNVCData(gU.dbEncode(ITM_PROD_GROUP.Text.Trim)) & ", " &
                    "ITM_SERIAL_NO_YN=" & lITM_SERIAL_NO_YN & "," &
                    "ITM_STACKABLE_YN=" & lITM_STACKABLE_YN & "," &
                    "ITEM_PRICE_CLASS=" & gU.convdbNVCData(gU.dbEncode(ITEM_PRICE_CLASS.SelectedValue)) & "," &
                    "ITM_CAP_REV=" & gU.convdbNVCData(gU.dbEncode(ITM_CAP_REV.SelectedValue)) & "," &
                    "ITM_CAP_VALUE=" & gU.convdbNVCData(gU.dbEncode(ITM_CAP_VALUE.SelectedValue)) & "," &
                    "ITM_INSP_YN=" & lITM_INSP_YN & "," &
                    "ITM_SCRAP_YN=" & lITM_SCRAP_YN & "," &
                    "ITM_NONSTOCK_YN=" & lITM_NONSTOCK_YN & "," &
                    "ITM_CC=" & lITM_CC & ", " &
                    "proj_no=" & gU.convdbNVCData(gU.dbEncode(proj_no.Text.Trim)) & ", " &
                    "ITM_SHELF_LIFE=" & gU.convdbNVCData(gU.dbEncode(ITM_SHELF_LIFE.Text.Trim)) & ", " &
                    "ITM_GP_CODE=" & gU.convdbNVCData(gU.dbEncode(ITM_GP_CODE.Text.Trim)) & ", " &
                    "ITM_DIV_CODE=" & gU.convdbNVCData(gU.dbEncode(ITM_DIV_CODE.Text.Trim)) & ", " &
                    "ITM_MFG=" & gU.convdbNVCData(gU.dbEncode(ITM_MFG.Text.Trim)) & ", " &
                    "ITM_STORES_CLASS=" & gU.convdbNVCData(gU.dbEncode(ITM_STORES_CLASS.Text.Trim)) & ", " &
                    "ITM_EMB=" & gU.convdbNVCData(gU.dbEncode(ITM_EMB.SelectedValue.Trim)) & ", " &
                    "ITM_HAZARD_CODE=" & gU.convdbNVCData(gU.dbEncode(ITM_HAZARD_CODE.Text.Trim)) & ", " &
                    "ITM_CRITICAL_YN=" & gU.convdbNVCData(gU.dbEncode(ITM_CRITICAL_YN.SelectedValue.Trim)) & ", " &
                    "ITM_MAX_STOCK=" & gU.convdbNVCData(gU.dbEncode(ITM_MAX_STOCK.Text.Trim)) & ", " &
                    "ITM_RESTRICTED_ITEM=" & gU.convdbNVCData(gU.dbEncode(ITM_RESTRICTED_ITEM.SelectedValue.Trim)) & ", " &
                    "ITM_DRAWING_NO=" & gU.convdbNVCData(gU.dbEncode(ITM_DRAWING_NO.Text.Trim)) & ", " &
                    "ITM_KEPT_DIV_CODE=" & gU.convdbNVCData(gU.dbEncode(ITM_KEPT_DIV_CODE.Text.Trim)) & ", " &
                    "ITM_EXPENSE_CODE=" & gU.convdbNVCData(gU.dbEncode(ITM_EXPENSE_CODE.Text.Trim)) & ", " &
                    "ITM_INT_ORDER_QTY=" & gU.convdbNVCData(gU.dbEncode(ITM_INT_ORDER_QTY.Text.Trim)) & ", " &
                    "ITM_1ST_INSP=" & gU.convdbNVCData(gU.dbEncode(ITM_1ST_INSP.Text.Trim)) & ", " &
                    "ITM_ORO_YN=" & gU.convdbNVCData(gU.dbEncode(ITM_ORO_YN.SelectedValue.Trim)) & ", " &
                    "ITM_ROP_APL=" & gU.convdbNVCData(gU.dbEncode(ITM_ROP_APL.Text.Trim)) & ", " &
                    "ITM_ROP_LAM=" & gU.convdbNVCData(gU.dbEncode(ITM_ROP_LAM.Text.Trim)) & ", " &
                    "ITM_ROP_CABLE=" & gU.convdbNVCData(gU.dbEncode(ITM_ROP_CABLE.Text.Trim)) & ", " &
                    "ITM_ROP_NP=" & gU.convdbNVCData(gU.dbEncode(ITM_ROP_NP.Text.Trim)) & ", " &
                    "ITM_TI_YTD=" & gU.convdbNVCData(gU.dbEncode(ITM_TI_YTD.Text.Trim)) & ", " &
                    "ITM_TI_LYR=" & gU.convdbNVCData(gU.dbEncode(ITM_TI_LYR.Text.Trim)) & ", " &
                    "ITM_TI_PY2=" & gU.convdbNVCData(gU.dbEncode(ITM_TI_PY2.Text.Trim)) & ", " &
                    "ITM_TI_PY3=" & gU.convdbNVCData(gU.dbEncode(ITM_TI_PY3.Text.Trim)) & ", " &
                    "ITM_UOM2=" & gU.convdbNVCData(gU.dbEncode(ITM_UOM2.SelectedValue.Trim)) & ", " &
                    "ITM_QTY2=" & gU.convdbNVCData(gU.dbEncode(ITM_QTY2.Text.Trim)) & ", " &
                    "ITM_DG_YN=" & lITM_DG_YN & ", " &
                    "ITM_REQ_STORE_COLD_YN=" & lITM_REQ_STORE_COLD_YN & ", " &
                    "ITM_REQ_STORE_HUM_DESC=" & gU.convdbNVCData(gU.dbEncode(ITM_REQ_STORE_HUM_DESC.Text.Trim)) & ", " &
                    "ITM_DG_CAT=" & gU.convdbNVCData(gU.dbEncode(ITM_DG_CAT.Text.Trim)) & ", " &
                    "ITM_REQ_STORE_HUM_YN=" & lITM_REQ_STORE_HUM_YN & ", " &
                    "ITM_REQ_STORE_AIRCON_YN=" & lITM_REQ_STORE_AIRCON_YN & ", " &
                    "ITM_CHE_CLASS=" & gU.convdbNVCData(gU.dbEncode(gU.getChkBoxListValue(ITM_CHE_CLASS))) & ", " &
                    "ITM_REQ_MSDS_YN=" & lITM_REQ_MSDS_YN & ", " &
                    "ITM_WEIGHT_TYPE=" & gU.convdbNVCData(gU.dbEncode(ITM_WEIGHT_TYPE.SelectedValue.Trim)) & ", " &
                    "ITM_MANUFACTURER=" & gU.convdbNVCData(gU.dbEncode(ITM_MANUFACTURER.Text.Trim)) & ", " &
                    "ITM_REPLACEMENT_COST=" & gU.convdbNVCData(gU.dbEncode(ITM_REPLACEMENT_COST.Text.Trim)) & ", " &
                    "ITM_COMMODITY_CODE=" & gU.convdbNVCData(gU.dbEncode(ITM_COMMODITY_CODE.Text.Trim)) & ", " &
                    "ITM_RESTRICTED_RMK=" & gU.convdbNVCData(gU.dbEncode(ITM_RESTRICTED_RMK.Text.Trim)) & ", "
                    'ITM_MANUFACTURER, ITM_REPLACEMENT_COST, ITM_COMMODITY_CODE, ITM_RESTRICTED_RMK

                    If Not ViewState("remove_image1") Is Nothing Then
                        tempRemoveFile1 = ViewState("remove_image1")
                    Else
                        tempRemoveFile1 = False
                    End If

                    If Not ViewState("remove_image2") Is Nothing Then
                        tempRemoveFile2 = ViewState("remove_image2")
                    Else
                        tempRemoveFile2 = False
                    End If

                    If Not tempRemoveFile1 Then
                        If nFileName1 <> "" Then sql_string = sql_string & "itm_picture1 = " & gU.convdbNVCData(gU.dbEncode(nFileName1)) & ", "
                    Else
                        sql_string = sql_string & "itm_picture1 = null, "
                    End If

                    If Not tempRemoveFile2 Then
                        If nFileName2 <> "" Then sql_string = sql_string & "itm_picture2 = " & gU.convdbNVCData(gU.dbEncode(nFileName2)) & ", "
                    Else
                        sql_string = sql_string & "itm_picture2 = null, "
                    End If

                    sql_string = sql_string & "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() where itm_code = '" & _
                                    gU.dbEncode(itm_code.Text.Trim) & "' and imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                                    "and pack_key = '" & gU.dbEncode(n_pack_key) & "' and storer_code = '" & gU.dbEncode(n_storer_code) & "'"

                    '"itm_name_ch = " & gU.dbEncode(itm_name_ch.Text) & "', " & _
                    '"itm_desc_ch = " & gU.dbEncode(itm_desc_ch.Text) & "', " & _

                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    itemSQL = "insert into wms_alt_vend_item " & _
                                              "(itm_code, imp_code, storer_code, " & _
                                              "pack_key, vnd_code, vnd_name, " & _
                                              "alv_vnd_itmcode, alv_date_added, aitm_name, aitm_desc, " & _
                                              "aitm_spec, aitm_uom, aitm_pcs_per_pack, aitm_qty_per_ctn, " & _
                                              "aitm_length, aitm_width, aitm_hight, " & _
                                              "aitm_cbm, aitm_vol, aitm_ref_curr, " & _
                                              "aitm_ref_price, aitm_remarks, aitm_pref_wh, " & _
                                              "aitm_pref_loc, aitm_picture1, aitm_picture2, " & _
                                              "aitm_net_weight,aitm_gross_weight,aitm_origin, " & _
                                              "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                              "values ( " & _
                                                 gU.convdbNVCData(gU.dbEncode(itm_code.Text.Trim)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(storer_code.SelectedValue.Trim)) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(pack_key.Text.Trim)) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("vnd_code").ToString.Trim, ""))) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("vnd_name").ToString.Trim, ""))) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("alv_vnd_itmcode").ToString.Trim, ""))) & ", " & _
                                                  "CONVERT(datetime, " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("alv_date_added").ToString.Trim, ""))) & "," & DDFORMAT & ") , " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_name").ToString.Trim, ""))) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_desc").ToString.Trim, ""))) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_spec").ToString.Trim, ""))) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(itm_UOM.SelectedValue, ""))) & ", " & _
                                                 gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_pcs_per_pack").ToString.Trim, "0")) & ", " & _
                                                 gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_qty_per_ctn").ToString.Trim, "NULL")) & ", " & _
                                                 gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_length").ToString.Trim, "0")) & ", " & _
                                                 gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_width").ToString.Trim, "0")) & ", " & _
                                                 gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_hight").ToString.Trim, "0")) & ", " & _
                                                 gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_cbm").ToString.Trim, "0")) & ", " & _
                                                 gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_vol").ToString.Trim, "0")) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_ref_curr").ToString.Trim, ""))) & ", " & _
                                                 gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_ref_price").ToString.Trim, "0")) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_remarks").ToString.Trim, ""))) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_pref_wh").ToString.Trim, ""))) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_pref_loc").ToString.Trim, ""))) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_picture1").ToString.Trim, ""))) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_picture2").ToString.Trim, ""))) & ", " & _
                                                 gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_net_weight").ToString.Trim, "0")) & ", " & _
                                                 gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_gross_weight").ToString.Trim, "0")) & ", " & _
                                                 gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_origin").ToString.Trim, ""))) & ", " & _
                                                 "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                    'aitm_name_ch, aitm_desc_ch, 
                                    'gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_name_ch").ToString.Trim, ""))) & ", " & _
                                    ' gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_desc_ch").ToString.Trim, ""))) & ", " & _
                                Case "D"
                                    itemSQL = "delete from wms_alt_vend_item where itm_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("itm_code").ToString.Trim, "")) & "' and imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                                                "and pack_key = '" & gU.dbEncode(gU.decodeNull(n_pack_key, "")) & "' and storer_code = '" & gU.dbEncode(gU.decodeNull(n_storer_code, "")) & "' " & _
                                                "and vnd_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("o_vnd_code").ToString.Trim, "")) & "'"
                                Case Else
                                    Dim appSQL As String = ""

                                    If gU.decodeNull(rows.Item("o_vnd_code").ToString.Trim, "").ToString.Trim <> gU.decodeNull(rows.Item("vnd_code").ToString.Trim, "").ToString.Trim Then
                                        If rows.Item("vnd_code").ToString.Trim <> "" Then
                                            appSQL = "vnd_code='" & gU.dbEncode(rows.Item("vnd_code").ToString.Trim) & "', "
                                        End If
                                    End If


                                    itemSQL = "update wms_alt_vend_item set " & _
                                                  "vnd_name = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("vnd_name").ToString.Trim, ""))) & ", " & _
                                                  "alv_vnd_itmcode = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("alv_vnd_itmcode").ToString.Trim, ""))) & ", " & _
                                                  "alv_date_added = CONVERT(datetime, " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("alv_date_added").ToString.Trim, ""))) & "," & DDFORMAT & "), " & _
                                                  "aitm_status = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_status").ToString.Trim, ""))) & ", " & _
                                                  "aitm_name = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_name").ToString.Trim, ""))) & ", " & _
                                                  "aitm_desc = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_desc").ToString.Trim, ""))) & ", " & _
                                                  "aitm_name_ch = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_name_ch").ToString.Trim, ""))) & ", " & _
                                                  "aitm_desc_ch = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_desc_ch").ToString.Trim, ""))) & ", " & _
                                                  "aitm_spec = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_spec").ToString.Trim, ""))) & ", " & _
                                                  "aitm_uom = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(itm_UOM.SelectedValue, ""))) & ", " & _
                                                  "aitm_pcs_per_pack = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_pcs_per_pack").ToString.Trim, "0")) & ", " & _
                                                  "aitm_qty_per_ctn = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_qty_per_ctn").ToString.Trim, "NULL")) & ", " & _
                                                  "aitm_length = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_length").ToString.Trim, "0")) & ", " & _
                                                  "aitm_width = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_width").ToString.Trim, "0")) & ", " & _
                                                  "aitm_hight = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_hight").ToString.Trim, "0")) & ", " & _
                                                  "aitm_cbm = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_cbm").ToString.Trim, "0")) & ", " & _
                                                  "aitm_vol = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_vol").ToString.Trim, "0")) & ", " & _
                                                  "aitm_ref_curr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_ref_curr").ToString.Trim, ""))) & ", " & _
                                                  "aitm_ref_price = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_ref_price").ToString.Trim, "0")) & ", " & _
                                                  "aitm_remarks = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_remarks").ToString.Trim, ""))) & ", " & _
                                                  "aitm_pref_wh = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_pref_wh").ToString.Trim, ""))) & ", " & _
                                                  "aitm_pref_loc = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_pref_loc").ToString.Trim, ""))) & ", " & _
                                                  "aitm_picture1 = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_picture1").ToString.Trim, ""))) & ", " & _
                                                  "aitm_picture2 = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_picture2").ToString.Trim, ""))) & ",  " & _
                                                  "aitm_net_weight = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_net_weight").ToString.Trim, "0")) & ", " & _
                                                  "aitm_gross_weight = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("aitm_gross_weight").ToString.Trim, "0")) & ", " & _
                                                  "aitm_origin = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("aitm_origin").ToString.Trim, ""))) & ",  " & _
                                                  appSQL & _
                                                  "sys_lub = '" & Session("usr_id") & "', " & _
                                                  "sys_lud = Getdate() " & _
                                                "where itm_code = '" & gU.dbEncode(itm_code.Text) & "' and imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                                                "and pack_key = '" & gU.dbEncode(gU.decodeNull(n_pack_key, "")) & "' and storer_code = '" & gU.dbEncode(gU.decodeNull(n_storer_code, "")) & "' " & _
                                                "and vnd_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("o_vnd_code").ToString.Trim, "")) & "'"

                            End Select
                            REM **********************

                            'Response.Write(itemSQL)
                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                            If gU.decodeNull(rows.Item("has_pic1").ToString.Trim.ToUpper, "") = "Y" Then
                                hasFile1.Add(gU.decodeNull(rows.Item("aitm_picture1").ToString.Trim, ""))
                            End If

                            If gU.decodeNull(rows.Item("has_pic2").ToString.Trim.ToUpper, "") = "Y" Then
                                hasFile2.Add(gU.decodeNull(rows.Item("aitm_picture2").ToString.Trim, ""))
                            End If


                            If gU.decodeNull(rows.Item("remove_pic1").ToString.Trim.ToUpper, "") = "Y" Then
                                removeFile1.Add(gU.decodeNull(rows.Item("old_pic1").ToString.Trim, ""))
                            End If

                            If gU.decodeNull(rows.Item("remove_pic2").ToString.Trim.ToUpper, "") = "Y" Then
                                removeFile2.Add(gU.decodeNull(rows.Item("old_pic2").ToString.Trim, ""))
                            End If
                        Next
                    End If
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then
                    gDB.amendData(sql_string, gConn, transaction)
                End If

                Dim paP As GlobalDBFunc.DBCmdPara

                Dim ImgDT As DataTable

                paP = New GlobalDBFunc.DBCmdPara
                sql_string = "Select itm_photo1, itm_photo2, itm_photo3 from wms_item " & _
                             " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(ViewState("storer_code"))) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(ViewState("wms_itm_code"))) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(ViewState("pack_key")))

                ImgDT = gDB.getDataTable(sql_string, gConn, transaction, , paP)


                If ITM_PHOTO1.HasFile Then
                    Using fs As Stream = ITM_PHOTO1.PostedFile.InputStream
                        Using br As New BinaryReader(fs)
                            Dim bytes As Byte() = br.ReadBytes(fs.Length)

                            paP = New GlobalDBFunc.DBCmdPara
                            sql_string = "update wms_item set itm_photo" & IMG_SEQ1.SelectedValue & "=" & paP.AP(bytes, SqlDbType.VarBinary) & _
                                         " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(ViewState("storer_code"))) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(ViewState("wms_itm_code"))) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(ViewState("pack_key")))

                            gDB.amendData(sql_string, gConn, transaction, paP)

                        End Using
                    End Using
                Else
                    If IMG_SEQ1.SelectedValue <> "1" Then
                        If ImgDT.Rows.Count > 0 AndAlso ImgDT.Rows(0).Item("ITM_PHOTO1").ToString.Trim <> "" Then
                            Dim bytes As Byte() = ImgDT.Rows(0).Item("ITM_PHOTO1")

                            paP = New GlobalDBFunc.DBCmdPara
                            sql_string = "update wms_item set itm_photo" & IMG_SEQ1.SelectedValue & "=" & paP.AP(bytes, SqlDbType.VarBinary) & _
                                         " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(ViewState("storer_code"))) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(ViewState("wms_itm_code"))) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(ViewState("pack_key")))

                            gDB.amendData(sql_string, gConn, transaction, paP)
                        End If
                    End If
                End If

                If ITM_PHOTO2.HasFile Then
                    Using fs As Stream = ITM_PHOTO2.PostedFile.InputStream
                        Using br As New BinaryReader(fs)
                            Dim bytes As Byte() = br.ReadBytes(fs.Length)

                            paP = New GlobalDBFunc.DBCmdPara
                            sql_string = "update wms_item set itm_photo" & IMG_SEQ2.SelectedValue & "=" & paP.AP(bytes, SqlDbType.VarBinary) & _
                                         " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(ViewState("storer_code"))) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(ViewState("wms_itm_code"))) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(ViewState("pack_key")))

                            gDB.amendData(sql_string, gConn, transaction, paP)

                        End Using
                    End Using
                Else
                    If IMG_SEQ2.SelectedValue <> "2" Then
                        If ImgDT.Rows.Count > 0 AndAlso ImgDT.Rows(0).Item("ITM_PHOTO2").ToString.Trim <> "" Then
                            Dim bytes As Byte() = ImgDT.Rows(0).Item("ITM_PHOTO2")

                            paP = New GlobalDBFunc.DBCmdPara
                            sql_string = "update wms_item set itm_photo" & IMG_SEQ2.SelectedValue & "=" & paP.AP(bytes, SqlDbType.VarBinary) & _
                                         " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(ViewState("storer_code"))) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(ViewState("wms_itm_code"))) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(ViewState("pack_key")))

                            gDB.amendData(sql_string, gConn, transaction, paP)
                        End If
                    End If
                End If

                If ITM_PHOTO3.HasFile Then
                    Using fs As Stream = ITM_PHOTO3.PostedFile.InputStream
                        Using br As New BinaryReader(fs)
                            Dim bytes As Byte() = br.ReadBytes(fs.Length)

                            paP = New GlobalDBFunc.DBCmdPara
                            sql_string = "update wms_item set itm_photo" & IMG_SEQ3.SelectedValue & "=" & paP.AP(bytes, SqlDbType.VarBinary) & _
                                         " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(ViewState("storer_code"))) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(ViewState("wms_itm_code"))) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(ViewState("pack_key")))

                            gDB.amendData(sql_string, gConn, transaction, paP)

                        End Using
                    End Using
                Else
                    If IMG_SEQ3.SelectedValue <> "3" Then
                        If ImgDT.Rows.Count > 0 AndAlso ImgDT.Rows(0).Item("ITM_PHOTO3").ToString.Trim <> "" Then
                            Dim bytes As Byte() = ImgDT.Rows(0).Item("ITM_PHOTO3")

                            paP = New GlobalDBFunc.DBCmdPara
                            sql_string = "update wms_item set itm_photo" & IMG_SEQ3.SelectedValue & "=" & paP.AP(bytes, SqlDbType.VarBinary) & _
                                         " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(ViewState("storer_code"))) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(ViewState("wms_itm_code"))) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(ViewState("pack_key")))

                            gDB.amendData(sql_string, gConn, transaction, paP)
                        End If
                    End If
                End If

                If Not transaction Is Nothing Then
                    transaction.Commit()
                    transaction = Nothing

                    Dim isSavePic1 As Boolean = False
                    Dim isSavePic2 As Boolean = False

                    If Not IO.Directory.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE")) Then IO.Directory.CreateDirectory(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE"))
                    If Not IO.Directory.Exists(SYSP_TEMP_DIR) Then IO.Directory.CreateDirectory(SYSP_TEMP_DIR)

                    If Not tempRemoveFile1 Then
                        If nUP1 Is Nothing Then
                            If itm_picture1_upload.HasFile Then
                                itm_picture1_upload.SaveAs(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & nFileName1)
                                isSavePic1 = True
                            End If
                        Else
                            nUP1.SaveAs(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & nFileName1)
                            isSavePic1 = True
                        End If

                        If isSavePic1 Then
                            If IO.File.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_filename1) Then

                                thumb = New ThumbGenerator
                                thumb.SetParams(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_filename1, 150, 180)
                                Dim unID As String = thumb.GetUniqueThumbName
                                Cache.Remove(unID)

                                IO.File.Delete(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_filename1)
                            End If

                            IO.File.Move(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & nFileName1, PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & nFileName1)
                        End If

                    Else
                        removeImage(old_filename1)
                        ViewState("remove_image1") = False
                    End If

                    If Not tempRemoveFile2 Then
                        If nUP2 Is Nothing Then
                            If itm_picture2_upload.HasFile Then
                                itm_picture2_upload.SaveAs(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & nFileName2)
                                isSavePic2 = True
                            End If
                        Else
                            nUP2.SaveAs(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & nFileName2)
                            isSavePic2 = True
                        End If

                        If isSavePic2 Then
                            If IO.File.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_filename2) Then

                                thumb = New ThumbGenerator
                                thumb.SetParams(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_filename2, 150, 180)
                                Dim unID As String = thumb.GetUniqueThumbName
                                Cache.Remove(unID)

                                IO.File.Delete(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_filename2)
                            End If

                            IO.File.Move(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & nFileName2, PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & nFileName2)
                        End If
                    Else
                        removeImage(old_filename2)
                        ViewState("remove_image2") = False
                    End If

                    For t As Integer = 0 To hasFile1.Count - 1
                        If IO.File.Exists(SYSP_TEMP_DIR & "\" & "temp_" & hasFile1.Item(t)) Then

                            If IO.File.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & hasFile1.Item(t)) Then

                                thumb = New ThumbGenerator
                                thumb.SetParams(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & hasFile1.Item(t), 150, 180)
                                Dim unID As String = thumb.GetUniqueThumbName
                                Cache.Remove(unID)

                                IO.File.Delete(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & hasFile1.Item(t))
                            End If

                            IO.File.Move(SYSP_TEMP_DIR & "\" & "temp_" & hasFile1.Item(t), PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & hasFile1.Item(t))
                        End If
                    Next

                    For n As Integer = 0 To hasFile2.Count - 1
                        If IO.File.Exists(SYSP_TEMP_DIR & "\" & "temp_" & hasFile2.Item(n)) Then

                            If IO.File.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & hasFile2.Item(n)) Then

                                thumb = New ThumbGenerator
                                thumb.SetParams(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & hasFile2.Item(n), 150, 180)
                                Dim unID As String = thumb.GetUniqueThumbName
                                Cache.Remove(unID)

                                IO.File.Delete(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & hasFile2.Item(n))
                            End If

                            IO.File.Move(SYSP_TEMP_DIR & "\" & "temp_" & hasFile2.Item(n), PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & hasFile2.Item(n))
                        End If
                    Next

                    For g As Integer = 0 To removeFile1.Count - 1
                        removeImage(removeFile1.Item(g))
                    Next

                    For h As Integer = 0 To removeFile2.Count - 1
                        removeImage(removeFile2.Item(h))
                    Next
                End If

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")

                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    'itm_code.Text = nextNo
                    ViewState("wms_itm_code") = nextNo.Trim
                    'itm_code.ReadOnly = True
                    'itm_code.BorderWidth = 0
                    'itm_code.BackColor = Drawing.Color.Transparent
                    'itm_code.ForeColor = Drawing.Color.Black
                    'itm_code.Font.Size = 10
                    REM **********************
                End If

                Session.Remove(itm_picture1_upload.ClientID)
                Session.Remove(itm_picture2_upload.ClientID)

                itm_picture1_upload.Visible = True
                itm_picture2_upload.Visible = True

                itm_picture1_upload_lit.Visible = False
                itm_picture2_upload_lit.Visible = False

                itm_picture1_edit.Visible = False
                itm_picture2_edit.Visible = False

                If flag <> "Y" Then uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                Call BindGV()
            Catch ex As Exception
                If Not transaction Is Nothing Then
                    transaction.Rollback()
                    transaction = Nothing
                End If

                Session.Remove(itm_picture1_upload.ClientID)
                Session.Remove(itm_picture2_upload.ClientID)

                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            Finally
                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try
        End If
    End Sub

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"
        Dim WhereStr As String = ""
        Dim pk_code As String = ""
        Dim pic_preview_yn As Boolean = False
        Dim pk_code2 As String = ""
        Dim str_code As String = ""
        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("wms_itm_code") <> "" Then
            pk_code = ViewState("wms_itm_code")
            pk_code2 = pack_key.Text
            str_code = storer_code.SelectedValue
        Else
            pk_code = Server.UrlDecode(Request("itm_code"))
            pk_code2 = Server.UrlDecode(Request("pack_key"))
            str_code = Server.UrlDecode(Request("storer_code"))
        End If
        REM **********************

        IMG_SEQ1.SelectedValue = 1
        IMG_SEQ2.SelectedValue = 2
        IMG_SEQ3.SelectedValue = 3
        PRE_IMG_SEQ1.Value = 1
        PRE_IMG_SEQ2.Value = 2
        PRE_IMG_SEQ3.Value = 3

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            'itm_code.ForeColor = Drawing.Color.Red
            itm_status.Text = "NEW"

            itm_picture1.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=&refresh=true&Height=180&userid=" & Session("usr_id")
            itm_picture2.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=&refresh=true&Height=180&userid=" & Session("usr_id")

            BtnCopy.Visible = False
            btnPrint.Visible = False
            LabelSize.Visible = False
            btnAttach.Visible = False
            'btnMSDS.Visible = False
            btnMsdsUp.Visible = False
            REM **********************
        Else
            BtnCopy.Visible = True
            btnPrint.Visible = True
            LabelSize.Visible = False
            btnAttach.Visible = True
            btnMsdsUp.Visible = True
            'btnMSDS.Visible = True
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header


            SQLString = "select wms_item.itm_code, wms_item.pack_key, wms_item.storer_code, " & _
                           "wms_item.itm_vend_itm_no, wms_item.itm_barcode, wms_item.itm_name, wms_item.itm_desc, wms_item.itm_cat, " & _
                           "wms_item.itm_type, wms_item.itm_brand, wms_item.itm_pref_wh, " & _
                           "wms_item.itm_pref_loc,wms_item.itm_pref_loc2, wms_item.itm_ref_curr, wms_item.itm_ref_price, wms_item.itm_price, " & _
                           "wms_item.itm_balance, wms_item.itm_flags, wms_item.itm_remarks, " & _
                           "wms_item.itm_series, wms_item.itm_series_no, wms_item.itm_model, wms_item.itm_parent, wms_item.ITM_KIT_QTY, " & _
                           "wms_item.itm_picture1, wms_item.itm_picture2, wms_item.itm_status, " & _
                           "wms_item.ITM_TEMP_YN,wms_item.ITM_TEMP_FR,wms_item.ITM_TEMP_TO,wms_item.itm_uom, " & _
                           "wms_item.sys_cb, wms_item.sys_lub, wms_item.sys_cd, wms_item.sys_lud, wms_item.itm_pcs_per_uom, " & _
                           "wms_item.itm_qty_of_unit, wms_item.itm_sku_no, " & _
                           "convert(varchar,ITM_PD_RCV_DATE," & DDFORMAT & ") as ITM_PD_RCV_DATE,wms_item.ITM_PD_CONTRACT_NO,wms_item.ITM_PD_ARR_NOTICE_NO,wms_item.ITM_PD_COND_OF_SPARES,wms_item.ITM_PD_ST_1_YEAR,wms_item.ITM_PD_ST_OVER_1_YEAR," & _
                           "wms_item.ITM_SIZE,wms_item.ITM_COLOR_CODE,wms_item.ITM_SERIAL_NO,wms_item.ITM_PART_NO,wms_item.ITM_PROD_NO,ITM_PROD_GROUP, " & _
                           "wms_item.ITM_SERIAL_NO_YN,wms_item.ITM_STACKABLE_YN,wms_item.ITM_INSP_YN,wms_item.ITM_SCRAP_YN,wms_item.ITEM_PRICE_CLASS, wms_item.ITM_CAP_REV, wms_item.ITM_CAP_VALUE, " & _
                           "wms_item.ITM_NONSTOCK_YN, wms_item.ITM_CC, convert(varchar,ITM_CC_DATE," & DDFORMAT & ") as ITM_CC_DATE, wms_item.ITM_SHELF_LIFE, wms_item.proj_no, " & _
                           "wms_item.ITM_GP_CODE	,wms_item.ITM_DIV_CODE	,wms_item.ITM_MFG	,wms_item.ITM_STORES_CLASS	,wms_item.ITM_EMB	,wms_item.ITM_HAZARD_CODE	," & _
                           "wms_item.ITM_CRITICAL_YN	,wms_item.ITM_MAX_STOCK	,wms_item.ITM_RESTRICTED_ITEM	,wms_item.ITM_DRAWING_NO	,wms_item.ITM_CC	,wms_item.ITM_KEPT_DIV_CODE	," & _
                           "wms_item.ITM_EXPENSE_CODE	,wms_item.ITM_INT_ORDER_QTY	, wms_item.ITM_1ST_INSP	, wms_item.ITM_ORO_YN	, wms_item.ITM_ROP_APL	,wms_item.ITM_ROP_LAM	," & _
                           "wms_item.ITM_ROP_CABLE	, wms_item.ITM_ROP_NP	, wms_item.ITM_TI_YTD	, wms_item.ITM_TI_LYR	, wms_item.ITM_TI_PY2	, wms_item.ITM_TI_PY3	," & _
                           "wms_item.ITM_UOM2	, wms_item.ITM_QTY2	, wms_item.ITM_DG_YN	, wms_item.ITM_REQ_STORE_HUM_DESC, wms_item.ITM_REQ_STORE_COLD_YN, wms_item.ITM_DG_CAT	, wms_item.ITM_REQ_STORE_HUM_YN	, wms_item.ITM_REQ_STORE_AIRCON_YN	, " & _
                           "wms_item.ITM_CHE_CLASS	, wms_item.ITM_REQ_MSDS_YN	, wms_item.ITM_WEIGHT_TYPE, " & _
                           "wms_item.ITM_MANUFACTURER, wms_item.ITM_REPLACEMENT_COST, wms_item.ITM_COMMODITY_CODE, wms_item.ITM_RESTRICTED_RMK " & _
                           "from wms_item where wms_item.itm_code = '" & gU.dbEncode(pk_code) & "' " & _
                           "and wms_item.imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                           "and wms_item.pack_key = '" & gU.dbEncode(pk_code2) & "' " & _
                           "and storer_code='" & gU.dbEncode(str_code) & "'"

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                ViewState("pack_key") = pk_code2
                ViewState("storer_code") = str_code
                ViewState("wms_itm_code") = pk_code

                btnAttach.Attributes.Add("onclick", "javascript:goToAttach('MAST_IM','" & imp_code & "||" & str_code & "||" & pk_code & "||" & pk_code2 & "','N');return false;")
                btnMsdsUp.Attributes.Add("onclick", "javascript:goToAttach('MAST_ITM_MSDS','" & imp_code & "||" & str_code & "||" & pk_code & "||" & pk_code2 & "','N');return false;")

                If Session("pagemode") <> "N" Then itm_code.Text = dt.Rows(0).Item("itm_code").ToString
                pack_key.Text = dt.Rows(0).Item("pack_key").ToString

                If dt.Rows(0).Item("storer_code").ToString = "" Then
                    uiFun.load_dropdown(storer_code, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STORER_CODE <> 'PD' AND STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(storer_code, "select STORER_CODE, STO_SHORTNAME AS STO_NAME from WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
                End If

                If dt.Rows(0).Item("ITM_SERIAL_NO_YN").ToString.Trim = "Y" Then
                    ITM_SERIAL_NO_YN.Checked = True
                End If

                If dt.Rows(0).Item("ITM_STACKABLE_YN").ToString.Trim = "Y" Then
                    'ITM_STACKABLE_YN.Checked = True
                    radLotExist.Checked = True
                    radLotNotExist.Checked = False
                ElseIf dt.Rows(0).Item("ITM_STACKABLE_YN").ToString.Trim = "N" Then
                    'ITM_STACKABLE_YN.Checked = True
                    radLotExist.Checked = False
                    radLotNotExist.Checked = True
                End If

                If dt.Rows(0).Item("ITM_INSP_YN").ToString.Trim = "Y" Then
                    ITM_INSP_YN.Checked = True
                End If

                'If dt.Rows(0).Item("ITM_SCRAP_YN").ToString.Trim = "Y" Then
                '    ITM_SCRAP_YN.Checked = True
                'End If

                If dt.Rows(0).Item("ITM_NONSTOCK_YN").ToString.Trim = "Y" Then
                    'ITM_NONSTOCK_YN.Checked = True
                End If

                'If ITM_NONSTOCK_YN.Checked = True Then
                '    ITM_PD_RCV_DATE.CssClass = "REQUIRED"
                '    ITM_PD_CONTRACT_NO.CssClass = "REQUIRED"
                '    ITM_PD_ARR_NOTICE_NO.CssClass = "REQUIRED"
                '    ITM_PD_COND_OF_SPARES.CssClass = "REQUIRED"
                '    ITM_PD_ST_1_YEAR.CssClass = "REQUIRED"
                '    ITM_PD_ST_OVER_1_YEAR.CssClass = "REQUIRED"
                'End If

                If dt.Rows(0).Item("ITM_CC").ToString.Trim = "Y" Then
                    ITM_CC.Checked = True
                End If

                storer_code.SelectedValue = dt.Rows(0).Item("storer_code").ToString

                itm_vend_itm_no.Value = dt.Rows(0).Item("itm_vend_itm_no").ToString
                ori_itm_vend_itm_no.Value = dt.Rows(0).Item("itm_vend_itm_no").ToString

                itm_barcode.Text = dt.Rows(0).Item("itm_barcode").ToString
                itm_name.Text = dt.Rows(0).Item("itm_name").ToString
                itm_desc.Text = dt.Rows(0).Item("itm_desc").ToString
                itm_cat.SelectedValue = dt.Rows(0).Item("itm_cat").ToString
                itm_type.selectedvalue = dt.Rows(0).Item("itm_type").ToString
                itm_brand.Text = dt.Rows(0).Item("itm_brand").ToString
                itm_pref_wh.SelectedValue = dt.Rows(0).Item("itm_pref_wh").ToString
                itm_pref_loc.Value = dt.Rows(0).Item("itm_pref_loc").ToString
                itm_pref_loc2.Value = dt.Rows(0).Item("itm_pref_loc2").ToString

                dsp_itm_pref_loc.Text = dt.Rows(0).Item("itm_pref_loc").ToString
                dsp_itm_pref_loc2.Text = dt.Rows(0).Item("itm_pref_loc2").ToString

                itm_ref_curr.Text = dt.Rows(0).Item("itm_ref_curr").ToString
                itm_ref_price.Text = cU.FormatDecimalwString(dt.Rows(0).Item("itm_ref_price").ToString)
                itm_price.Text = cU.FormatDecimalwString(dt.Rows(0).Item("itm_price").ToString)
                'itm_balance.Text = cU.FormatIntegerString(dt.Rows(0).Item("itm_balance").ToString)

                itm_UOM.SelectedValue = dt.Rows(0).Item("itm_uom").ToString.Trim
                itm_pcs_per_uom.Text = dt.Rows(0).Item("itm_pcs_per_uom").ToString.Trim
                itm_qty_of_unit.Text = dt.Rows(0).Item("itm_qty_of_unit").ToString.Trim
                itm_sku_no.Text = dt.Rows(0).Item("itm_sku_no").ToString.Trim
                ITM_CC_DATE.Text = dt.Rows(0).Item("ITM_CC_DATE").ToString.Trim
                itm_status.Text = gU.decodeNullOrEmpty(dt.Rows(0).Item("itm_status").ToString, "NEW")

                itm_picture1.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & dt.Rows(0).Item("itm_picture1").ToString & "&refresh=true&Height=180&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                itm_picture2.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & dt.Rows(0).Item("itm_picture2").ToString & "&refresh=true&Height=180&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")

                If dt.Rows(0).Item("ITM_TEMP_YN").ToString = "Y" Then
                    ITM_TEMP_YN.Checked = True
                Else
                    ITM_TEMP_YN.Checked = False
                End If

                ITM_TEMP_FR.Text = dt.Rows(0).Item("ITM_TEMP_FR").ToString
                ITM_TEMP_TO.Text = dt.Rows(0).Item("ITM_TEMP_TO").ToString

                ITM_PD_RCV_DATE.Text = dt.Rows(0).Item("ITM_PD_RCV_DATE").ToString.Trim
                ITM_PD_CONTRACT_NO.Text = dt.Rows(0).Item("ITM_PD_CONTRACT_NO").ToString.Trim
                ITM_PD_ARR_NOTICE_NO.Text = dt.Rows(0).Item("ITM_PD_ARR_NOTICE_NO").ToString.Trim
                ITM_PD_COND_OF_SPARES.Text = dt.Rows(0).Item("ITM_PD_COND_OF_SPARES").ToString.Trim
                ITM_PD_ST_1_YEAR.Text = dt.Rows(0).Item("ITM_PD_ST_1_YEAR").ToString.Trim
                ITM_PD_ST_OVER_1_YEAR.Text = dt.Rows(0).Item("ITM_PD_ST_OVER_1_YEAR").ToString.Trim

                ITM_SIZE.Text = dt.Rows(0).Item("ITM_SIZE").ToString.Trim
                ITM_COLOR_CODE.Text = dt.Rows(0).Item("ITM_COLOR_CODE").ToString.Trim
                ITM_SERIAL_NO.Text = dt.Rows(0).Item("ITM_SERIAL_NO").ToString.Trim
                ITM_PART_NO.Text = dt.Rows(0).Item("ITM_PART_NO").ToString.Trim
                ITM_PROD_NO.Text = dt.Rows(0).Item("ITM_PROD_NO").ToString.Trim
                ITM_PROD_GROUP.Text = dt.Rows(0).Item("ITM_PROD_GROUP").ToString.Trim

                ITEM_PRICE_CLASS.SelectedValue = dt.Rows(0).Item("ITEM_PRICE_CLASS").ToString.Trim
                ITM_CAP_REV.SelectedValue = dt.Rows(0).Item("ITM_CAP_REV").ToString.Trim
                ITM_CAP_VALUE.SelectedValue = dt.Rows(0).Item("ITM_CAP_VALUE").ToString.Trim

                proj_no.Text = dt.Rows(0).Item("PROJ_NO").ToString.Trim
                ITM_SHELF_LIFE.Text = dt.Rows(0).Item("ITM_SHELF_LIFE").ToString.Trim
                ITM_GP_CODE.Text = dt.Rows(0).Item("ITM_GP_CODE").ToString.Trim
                ITM_DIV_CODE.Text = dt.Rows(0).Item("ITM_DIV_CODE").ToString.Trim
                ITM_MFG.Text = dt.Rows(0).Item("ITM_MFG").ToString.Trim
                ITM_STORES_CLASS.Text = dt.Rows(0).Item("ITM_STORES_CLASS").ToString.Trim
                If dt.Rows(0).Item("ITM_EMB").ToString <> "" Then
                    ITM_EMB.SelectedValue = dt.Rows(0).Item("ITM_EMB").ToString
                End If
                ITM_HAZARD_CODE.Text = dt.Rows(0).Item("ITM_HAZARD_CODE").ToString.Trim
                If dt.Rows(0).Item("ITM_CRITICAL_YN").ToString <> "" Then
                    ITM_CRITICAL_YN.SelectedValue = dt.Rows(0).Item("ITM_CRITICAL_YN").ToString
                End If
                ITM_MAX_STOCK.Text = dt.Rows(0).Item("ITM_MAX_STOCK").ToString.Trim
                If dt.Rows(0).Item("ITM_RESTRICTED_ITEM").ToString <> "" Then
                    ITM_RESTRICTED_ITEM.SelectedValue = dt.Rows(0).Item("ITM_RESTRICTED_ITEM").ToString
                End If
                ITM_DRAWING_NO.Text = dt.Rows(0).Item("ITM_DRAWING_NO").ToString.Trim
                ITM_KEPT_DIV_CODE.Text = dt.Rows(0).Item("ITM_KEPT_DIV_CODE").ToString.Trim
                ITM_EXPENSE_CODE.Text = dt.Rows(0).Item("ITM_EXPENSE_CODE").ToString.Trim
                ITM_INT_ORDER_QTY.Text = dt.Rows(0).Item("ITM_INT_ORDER_QTY").ToString.Trim
                ITM_1ST_INSP.Text = dt.Rows(0).Item("ITM_1ST_INSP").ToString.Trim
                If dt.Rows(0).Item("ITM_ORO_YN").ToString <> "" Then
                    ITM_ORO_YN.SelectedValue = dt.Rows(0).Item("ITM_ORO_YN").ToString
                End If
                ITM_ROP_APL.Text = dt.Rows(0).Item("ITM_ROP_APL").ToString.Trim
                ITM_ROP_LAM.Text = dt.Rows(0).Item("ITM_ROP_LAM").ToString.Trim
                ITM_ROP_CABLE.Text = dt.Rows(0).Item("ITM_ROP_CABLE").ToString.Trim
                ITM_ROP_NP.Text = dt.Rows(0).Item("ITM_ROP_NP").ToString.Trim
                ITM_TI_YTD.Text = dt.Rows(0).Item("ITM_TI_YTD").ToString.Trim
                ITM_TI_LYR.Text = dt.Rows(0).Item("ITM_TI_LYR").ToString.Trim
                ITM_TI_PY2.Text = dt.Rows(0).Item("ITM_TI_PY2").ToString.Trim
                ITM_TI_PY3.Text = dt.Rows(0).Item("ITM_TI_PY3").ToString.Trim
                ITM_UOM2.SelectedValue = dt.Rows(0).Item("ITM_UOM2").ToString.Trim
                ITM_QTY2.Text = dt.Rows(0).Item("ITM_QTY2").ToString.Trim
                If dt.Rows(0).Item("ITM_DG_YN").ToString.Trim = "Y" Then
                    ITM_DG_YN.Checked = True
                End If
                If dt.Rows(0).Item("ITM_REQ_STORE_COLD_YN").ToString.Trim = "Y" Then
                    ITM_REQ_STORE_COLD_YN.Checked = True
                End If
                ITM_REQ_STORE_HUM_DESC.Text = dt.Rows(0).Item("ITM_REQ_STORE_HUM_DESC").ToString.Trim
                ITM_DG_CAT.Text = dt.Rows(0).Item("ITM_DG_CAT").ToString.Trim
                If dt.Rows(0).Item("ITM_REQ_STORE_HUM_YN").ToString.Trim = "Y" Then
                    ITM_REQ_STORE_HUM_YN.Checked = True
                End If
                If dt.Rows(0).Item("ITM_REQ_STORE_AIRCON_YN").ToString.Trim = "Y" Then
                    ITM_REQ_STORE_AIRCON_YN.Checked = True
                End If
                'ITM_CHE_CLASS

                uiFun.load_checkboxList(ITM_CHE_CLASS, "SELECT COLC_CODE, COLC_ENG_VALUE, COLC_DISPLAY_SEQ FROM WMS_COL_CODE WHERE COLC_TABCOL = 'WMS_ITEM.ITM_CHE_CLASS' ORDER BY COLC_DISPLAY_SEQ", "COLC_CODE", "COLC_ENG_VALUE", "")
                gU.setChkBoxListByValue(ITM_CHE_CLASS, dt.Rows(0).Item("ITM_CHE_CLASS").ToString.Trim)

                If dt.Rows(0).Item("ITM_REQ_MSDS_YN").ToString.Trim = "Y" Then
                    ITM_REQ_MSDS_YN.Checked = True
                End If
                If dt.Rows(0).Item("ITM_WEIGHT_TYPE").ToString <> "" Then
                    ITM_WEIGHT_TYPE.SelectedValue = dt.Rows(0).Item("ITM_WEIGHT_TYPE").ToString
                End If

                If dt.Rows(0).Item("itm_picture1").ToString <> "" Then
                    itm_picture1.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & dt.Rows(0).Item("itm_picture1").ToString & "&refresh=true&ds=true&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                    itm_picture1.Target = "_new1"

                    preview_pic1.Visible = True
                    itm_picture1_remove.Visible = True
                    itm_picture1.Enabled = True
                    pic_preview_yn = True
                Else
                    preview_pic1.Visible = False
                    itm_picture1_remove.Visible = False
                    itm_picture1.Enabled = False
                End If

                If dt.Rows(0).Item("itm_picture2").ToString <> "" Then
                    itm_picture2.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & dt.Rows(0).Item("itm_picture2").ToString & "&refresh=true&ds=true&Height=180&width=180&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                    itm_picture2.Target = "_new2"

                    preview_pic2.Visible = True
                    itm_picture2_remove.Visible = True
                    itm_picture2.Enabled = True
                    pic_preview_yn = True
                Else
                    preview_pic2.Visible = False
                    itm_picture2_remove.Visible = False
                    itm_picture2.Enabled = False
                End If

                If pic_preview_yn Then
                    pic_tr.Visible = True
                    remove_tr.Visible = True
                Else
                    pic_tr.Visible = False
                    remove_tr.Visible = False
                End If

                ViewState("image1_name") = dt.Rows(0).Item("itm_picture1").ToString
                ViewState("image2_name") = dt.Rows(0).Item("itm_picture2").ToString

                If dt.Rows(0).Item("itm_flags").ToString = "" Then
                    itm_flags.SelectedIndex = -1
                Else

                    If dt.Rows(0).Item("itm_flags").ToString.Contains(",") Then
                        Dim flags_list As String()
                        Dim tempFlags As String = dt.Rows(0).Item("itm_flags").ToString.ToUpper
                        flags_list = tempFlags.Split(",")

                        Dim il As IList = flags_list

                        For k As Integer = 0 To itm_flags.Items.Count - 1
                            If il.Contains(itm_flags.Items(k).Value) Then
                                itm_flags.Items(k).Selected = True
                            End If

                        Next

                    Else
                        itm_flags.SelectedValue = dt.Rows(0).Item("itm_flags").ToString
                    End If

                End If

                ITM_MANUFACTURER.Text = dt.Rows(0).Item("ITM_MANUFACTURER").ToString
                ITM_REPLACEMENT_COST.Text = dt.Rows(0).Item("ITM_REPLACEMENT_COST").ToString
                ITM_COMMODITY_CODE.Text = dt.Rows(0).Item("ITM_COMMODITY_CODE").ToString
                ITM_RESTRICTED_RMK.Text = dt.Rows(0).Item("ITM_RESTRICTED_RMK").ToString

                itm_remarks.Text = dt.Rows(0).Item("itm_remarks").ToString
                itm_series.Text = dt.Rows(0).Item("itm_series").ToString
                itm_series_no.Text = dt.Rows(0).Item("itm_series_no").ToString
                itm_model.Text = dt.Rows(0).Item("itm_model").ToString
                itm_parent.Text = dt.Rows(0).Item("itm_parent").ToString
                ITM_KIT_QTY.text = dt.Rows(0).Item("ITM_KIT_QTY").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If ViewState("remove_image1") Then
                    itm_picture1_upload.Enabled = False
                    itm_picture1_remove.Enabled = False
                    itm_picture1.Enabled = False
                    preview_pic1.Enabled = False
                Else
                    itm_picture1_upload.Enabled = True
                    itm_picture1_remove.Enabled = True
                    itm_picture1.Enabled = True
                    preview_pic1.Enabled = True
                End If

                If ViewState("remove_image2") Then
                    itm_picture2_upload.Enabled = False
                    itm_picture2_remove.Enabled = False
                    itm_picture2.Enabled = False
                    preview_pic2.Enabled = False

                Else
                    itm_picture2_upload.Enabled = True
                    itm_picture2_remove.Enabled = True
                    itm_picture2.Enabled = True
                    preview_pic2.Enabled = True
                End If

                Dim balSQL As String = ""

                balSQL = "select ISNULL(SUM(ILOC_BAL_QTY),0) as value from wms_item_loc_bal where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(str_code) & "' and ITM_CODE='" & gU.dbEncode(pk_code) & "' and pack_key ='" & gU.dbEncode(pk_code2) & "'"
                itm_balance.Text = gU.decodeEmptyCInt(DB.getValueFromSQL(balSQL), 0)

                If dt.Rows(0).Item("itm_type").ToString <> "CABLE" Then
                    btnPrint.OnClientClick = "OpenItemLbls();"
                    btnPrints.OnClientClick = "OpenItemLbls2();"
                Else
                    SQLString = "Select count(*) as bal_count from WMS_ITEM_LOC_BAL_S where itm_code = '" & gU.dbEncode(pk_code) & "' " & _
                                "and imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                                "and pack_key = '" & gU.dbEncode(pk_code2) & "' " & _
                                "and storer_code='" & gU.dbEncode(str_code) & "' " & _
                                "and ILBS_QTY2 > 0 "


                    Dim balCount As Integer = 0

                    balCount = gU.decodeEmptyCInt(DB.getValueFromSQL(SQLString), 0)

                    If balCount > 0 Then
                        btnPrint.OnClientClick = "OpenItemLblsC('S');"
                    Else
                        btnPrint.OnClientClick = "OpenItemLblsC('N');"
                    End If


                End If


                PhotoTR.Visible = False
                photoImgTR.Visible = False
                photoDelTR.Visible = False

                Dim imgDT As DataTable
                SQLString = "Select itm_photo1, itm_photo2, itm_photo3 " & _
                            "from wms_item where wms_item.itm_code = '" & gU.dbEncode(pk_code) & "' " & _
                            "and wms_item.imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                            "and wms_item.pack_key = '" & gU.dbEncode(pk_code2) & "' " & _
                            "and storer_code='" & gU.dbEncode(str_code) & "'"

                imgDT = gDB.getDataTable(SQLString)

                If imgDT.Rows.Count > 0 Then
                    PhotoTR.Visible = True
                    photoImgTR.Visible = True
                    photoDelTR.Visible = True

                    Dim base64String As String
                    Dim imgByte As Byte()

                    If imgDT.Rows(0).Item("itm_photo1").ToString.Trim <> "" Then
                        pre_ITM_PHOTO1.Visible = True
                        img_ITM_PHOTO1.Visible = True
                        del_ITM_PHOTO1.Visible = True

                        imgByte = imgDT.Rows(0).Item("itm_photo1")
                        Session("Photo1Byte") = imgDT.Rows(0).Item("itm_photo1")
                        base64String = Convert.ToBase64String(imgByte, 0, imgByte.Length)
                        img_ITM_PHOTO1.ImageUrl = Convert.ToString("data:image/png;base64,") & base64String

                    Else
                        Session("Photo1Byte") = Nothing
                        pre_ITM_PHOTO1.Visible = False
                        img_ITM_PHOTO1.Visible = False
                        del_ITM_PHOTO1.Visible = False
                    End If

                    If imgDT.Rows(0).Item("itm_photo2").ToString.Trim <> "" Then
                        pre_ITM_PHOTO2.Visible = True
                        img_ITM_PHOTO2.Visible = True
                        del_ITM_PHOTO2.Visible = True

                        imgByte = imgDT.Rows(0).Item("itm_photo2")
                        Session("Photo2Byte") = imgDT.Rows(0).Item("itm_photo2")
                        base64String = Convert.ToBase64String(imgByte, 0, imgByte.Length)
                        img_ITM_PHOTO2.ImageUrl = Convert.ToString("data:image/png;base64,") & base64String

                    Else
                        Session("Photo2Byte") = Nothing
                        pre_ITM_PHOTO2.Visible = False
                        img_ITM_PHOTO2.Visible = False
                        del_ITM_PHOTO2.Visible = False
                    End If

                    If imgDT.Rows(0).Item("itm_PHOTO3").ToString.Trim <> "" Then
                        pre_ITM_PHOTO3.Visible = True
                        img_ITM_PHOTO3.Visible = True
                        del_ITM_PHOTO3.Visible = True

                        imgByte = imgDT.Rows(0).Item("itm_PHOTO3")
                        Session("PHOTO3Byte") = imgDT.Rows(0).Item("itm_PHOTO3")
                        base64String = Convert.ToBase64String(imgByte, 0, imgByte.Length)
                        img_ITM_PHOTO3.ImageUrl = Convert.ToString("data:image/png;base64,") & base64String

                    Else
                        Session("PHOTO3Byte") = Nothing
                        pre_ITM_PHOTO3.Visible = False
                        img_ITM_PHOTO3.Visible = False
                        del_ITM_PHOTO3.Visible = False
                    End If
                End If
                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "select IMP_CODE, STORER_CODE, PACK_KEY, ITM_CODE, VND_CODE, VND_NAME, ALV_VND_ITMCODE, " & _
                        "AITM_STATUS, AITM_NAME, AITM_DESC, AITM_NAME_CH, AITM_DESC_CH, " & _
                        "AITM_SPEC, AITM_UOM, AITM_PCS_PER_PACK, aitm_qty_per_ctn, AITM_LENGTH, AITM_WIDTH, AITM_HIGHT, " & _
                        "AITM_CBM, AITM_VOL, AITM_REF_CURR, AITM_REF_PRICE, AITM_REMARKS, AITM_PREF_WH, " & _
                        "AITM_PREF_LOC, AITM_PICTURE1, AITM_PICTURE2, " & _
                        "SYS_LUB, SYS_CB, SYS_LUD, SYS_CD, " & _
                        "CONVERT(nvarchar(30), ALV_DATE_ADDED, " & DDFORMAT & ") as ALV_DATE_ADDED, " & _
                        "AITM_NET_WEIGHT, AITM_GROSS_WEIGHT, AITM_ORIGIN, " & _
                        "'U' as mflag, wms_alt_vend_item.vnd_code as o_vnd_code, " & _
                        "'N' as has_pic1, 'N' as has_pic2, " & _
                        "wms_alt_vend_item.aitm_picture1 as old_pic1, wms_alt_vend_item.aitm_picture2 as old_pic2, " & _
                        "'N' as remove_pic1, 'N' as remove_pic2 " & _
                    "from wms_alt_vend_item " & _
                    "where wms_alt_vend_item.itm_code = '" & gU.dbEncode(pk_code) & "' " & _
                    "and wms_alt_vend_item.imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                    "and storer_code = '" & gU.dbEncode(ViewState("storer_code")) & "' " & _
                    "and pack_key = '" & gU.dbEncode(ViewState("pack_key")) & "'"

        SQLString = SQLString & " order by sys_lud"



        REM **********************
        dt = gDB.getDataTable(SQLString)

        GridView1.DataKeyNames = New String() {"itm_code", "pack_key", "storer_code", "vnd_code"}
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If

        ViewState("im_dt") = dt
        Session("avi_dt") = dt

        GridView1.DataBind()

        REM **********************
    End Sub

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        itm_status.Text = "CANCELLED"

        Session.Remove(itm_picture1_upload.ClientID)
        Session.Remove(itm_picture2_upload.ClientID)

        Call save()
        ar.sec_write = "N"
        CancelBtn.Visible = False
        ar.hideForm(Me)
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub itm_picture1_remove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itm_picture1_remove.Click
        ViewState("remove_image1") = True
        Session.Remove(itm_picture1_upload.ClientID)
        itm_picture1_upload.Enabled = False
        itm_picture1_remove.Enabled = False
        itm_picture1.Enabled = False
        preview_pic1.Enabled = False
        itm_picture1_upload.Visible = True
        itm_picture1_edit.Visible = False
        itm_picture1_upload_lit.Text = ""
        itm_picture1_upload_lit.Visible = False

    End Sub

    Protected Sub itm_picture2_remove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itm_picture2_remove.Click

        'If ViewState("remove_image1") Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", "You must  cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_itm_code.Text & "不能空白!", Session("gLang"))
        '    End If
        'End If

        ViewState("remove_image2") = True
        Session.Remove(itm_picture2_upload.ClientID)
        itm_picture2_upload.Enabled = False
        itm_picture2_remove.Enabled = False
        itm_picture2.Enabled = False
        preview_pic2.Enabled = False
        itm_picture2_upload.Visible = True
        itm_picture2_edit.Visible = False
        itm_picture2_upload_lit.Text = ""
        itm_picture2_upload_lit.Visible = False

    End Sub

    Private Sub removeImage(ByVal FileName As String)
        If FileName <> "" Then
            Try
                If IO.File.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & FileName) Then

                    thumb = New ThumbGenerator
                    thumb.SetParams(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & FileName, 150, 180)
                    Dim unID As String = thumb.GetUniqueThumbName
                    Cache.Remove(unID)

                    IO.File.Delete(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & FileName)
                End If

            Catch ex As Exception
                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            End Try
        End If
    End Sub

    Protected Sub itm_picture1_edit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itm_picture1_edit.Click
        Session.Remove(itm_picture1_upload.ClientID)
        itm_picture1_upload.Visible = True
        itm_picture1_upload_lit.Text = ""
        itm_picture1_upload_lit.Visible = False
        itm_picture1_edit.Visible = False
    End Sub

    Protected Sub itm_picture2_edit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itm_picture2_edit.Click
        Session.Remove(itm_picture2_upload.ClientID)
        itm_picture2_upload.Visible = True
        itm_picture2_upload_lit.Text = ""
        itm_picture2_upload_lit.Visible = False
        itm_picture2_edit.Visible = False
    End Sub

    Protected Sub BtnCopy_Click(sender As Object, e As System.EventArgs) Handles BtnCopy.Click
        CloneItem()
    End Sub

    Protected Sub CloneItem()
        Dim SuccessFlag As Boolean = False
        If validateAll() Then
            save("Y")


            Dim gConn As SqlConnection = gDB.getConnection
            Dim updateSQL As String = ""
            Dim SQLstring As String = ""
            Dim transaction As SqlTransaction = Nothing


            Dim i_code As String = ""
            Dim s_code As String = ""
            Dim p_key As String = ""
            Dim max_seq As String = ""
            Try
                gConn = gDB.getConnection()
                transaction = gConn.BeginTransaction()


                i_code = itm_code.Text
                s_code = storer_code.SelectedValue
                p_key = pack_key.Text

                SQLstring = "Select Max(CONVERT(int, PACK_KEY)+1) as pack_key from wms_item " & _
                            " Where IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "' AND STORER_CODE='" & gU.dbEncode(s_code) & "' AND ITM_CODE='" & gU.dbEncode(i_code) & "' "

                max_seq = gU.decodeNullOrEmpty(DB.getValueFromSQL(SQLstring, gConn, transaction), "")


                If max_seq <> "" Then

                    updateSQL = " Insert into wms_item (" & _
                                "IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ITM_STATUS, ITM_BARCODE, ITM_NAME, ITM_DESC, ITM_NAME_CH, ITM_DESC_CH, ITM_CAT, ITM_TYPE, ITM_BRAND, ITM_SERIES, ITM_SERIES_NO, " & _
                                "ITM_MODEL, ITM_PARENT, ITM_KIT_QTY, ITM_PREF_WH, ITM_PREF_LOC, ITM_REF_CURR, ITM_REF_PRICE, ITM_PRICE, ITM_BALANCE, ITM_BALANCE_CBM, ITM_BALANCE_KG, ITM_FLAGS, ITM_REMARKS, ITM_PICTURE1, ITM_PICTURE2,  " & _
                                "SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, ITM_PARENT_CODE, ITM_SPEC, ITM_SIZE_L, ITM_SIZE_W, ITM_SIZE_H, ITM_SCAN_CODE, ITM_VEND_ITM_NO, ITM_VEND_CODE, ITM_WHS_CODE, ITM_PROD_GROUP,  " & _
                                "ITM_PD_RCV_DATE,ITM_PD_CONTRACT_NO,ITM_PD_ARR_NOTICE_NO,ITM_PD_COND_OF_SPARES,ITM_PD_ST_1_YEAR,ITM_PD_ST_OVER_1_YEAR," & _
                                "ITM_SIZE, ITM_TEMP_YN, ITM_TEMP_FR_CHAR, ITM_TEMP_TO_CHAR, ITM_UOM, ITM_TEMP_FR, ITM_TEMP_TO, ITM_QTY_OF_UNIT, ITM_PCS_PER_UOM, ITM_SKU_NO, ITM_COLOR_CODE, ITM_SERIAL_NO,  " & _
                                "ITM_PART_NO, ITM_PROD_NO, PROJ_NO, ITM_SERIAL_NO_YN, ITM_STACKABLE_YN, ITM_INSP_YN, ITM_SHELF_LIFE, " & _
                                "ITM_GP_CODE, ITM_DIV_CODE, ITM_MFG,ITM_STORES_CLASS, ITM_EMB,ITM_HAZARD_CODE,ITM_CRITICAL_YN, " & _
                                "ITM_MAX_STOCK, ITM_RESTRICTED_ITEM,ITM_DRAWING_NO, ITM_CC, ITM_KEPT_DIV_CODE, ITM_EXPENSE_CODE, " & _
                                "ITM_INT_ORDER_QTY, ITM_1ST_INSP,ITM_ORO_YN, ITM_ROP_APL,ITM_ROP_LAM, ITM_ROP_CABLE, ITM_ROP_NP, " & _
                                "ITM_TI_YTD, ITM_TI_LYR,ITM_TI_PY2,ITM_TI_PY3,ITM_UOM2,ITM_QTY2, ITM_DG_YN, " & _
                                "ITM_REQ_STORE_COLD_YN,ITM_REQ_STORE_HUM_DESC,ITM_DG_CAT, " & _
                                "ITM_REQ_STORE_HUM_YN,ITM_REQ_STORE_AIRCON_YN, ITM_CHE_CLASS, ITM_REQ_MSDS_YN, ITM_WEIGHT_TYPE," & _
                                "ITM_MANUFACTURER, ITM_REPLACEMENT_COST, ITM_COMMODITY_CODE, ITM_RESTRICTED_RMK, itm_msds, itm_photo1, itm_photo2, itm_photo3 " & _
                                ") ( " & _
                                " SELECT wms_item.IMP_CODE,  wms_item.STORER_CODE,  wms_item.ITM_CODE,  '" & max_seq & "', " & _
                                "  wms_item.ITM_STATUS,  wms_item.ITM_BARCODE,  wms_item.ITM_NAME,  wms_item.ITM_DESC, " & _
                                "  wms_item.ITM_NAME_CH,  wms_item.ITM_DESC_CH,  wms_item.ITM_CAT,  wms_item.ITM_TYPE, " & _
                                "  wms_item.ITM_BRAND,  wms_item.ITM_SERIES,  wms_item.ITM_SERIES_NO,  wms_item.ITM_MODEL, " & _
                                "  wms_item.ITM_PARENT,  wms_item.ITM_KIT_QTY, wms_item.ITM_PREF_WH,  wms_item.ITM_PREF_LOC,  wms_item.ITM_REF_CURR, " & _
                                "  wms_item.ITM_REF_PRICE,  wms_item.ITM_PRICE,  wms_item.ITM_BALANCE,  wms_item.ITM_BALANCE_CBM,  wms_item.ITM_BALANCE_KG, " & _
                                "  wms_item.ITM_FLAGS,  wms_item.ITM_REMARKS,  wms_item.ITM_PICTURE1,  wms_item.ITM_PICTURE2, " & _
                                "  '" & Session("usr_id") & "',  Getdate(),  Getdate(),  '" & Session("usr_id") & "', " & _
                                "  wms_item.ITM_PARENT_CODE,  wms_item.ITM_SPEC,  wms_item.ITM_SIZE_L,  wms_item.ITM_SIZE_W, " & _
                                "  wms_item.ITM_SIZE_H,  wms_item.ITM_SCAN_CODE,  wms_item.ITM_VEND_ITM_NO,  wms_item.ITM_VEND_CODE, " & _
                                "  wms_item.ITM_WHS_CODE,  wms_item.ITM_PROD_GROUP,  " & _
                                "  wms_item.ITM_PD_RCV_DATE,wms_item.ITM_PD_CONTRACT_NO,wms_item.ITM_PD_ARR_NOTICE_NO,wms_item.ITM_PD_COND_OF_SPARES,wms_item.ITM_PD_ST_1_YEAR,wms_item.ITM_PD_ST_OVER_1_YEAR, " & _
                                " wms_item.ITM_SIZE,  wms_item.ITM_TEMP_YN,  wms_item.ITM_TEMP_FR_CHAR,  wms_item.ITM_TEMP_TO_CHAR, wms_item.ITM_UOM, " & _
                                " wms_item.ITM_TEMP_FR, wms_item.ITM_TEMP_TO,ITM_QTY_OF_UNIT, wms_item.itm_pcs_per_uom, wms_item.itm_sku_no, " & _
                                " wms_item.ITM_COLOR_CODE,wms_item.ITM_SERIAL_NO,wms_item.ITM_PART_NO,wms_item.ITM_PROD_NO, wms_item.proj_no, NULL, NULL, NULL, wms_item.ITM_SHELF_LIFE, " & _
                                "ITM_GP_CODE, " & _
                                "ITM_DIV_CODE, " & _
                                "ITM_MFG, " & _
                                "ITM_STORES_CLASS, " & _
                                "ITM_EMB, " & _
                                "ITM_HAZARD_CODE, " & _
                                "ITM_CRITICAL_YN, " & _
                                "ITM_MAX_STOCK, " & _
                                "ITM_RESTRICTED_ITEM, " & _
                                "ITM_DRAWING_NO, " & _
                                "ITM_CC, " & _
                                "ITM_KEPT_DIV_CODE, " & _
                                "ITM_EXPENSE_CODE, " & _
                                "ITM_INT_ORDER_QTY, " & _
                                "ITM_1ST_INSP, " & _
                                "ITM_ORO_YN, " & _
                                "ITM_ROP_APL, " & _
                                "ITM_ROP_LAM, " & _
                                "ITM_ROP_CABLE, " & _
                                "ITM_ROP_NP, " & _
                                "ITM_TI_YTD, " & _
                                "ITM_TI_LYR, " & _
                                "ITM_TI_PY2, " & _
                                "ITM_TI_PY3, " & _
                                "ITM_UOM2, " & _
                                "ITM_QTY2, " & _
                                "ITM_DG_YN, " & _
                                "ITM_REQ_STORE_COLD_YN, " & _
                                "ITM_REQ_STORE_HUM_DESC, " &
                                "ITM_DG_CAT, " & _
                                "ITM_REQ_STORE_HUM_YN, " & _
                                "ITM_REQ_STORE_AIRCON_YN, " & _
                                "ITM_CHE_CLASS, " & _
                                "ITM_REQ_MSDS_YN, " & _
                                "ITM_WEIGHT_TYPE, " & _
                                "ITM_MANUFACTURER, ITM_REPLACEMENT_COST, ITM_COMMODITY_CODE, ITM_RESTRICTED_RMK, itm_msds, itm_photo1, itm_photo2, itm_photo3 " & _
                                " FROM wms_item " & _
                                " WHERE IMP_CODE='" & Session("IMP_CODE") & "' AND STORER_CODE='" & gU.dbEncode(s_code) & "' AND ITM_CODE='" & gU.dbEncode(i_code) & "' AND PACK_KEY='" & gU.dbEncode(p_key) & "'" & _
                                ")"

                    gDB.amendData(updateSQL, gConn, transaction)


                    If Not transaction Is Nothing Then
                        transaction.Commit()
                        transaction = Nothing
                    End If

                    SuccessFlag = True
                Else

                    If Not transaction Is Nothing Then
                        transaction.Rollback()
                        transaction = Nothing
                    End If

                    uiFun.displayMsg(Me, "", "Error Occured when creating new item. Please try again later.", Session("gLang"))
                    Exit Sub

                End If

            Catch ex As Exception
                If Not transaction Is Nothing Then
                    transaction.Rollback()
                    transaction = Nothing
                End If

                Session.Remove(itm_picture1_upload.ClientID)
                Session.Remove(itm_picture2_upload.ClientID)

                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            Finally
                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try

            If SuccessFlag Then
                Dim rmtPost As New RemotePost

                rmtPost.Add("STORER_CODE", s_code)
                rmtPost.Add("ITM_CODE", i_code)
                rmtPost.Add("PACK_KEY", max_seq)
                rmtPost.alertMsg = "New Item was cloned Successfully. Forwarding to New Item."
                rmtPost.Url = "im_master.aspx"
                rmtPost.Post()
            End If
        End If

    End Sub

    Private Sub setPageCtrlAccess()

        exceptionEditList = New List(Of String)

        exceptionEditList.Add("btnPrint")
        exceptionEditList.Add("LabelSize")
    End Sub

    Protected Sub btnMSDS_Click(sender As Object, e As System.EventArgs) Handles btnMSDS.Click
        btnMSDS_ModalPopupExtender.Show()
        BindMSDS()
    End Sub

    Protected Sub btnPUpload_Click(sender As Object, e As System.EventArgs) Handles btnPUpload.Click
        If uploadMSDSPhoto() Then
            uiFun.displayMsgNew(UDMSDSP, "", "Upload Successfully!", Session("gLang"))
            BindMSDS()
        End If
        btnMSDS_ModalPopupExtender.Show()
    End Sub

    Private Sub BindMSDS()

        Dim selectSQL As String = "Select itm_msds from wms_item " & _
                                  " WHERE IMP_CODE='" & Session("IMP_CODE") & "' AND STORER_CODE='" & gU.dbEncode(ViewState("storer_code")) & "' AND ITM_CODE='" & gU.dbEncode(ViewState("wms_itm_code")) & "' AND PACK_KEY='" & gU.dbEncode(ViewState("pack_key")) & "'"
        Dim tempDT As DataTable = gDB.getDataTable(selectSQL)

        If tempDT.Rows.Count > 0 Then
            If tempDT.Rows(0).Item("itm_msds").ToString <> "" Then
                MSDS_TR.Visible = True
                Dim imgByte As Byte()

                imgByte = tempDT.Rows(0).Item("itm_msds")
                Dim base64String As String = Convert.ToBase64String(imgByte, 0, imgByte.Length)
                MSDS_IMG.ImageUrl = Convert.ToString("data:image/png;base64,") & base64String
                Session("ShowImgByte") = imgByte
            Else
                MSDS_TR.Visible = False
            End If

        Else
            MSDS_TR.Visible = False
        End If

    End Sub

    Private Function uploadMSDSPhoto() As Boolean
        Dim successFlag As Boolean = False

        If ITM_MSDS.HasFile Then
            Dim fileExt As String = Path.GetExtension(ITM_MSDS.FileName)

            If Not (UCase(fileExt) = ".JPG" OrElse UCase(fileExt) = ".GIF" OrElse UCase(fileExt) = ".PNG") Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "MSDS Photo must be image files!(*.jpg, *.png, *.gif)", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "MSDS Photo must be image files!(*.jpg, *.png, *.gif)", Session("gLang"))
                End If
                Return False
            End If


            If ITM_MSDS.PostedFile.ContentLength > fileSize Then
                uiFun.displayMsgNew(UDMSDSP, "", "MSDS Photo Size cannot larger than " & fileSize & "KB!", Session("gLang"))
                btnMSDS_ModalPopupExtender.Show()

            Else
                Using fs As Stream = ITM_MSDS.PostedFile.InputStream
                    Using br As New BinaryReader(fs)
                        Dim bytes As Byte() = br.ReadBytes(fs.Length)
                        Dim paP As GlobalDBFunc.DBCmdPara
                        Dim gConn As SqlConnection = gDB.getConnection
                        Dim transaction As SqlTransaction = Nothing

                        Try
                            gConn = gDB.getConnection()
                            transaction = gConn.BeginTransaction()


                            paP = New GlobalDBFunc.DBCmdPara
                            Dim updateSQL As String = "update wms_item set itm_msds=" & paP.AP(bytes, SqlDbType.VarBinary) & ", sys_lud=getdate(), sys_lub=" & paP.AP(Session("usr_id")) & _
                                                      " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(ViewState("storer_code"))) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(ViewState("wms_itm_code"))) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(ViewState("pack_key")))

                            gDB.amendData(updateSQL, gConn, transaction, paP)

                            transaction.Commit()
                            successFlag = True

                        Catch ex As Exception
                            transaction.Rollback()
                            transaction = Nothing
                            'Response.Write(ex.Message)
                            uiFun.displayMsgNew(UDMSDSP, "1008", "", Session("gLang"))
                        Finally

                            If gConn IsNot Nothing Then
                                If gConn.State = ConnectionState.Open Then
                                    gConn.Close()
                                    gConn.Dispose()
                                End If
                            End If

                        End Try
                    End Using
                End Using

            End If

        Else
            uiFun.displayMsgNew(UDMSDSP, "", "No file has been selected", Session("gLang"))
            btnMSDS_ModalPopupExtender.Show()
        End If

        Return successFlag

    End Function

    Protected Sub btnDelMSDS_Click(sender As Object, e As System.EventArgs) Handles btnDelMSDS.Click
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim gConn As SqlConnection = gDB.getConnection
        Dim transaction As SqlTransaction = Nothing
        Dim successFlag As Boolean = False

        Try
            gConn = gDB.getConnection()
            transaction = gConn.BeginTransaction()

            paP = New GlobalDBFunc.DBCmdPara
            Dim updateSQL As String = "update wms_item set itm_msds=NULL, sys_lud=getdate(), sys_lub=" & paP.AP(Session("usr_id")) & _
                                      " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(ViewState("storer_code"))) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(ViewState("wms_itm_code"))) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(ViewState("pack_key")))

            gDB.amendData(updateSQL, gConn, transaction, paP)

            transaction.Commit()
            successFlag = True

        Catch ex As Exception
            transaction.Rollback()
            transaction = Nothing
            'Response.Write(ex.Message)
            uiFun.displayMsgNew(UDMSDSP, "1008", "", Session("gLang"))
        Finally

            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If

        End Try

        If successFlag Then
            BindMSDS()
            uiFun.displayMsgNew(UDMSDSP, "", "MSDS Photo has been Deleted!", Session("gLang"))
        End If

        btnMSDS_ModalPopupExtender.Show()
    End Sub

    Protected Sub img_ITM_PHOTO1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles img_ITM_PHOTO1.Click
        Session("ShowImgByte") = Session("Photo1Byte")
        ScriptManager.RegisterStartupScript(UDP_IMGBTN1, UDP_IMGBTN1.GetType, "openIMG", "window.open('../../ShowImg.aspx', 'NewImage', ""titlebar=0,location=0,menubar=0,toolbar=0,resizable=yes,scrollbars=yes,personalbar=0,status=1,width=' + screen.width + ',height=' + screen.height"").focus();", True)

    End Sub

    Protected Sub img_ITM_PHOTO2_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles img_ITM_PHOTO2.Click
        Session("ShowImgByte") = Session("Photo2Byte")
        ScriptManager.RegisterStartupScript(UDP_IMGBTN2, UDP_IMGBTN2.GetType, "openIMG", "window.open('../../ShowImg.aspx', 'NewImage', ""titlebar=0,location=0,menubar=0,toolbar=0,resizable=yes,scrollbars=yes,personalbar=0,status=1,width=' + screen.width + ',height=' + screen.height"").focus();", True)
    End Sub

    Protected Sub img_ITM_PHOTO3_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles img_ITM_PHOTO3.Click
        Session("ShowImgByte") = Session("Photo3Byte")
        ScriptManager.RegisterStartupScript(UDP_IMGBTN3, UDP_IMGBTN3.GetType, "openIMG", "window.open('../../ShowImg.aspx', 'NewImage', ""titlebar=0,location=0,menubar=0,toolbar=0,resizable=yes,scrollbars=yes,personalbar=0,status=1,width=' + screen.width + ',height=' + screen.height"").focus();", True)
    End Sub

    Protected Sub del_ITM_PHOTO1_Click(sender As Object, e As System.EventArgs) Handles del_ITM_PHOTO1.Click
        RemovePhoto(1)
    End Sub

    Protected Sub del_ITM_PHOTO2_Click(sender As Object, e As System.EventArgs) Handles del_ITM_PHOTO2.Click
        RemovePhoto(2)
    End Sub

    Protected Sub del_ITM_PHOTO3_Click(sender As Object, e As System.EventArgs) Handles del_ITM_PHOTO3.Click
        RemovePhoto(3)
    End Sub

    Private Sub RemovePhoto(ByVal seq As Integer)
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim gConn As SqlConnection = gDB.getConnection
        Dim transaction As SqlTransaction = Nothing
        Dim successFlag As Boolean = False

        Try
            gConn = gDB.getConnection()
            transaction = gConn.BeginTransaction()

            paP = New GlobalDBFunc.DBCmdPara
            Dim updateSQL As String = "update wms_item set itm_photo" & seq & "=NULL, sys_lud=getdate(), sys_lub=" & paP.AP(Session("usr_id")) & _
                                      " WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(gU.dbEncode(ViewState("storer_code"))) & " AND ITM_CODE=" & paP.AP(gU.dbEncode(ViewState("wms_itm_code"))) & " AND PACK_KEY=" & paP.AP(gU.dbEncode(ViewState("pack_key")))

            gDB.amendData(updateSQL, gConn, transaction, paP)

            transaction.Commit()
            successFlag = True

        Catch ex As Exception
            transaction.Rollback()
            transaction = Nothing
            'Response.Write(ex.Message)
            uiFun.displayMsgNew(UDMSDSP, "1008", "", Session("gLang"))
        Finally

            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If

        End Try

        If successFlag Then
            Session("Photo" & seq & "Byte") = Nothing
            DirectCast(FindControl("pre_ITM_PHOTO" & seq), Label).Visible = False
            DirectCast(FindControl("img_ITM_PHOTO" & seq), ImageButton).ImageUrl = ""
            DirectCast(FindControl("img_ITM_PHOTO" & seq), ImageButton).Visible = False
            DirectCast(FindControl("del_ITM_PHOTO" & seq), LinkButton).Visible = False
            uiFun.displayMsg(Me, "", "Photo has been removed", Session("gLang"))
        End If
    End Sub

    Protected Sub IMG_SEQ1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles IMG_SEQ1.SelectedIndexChanged
        If IMG_SEQ2.SelectedValue = IMG_SEQ1.SelectedValue Then
            IMG_SEQ2.SelectedValue = PRE_IMG_SEQ1.Value
            PRE_IMG_SEQ2.Value = PRE_IMG_SEQ1.Value
        End If

        If IMG_SEQ3.SelectedValue = IMG_SEQ1.SelectedValue Then
            IMG_SEQ3.SelectedValue = PRE_IMG_SEQ1.Value
            PRE_IMG_SEQ3.Value = PRE_IMG_SEQ1.Value
        End If

        PRE_IMG_SEQ1.Value = IMG_SEQ1.SelectedValue
    End Sub

    Protected Sub IMG_SEQ2_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles IMG_SEQ2.SelectedIndexChanged
        If IMG_SEQ1.SelectedValue = IMG_SEQ2.SelectedValue Then
            IMG_SEQ1.SelectedValue = PRE_IMG_SEQ2.Value
            PRE_IMG_SEQ1.Value = PRE_IMG_SEQ2.Value
        End If

        If IMG_SEQ3.SelectedValue = IMG_SEQ2.SelectedValue Then
            IMG_SEQ3.SelectedValue = PRE_IMG_SEQ2.Value
            PRE_IMG_SEQ3.Value = PRE_IMG_SEQ2.Value
        End If

        PRE_IMG_SEQ2.Value = IMG_SEQ2.SelectedValue
    End Sub

    Protected Sub IMG_SEQ3_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles IMG_SEQ3.SelectedIndexChanged
        If IMG_SEQ2.SelectedValue = IMG_SEQ3.SelectedValue Then
            IMG_SEQ2.SelectedValue = PRE_IMG_SEQ3.Value
            PRE_IMG_SEQ2.Value = PRE_IMG_SEQ3.Value
        End If

        If IMG_SEQ1.SelectedValue = IMG_SEQ3.SelectedValue Then
            IMG_SEQ1.SelectedValue = PRE_IMG_SEQ3.Value
            PRE_IMG_SEQ1.Value = PRE_IMG_SEQ3.Value
        End If

        PRE_IMG_SEQ3.Value = IMG_SEQ3.SelectedValue
    End Sub
End Class
