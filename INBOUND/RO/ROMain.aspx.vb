Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Text

Partial Class INBOUND_RO_ROMain
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private dl As New DocLink
    Private moduleAction As String = ""
    Private imp_code As String = ""
    Private DDFORMAT As String = gU.getConfig("DDFORMATNO")
    Private DDFormat2 As String = gU.getConfig("DDFORMAT2")
    Private dt As New DataTable
    Private dtUOM, dtUOM2 As DataTable
    Private gvCol() As String = {"CB_COPY", "ROD_DISP_SEQ", "ITM_SKU_NO", "ROD_ITM_NAME", "ITM_DESC", "ROD_PALLET_NO", "ROD_CARTON_NO",
                                 "ROD_BATCH_NO", "ROD_REF_NO", "ROD_EXPIRY_DATE", "ROD_MANU_DATE", "ROD_ITM_PARENT", "ROD_SERIES_NO", "ROD_QTY",
                                 "ROD_OS_QTY", "ROD_UOM", "ROD_PCS_PER_UOM", "ROD_TOT_PCS", "ROD_LENGTH", "ROD_WIDTH", "ROD_HEIGHT", "ROD_KG", "ROD_CBM"}
    Dim SubTotalPCS As Double = 0.0
    Private exceptionEditList As List(Of String)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("IB_RO", Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            ViewState("STO_BATCH_FIELD_REF") = ""

            uiFun.load_dropdown(RO_WH_CODE, "select DISTINCT WH_MAIN_WH AS CODE, WH_MAIN_WH AS NAME from WMS_WAREHOUSE where imp_code='" & Session("imp_code") & "' ORDER BY 2", "CODE", "NAME", , Session("gSelectLabel"))

            uiFun.load_dropdown(PRJ_CODE, "select PRJ_CODE, PRJ_NAME from WMS_PROJECT ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(RO_SHIP_MODE, "select colc_code, colc_eng_value as colc_value from WMS_col_code where colc_tabcol='WMS_DELV_ORDER.DO_SHIP_MODE' ORDER BY colc_display_seq", "colc_code", "colc_value", , Session("gSelectLabel"))

            uiFun.load_dropdown(RO_CTRY_ORIGIN, "select ctry_code, ctry_name from WMS_country ORDER BY 2", "ctry_code", "ctry_name", , Session("gSelectLabel"))

            uiFun.load_dropdown(PO_TYPE, "Select PO_TYPE_CODE,PO_TYPE_DESC from PO_TYPE_MASTER ORDER BY 2", "PO_TYPE_CODE", "PO_TYPE_DESC", , Session("gSelectLabel"))

            uiFun.load_dropdown(PO_CAT, "Select PO_CATEGORY_CODE,PO_CATEGORY_DESC from PO_CATEGORY_MASTER ORDER BY 2", "PO_CATEGORY_CODE", "PO_CATEGORY_DESC", , Session("gSelectLabel"))

            If Session("pagemode") = "N" Then
                uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            End If
        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            btnConfirm.Visible = False
            If STORER_CODE.selectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
                'ViewState("STORER_CODE") = Session("usr_pref_storer")
                If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
                    STORER_CODE.Enabled = False
                End If
                Call setDefStorerInfo()
            End If
            If RO_DATE.Text = "" Then
                RO_DATE.Text = Now.ToString("dd/MM/yyyy")
            End If
            'btnConfirm.Visible = True
        End If

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Value = "WMS Purchase Order Maintenance"
            lbl_ImageHd.Text = "Order Items"
            lbl_RO_CODE.Text = "WPO Code"
            lbl_RO_STATUS.Text = "Status"
            lbl_STORER_CODE.Text = "Organizations"
            lbl_RO_DATE.Text = "Date"
            lbl_RO_RCV_BY.Text = "Received By"
            lbl_PRJ_CODE.Text = "Project Code"
            lbl_PO_TYPE.Text = "PO Type"
            lbl_RO_REF_NO.Text = "Contract No."
            lbl_RO_BATCH_NO.Text = "Batch No."
            lbl_RO_ETD.Text = "ETD"
            lbl_RO_ETA.Text = "ETA"
            lbl_RO_ISSUED_BY.Text = "Issued By"
            lbl_RO_STORER_TEL.Text = "Storer's Tel."
            lbl_RO_STORER_FAX.Text = "Storer's Fax"
            lbl_RO_STORER_EMAIL.Text = "Storer's Email"
            lbl_RO_STORER_CONT.Text = "Storer's Contact"
            lbl_RO_REM.Text = "Remarks"
            lbl_RO_SHIP_MODE.Text = "Ship Mode"
            lbl_RO_EDI_PO_NO.Text = "EBS PO No."
            'lbl_RO_TRACK_NO.Text = "Tracking No.:"
            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Close Order"
            'newrow.Text = "Add"
            btnNOTE.Text = "Notes"
            BtnCopy.Text = "Copy WPO"
            BtnGenBatchNo.Text = "Generate Batch No."
            cSBBtn.Text = "Check Stock Balance"
            btnConfirm.Text = "Confirm"
            btnDSCP.Text = "Print Discrepancy Report"
            btnPrintOrd.Value = "Print Order"
            btnPrintLbl.Value = "Print Item Label"

            btnConfirm.Text = "Confirm"
            selectItemBtn.Text = "Select Item"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to close this order?"");"
            reOpenBtn.OnClientClick = "return confirm(""Are you sure to re-open this order?"");"
            btnConfirm.OnClientClick = "return confirm(""Are you sure to confirm this order?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this order?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this order?"");"
            If Session("pagemode") = "N" Then
                RO_CODE.Text = "[No. will be auto generated]"
            End If
        ElseIf Session("gLang") = "C" Then
            lheader.Value = "添貨指令維護"
            lbl_ImageHd.Text = "訂單物件"
            lbl_RO_CODE.Text = "添貨指令號碼"
            lbl_RO_STATUS.Text = "狀態"
            lbl_STORER_CODE.Text = "部門"
            lbl_RO_DATE.Text = "日期"
            lbl_RO_RCV_BY.Text = "收貨人"
            lbl_PRJ_CODE.Text = "項目"
            lbl_PO_TYPE.Text = "類型"
            lbl_RO_REF_NO.Text = "貨主參考編號"
            lbl_RO_BATCH_NO.Text = "批次號"
            lbl_RO_ETD.Text = "預計出發時間"
            lbl_RO_ETA.Text = "預計到達時間"
            lbl_RO_ISSUED_BY.Text = "簽發者"
            lbl_RO_STORER_TEL.Text = "貨主電話"
            lbl_RO_STORER_FAX.Text = "貨主傳真"
            lbl_RO_STORER_EMAIL.Text = "貨主電郵"
            lbl_RO_STORER_CONT.Text = "貨主聯絡人"
            lbl_RO_REM.Text = "備註"
            lbl_RO_SHIP_MODE.Text = "運輸方式"
            lbl_RO_EDI_PO_NO.Text = "單號"
            'lbl_RO_TRACK_NO.Text = "追查編號:"
            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "儲存"
            saveBtn2.Text = "儲存"
            btnPrintOrd.Value = "打印訂單"
            btnPrintLbl.Value = "打印物品標籤"
            btnNOTE.Text = "工單"
            BtnCopy.Text = "複製訂單"
            BtnGenBatchNo.Text = "生成批號"
            cSBBtn.Text = "檢查貨品存庫"
            btnConfirm.Text = "確認"
            btnDSCP.Text = "差異報告"

            btnConfirm.Text = "確認"
            selectItemBtn.Text = "選擇物品"
            CancelBtn.OnClientClick = "return confirm(""確定關閉添貨指令?"");"
            btnConfirm.OnClientClick = "return confirm(""確定確認添貨指令?"");"
            saveBtn1.OnClientClick = "return confirm(""確定儲存添貨指令?"");"
            saveBtn2.OnClientClick = "return confirm(""確定儲存添貨指令?"");"
            CancelBtn.Text = "關閉添貨指令"
            'newrow.Text = "新增"

            lbl_RO_WH_CODE.Text = "子庫存"
            lbl_RO_CTRY_ORIGIN.Text = "來源國家"
            lbl_RO_SEAL_NO.Text = "封裝號碼"
            lbl_RO_CONTAINER_NO.Text = "貨櫃號碼"
            lbl_RO_CUST_INV_NO.Text = "客戶單據號碼"
            lbl_RO_CUST_INV_NO.Text = "運送模式"
        End If
        REM **********************

        REM **********************
        REM Additional CSS
        RO_DATE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            STORER_CODE.CssClass = "REQUIRED"
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("RO_CODE") = ""
            ViewState("STORER_CODE") = ""
            ViewState("FrmUP") = ""

            ViewState("RO_CODE") = Server.UrlDecode(Request("RO_CODE"))
            If Session("pagemode") <> "N" Then ViewState("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))
            ViewState("FrmUP") = Server.UrlDecode(Request("FrmUP"))

            Call BindGV()
        Else
            dt = ViewState("dt")
        End If

        'Set Access Right
        If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
            If Session("usr_pref_storer") <> STORER_CODE.Text Then
                Response.End()
            End If
        End If

        If moduleAction = "SELECTIM" Then
            addItemtoRO()
        End If

        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME AS STO_NAME from WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
        End If

        Call changeLabel()

        Dim colIdx_StartWith As Integer = 1

        ar.addColDef("ROD_DISP_SEQ", "rod_disp_seq", colIdx_StartWith)
        ar.addColDef("ROD_ITM_NAME", "rod_itm_name", colIdx_StartWith)
        ar.addColDef("ROD_PALLET_NO", "rod_pallet_no", colIdx_StartWith)
        ar.addColDef("ROD_CARTON_NO", "rod_carton_no", colIdx_StartWith)
        ar.addColDef("ROD_DOC_NO", "rod_doc_no", colIdx_StartWith)
        ar.addColDef("ROD_SERIES_NO", "rod_series_no", colIdx_StartWith)
        ar.addColDef("ROD_BATCH_NO", "rod_batch_no", colIdx_StartWith)
        ar.addColDef("ROD_REF_NO", "rod_ref_no", colIdx_StartWith)
        ar.addColDef("ROD_ITM_PARENT", "rod_itm_parent", colIdx_StartWith)
        'ar.addColDef("ROD_LOCATION", "rod_location", colIdx_StartWith)
        'ar.addColDef("ROD_SERIES_NO", "rod_series_no", colIdx_StartWith)
        ar.addColDef("ROD_QTY", "rod_qty", colIdx_StartWith)
        ar.addColDef("ROD_OS_QTY", "ROD_OS_QTY", colIdx_StartWith)
        ar.addColDef("ROD_UOM", "rod_uom", colIdx_StartWith)
        ar.addColDef("ROD_PCS_PER_UOM", "rod_pcs_per_uom", colIdx_StartWith)
        ar.addColDef("ROD_TOT_PCS", "rod_tot_pcs", colIdx_StartWith)
        'ar.addColDef("ROD_LENGTH", "rod_length", colIdx_StartWith)
        'ar.addColDef("ROD_WIDTH", "rod_width", colIdx_StartWith)
        'ar.addColDef("ROD_HEIGHT", "rod_height", colIdx_StartWith)
        'ar.addColDef("ROD_KG", "rod_kg", colIdx_StartWith)
        'ar.addColDef("ROD_CBM", "rod_cbm", colIdx_StartWith)

        'For i As Integer = 0 To GridView1.Rows.Count - 1
        '    Response.Write(i & "=" & CType(GridView1.Rows(i).FindControl("rod_itm_name"), TextBox).Text & "<BR>")
        'Next

        cm = New CommonMenu("RO", lheader.Value, RO_CODE.Text)
        cm.parentDir = "../../"
        cm.haveCheckList = "N"
        cm.haveAttachments = "N"
        cm.haveNotes = "N"
        cm.haveTasks = "N"
        cm.haveEmail = "N"
        cm.haveHistory = "N"

        cm.genCM(cmBar)

        Call dl.genLinkBar(linkBar, STORER_CODE.SelectedValue, "RO", RO_CODE.Text, "../../")

        setPageCtrlAccess()

        REM Select RO button
        If RO_STATUS.Text = "CLOSED" Then
            selectItemBtn.Attributes.Add("onclick", "ItemLookUp('" & STORER_CODE.SelectedValue & "');")
            cSBBtn.Attributes.Add("onclick", "checkSB('" & STORER_CODE.SelectedValue & "');")
        Else
            selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value);")
            cSBBtn.Attributes.Add("onclick", "checkSB(document.myform." & STORER_CODE.ClientID & ".value);")
        End If

        REM ****************************************************************
        If RO_STATUS.Text = "CLOSED" Then
            'ar.sec_write = "N"
            ar.sec_viewMode = "Y"
            CancelBtn.Visible = False
            btnConfirm.Visible = False
        ElseIf RO_STATUS.Text = "CONFIRMED" Then
            'ar.sec_write = "N"
            RO_CONTAINER_NO.CssClass = "REQUIRED"
            btnConfirm.Visible = False
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)
        'ar.hideGVForStorer(GridView1, STORER_CODE.Text, "RO", "WMS_REPLENISH_D")

        If ar.sec_viewMode = "Y" Then
            If GridView1.Rows.Count > 0 Then
                For i = 0 To GridView1.Rows.Count - 1
                    GridView1.Rows(i).FindControl("vw_itm_sku_no").Visible = False
                    GridView1.Rows(i).FindControl("vw_rod_itm_name").Visible = False
                Next
            End If
        End If

        If RO_STATUS.Text = "CLOSED" Then
            reOpenBtn.Visible = True
            reOpenBtn.Enabled = True
        Else
            reOpenBtn.Visible = False
            reOpenBtn.Enabled = False
        End If

        If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
            reOpenBtn.Visible = False
            BtnCopy.Visible = False
        End If

        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('IB_RO','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("RO_CODE") & "','N');")

        'ar.setFieldCustomize(Me, "IB_RO", STORER_CODE.SelectedValue, "WMS_REPLENISH")

        'If GridView1.Rows.Count > 0 Then
        '    ar.setGVCustomize(GridView1, "IB_RO", STORER_CODE.SelectedValue, "WMS_REPLENISH_D", gvCol)
        'End If


        Dim cIndex As Integer = 0

        If GridView1.Rows.Count > 0 Then
            For v As Integer = 0 To GridView1.Columns.Count - 1
                If GridView1.Columns(v).Visible Then
                    If GridView1.Columns(v).AccessibleHeaderText = "rod_tot_pcs" Then
                        Exit For
                    Else
                        cIndex += 1
                    End If
                End If
            Next
        End If

        ViewState("cIndex") = cIndex
    End Sub

    Private Function customizectrl(ByVal ctl As Control, ByRef ctrlArrayList As ArrayList) As Boolean
        customizectrl = False
    End Function

    Private Function page_customizectrl(ByVal ctl As Control) As Boolean
        page_customizectrl = False

        If ctl.ID = "cSBBtn" Then
            CType(ctl, Button).Enabled = True
            page_customizectrl = True
        End If
    End Function

    Protected Sub GridView1_DataBinding(sender As Object, e As System.EventArgs) Handles GridView1.DataBinding
        Dim selectSql As String

        selectSql = "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1"

        dtUOM = gDB.getDataTable(selectSql)

        selectSql = "select colc_code, colc_eng_value as colc_value from wms_col_code where colc_tabcol = 'WMS_ITEM.ITM_UOM2' order by COLC_DISPLAY_SEQ"

        dtUOM2 = gDB.getDataTable(selectSql)
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Dim i As Integer
        If e.CommandName = "SplitItem" Then
        End If
    End Sub

    Protected Sub GridView1_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Dim gvColdt As DataTable
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim selectSQL As String = ""

        'paP = New GlobalDBFunc.DBCmdPara
        'selectSQL = "Select upper(FLDO_FILED_CTRL) as FLDO_FILED_CTRL from WMS_FIELD_OPTION " & _
        '            "Where imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(ViewState("STORER_CODE")) & " AND FUN_CODE='IB_RO' AND ISNULL(FLDO_FIELD_OPTION, '') <> 'Y' AND FLDO_FIELD_TYPE is null " & _
        '            " AND FLDO_TABLE_NAME='WMS_REPLENISH_D' "

        'gvColdt = gDB.getDataTable(selectSQL, , , , paP)

        'If gvColdt.Rows.Count > 0 Then
        If 1 = 3 Then
            Select Case e.Row.RowType
                Case DataControlRowType.Header
                    For x = 0 To gvCol.Length - 1
                        If gvColdt.Select("FLDO_FILED_CTRL='" & gvCol(x) & "'").Length > 0 Then
                            e.Row.Cells(x).Visible = False
                        End If
                    Next

                Case DataControlRowType.DataRow
                    'For i = 0 To gvColdt.Rows.Count - 1
                    '    ctrl = e.Row.FindControl(gvColdt.Rows(i).Item("FLDO_FILED_CTRL").ToString.Trim)

                    '    If Not ctrl Is Nothing Then
                    '        ctrl.Visible = False
                    '    End If
                    'Next
                    Dim ctrlType As String = ""
                    For x = 0 To gvCol.Length - 1
                        If gvColdt.Select("FLDO_FILED_CTRL='" & gvCol(x) & "'").Length > 0 Then

                            If e.Row.Cells(x).Controls.Count > 0 Then
                                For Each ctl As Control In e.Row.Cells(x).Controls
                                    'Set control invisible
                                    Select Case TypeName(ctl).ToUpper
                                        Case "TEXTBOX"
                                            DirectCast(ctl, TextBox).Enabled = False
                                        Case "BUTTON"
                                            DirectCast(ctl, Button).Enabled = False

                                        Case "IMAGEBUTTON"
                                            DirectCast(ctl, ImageButton).Enabled = False

                                        Case "DROPDOWNLIST"
                                            DirectCast(ctl, DropDownList).Enabled = False

                                        Case "COMBOBOX"
                                            DirectCast(ctl, AjaxControlToolkit.ComboBox).Enabled = False

                                        Case "CHECKBOX"
                                            DirectCast(ctl, CheckBox).Enabled = False

                                        Case "CHECKBOXLIST"
                                            For Z = 0 To DirectCast(ctl, CheckBoxList).Items.Count - 1
                                                DirectCast(ctl, CheckBoxList).Items(Z).Enabled = False
                                            Next

                                        Case "IMAGE"
                                            DirectCast(ctl, Image).Visible = False
                                    End Select
                                Next
                            End If

                            e.Row.Cells(x).Visible = False
                        End If
                    Next
                Case DataControlRowType.Footer
                    For x = 0 To gvCol.Length - 1
                        If gvColdt.Select("FLDO_FILED_CTRL='" & gvCol(x) & "'").Length > 0 Then
                            e.Row.Cells(x).Visible = False
                        End If
                    Next
            End Select
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("rod_disp_seq"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_disp_seq").ToString.Trim

                CType(e.Row.FindControl("rod_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_pallet_no").ToString.Trim
                CType(e.Row.FindControl("rod_carton_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_carton_no").ToString.Trim
                'CType(e.Row.FindControl("rod_batch_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_batch_no").ToString.Trim
                uiFun.load_ComboBox(CType(e.Row.FindControl("ROD_BATCH_NO"), AjaxControlToolkit.ComboBox), _
                                    "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' order by 1", "dc_date_code", "dc_date_code", , "N/A", , False)

                CType(e.Row.FindControl("rod_batch_no"), AjaxControlToolkit.ComboBox).SelectedValue = DataBinder.Eval(e.Row.DataItem, "rod_batch_no").ToString.Trim
                CType(e.Row.FindControl("rod_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "rod_seq").ToString.Trim
                CType(e.Row.FindControl("rod_ref_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_ref_no").ToString.Trim
                CType(e.Row.FindControl("rod_itm_parent"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_itm_parent").ToString.Trim
                CType(e.Row.FindControl("rod_itm_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "rod_itm_code").ToString.Trim
                CType(e.Row.FindControl("vw_itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("rod_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_pack_key").ToString.Trim
                CType(e.Row.FindControl("rod_itm_name"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_itm_name").ToString.Trim
                CType(e.Row.FindControl("vw_rod_itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_itm_name").ToString.Trim
                CType(e.Row.FindControl("rod_os_qty"), Label).Text = gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "rod_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "rod_post_qty").ToString.Trim, 0)
                'CType(e.Row.FindControl("dsp_rod_vnd_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_vnd_code").ToString.Trim
                CType(e.Row.FindControl("rod_vnd_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "rod_vnd_code").ToString.Trim
                CType(e.Row.FindControl("dsp_rod_vnd_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_vnd_code").ToString.Trim

                Dim nDropDown As DropDownList
                'Dim nDropDown As DropDownList = CType(e.Row.FindControl("rod_uom"), DropDownList)
                'uiFun.load_dropdown(nDropDown, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
                'nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "rod_uom").ToString.Trim

                uiFun.load_dropdown(CType(e.Row.FindControl("rod_uom"), DropDownList), dtUOM, "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "rod_uom").ToString.Trim)

                uiFun.load_dropdown(CType(e.Row.FindControl("rod_uom2"), DropDownList), dtUOM2, "colc_code", "colc_value", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "rod_uom2").ToString.Trim)

                nDropDown = CType(e.Row.FindControl("rod_wh_code"), DropDownList)
                uiFun.load_dropdown(nDropDown, "select DISTINCT WH_MAIN_WH AS CODE, WH_MAIN_WH AS NAME from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' order by 1", "CODE", "NAME", , Session("gSelectLabel"))
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "rod_wh_code").ToString.Trim

                nDropDown = CType(e.Row.FindControl("rod_location_code"), DropDownList)
                uiFun.load_dropdown(nDropDown, "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & DataBinder.Eval(e.Row.DataItem, "rod_wh_code").ToString.Trim & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "rod_location_code").ToString.Trim

                'nDropDown = CType(e.Row.FindControl("rod_curr"), DropDownList)
                'uiFun.load_dropdownBy_ColCode(nDropDown, "CURRENCY", Session("gSelectLabel"))
                'nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "rod_curr").ToString.Trim

                CType(e.Row.FindControl("ROD_DOC_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_DOC_NO").ToString.Trim
                CType(e.Row.FindControl("ROD_SERIES_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_SERIES_NO").ToString.Trim



                CType(e.Row.FindControl("rod_qty"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_qty").ToString.Trim)
                CType(e.Row.FindControl("rod_pcs_per_uom"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_pcs_per_uom").ToString.Trim)
                CType(e.Row.FindControl("rod_tot_pcs"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_tot_pcs").ToString.Trim)
                'CType(e.Row.FindControl("rod_length"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_length").ToString.Trim)
                'CType(e.Row.FindControl("rod_width"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_width").ToString.Trim)
                'CType(e.Row.FindControl("rod_height"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_height").ToString.Trim)
                'CType(e.Row.FindControl("rod_kg"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_kg").ToString.Trim)

                'If IsDBNull(DataBinder.Eval(e.Row.DataItem, "rod_cbm")) Then
                '    CType(e.Row.FindControl("rod_cbm"), TextBox).Text = "0"
                'Else
                '    CType(e.Row.FindControl("rod_cbm"), TextBox).Text = CDbl(DataBinder.Eval(e.Row.DataItem, "rod_cbm")).ToString("###,###,##0.0000")
                'End If

                'CType(e.Row.FindControl("rod_location"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_location").ToString.Trim
                'CType(e.Row.FindControl("rod_series_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_series_no").ToString.Trim
                REM **********************

                CType(e.Row.FindControl("ROD_EXPIRY_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_EXPIRY_DATE").ToString.Trim

                    CType(e.Row.FindControl("ROD_MANU_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ROD_EXPIRY_DATE").ToString.Trim
                    If DataBinder.Eval(e.Row.DataItem, "ROD_EXPIRY_DATE").ToString.Trim = "" Then
                        CType(e.Row.FindControl("ROD_MANU_DATE"), Label).Text = ""
                    Else
                        Dim d As DateTime = DateTime.ParseExact(CType(e.Row.FindControl("ROD_MANU_DATE"), Label).Text, "dd/MM/yyyy", CultureInfo.InvariantCulture)
                        Dim reformatted As String = d.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                        CType(e.Row.FindControl("ROD_MANU_DATE"), Label).Text = reformatted
                    End If

                    'CType(e.Row.FindControl("ROD_EXPIRY_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_EXPIRY_DATE").ToString.Trim
                    'CType(e.Row.FindControl("ROD_MANU_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_MANU_DATE").ToString.Trim

                    Dim jsStr As String = "sumSubTotalPrice('" & CType(e.Row.FindControl("rod_qty"), TextBox).ClientID & "');"
                    ''" & CType(e.Row.FindControl("ROD_UNIT_PRICE"), TextBox).ClientID & "',
                    ','" & CType(e.Row.FindControl("PRICE_SUB_TOTAL"), TextBox).ClientID & "'

                    CType(e.Row.FindControl("rod_qty"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("rod_tot_pcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("rod_pcs_per_uom"), TextBox).ClientID & ".value * this.value;sumPCStotal();" & jsStr)
                    CType(e.Row.FindControl("rod_pcs_per_uom"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("rod_tot_pcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("rod_qty"), TextBox).ClientID & ".value * this.value;")


                    'CType(e.Row.FindControl("rod_length"), TextBox).Attributes("onkeypress") = "return maskKey(event)"
                    'CType(e.Row.FindControl("rod_width"), TextBox).Attributes("onkeypress") = "return maskKey(event)"
                    'CType(e.Row.FindControl("rod_height"), TextBox).Attributes("onkeypress") = "return maskKey(event)"

                    If DataBinder.Eval(e.Row.DataItem, "cb_copy").ToString.Trim = "1" Then
                        CType(e.Row.FindControl("cb_copy"), CheckBox).Checked = True
                    End If

                    SumPCS(cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_tot_pcs").ToString.Trim))

                    CType(e.Row.FindControl("rod_tot_pcs"), TextBox).Attributes("onkeyup") = "sumPCStotal()"

                    'CType(e.Row.FindControl("rod_length"), TextBox).Attributes.Add("onkeyup", _
                    '         "document.forms[0]." & CType(e.Row.FindControl("rod_cbm"), TextBox).ClientID & ".value = " & _
                    '         "(this.value * " & _
                    '         "document.forms[0]." & CType(e.Row.FindControl("rod_width"), TextBox).ClientID & ".value * " & _
                    '         "document.forms[0]." & CType(e.Row.FindControl("rod_height"), TextBox).ClientID & ".value) / 1000000")


                    'CType(e.Row.FindControl("rod_width"), TextBox).Attributes.Add("onkeyup", _
                    '         "document.forms[0]." & CType(e.Row.FindControl("rod_cbm"), TextBox).ClientID & ".value = " & _
                    '         "(document.forms[0]." & CType(e.Row.FindControl("rod_length"), TextBox).ClientID & ".value * " & _
                    '         "this.value * " & _
                    '         "document.forms[0]." & CType(e.Row.FindControl("rod_height"), TextBox).ClientID & ".value) / 1000000")


                    'CType(e.Row.FindControl("rod_height"), TextBox).Attributes.Add("onkeyup", _
                    '         "document.forms[0]." & CType(e.Row.FindControl("rod_cbm"), TextBox).ClientID & ".value = " & _
                    '         "(document.forms[0]." & CType(e.Row.FindControl("rod_length"), TextBox).ClientID & ".value * " & _
                    '         "document.forms[0]." & CType(e.Row.FindControl("rod_width"), TextBox).ClientID & ".value * " & _
                    '         "this.value) / 1000000")

                    'CType(e.Row.FindControl("rod_curr"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_EXPIRY_DATE").ToString.Trim

                    CType(e.Row.FindControl("rod_status"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_status").ToString.Trim
                    'CType(e.Row.FindControl("ROD_ON_BEHALF"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ROD_ON_BEHALF").ToString.Trim

                    CType(e.Row.FindControl("rod_qty2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_qty2").ToString.Trim
                'CType(e.Row.FindControl("ROD_CURR_RATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_CURR_RATE").ToString.Trim
                'CType(e.Row.FindControl("ROD_UNIT_PRICE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_UNIT_PRICE").ToString.Trim
                'CType(e.Row.FindControl("ROD_UNIT_PRICE"), TextBox).Attributes.Add("onchange", jsStr)
                'CType(e.Row.FindControl("ROD_HKD_EQU"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_HKD_EQU").ToString.Trim

                'CType(e.Row.FindControl("PRICE_SUB_TOTAL"), TextBox).Text = gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "ROD_UNIT_PRICE").ToString.Trim, 0) * gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "ROD_QTY").ToString.Trim, 0)
                'CType(e.Row.FindControl("PRICE_SUB_TOTAL"), TextBox).Attributes.Add("readonly", "readonly")
                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)


                    If DataBinder.Eval(e.Row.DataItem, "ROD_INSP_REQ").ToString.Trim = "Y" Then
                        CType(e.Row.FindControl("ROD_INSP_REQ"), CheckBox).Checked = True
                    Else
                        CType(e.Row.FindControl("ROD_INSP_REQ"), CheckBox).Checked = False
                    End If

                    'nButton.ID = "btnDelete" & e.Row.RowIndex

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

                    If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "CN" OrElse DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "NSCN" Then

                        For k As Integer = 0 To e.Row.Cells.Count - 1
                            e.Row.Cells(k).CssClass = "GV_COPY"
                        Next
                    End If

                    If PO_TYPE.SelectedValue = "PO" Then
                        CType(e.Row.FindControl("vw_itm_sku_no"), Label).Visible = True
                        CType(e.Row.FindControl("itm_sku_no"), TextBox).Visible = False
                        CType(e.Row.FindControl("rod_itm_name"), TextBox).Visible = False
                        CType(e.Row.FindControl("vw_rod_itm_name"), Label).Visible = True

                        CType(e.Row.FindControl("rod_uom"), DropDownList).Enabled = False
                        CType(e.Row.FindControl("rod_uom2"), DropDownList).Enabled = False
                    Else
                        CType(e.Row.FindControl("vw_itm_sku_no"), Label).Visible = False
                        CType(e.Row.FindControl("itm_sku_no"), TextBox).Visible = True
                        CType(e.Row.FindControl("rod_itm_name"), TextBox).Visible = True
                        CType(e.Row.FindControl("vw_rod_itm_name"), Label).Visible = False

                        If xFlag = "NSN" OrElse xFlag = "NSCN" Then
                            CType(e.Row.FindControl("rod_uom"), DropDownList).Enabled = True
                            CType(e.Row.FindControl("itm_sku_no"), TextBox).Enabled = True
                            CType(e.Row.FindControl("rod_uom2"), DropDownList).Enabled = True
                            CType(e.Row.FindControl("rod_itm_name"), TextBox).Enabled = True
                        Else
                            CType(e.Row.FindControl("rod_uom"), DropDownList).Enabled = False
                            CType(e.Row.FindControl("rod_uom2"), DropDownList).Enabled = False
                            CType(e.Row.FindControl("itm_sku_no"), TextBox).Visible = False
                            CType(e.Row.FindControl("rod_itm_name"), TextBox).Visible = False
                            CType(e.Row.FindControl("vw_itm_sku_no"), Label).Visible = True
                            CType(e.Row.FindControl("vw_rod_itm_name"), Label).Visible = True
                        End If

                    End If

                'If Not ar.hasBtnRight("BT_RO_PRICE_EDIT") Then
                '    CType(e.Row.FindControl("ROD_CURR_RATE"), TextBox).Attributes.Add("readonly", "readonly")
                '    CType(e.Row.FindControl("ROD_CURR_RATE"), TextBox).CssClass = "READONLY"
                '    CType(e.Row.FindControl("ROD_UNIT_PRICE"), TextBox).Attributes.Add("readonly", "readonly")
                '    CType(e.Row.FindControl("ROD_UNIT_PRICE"), TextBox).CssClass = "READONLY"
                '    CType(e.Row.FindControl("rod_curr"), DropDownList).Enabled = False
                'End If

                'ar.hideGVForStorer(GridView1, STORER_CODE.Text, "RO", "WMS_REPLENISH_D", e)
            Case DataControlRowType.Footer
                        CType(e.Row.FindControl("rod_sub_total_pcs"), Label).Text = GetTotal()
                End Select
    End Sub

    'Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
    '    If STORER_CODE.SelectedValue <> "" Then

    '        If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
    '            Dim rows_count As Integer = 0
    '            REM **********************
    '            REM Modify Here
    '            Dim seq_string As String = "select max(CONVERT(int,rod_seq)) + 1 from wms_replenish_d " &
    '                                        "where imp_code = '" & gU.dbEncode(imp_code.Trim.ToString) & "' " &
    '                                        "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
    '                                        "and ro_code = '" & gU.dbEncode(RO_CODE.Text) & "' "
    '            REM **********************
    '            Dim nS_dt As New DataTable
    '            nS_dt = gDB.getDataTable(seq_string)
    '            Dim next_seq_no, temp_no As String
    '            Dim temp_seq_no As Integer = 1

    '            next_seq_no = ""

    '            If nS_dt.Rows.Count > 0 Then
    '                next_seq_no = nS_dt.Rows(0).Item(0).ToString()
    '            Else
    '                temp_no = "1"
    '            End If

    '            If next_seq_no = "" Then next_seq_no = "1"

    '            If ViewState("n_cur_seq") = "" Then
    '                ViewState("n_cur_seq") = next_seq_no
    '            Else
    '                temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
    '                ViewState("n_cur_seq") = temp_seq_no.ToString
    '            End If

    '            Dim newRow As DataRow
    '            newRow = dt.NewRow

    '            newRow.Item("mFlag") = "NSN"
    '            newRow.Item("rod_seq") = ViewState("n_cur_seq").ToString
    '            newRow.Item("rod_wh_code") = RO_WH_CODE.SelectedValue
    '            newRow.Item("ROD_STATUS") = "NEW"
    '            newRow.Item("ROD_DISP_SEQ") = dt.Rows.Count + 1
    '            newRow.Item("ROD_PACK_KEY") = "1"
    '            newRow.Item("rod_pcs_per_uom") = 1

    '            dt.Rows.Add(newRow)
    '            dt.AcceptChanges()

    '            ViewState("dt") = dt
    '            GridView1.DataSource = dt
    '            GridView1.DataBind()

    '            Dim cIndex As Integer = 0

    '            If GridView1.Rows.Count > 0 Then
    '                For v As Integer = 0 To GridView1.Columns.Count - 1
    '                    If GridView1.Columns(v).Visible Then
    '                        If GridView1.Columns(v).AccessibleHeaderText = "rod_tot_pcs" Then
    '                            Exit For
    '                        Else
    '                            cIndex += 1
    '                        End If
    '                    End If
    '                Next
    '            End If

    '            ViewState("cIndex") = cIndex

    '        End If
    '    Else
    '        uiFun.displayMsg(Me, "", "Please Select Storer Code.", Session("gLang"))

    '    End If
    'End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""
        Dim i As Integer

        'If RO_CODE.Text.Trim = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_RO_CODE.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_RO_CODE.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        If RO_STATUS.Text = "CONFIRMED" Then
            If String.IsNullOrWhiteSpace(RO_CONTAINER_NO.Text) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Please enter Container No.!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "Please enter Container No.!", Session("gLang"))
                End If
                Return False
            End If
        End If

        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If RO_WH_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_RO_WH_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_RO_WH_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If RO_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_RO_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_RO_DATE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        ElseIf Not gU.isValidDate(RO_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_RO_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_RO_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If RO_ETD.Text.Trim <> "" And Not gU.isValidDate(RO_ETD.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_RO_ETD.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_RO_ETD.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If RO_ETA.Text.Trim <> "" And Not gU.isValidDate(RO_ETA.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_RO_ETA.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_RO_ETA.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If String.IsNullOrWhiteSpace(RO_EDI_PO_NO.Text) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Please enter PO#!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "Please enter PO#!", Session("gLang"))
            End If
            Return False
        End If

        If PO_TYPE.SelectedValue <> "PO" Then
            If String.IsNullOrWhiteSpace(RO_REF_NO.Text) Then
                'If Session("gLang") = "E" Then
                '    uiFun.displayMsg(Me, "", "Please enter Contract No.!", Session("gLang"))
                'Else
                '    uiFun.displayMsg(Me, "", "Please enter Contract No.!", Session("gLang"))
                'End If
                'Return False
            Else
                Dim tempCount As Integer = 0
                selectSql = "Select count(*) from wms_replenish where imp_code='" & gU.dbEncode(imp_code) & "' " &
                            "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and upper(RO_REF_NO)=upper('" & gU.dbEncode(RO_REF_NO.Text.Trim) & "')"

                If Session("pagemode") <> "N" Then
                    selectSql &= " AND RO_CODE <> '" & ViewState("RO_CODE") & "'"
                End If

                tempCount = gU.decodeEmptyCInt(DB.getValueFromSQL(selectSql), 0)

                If tempCount > 0 Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Contract No. already exists. Please enter another Contract No.!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "Contract No. already exists. Please enter another Contract No.!", Session("gLang"))
                    End If
                    Return False
                End If

            End If
        End If

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1

                If CType(GridView1.Rows(i).FindControl("ROD_EXPIRY_DATE"), TextBox).Text <> "" AndAlso Not gU.isValidDate(CType(GridView1.Rows(i).FindControl("ROD_EXPIRY_DATE"), TextBox).Text.Trim) Then
                    uiFun.displayMsg(Me, "", "Invalid Expiry Date!", Session("gLang"))
                    CType(GridView1.Rows(i).FindControl("ROD_EXPIRY_DATE"), TextBox).Focus()
                    Return False
                End If

                'If CType(GridView1.Rows(i).FindControl("ROD_MANU_DATE"), TextBox).Text <> "" AndAlso Not gU.isValidDate(CType(GridView1.Rows(i).FindControl("ROD_MANU_DATE"), TextBox).Text.Trim) Then
                '    uiFun.displayMsg(Me, "", "Invalid Manufactory Date!", Session("gLang"))
                '    CType(GridView1.Rows(i).FindControl("ROD_MANU_DATE"), TextBox).Focus()
                '    Return False
                'End If

                'Not String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("ROD_MANU_DATE"), TextBox).Text) AndAlso

                If Not String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("ROD_EXPIRY_DATE"), TextBox).Text) Then
                    Dim customDateTimeFormat As DateTimeFormatInfo = New DateTimeFormatInfo()
                    customDateTimeFormat.DateSeparator = "/"
                    customDateTimeFormat.TimeSeparator = ":"
                    customDateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
                    customDateTimeFormat.LongDatePattern = "dd/MM/yyyy"
                    customDateTimeFormat.ShortTimePattern = "HH:mm"
                    customDateTimeFormat.LongTimePattern = "HH:mm"
                    customDateTimeFormat.FullDateTimePattern = "dd/MM/yyyy HH:mm"

                    'Dim manu_date, exp_date As Date
                    Dim exp_date As Date

                    'manu_date = Convert.ToDateTime(CType(GridView1.Rows(i).FindControl("ROD_MANU_DATE"), Label).Text, customDateTimeFormat)
                    exp_date = Convert.ToDateTime(CType(GridView1.Rows(i).FindControl("ROD_EXPIRY_DATE"), TextBox).Text, customDateTimeFormat)

                    'If manu_date > exp_date Then
                    '    If Session("gLang") = "E" Then
                    '        uiFun.displayMsg(Me, "", "Sotck No.(" & CType(GridView1.Rows(i).FindControl("ITM_SKU_NO"), TextBox).Text & "): " & "Manu. Date cannot later then Expiry Date!", Session("gLang"))
                    '    Else
                    '        uiFun.displayMsg(Me, "", "Sotck No.(" & CType(GridView1.Rows(i).FindControl("ITM_SKU_NO"), TextBox).Text & "): " & "Manu. Date cannot later then Expiry Date!", Session("gLang"))
                    '    End If
                    '    Return False
                    'End If
                End If



                If DirectCast(GridView1.Rows(i).FindControl("ROD_WH_CODE"), DropDownList).SelectedValue = "" Then
                    uiFun.displayMsg(Me, "", "Please select Warehouse for items!", Session("gLang"))
                    DirectCast(GridView1.Rows(i).FindControl("ROD_WH_CODE"), DropDownList).Focus()
                    Return False
                End If

                'If DirectCast(GridView1.Rows(i).FindControl("ITM_SKU_NO"), TextBox).Text = "" Then
                '    uiFun.displayMsg(Me, "", "Please Enter SKU No.!", Session("gLang"))
                '    DirectCast(GridView1.Rows(i).FindControl("ITM_SKU_NO"), DropDownList).Focus()
                '    Return False
                'End If
                'Updated by Aswin 10/07/2020
                If DirectCast(GridView1.Rows(i).FindControl("rod_itm_name"), TextBox).Text = "" Then
                    uiFun.displayMsg(Me, "", "Please Enter Item Name!", Session("gLang"))
                    DirectCast(GridView1.Rows(i).FindControl("rod_itm_name"), TextBox).Focus()
                    Return False
                End If

                If uiFun.gvValidate(Me, dt, "rod_qty", "QTY",
                                 CType(GridView1.Rows(i).FindControl("rod_qty"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "rod_pcs_per_uom", "PCS per UOM",
                                 CType(GridView1.Rows(i).FindControl("rod_pcs_per_uom"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "rod_tot_pcs", "Total Number",
                                 CType(GridView1.Rows(i).FindControl("rod_tot_pcs"), TextBox).Text) = False Then Return False

                'If uiFun.gvValidate(Me, dt, "rod_length", "Length",
                '                 CType(GridView1.Rows(i).FindControl("rod_length"), TextBox).Text) = False Then Return False

                'If uiFun.gvValidate(Me, dt, "rod_width", "Width",
                '                 CType(GridView1.Rows(i).FindControl("rod_width"), TextBox).Text) = False Then Return False

                'If uiFun.gvValidate(Me, dt, "rod_height", "Height",
                '                 CType(GridView1.Rows(i).FindControl("rod_height"), TextBox).Text) = False Then Return False

                'If uiFun.gvValidate(Me, dt, "rod_kg", "Weight",
                '                 CType(GridView1.Rows(i).FindControl("rod_kg"), TextBox).Text) = False Then Return False

                'If uiFun.gvValidate(Me, dt, "rod_cbm", "CBM",
                '                 CType(GridView1.Rows(i).FindControl("rod_cbm"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "rod_qty2", "Qty2",
                                 CType(GridView1.Rows(i).FindControl("rod_qty2"), TextBox).Text) = False Then Return False

                ''If uiFun.gvValidate(Me, dt, "ROD_CURR_RATE", "Currency Rate",
                '                CType(GridView1.Rows(i).FindControl("ROD_CURR_RATE"), TextBox).Text) = False Then Return False

                'If uiFun.gvValidate(Me, dt, "ROD_UNIT_PRICE", "Unit Price",
                '                CType(GridView1.Rows(i).FindControl("ROD_UNIT_PRICE"), TextBox).Text) = False Then Return False

                'If uiFun.gvValidate(Me, dt, "ROD_HKD_EQU", "HKD Equivalent",
                '                CType(GridView1.Rows(i).FindControl("ROD_HKD_EQU"), TextBox).Text) = False Then Return False

                If PO_TYPE.SelectedValue = "CPO" Then
                    Dim ro_code = RO_CODE_HF.Value.Trim.Split("-")(0)

                    Dim rod_seq = CType(GridView1.Rows(i).FindControl("rod_seq"), HiddenField).Value.ToString.Trim
                    Dim cpo_qty = Convert.ToDouble(CType(GridView1.Rows(i).FindControl("rod_qty"), TextBox).Text.Trim)
                    Dim SqlStr As String = "select IsNUll(ROD_QTY,0)ROD_QTY  from WMS_REPLENISH_D Where ROD_SEQ = '" + rod_seq + "' and RO_CODE='" + ro_code + "' and STORER_CODE='" + STORER_CODE.SelectedValue + "' and IMP_CODE='" + imp_code + "'"
                    Dim dtTemp As DataTable = gDB.getDataTable(SqlStr)
                    If dtTemp.Rows.Count > 0 Then
                        Dim po_qty = Convert.ToDouble(dtTemp.Rows(0)("ROD_QTY").ToString.Trim)
                        If (po_qty < cpo_qty) Then
                            uiFun.displayMsg(Me, "", "CPO qty is greater than PO qty!", Session("gLang"))
                            DirectCast(GridView1.Rows(i).FindControl("rod_qty"), TextBox).Focus()
                            Return False
                        End If
                    Else
                        uiFun.displayMsg(Me, "", "Data not found in PO!", Session("gLang"))
                        Return False
                    End If
                End If
            Next
        End If

        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim nextNo As String = ""
        Dim nextSEQ As String = ""
        Dim gConn As SqlConnection
        Dim dupSQL As String
        Dim dupTbl As New DataTable
        Dim updateSQL As String = ""
        Dim InsertIMSQL As String = ""
        Dim onBehalfYN As String = "N"
        Dim rod_batch_no As String
        Dim shelfLife As Integer = 0
        'Dim expected_EXP_DATE As String = ""
        Dim newITM_CODE As String = ""
        Dim newSKU_NO As String = ""

        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                imp_code = Session("imp_code")
                If PO_TYPE.SelectedValue.ToString = "CPO" Then
                    If validateCPO() = False Then
                        uiFun.displayMsg(Me, "", "CPO Qty exceeds PO Qty", Session("gLang"))
                        Return

                    End If
                End If
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("RO", gConn, transaction)
                    'nextNo = RO_CODE.Text
                    REM **********************

                    If RO_TRACK_NO.Value = "" Then
                        RO_TRACK_NO.Value = DB.getDocNo("TRACKNO", gConn, transaction)
                    End If
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    dupSQL = "select 1 from wms_replenish " &
                            "where ro_code = '" & gU.dbEncode(nextNo) & "' " &
                            "and imp_code='" & gU.dbEncode(imp_code) & "' " &
                            "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into wms_replenish ( " &
                                        "ro_code, imp_code, storer_code, " &
                                        "po_type,po_cat, " &
                                        "prj_code, ro_status, ro_type, " &
                                        "ro_date, ro_etd, ro_eta, " &
                                        "ro_issued_by, ro_rcv_by, ro_ref_no, " &
                                        "ro_batch_no, ro_storer_cont, ro_storer_tel, " &
                                        "ro_storer_fax, ro_storer_email, ro_dest, ro_rem, ro_track_no, " &
                                        "RO_SEAL_NO,RO_CONTAINER_NO, RO_CUST_INV_NO, RO_SHIP_MODE, RO_CTRY_ORIGIN,RO_EDI_PO_NO, RO_WH_CODE, " &
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                        "values ( " &
                                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(imp_code)) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(PO_TYPE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(PO_CAT.SelectedValue)) & "," &
                                        gU.convdbNVCData(gU.dbEncode(PRJ_CODE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(RO_STATUS.Text)) & "," & gU.convdbNVCData(gU.dbEncode(PO_TYPE.SelectedValue)) & ", " &
                                        gU.convdbDate(gU.dbEncode(RO_DATE.Text)) & "," & gU.convdbDate(gU.dbEncode(RO_ETD.Text)) & "," & gU.convdbDate(gU.dbEncode(RO_ETA.Text)) & ", " &
                                        gU.convdbNVCData(gU.dbEncode(RO_ISSUED_BY.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_RCV_BY.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_REF_NO.Text)) & ", " &
                                        gU.convdbNVCData(gU.dbEncode(RO_BATCH_NO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_STORER_CONT.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_STORER_TEL.Text)) & ", " &
                                        gU.convdbNVCData(gU.dbEncode(RO_STORER_FAX.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_STORER_EMAIL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_DEST.Text)) & ", " &
                                        gU.convdbNVCData(gU.dbEncode(RO_REM.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_TRACK_NO.Value)) & "," &
                                        gU.convdbNVCData(gU.dbEncode(RO_SEAL_NO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_CONTAINER_NO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_CUST_INV_NO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_SHIP_MODE.SelectedValue)) & "," &
                                        gU.convdbNVCData(gU.dbEncode(RO_CTRY_ORIGIN.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(RO_EDI_PO_NO.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(RO_WH_CODE.SelectedValue)) & "," &
                                        "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                        REM **********************

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then

                            uiFun.reOrderDetails(dt, "rod_disp_seq")

                            For Each rows As DataRow In dt.Rows
                                itemSQL = ""
                                REM **********************
                                REM Modify Here
                                rod_batch_no = ""
                                rod_batch_no = returnBatchNO(ViewState("STO_BATCH_FIELD_REF"), rows.Item("rod_batch_no").ToString.Trim, gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""), gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))

                                If rows.Item("rod_wh_code").ToString.Trim <> RO_WH_CODE.SelectedValue Then
                                    onBehalfYN = "Y"
                                Else
                                    onBehalfYN = "N"
                                End If

                                Select Case rows.Item("mFlag")
                                    Case "N", "CN"
                                        If rod_batch_no <> "" Then
                                            dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(rod_batch_no) & "'"
                                            dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                            If dupTbl.Rows.Count <= 0 Then

                                                updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                            " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" &
                                                            "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                            "'" & gU.dbEncode(rod_batch_no) & "'," &
                                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                                gDB.amendData(updateSQL, gConn, transaction)
                                            End If
                                        End If

                                        shelfLife = gU.decodeEmptyCInt(DB.getValueFromSQL("Select ITM_SHELF_LIFE from wms_item where imp_code='" & imp_code & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and itm_code='" & gU.dbEncode(rows.Item("rod_itm_code").ToString.Trim) & "'", gConn, transaction), 0)

                                        'If rows.Item("ROD_EXPIRY_DATE").ToString.Trim = "" AndAlso rows.Item("ROD_MANU_DATE").ToString.Trim <> "" AndAlso shelfLife > 0 Then
                                        '    expected_EXP_DATE = gU.decodeNullOrEmpty(DB.getValueFromSQL("select convert(varchar, convert(datetime,'" & rows.Item("ROD_MANU_DATE").ToString.Trim & "',103) + " & shelfLife & ",103)"), "")
                                        'Else
                                        '    expected_EXP_DATE = rows.Item("ROD_EXPIRY_DATE").ToString.Trim
                                        'End If


                                        itemSQL = "insert into wms_replenish_d ( " &
                                                    "ro_code, imp_code, storer_code, rod_seq, rod_disp_seq, " &
                                                    "rod_pallet_no, rod_carton_no, rod_batch_no, rod_ref_no, " &
                                                    "rod_itm_parent, rod_itm_code, rod_pack_key, rod_itm_name, " &
                                                    "rod_qty, rod_uom, rod_pcs_per_uom, rod_tot_pcs, rod_length, " &
                                                    "rod_width, rod_height, rod_kg, rod_cbm,rod_series_no,rod_doc_no, rod_vnd_code, ROD_EXPIRY_DATE, ROD_MANU_DATE," &
                                                    " ROD_STATUS,ROD_UOM2, ROD_QTY2,ROD_CURR, ROD_CURR_RATE, ROD_UNIT_PRICE, ROD_ON_BEHALF, ROD_WH_CODE, ROD_LOCATION_CODE, ROD_HKD_EQU, ROD_INSP_REQ, " &
                                                    " sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                    "values ( " &
                                                    gU.convdbNVCData(gU.dbEncode(nextNo)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_seq").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_disp_seq").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pallet_no").ToString.Trim, "000"))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_carton_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(rod_batch_no)) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_ref_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_parent").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_code").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_pack_key").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_name").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_qty").ToString.Trim, "0")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_uom").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pcs_per_uom").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_tot_pcs").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_length").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_width").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_height").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_kg").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_cbm").ToString.Trim, "0")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_series_no").ToString.Trim, ""))) & ", " &
                                                     gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_doc_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_vnd_code").ToString.Trim, ""))) & ", " &
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                                    "'NEW'," & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_UOM2").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_QTY2").ToString.Trim, "0")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_CURR").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_CURR_RATE").ToString.Trim, "0")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_UNIT_PRICE").ToString.Trim, "0")) & ", " &
                                                    "'" & onBehalfYN & "'," &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_WH_CODE").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_LOCATION_CODE").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_HKD_EQU").ToString.Trim, "0")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_INSP_REQ").ToString.Trim, ""))) & ", " &
                                                    "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                        'gU.convdbDate(gU.dbEncode(gU.decodeNull(expected_EXP_DATE, ""))) & ", " &
                                    Case "NSN", "NSCN"
                                        newITM_CODE = DB.getDocNo("NONSTOCKITM", gConn, transaction)
                                        newSKU_NO = ""

                                        If rows.Item("itm_sku_no").ToString.Trim = "" Then newSKU_NO = newITM_CODE Else newSKU_NO = rows.Item("itm_sku_no").ToString.Trim

                                        If rod_batch_no <> "" Then
                                            dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(rod_batch_no) & "'"
                                            dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                            If dupTbl.Rows.Count <= 0 Then

                                                updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                            " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" &
                                                            "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                            "'" & gU.dbEncode(rod_batch_no) & "'," &
                                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                                gDB.amendData(updateSQL, gConn, transaction)
                                            End If
                                        End If
                                        InsertIMSQL = " INSERT INTO WMS_ITEM " &
                                                        " (IMP_CODE, STORER_CODE, PACK_KEY, ITM_CODE,  ITM_STATUS, ITM_NAME, ITM_DESC, ITM_VEND_CODE, ITM_SKU_NO,ITM_UOM, ITM_PCS_PER_UOM, ITM_NONSTOCK_YN, " &
                                                        " ITM_SIZE_L, ITM_SIZE_W, ITM_SIZE_H, PROJ_NO,  SYS_CD, SYS_CB,SYS_LUD, SYS_LUB) " &
                                                        " VALUES('" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ",'1','" & gU.dbEncode(newITM_CODE) & "','NEW'," &
                                                        gU.convdbNVCData(gU.dbEncode(rows.Item("ROD_ITM_NAME").ToString.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(rows.Item("ROD_ITM_NAME").ToString.Trim)) & ",'DEF_VEND'," & gU.convdbNVCData(gU.dbEncode(newSKU_NO)) & "," &
                                                        gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_uom").ToString.Trim, ""))) & ", " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pcs_per_uom").ToString.Trim, "0")) & ",'Y', " &
                                                        gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_length").ToString.Trim, "0")) & ", " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_width").ToString.Trim, "0")) & ", " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_height").ToString.Trim, "0")) & ", " &
                                                        "'" & gU.dbEncode(RO_REF_NO.Text.Trim) & "'," &
                                                        "GetDate(),'" & Session("usr_id") & "',GetDate(),'" & Session("usr_id") & "') "

                                        gDB.amendData(InsertIMSQL, gConn, transaction)

                                        InsertIMSQL = " INSERT INTO WMS_ALT_VEND_ITEM " &
                                                  " (IMP_CODE, STORER_CODE, PACK_KEY, ITM_CODE, VND_CODE, VND_NAME, ALV_DATE_ADDED, AITM_STATUS," &
                                                  " AITM_UOM, AITM_PCS_PER_PACK,AITM_QTY_PER_CTN," &
                                                  " sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                  " VALUES ('" & imp_code & "', " & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", '1','" & gU.dbEncode(newITM_CODE) & "', 'DEF_VEND', 'Default Vendor', getdate(), 'Active'," &
                                                  " 'PCS', 1, 1," &
                                                  "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                        gDB.amendData(InsertIMSQL, gConn, transaction)

                                        itemSQL = "insert into wms_replenish_d ( " &
                                                  "ro_code, imp_code, storer_code, rod_seq, rod_disp_seq, " &
                                                  "rod_pallet_no, rod_carton_no, rod_batch_no, rod_ref_no, " &
                                                  "rod_itm_parent, rod_itm_code, rod_pack_key, rod_itm_name, " &
                                                  "rod_qty, rod_uom, rod_pcs_per_uom, rod_tot_pcs, rod_length, " &
                                                  "rod_width, rod_height, rod_kg, rod_cbm,rod_series_no,rod_doc_no, rod_vnd_code, ROD_EXPIRY_DATE, ROD_MANU_DATE," &
                                                  " ROD_STATUS,ROD_UOM2, ROD_QTY2,ROD_CURR, ROD_CURR_RATE, ROD_UNIT_PRICE, ROD_ON_BEHALF, ROD_WH_CODE, ROD_LOCATION_CODE, ROD_HKD_EQU, ROD_INSP_REQ, " &
                                                  " sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                  "values ( " &
                                                  gU.convdbNVCData(gU.dbEncode(nextNo)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_seq").ToString.Trim, ""))) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_disp_seq").ToString.Trim, ""))) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pallet_no").ToString.Trim, "000"))) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_carton_no").ToString.Trim, ""))) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(rod_batch_no)) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_ref_no").ToString.Trim, ""))) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_parent").ToString.Trim, ""))) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(newITM_CODE)) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_pack_key").ToString.Trim, "1"))) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_name").ToString.Trim, ""))) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_qty").ToString.Trim, "0")) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_uom").ToString.Trim, ""))) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pcs_per_uom").ToString.Trim, "0")) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_tot_pcs").ToString.Trim, "0")) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_length").ToString.Trim, "0")) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_width").ToString.Trim, "0")) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_height").ToString.Trim, "0")) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_kg").ToString.Trim, "0")) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_cbm").ToString.Trim, "0")) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_series_no").ToString.Trim, ""))) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_doc_no").ToString.Trim, ""))) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_vnd_code").ToString.Trim, ""))) & ", " &
                                                  gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                                  gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                                  "'NEW'," & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_UOM2").ToString.Trim, ""))) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_QTY2").ToString.Trim, "0")) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_CURR").ToString.Trim, ""))) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_CURR_RATE").ToString.Trim, "0")) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_UNIT_PRICE").ToString.Trim, "0")) & ", " &
                                                  "'" & onBehalfYN & "'," &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_WH_CODE").ToString.Trim, ""))) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_LOCATION_CODE").ToString.Trim, ""))) & ", " &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_HKD_EQU").ToString.Trim, "0")) & ", " &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_INSP_REQ").ToString.Trim, ""))) & ", " &
                                                  "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                        'gU.convdbDate(gU.dbEncode(gU.decodeNull(expected_EXP_DATE, ""))) & ", " &

                                End Select
                                REM **********************

                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                            Next
                        End If

                        If Session("usr_type") = "C" OrElse Session("usr_type") = "T" Then
                            SendCustROMail(nextNo, STORER_CODE.SelectedValue)
                        End If

                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Replenishment Order!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "添貨指令資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_replenish set " &
                                    "prj_code = " & gU.convdbNVCData(gU.dbEncode(PRJ_CODE.Text)) & ", " &
                                    "ro_status = " & gU.convdbNVCData(gU.dbEncode(RO_STATUS.Text)) & ", " &
                                    "po_type = " & gU.convdbNVCData(gU.dbEncode(PO_TYPE.SelectedValue)) & ", " &
                                    "po_cat = " & gU.convdbNVCData(gU.dbEncode(PO_CAT.SelectedValue)) & ", " &
                                    "ro_date = " & gU.convdbDate(gU.dbEncode(RO_DATE.Text)) & ", " &
                                    "ro_etd = " & gU.convdbDate(gU.dbEncode(RO_ETD.Text)) & ", " &
                                    "ro_eta = " & gU.convdbDate(gU.dbEncode(RO_ETA.Text)) & ", " &
                                    "ro_issued_by = " & gU.convdbNVCData(gU.dbEncode(RO_ISSUED_BY.Text)) & ", " &
                                    "ro_rcv_by = " & gU.convdbNVCData(gU.dbEncode(RO_RCV_BY.Text)) & ", " &
                                    "ro_ref_no = " & gU.convdbNVCData(gU.dbEncode(RO_REF_NO.Text)) & ", " &
                                    "ro_batch_no = " & gU.convdbNVCData(gU.dbEncode(RO_BATCH_NO.Text)) & ", " &
                                    "ro_storer_cont = " & gU.convdbNVCData(gU.dbEncode(RO_STORER_CONT.Text)) & ", " &
                                    "ro_storer_tel = " & gU.convdbNVCData(gU.dbEncode(RO_STORER_TEL.Text)) & ", " &
                                    "ro_storer_fax = " & gU.convdbNVCData(gU.dbEncode(RO_STORER_FAX.Text)) & ", " &
                                    "ro_storer_email = " & gU.convdbNVCData(gU.dbEncode(RO_STORER_EMAIL.Text)) & ", " &
                                    "ro_dest = " & gU.convdbNVCData(gU.dbEncode(RO_DEST.Text)) & ", " &
                                    "ro_rem = " & gU.convdbNVCData(gU.dbEncode(RO_REM.Text)) & ", " &
                                    "ro_track_no = " & gU.convdbNVCData(gU.dbEncode(RO_TRACK_NO.Value)) & ", " &
                                    "RO_SEAL_NO = " & gU.convdbNVCData(gU.dbEncode(RO_SEAL_NO.Text)) & ", " &
                                    "RO_CONTAINER_NO = " & gU.convdbNVCData(gU.dbEncode(RO_CONTAINER_NO.Text)) & ", " &
                                    "RO_CUST_INV_NO = " & gU.convdbNVCData(gU.dbEncode(RO_CUST_INV_NO.Text)) & ", " &
                                    "RO_SHIP_MODE = " & gU.convdbNVCData(gU.dbEncode(RO_SHIP_MODE.SelectedValue)) & ", " &
                                    "RO_CTRY_ORIGIN = " & gU.convdbNVCData(gU.dbEncode(RO_CTRY_ORIGIN.SelectedValue)) & ", " &
                                    "RO_EDI_PO_NO=" & gU.convdbNVCData(gU.dbEncode(RO_EDI_PO_NO.Text.Trim)) & ", " &
                                    "RO_WH_CODE=" & gU.convdbNVCData(gU.dbEncode(RO_WH_CODE.SelectedValue)) & ", " &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                    "where imp_code = '" & gU.dbEncode(imp_code) & "' " &
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                    "and ro_code = '" & gU.dbEncode(RO_CODE.Text) & "' "

                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then

                        uiFun.reOrderDetails(dt, "rod_disp_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""

                            rod_batch_no = ""
                            rod_batch_no = returnBatchNO(ViewState("STO_BATCH_FIELD_REF"), rows.Item("rod_batch_no").ToString.Trim, gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""), gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))

                            shelfLife = gU.decodeEmptyCInt(DB.getValueFromSQL("Select ITM_SHELF_LIFE from wms_item where imp_code='" & imp_code & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and itm_code='" & gU.dbEncode(rows.Item("rod_itm_code").ToString.Trim) & "'", gConn, transaction), 0)

                            'If rows.Item("ROD_EXPIRY_DATE").ToString.Trim = "" AndAlso rows.Item("ROD_MANU_DATE").ToString.Trim <> "" AndAlso shelfLife > 0 Then
                            '    expected_EXP_DATE = gU.decodeNullOrEmpty(DB.getValueFromSQL("select convert(varchar, convert(datetime,'" & rows.Item("ROD_MANU_DATE").ToString.Trim & "',103) + " & shelfLife & ",103)"), "")
                            'Else
                            '    expected_EXP_DATE = rows.Item("ROD_EXPIRY_DATE").ToString.Trim
                            'End If

                            If rows.Item("rod_wh_code").ToString.Trim <> RO_WH_CODE.SelectedValue Then
                                onBehalfYN = "Y"
                            Else
                                onBehalfYN = "N"
                            End If

                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N", "CN"
                                    If rod_batch_no <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(rod_batch_no) & "'"
                                        dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTbl.Rows.Count <= 0 Then

                                            updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values (" &
                                                        "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                        "'" & gU.dbEncode(rod_batch_no) & "'," &
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(updateSQL, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "insert into wms_replenish_d ( " &
                                              "ro_code, imp_code, storer_code, rod_seq, rod_disp_seq, " &
                                              "rod_pallet_no, rod_carton_no, rod_batch_no, rod_ref_no, " &
                                              "rod_itm_parent, rod_itm_code, rod_pack_key, rod_itm_name, " &
                                              "rod_qty, rod_uom, rod_pcs_per_uom, rod_tot_pcs, rod_length, " &
                                              "rod_width, rod_height, rod_kg, rod_cbm, rod_series_no,rod_doc_no, rod_vnd_code, ROD_EXPIRY_DATE, ROD_MANU_DATE," &
                                              " ROD_STATUS,ROD_UOM2, ROD_QTY2,ROD_CURR, ROD_CURR_RATE, ROD_UNIT_PRICE, ROD_ON_BEHALF, ROD_WH_CODE, ROD_LOCATION_CODE, ROD_HKD_EQU, ROD_INSP_REQ, " &
                                              " sys_cb, sys_cd, sys_lub, sys_lud) " &
                                              "values ( " &
                                              gU.convdbNVCData(gU.dbEncode(RO_CODE.Text.Trim)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_seq").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_disp_seq").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pallet_no").ToString.Trim, "000"))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_carton_no").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(rod_batch_no)) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_ref_no").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_parent").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_code").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_pack_key").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_name").ToString.Trim, ""))) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_qty").ToString.Trim, "0")) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_uom").ToString.Trim, ""))) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pcs_per_uom").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_tot_pcs").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_length").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_width").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_height").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_kg").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_cbm").ToString.Trim, "0")) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_series_no").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_doc_no").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_vnd_code").ToString.Trim, ""))) & ", " &
                                               gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                              gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                              "'NEW'," & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_UOM2").ToString.Trim, ""))) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_QTY2").ToString.Trim, "0")) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_CURR").ToString.Trim, ""))) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_CURR_RATE").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_UNIT_PRICE").ToString.Trim, "0")) & ", " &
                                              "'" & onBehalfYN & "'," &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_WH_CODE").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_LOCATION_CODE").ToString.Trim, ""))) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_HKD_EQU").ToString.Trim, "0")) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_INSP_REQ").ToString.Trim, ""))) & ", " &
                                              "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                    'gU.convdbDate(gU.dbEncode(gU.decodeNull(expected_EXP_DATE, ""))) & ", " &

                                Case "D"

                                    itemSQL = "delete from wms_replenish_d " &
                                            "where imp_code = '" & gU.dbEncode(imp_code) & "' " &
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "and ro_code = '" & gU.dbEncode(RO_CODE.Text.Trim) & "' " &
                                            "and rod_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("rod_seq").ToString.Trim, "")) & "' "

                                Case "NSN", "NSCN"
                                    newITM_CODE = DB.getDocNo("NONSTOCKITM", gConn, transaction)
                                    newSKU_NO = ""

                                    If rows.Item("itm_sku_no").ToString.Trim = "" Then newSKU_NO = newITM_CODE Else newSKU_NO = rows.Item("itm_sku_no").ToString.Trim

                                    If rod_batch_no <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(rod_batch_no) & "'"
                                        dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTbl.Rows.Count <= 0 Then

                                            updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values (" &
                                                        "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                        "'" & gU.dbEncode(rod_batch_no) & "'," &
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(updateSQL, gConn, transaction)
                                        End If
                                    End If

                                    InsertIMSQL = " INSERT INTO WMS_ITEM " &
                                                  " (IMP_CODE, STORER_CODE, PACK_KEY, ITM_CODE,  ITM_STATUS, ITM_NAME, ITM_DESC, ITM_VEND_CODE, ITM_SKU_NO,ITM_UOM, ITM_PCS_PER_UOM, ITM_NONSTOCK_YN, " &
                                                  " ITM_SIZE_L, ITM_SIZE_W, ITM_SIZE_H, PROJ_NO,  SYS_CD, SYS_CB,SYS_LUD, SYS_LUB) " &
                                                  " VALUES('" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ",'1','" & gU.dbEncode(newITM_CODE) & "','NEW'," &
                                                  gU.convdbNVCData(gU.dbEncode(rows.Item("ROD_ITM_NAME").ToString.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(rows.Item("ROD_ITM_NAME").ToString.Trim)) & ",'DEF_VEND'," & gU.convdbNVCData(newSKU_NO) & "," &
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_uom").ToString.Trim, ""))) & ", " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pcs_per_uom").ToString.Trim, "0")) & ", 'Y'," &
                                                  gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_length").ToString.Trim, "0")) & ", " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_width").ToString.Trim, "0")) & ", " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_height").ToString.Trim, "0")) & ", " &
                                                  "'" & gU.dbEncode(RO_REF_NO.Text.Trim) & "'," &
                                                  "GetDate(),'" & Session("usr_id") & "',GetDate(),'" & Session("usr_id") & "') "

                                    gDB.amendData(InsertIMSQL, gConn, transaction)

                                    InsertIMSQL = " INSERT INTO WMS_ALT_VEND_ITEM " &
                                                  " (IMP_CODE, STORER_CODE, PACK_KEY, ITM_CODE, VND_CODE, VND_NAME, ALV_DATE_ADDED, AITM_STATUS," &
                                                  " AITM_UOM, AITM_PCS_PER_PACK,AITM_QTY_PER_CTN," &
                                                  " sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                  " VALUES ('" & imp_code & "', " & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", '1','" & gU.dbEncode(newITM_CODE) & "', 'DEF_VEND', 'Default Vendor', getdate(), 'Active'," &
                                                  " 'PCS', 1, 1," &
                                                  "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                    gDB.amendData(InsertIMSQL, gConn, transaction)

                                    itemSQL = "insert into wms_replenish_d ( " &
                                              "ro_code, imp_code, storer_code, rod_seq, rod_disp_seq, " &
                                              "rod_pallet_no, rod_carton_no, rod_batch_no, rod_ref_no, " &
                                              "rod_itm_parent, rod_itm_code, rod_pack_key, rod_itm_name, " &
                                              "rod_qty, rod_uom, rod_pcs_per_uom, rod_tot_pcs, rod_length, " &
                                              "rod_width, rod_height, rod_kg, rod_cbm,rod_series_no,rod_doc_no, rod_vnd_code, ROD_EXPIRY_DATE, ROD_MANU_DATE," &
                                              " ROD_STATUS,ROD_UOM2, ROD_QTY2,ROD_CURR, ROD_CURR_RATE, ROD_UNIT_PRICE, ROD_ON_BEHALF, ROD_WH_CODE, ROD_LOCATION_CODE, ROD_HKD_EQU, ROD_INSP_REQ, " &
                                              " sys_cb, sys_cd, sys_lub, sys_lud) " &
                                              "values ( " &
                                              gU.convdbNVCData(gU.dbEncode(RO_CODE.Text.Trim)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_seq").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_disp_seq").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pallet_no").ToString.Trim, "000"))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_carton_no").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(rod_batch_no)) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_ref_no").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_parent").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(newITM_CODE)) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_pack_key").ToString.Trim, "1"))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_name").ToString.Trim, ""))) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_qty").ToString.Trim, "0")) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_uom").ToString.Trim, ""))) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pcs_per_uom").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_tot_pcs").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_length").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_width").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_height").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_kg").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_cbm").ToString.Trim, "0")) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_series_no").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_doc_no").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_vnd_code").ToString.Trim, ""))) & ", " &
                                              gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                              gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                              "'NEW'," & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_UOM2").ToString.Trim, ""))) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_QTY2").ToString.Trim, "0")) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_CURR").ToString.Trim, ""))) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_CURR_RATE").ToString.Trim, "0")) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_UNIT_PRICE").ToString.Trim, "0")) & ", " &
                                              "'" & onBehalfYN & "'," &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_WH_CODE").ToString.Trim, ""))) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_LOCATION_CODE").ToString.Trim, ""))) & ", " &
                                              gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_HKD_EQU").ToString.Trim, "0")) & ", " &
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_INSP_REQ").ToString.Trim, ""))) & ", " &
                                              "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                    'gU.convdbDate(gU.dbEncode(gU.decodeNull(expected_EXP_DATE, ""))) & ", " &

                                Case Else
                                    If rod_batch_no <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(rod_batch_no) & "'"
                                        dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTbl.Rows.Count <= 0 Then

                                            updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" &
                                                        "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                        "'" & gU.dbEncode(rod_batch_no) & "'," &
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"
                                            gDB.amendData(updateSQL, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "update wms_replenish_d set " &
                                                "rod_disp_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_disp_seq").ToString.Trim, ""))) & ", " &
                                                "rod_pallet_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pallet_no").ToString.Trim, "000"))) & ", " &
                                                "rod_carton_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_carton_no").ToString.Trim, ""))) & ", " &
                                                "rod_batch_no = " & gU.convdbNVCData(gU.dbEncode(rod_batch_no)) & ", " &
                                                "rod_ref_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_ref_no").ToString.Trim, ""))) & ", " &
                                                "rod_itm_parent = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_parent").ToString.Trim, ""))) & ", " &
                                                "rod_itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_code").ToString.Trim, ""))) & ", " &
                                                "rod_pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_pack_key").ToString.Trim, ""))) & ", " &
                                                "rod_itm_name = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_name").ToString.Trim, ""))) & ", " &
                                                "rod_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_qty").ToString.Trim, "0")) & ", " &
                                                "rod_uom = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_uom").ToString.Trim, ""))) & ", " &
                                                "rod_pcs_per_uom = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pcs_per_uom").ToString.Trim, "0")) & ", " &
                                                "rod_tot_pcs = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_tot_pcs").ToString.Trim, "0")) & ", " &
                                                "rod_length = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_length").ToString.Trim, "0")) & ", " &
                                                "rod_width = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_width").ToString.Trim, "0")) & ", " &
                                                "rod_height = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_height").ToString.Trim, "0")) & ", " &
                                                "rod_kg = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_kg").ToString.Trim, "0")) & ", " &
                                                "rod_cbm = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_cbm").ToString.Trim, "0")) & ", " &
                                                "rod_series_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_series_no").ToString.Trim, ""))) & ", " &
                                                "rod_doc_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_doc_no").ToString.Trim, ""))) & ", " &
                                                "rod_vnd_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_vnd_code").ToString.Trim, ""))) & ", " &
                                                "ROD_EXPIRY_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                                "ROD_MANU_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                                "ROD_UOM2=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_UOM2").ToString.Trim, ""))) & ", " &
                                                "ROD_QTY2=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_QTY2").ToString.Trim, "0")) & "," &
                                                "ROD_CURR=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_CURR").ToString.Trim, ""))) & ", " &
                                                "ROD_CURR_RATE=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_CURR_RATE").ToString.Trim, "0")) & ", " &
                                                "ROD_UNIT_PRICE=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_UNIT_PRICE").ToString.Trim, "0")) & "," &
                                                "ROD_ON_BEHALF='" & onBehalfYN & "'," &
                                                "ROD_WH_CODE=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_WH_CODE").ToString.Trim, ""))) & "," &
                                                "ROD_LOCATION_CODE=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_LOCATION_CODE").ToString.Trim, ""))) & "," &
                                                "ROD_HKD_EQU=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ROD_HKD_EQU").ToString.Trim, "0")) & "," &
                                                "ROD_INSP_REQ = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ROD_INSP_REQ").ToString.Trim, ""))) & "," &
                                                "sys_lub = '" & Session("usr_id") & "', " &
                                                "sys_lud = Getdate() " &
                                                "where imp_code = '" & gU.dbEncode(imp_code) & "' " &
                                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                "and ro_code = '" & gU.dbEncode(RO_CODE.Text.Trim) & "' " &
                                                "and rod_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("rod_seq").ToString.Trim, "")) & "'"
                            End Select
                            REM **********************

                            'Response.Write(itemSQL & "<BR>")
                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                        Next
                    End If
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    RO_CODE.Text = nextNo
                    ViewState("RO_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    RO_CODE.ForeColor = Drawing.Color.Black
                    RO_CODE.Font.Size = 10
                    REM **********************
                End If

                If flag <> "Y" Then uiFun.displayMsg(Me, "1007", "", Session("gLang"))

                Call BindGV()

                changeLabel()

            Catch ex As Exception
                transaction.Rollback()
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
        'Dim pk_code As String = ""
        Dim pk_code, storerCode As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("RO_CODE") <> "" Then
            'pk_code = ViewState("RO_CODE")
            pk_code = ViewState("RO_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            'pk_code = Request("RO_CODE")
            pk_code = Server.UrlDecode(Request("RO_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))
        End If
        REM **********************

        RO_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            RO_CODE.ForeColor = Drawing.Color.Red
            RO_STATUS.Text = "NEW"
            btnPrintLbl.Visible = False
            btnPrintOrd.Visible = False
            btnAttach.Visible = False
            BtnCopy.Visible = False
            BtnGenBatchNo.Visible = False
            btnPrintPalletLabel.Visible = False
            btnNOTE.Visible = False
            If PO_TYPE.SelectedValue = "PO" Then
                selectItemBtn.Visible = True
                'newrow.Visible = False
            Else
                selectItemBtn.Visible = True
                'newrow.Visible = True
            End If


            ViewState("STORER_CODE") = STORER_CODE.SelectedValue

            REM **********************
        Else
            btnPrintLbl.Visible = True
            btnPrintOrd.Visible = True
            btnAttach.Visible = True
            PO_TYPE.Enabled = False
            BtnCopy.Visible = True
            BtnGenBatchNo.Visible = True
            btnPrintPalletLabel.Visible = True
            btnNOTE.Visible = True
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = " SELECT wms_replenish.IMP_CODE,  wms_replenish.STORER_CODE,  wms_replenish.RO_CODE, " &
                        "  wms_replenish.PRJ_CODE,  wms_replenish.RO_STATUS,  wms_replenish.RO_TYPE,wms_replenish.PO_CAT, " &
                        "  CONVERT(varchar, wms_replenish.RO_DATE," & DDFORMAT & ") as RO_DATE,  CONVERT(varchar, wms_replenish.RO_ETD," & DDFORMAT & ") as RO_ETD,  CONVERT(varchar, wms_replenish.RO_ETA," & DDFORMAT & ") as RO_ETA , " &
                        "  wms_replenish.RO_ISSUED_BY,  wms_replenish.RO_RCV_BY,  wms_replenish.RO_REF_NO, " &
                        "  wms_replenish.RO_BATCH_NO,  wms_replenish.RO_STORER_CONT,  wms_replenish.RO_STORER_TEL, " &
                        "  wms_replenish.RO_STORER_FAX,  wms_replenish.RO_STORER_EMAIL,  wms_replenish.RO_DEST, " &
                        "  wms_replenish.RO_REM,  wms_replenish.RO_TRACK_NO,  wms_replenish.SYS_LUB, " &
                        "  wms_replenish.SYS_LUD,  wms_replenish.SYS_CD,  wms_replenish.SYS_CB,  wms_storer.STO_BATCH_FIELD_REF, " &
                        "  wms_replenish.RO_SHIP_MODE,  wms_replenish.RO_SEAL_NO,  wms_replenish.RO_CONTAINER_NO,  wms_replenish.RO_CUST_INV_NO,wms_replenish.RO_CTRY_ORIGIN,wms_replenish.RO_EDI_PO_NO, wms_replenish.RO_WH_CODE " &
                        " FROM wms_replenish " &
                        " left outer join wms_storer " &
                        " on wms_storer.storer_code = wms_replenish.storer_code " &
                        "AND wms_storer.imp_code = wms_replenish.imp_code " &
                        "where wms_replenish.ro_code = '" & gU.dbEncode(pk_code) & "' " &
                        "and wms_replenish.imp_code = '" & imp_code & "' " &
                        "and wms_replenish.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)
            If dt.Rows.Count > 0 Then
                'imp_code = dt.Rows(0).Item("IMP_CODE").ToString
                ViewState("STO_BATCH_FIELD_REF") = dt.Rows(0).Item("STO_BATCH_FIELD_REF").ToString


                RO_CODE.Text = dt.Rows(0).Item("RO_CODE").ToString
                RO_CODE_HF.Value = dt.Rows(0).Item("RO_CODE").ToString
                RO_STATUS.Text = dt.Rows(0).Item("RO_STATUS").ToString
                'STORER_CODE.Text = dt.Rows(0).Item("STORER_CODE").ToString
                'RO_DATE.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("RO_DATE").ToString)
                RO_DATE.Text = dt.Rows(0).Item("RO_DATE").ToString
                RO_RCV_BY.Text = dt.Rows(0).Item("RO_RCV_BY").ToString
                PRJ_CODE.SelectedValue = dt.Rows(0).Item("PRJ_CODE").ToString
                PO_TYPE.SelectedValue = gU.decodeNullOrEmpty(dt.Rows(0).Item("RO_TYPE").ToString, "PO")
                PO_CAT.SelectedValue = dt.Rows(0).Item("PO_CAT").ToString
                RO_REF_NO.Text = dt.Rows(0).Item("RO_REF_NO").ToString
                RO_BATCH_NO.Text = dt.Rows(0).Item("RO_BATCH_NO").ToString
                'RO_ETD.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("RO_ETD").ToString)
                'RO_ETA.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("RO_ETA").ToString)
                RO_ETD.Text = dt.Rows(0).Item("RO_ETD").ToString
                RO_ETA.Text = dt.Rows(0).Item("RO_ETA").ToString
                RO_ISSUED_BY.Text = dt.Rows(0).Item("RO_ISSUED_BY").ToString
                RO_STORER_TEL.Text = dt.Rows(0).Item("RO_STORER_TEL").ToString
                RO_STORER_FAX.Text = dt.Rows(0).Item("RO_STORER_FAX").ToString
                RO_STORER_EMAIL.Text = dt.Rows(0).Item("RO_STORER_EMAIL").ToString
                RO_STORER_CONT.Text = dt.Rows(0).Item("RO_STORER_CONT").ToString
                RO_DEST.Text = dt.Rows(0).Item("RO_DEST").ToString
                RO_REM.Text = dt.Rows(0).Item("RO_REM").ToString
                RO_TRACK_NO.Value = dt.Rows(0).Item("ro_track_no").ToString

                RO_CONTAINER_NO.Text = dt.Rows(0).Item("RO_CONTAINER_NO").ToString
                RO_SEAL_NO.Text = dt.Rows(0).Item("RO_SEAL_NO").ToString
                RO_CUST_INV_NO.Text = dt.Rows(0).Item("RO_CUST_INV_NO").ToString
                RO_SHIP_MODE.SelectedValue = dt.Rows(0).Item("RO_SHIP_MODE").ToString

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                RO_CTRY_ORIGIN.SelectedValue = dt.Rows(0).Item("RO_CTRY_ORIGIN").ToString
                RO_EDI_PO_NO.Text = dt.Rows(0).Item("RO_EDI_PO_NO").ToString
                RO_WH_CODE.SelectedValue = dt.Rows(0).Item("RO_WH_CODE").ToString
                'PRJ_NAME.Text = DB.getValueFromSQL("select PRJ_NAME from WMS_PROJECT where PRJ_CODE = '" & gU.dbEncode(dt.Rows(0).Item("PRJ_CODE").ToString) & "' ")

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME AS STO_NAME from WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
                End If
                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail

        SQLString = "SELECT wms_replenish_d.IMP_CODE,  wms_replenish_d.STORER_CODE, " &
                    " wms_replenish_d.RO_CODE,  wms_replenish_d.ROD_SEQ, " &
                    " wms_replenish_d.ROD_DISP_SEQ,  wms_replenish_d.ROD_CUST_CODE, " &
                    " wms_replenish_d.ROD_PALLET_NO,  wms_replenish_d.ROD_CARTON_NO," &
                    " wms_replenish_d.ROD_BATCH_NO,  wms_replenish_d.ROD_REF_NO," &
                    " wms_replenish_d.ROD_ITM_PARENT,  wms_replenish_d.ROD_ITM_CODE," &
                    " wms_replenish_d.ROD_PACK_KEY,  wms_replenish_d.ROD_ITM_NAME, wms_replenish_d.rod_post_qty, " &
                    " wms_replenish_d.ROD_SERIES_NO,wms_replenish_d.ROD_DOC_NO,  wms_replenish_d.ROD_QTY," &
                    " wms_replenish_d.ROD_UOM,  wms_replenish_d.ROD_PCS_PER_UOM, " &
                    " wms_replenish_d.ROD_TOT_PCS,  wms_replenish_d.ROD_LENGTH," &
                    " wms_replenish_d.ROD_WIDTH,  wms_replenish_d.ROD_HEIGHT," &
                    " wms_replenish_d.ROD_KG,  wms_replenish_d.ROD_CBM," &
                    " wms_replenish_d.ROD_VND_CODE, CONVERT(varchar, ROD_MANU_DATE," & DDFORMAT & ") as ROD_MANU_DATE, " &
                    " CONVERT(varchar, ROD_EXPIRY_DATE, " & DDFORMAT & ") as ROD_EXPIRY_DATE , " &
                    " wms_item.itm_sku_no, wms_item.itm_desc, wms_replenish_d.ROD_STATUS,wms_replenish_d.ROD_UOM2,wms_replenish_d.ROD_QTY2,wms_replenish_d.ROD_CURR," &
                    " wms_replenish_d.ROD_CURR_RATE,wms_replenish_d.ROD_UNIT_PRICE,wms_replenish_d.ROD_ON_BEHALF,wms_replenish_d.ROD_WH_CODE,wms_replenish_d.ROD_LOCATION_CODE,wms_replenish_d.ROD_HKD_EQU, " &
                    " wms_replenish_d.ROD_INSP_REQ, " &
                    " 'U' as mFlag, '0' as cb_copy, 0 as PRICE_SUB_TOTAL " &
                    "from wms_replenish_d left outer join wms_item " &
                    "ON wms_replenish_d.imp_code = wms_item.imp_code " &
                    "and wms_replenish_d.storer_code = wms_item.storer_code " &
                    "and wms_replenish_d.rod_itm_code = wms_item.itm_code " &
                    "and wms_replenish_d.rod_pack_key = wms_item.pack_key " &
                    "where wms_replenish_d.ro_code = '" & gU.dbEncode(pk_code) & "' " &
                    "and wms_replenish_d.storer_code = '" & gU.dbEncode(storerCode) & "' " &
                    "and wms_replenish_d.imp_code = '" & imp_code & "' "

        SQLString = SQLString & " order by wms_replenish_d.rod_disp_seq"

        REM **********************
        dt = gDB.getDataTable(SQLString)
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If
        ViewState("dt") = dt
        GridView1.DataBind()

        If PO_TYPE.SelectedValue = "PO" Then
            RO_REF_NO.CssClass = ""
            selectItemBtn.Visible = True
            'newrow.Visible = False
        Else
            'RO_REF_NO.CssClass = "REQUIRED"
            selectItemBtn.Visible = True
            'newrow.Visible = True
        End If
        REM **********************
    End Sub

    Protected Sub reOpenBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles reOpenBtn.Click
        'Call save("Y")
        Dim SuccessFlag As Boolean = False
        Dim updtSql, SQLString As String

        Dim gConn As SqlConnection

        Dim dt As DataTable

        gConn = gDB.getConnection()

        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try

            updtSql = "update WMS_REPLENISH " &
            "set RO_STATUS = 'NEW', " &
            "sys_lub = '" & Session("usr_id") & "', " &
            "sys_lud = Getdate() " &
            "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " &
            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
            "and RO_CODE = '" & gU.dbEncode(RO_CODE.Text) & "' " &
            "AND NOT EXISTS (SELECT 1 FROM WMS_REPLENISH_D D WHERE ISNULL(ROD_POST_QTY,0) <> 0 " &
            "AND D.IMP_CODE = WMS_REPLENISH.IMP_CODE " &
            "AND D.STORER_CODE = WMS_REPLENISH.STORER_CODE " &
            "AND D.RO_CODE = WMS_REPLENISH.RO_CODE) "

            gDB.amendData(updtSql, gConn, transaction)

            updtSql = "update WMS_REPLENISH " &
                        "set RO_STATUS = 'PARTIAL', " &
                        "sys_lub = '" & Session("usr_id") & "', " &
                        "sys_lud = Getdate() " &
                        "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " &
                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        "and RO_CODE = '" & gU.dbEncode(RO_CODE.Text) & "' " &
                        "AND EXISTS (SELECT 1 FROM WMS_REPLENISH_D D WHERE ISNULL(ROD_POST_QTY,0) <> 0 " &
                        "AND D.IMP_CODE = WMS_REPLENISH.IMP_CODE " &
                        "AND D.STORER_CODE = WMS_REPLENISH.STORER_CODE " &
                        "AND D.RO_CODE = WMS_REPLENISH.RO_CODE) "

            gDB.amendData(updtSql, gConn, transaction)

            transaction.Commit()

            SuccessFlag = True
            'SQLString = "select RO_STATUS FROM WMS_REPLENISH " & _
            '                        "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " & _
            '                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
            '                        "and RO_CODE = '" & gU.dbEncode(RO_CODE.Text) & "' "
            'dt = gDB.getDataTable(SQLString)
            'If dt.Rows.Count > 0 Then
            '    RO_STATUS.Text = dt.Rows(0).Item(0).ToString.Trim
            'End If

            'RO_STATUS.Text = "NEW"

            'CancelBtn.Visible = true
            'reOpenBtn.Visible = False

            'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)

            'uiFun.displayMsg(Me, "", "PO has been re-opened.", Session("gLang"))

        Catch ex As Exception
            transaction.Rollback()
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

            rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
            rmtPost.Add("RO_CODE", RO_CODE.Text.Trim)
            rmtPost.alertMsg = "PO has been re-opened."
            rmtPost.Target = "_self"
            rmtPost.Url = "ROMain.aspx"
            rmtPost.Post()

        End If

    End Sub

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        'Call save("Y")

        Dim cancelSql As String = "update wms_replenish " &
                     "set ro_status = 'CLOSED', " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(imp_code.Trim) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and ro_code = '" & gU.dbEncode(RO_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()

        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()

            RO_STATUS.Text = "CLOSED"

            ar.sec_viewMode = "Y"
            CancelBtn.Visible = False
            ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)

            uiFun.displayMsg(Me, "", "PO has been CLOSED.", Session("gLang"))

        Catch ex As Exception
            transaction.Rollback()
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
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub btnConfirm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConfirm.Click
        RO_STATUS.Text = "CONFIRMED"
        RO_STATUS.ForeColor = Drawing.Color.Blue
        btnConfirm.Enabled = False
        Call save()
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        dt.AcceptChanges()
    End Sub

    Protected Sub STORER_CODE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles STORER_CODE.SelectedIndexChanged
        Call setDefStorerInfo()
    End Sub

    Protected Sub setDefStorerInfo()
        ViewState("STORER_CODE") = STORER_CODE.SelectedValue

        Dim SQLString As String = ""
        Dim ldt As New DataTable

        SQLString = "SELECT * from WMS_STORER " &
                    "where STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                    "and IMP_CODE = '" & gU.dbEncode(imp_code) & "'"

        REM **********************
        ldt = gDB.getDataTable(SQLString)
        If ldt.Rows.Count > 0 Then
            RO_STORER_CONT.Text = ldt.Rows(0).Item("STO_CONT_PER_ORD").ToString()
            RO_STORER_TEL.Text = ldt.Rows(0).Item("STO_CONT_TEL_ORD").ToString()
            RO_STORER_FAX.Text = ldt.Rows(0).Item("STO_FAX").ToString()
            RO_STORER_EMAIL.Text = ldt.Rows(0).Item("STO_CONT_EMAIL_ORD").ToString()

            ViewState("STO_BATCH_FIELD_REF") = ldt.Rows(0).Item("STO_BATCH_FIELD_REF").ToString.Trim
        End If
        ldt = Nothing
    End Sub

    Protected Sub addItemtoRO()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(convert(int,rod_seq)) + 1 from WMS_REPLENISH_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and RO_CODE = '" & gU.dbEncode(RO_CODE.Text.Trim) & "' "
            REM **********************
            Dim nS_dt As New DataTable
            nS_dt = gDB.getDataTable(seq_string)
            Dim next_seq_no As String
            Dim temp_seq_no As Integer = 1

            next_seq_no = ""

            If nS_dt.Rows.Count > 0 Then
                next_seq_no = nS_dt.Rows(0).Item(0).ToString()
            End If

            If next_seq_no = "" Then next_seq_no = "1"

            Dim SQLString As String
            Dim RO_dt As DataTable

            Dim itemPackList As String = ""
            Dim itemListarray As String()
            Dim packKeyListarray As String()
            Dim qtyListarray As String()

            itemListarray = Split(itemList.Value, ", ")
            packKeyListarray = Split(packKeyList.Value, ", ")
            qtyListarray = Split(qtyList.Value, ", ")

            Dim qtyitemDict As New Dictionary(Of String, String)

            If qtyList.Value <> "" Then
                For i As Integer = 0 To itemListarray.Count - 1
                    qtyitemDict.Add(itemListarray(i) & "_000_" & packKeyListarray(i), qtyListarray(i))
                Next
            End If

            If itemListarray.Count = 0 Then
                If itemList.Value <> "" Then
                    itemPackList = Server.HtmlDecode(itemList.Value) & "_000_" & Server.HtmlDecode(packKeyList.Value)
                End If
            Else
                For i = 0 To itemListarray.Count - 1
                    itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemListarray(i)) & "_000_" & Server.HtmlDecode(packKeyListarray(i)))
                Next
            End If

            SQLString = ""
            SQLString += "select wms_item.*, wms_alt_vend_item.* "
            SQLString += "from wms_item left outer join wms_alt_vend_item on "
            SQLString += "wms_item.imp_code = wms_alt_vend_item.imp_code "
            SQLString += "and wms_item.storer_code = wms_alt_vend_item.storer_code "
            SQLString += "and wms_item.itm_code = wms_alt_vend_item.itm_code "
            SQLString += "and wms_item.pack_key = wms_alt_vend_item.pack_key "
            SQLString += "where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
            SQLString += "and wms_item.imp_code = '" & Session("IMP_CODE") & "' "
            SQLString += "and wms_item.itm_code + '_000_' + wms_item.pack_key + '_000_' + ISNULL(wms_alt_vend_item.vnd_code,'000') in ('" & Replace(itemPackList, ", ", "', '") & "') "

            SQLString = SQLString & " order by wms_item.itm_code, wms_item.pack_key, wms_alt_vend_item.vnd_code"
            REM **********************
            RO_dt = gDB.getDataTable(SQLString)

            For i As Integer = 0 To RO_dt.Rows.Count - 1
                If ViewState("n_cur_seq") = "" Then
                    ViewState("n_cur_seq") = next_seq_no
                Else
                    temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                    ViewState("n_cur_seq") = temp_seq_no.ToString
                End If

                dt.Rows.Add()

                rows_count = dt.Rows.Count

                Dim itmQty As Integer = 0

                If qtyitemDict.ContainsKey(RO_dt.Rows(i).Item("ITM_CODE").ToString & "_000_" & RO_dt.Rows(i).Item("PACK_KEY").ToString & "_000_" & gU.decodeNullOrEmpty(RO_dt.Rows(i).Item("vnd_code").ToString, "000")) Then
                    itmQty = gU.decodeEmptyCInt(qtyitemDict(RO_dt.Rows(i).Item("ITM_CODE").ToString & "_000_" & RO_dt.Rows(i).Item("PACK_KEY").ToString & "_000_" & gU.decodeNullOrEmpty(RO_dt.Rows(i).Item("vnd_code").ToString, "000")), 0)
                End If
                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("ROD_SEQ") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("ROD_DISP_SEQ") = rows_count
                'dt.Rows(rows_count - 1).Item("ROD_PALLET_NO") = ""
                'dt.Rows(rows_count - 1).Item("ROD_CARTON_NO") = ""
                'dt.Rows(rows_count - 1).Item("ROD_BATCH_NO") = ""
                'dt.Rows(rows_count - 1).Item("ROD_REF_NO") = ""
                dt.Rows(rows_count - 1).Item("ROD_ITM_PARENT") = RO_dt.Rows(i).Item("ITM_PARENT")
                dt.Rows(rows_count - 1).Item("ROD_ITM_CODE") = RO_dt.Rows(i).Item("ITM_CODE")
                dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = RO_dt.Rows(i).Item("ITM_SKU_NO")

                dt.Rows(rows_count - 1).Item("ITM_DESC") = RO_dt.Rows(i).Item("ITM_DESC")

                dt.Rows(rows_count - 1).Item("ROD_PACK_KEY") = RO_dt.Rows(i).Item("PACK_KEY")
                dt.Rows(rows_count - 1).Item("ROD_ITM_NAME") = RO_dt.Rows(i).Item("ITM_NAME")
                'dt.Rows(rows_count - 1).Item("ROD_QTY") = RO_dt.Rows(i).Item("ITM_BALANCE")
                'dt.Rows(rows_count - 1).Item("ROD_UOM") = RO_dt.Rows(i).Item("AITM_UOM")                
                dt.Rows(rows_count - 1).Item("ROD_QTY") = itmQty
                'dt.Rows(rows_count - 1).Item("ROD_UOM") = "PACK"

                dt.Rows(rows_count - 1).Item("ROD_UOM") = RO_dt.Rows(i).Item("ITM_UOM")

                dt.Rows(rows_count - 1).Item("ROD_UOM2") = RO_dt.Rows(i).Item("ITM_UOM2")
                dt.Rows(rows_count - 1).Item("ROD_QTY2") = RO_dt.Rows(i).Item("ITM_QTY2")

                'dt.Rows(rows_count - 1).Item("ROD_PCS_PER_UOM") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AITM_PCS_PER_PACK").ToString, CInt(RO_dt.Rows(i).Item("PACK_KEY").ToString))
                dt.Rows(rows_count - 1).Item("ROD_PCS_PER_UOM") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ITM_PCS_PER_UOM").ToString, CInt(RO_dt.Rows(i).Item("PACK_KEY").ToString))

                'dt.Rows(rows_count - 1).Item("ROD_TOT_PCS") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ITM_BALANCE").ToString, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AITM_PCS_PER_PACK").ToString, CInt(RO_dt.Rows(i).Item("PACK_KEY").ToString))
                dt.Rows(rows_count - 1).Item("ROD_TOT_PCS") = itmQty * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ITM_PCS_PER_UOM").ToString, CInt(RO_dt.Rows(i).Item("PACK_KEY").ToString))

                'dt.Rows(rows_count - 1).Item("ROD_LENGTH") = RO_dt.Rows(i).Item("AITM_LENGTH")
                'dt.Rows(rows_count - 1).Item("ROD_WIDTH") = RO_dt.Rows(i).Item("AITM_WIDTH")
                'dt.Rows(rows_count - 1).Item("ROD_HEIGHT") = RO_dt.Rows(i).Item("AITM_HIGHT")
                'dt.Rows(rows_count - 1).Item("ROD_KG") = RO_dt.Rows(i).Item("AITM_VOL")
                'dt.Rows(rows_count - 1).Item("ROD_CBM") = RO_dt.Rows(i).Item("AITM_CBM")
                'dt.Rows(rows_count - 1).Item("ROD_CBM") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AITM_LENGTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AITM_WIDTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AITM_HIGHT").ToString.Trim, 0) / 1000000
                'dt.Rows(rows_count - 1).Item("ROD_SERIES_NO") = RO_dt.Rows(i).Item("ITM_SERIES_NO")
                dt.Rows(rows_count - 1).Item("ROD_VND_CODE") = RO_dt.Rows(i).Item("vnd_code")

                dt.Rows(rows_count - 1).Item("ROD_WH_CODE") = RO_WH_CODE.SelectedValue

                dt.Rows(rows_count - 1).Item("ROD_INSP_REQ") = RO_dt.Rows(i).Item("ITM_INSP_YN")

                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

            Next
            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub

    Protected Sub changeLabel()
        REM **********************
        REM Use for re-create the label to change the Langauge
        REM Modify Here
        Call cU.newChangeGVLabel(GridView1, "", "")
        Call cU.newChangeGVLabel(GridView1, "No.", "編號")
        Call cU.newChangeGVLabel(GridView1, "Item Code", "貨物編號")
        Call cU.newChangeGVLabel(GridView1, "Item Name", "物件名稱")
        Call cU.newChangeGVLabel(GridView1, "Pack Key", "封裝內碼")
        Call cU.newChangeGVLabel(GridView1, "Distribution ID", "Distribution ID")
        Call cU.newChangeGVLabel(GridView1, "Vendor Code", "供應商號碼")
        Call cU.newChangeGVLabel(GridView1, "Transaction ID", "Transaction ID")
        Call cU.newChangeGVLabel(GridView1, "PO Header ID", "PO Header ID")
        Call cU.newChangeGVLabel(GridView1, "Carton No.", "外箱編號")
        Call cU.newChangeGVLabel(GridView1, "Lot", "批號")
        Call cU.newChangeGVLabel(GridView1, "Reference", "參考編號")
        Call cU.newChangeGVLabel(GridView1, "Lot No.", "批號")
        Call cU.newChangeGVLabel(GridView1, "Expiry Date", "失效日期")
        Call cU.newChangeGVLabel(GridView1, "Parent", "親")
        'Call cU.newChangeGVLabel(GridView1, "Location", "位置")
        'Call cU.newChangeGVLabel(GridView1, "Series No,", "系列編號")
        Call cU.newChangeGVLabel(GridView1, "Qty", "數量")
        Call cU.newChangeGVLabel(GridView1, "OS Qty", "餘數量")
        Call cU.newChangeGVLabel(GridView1, "UOM", "單位")
        Call cU.newChangeGVLabel(GridView1, "Number Per UOM", "單位件數")
        Call cU.newChangeGVLabel(GridView1, "Total Number", "總件數")
        'Call cU.newChangeGVLabel(GridView1, "Length", "長")
        'Call cU.newChangeGVLabel(GridView1, "Width", "寬")
        'Call cU.newChangeGVLabel(GridView1, "Height", "高")
        'Call cU.newChangeGVLabel(GridView1, "Weight(kg)", "重量(kg)")
        'Call cU.newChangeGVLabel(GridView1, "CBM", "體積")
        Call cU.newChangeGVLabel(GridView1, "Status", "狀態")
        Call cU.newChangeGVLabel(GridView1, "UOM2", "單位2")
        Call cU.newChangeGVLabel(GridView1, "Qty2", "數量2")
        'Call cU.newChangeGVLabel(GridView1, "Currency", "貨幣")
        'Call cU.newChangeGVLabel(GridView1, "Rate", "率")
        'Call cU.newChangeGVLabel(GridView1, "Unit Price", "單價")
        'Call cU.newChangeGVLabel(GridView1, "Sub Total", "共")
        'Call cU.newChangeGVLabel(GridView1, "HKD Equivalent", "港幣等價")
        'Call cU.newChangeGVLabel(GridView1, "On Behalf", "代表")
        Call cU.newChangeGVLabel(GridView1, "Subinventory", "子庫存")
        Call cU.newChangeGVLabel(GridView1, "Location Code", "位置碼")
        Call cU.newChangeGVLabel(GridView1, "Insp. Req", "需要驗貨")
        Call cU.newChangeGVLabel(GridView1, "", "")
        REM **********************
    End Sub

    Protected Sub btnCopyItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopyItem.Click
        If validateAll() Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then

                For x As Integer = 0 To dt.Rows.Count - 1
                    If dt.Rows(x).Item("cb_copy").ToString = "1" Then
                        Dim seq_string As String = "select MAX(convert(int,rod_seq)) + 1 from WMS_REPLENISH_D " &
                                       "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " &
                                       "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                       "and RO_CODE = '" & gU.dbEncode(RO_CODE.Text.Trim) & "' "
                        REM **********************
                        Dim nS_dt As New DataTable
                        nS_dt = gDB.getDataTable(seq_string)
                        Dim next_seq_no As String
                        Dim temp_seq_no As Integer = 1

                        next_seq_no = ""

                        If nS_dt.Rows.Count > 0 Then
                            next_seq_no = nS_dt.Rows(0).Item(0).ToString()
                        End If

                        If next_seq_no = "" Then next_seq_no = "1"

                        If ViewState("n_cur_seq") = "" Then
                            ViewState("n_cur_seq") = next_seq_no
                        Else
                            temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                            ViewState("n_cur_seq") = temp_seq_no.ToString
                        End If

                        Dim newRow As DataRow

                        newRow = dt.NewRow

                        newRow.Item("ROD_SEQ") = ViewState("n_cur_seq").ToString
                        newRow.Item("ROD_DISP_SEQ") = dt.Rows.Count + 1
                        newRow.Item("ROD_ITM_PARENT") = dt.Rows(x).Item("ROD_ITM_PARENT")
                        newRow.Item("ROD_ITM_CODE") = dt.Rows(x).Item("ROD_ITM_CODE")
                        newRow.Item("ROD_PACK_KEY") = dt.Rows(x).Item("ROD_PACK_KEY")
                        newRow.Item("ROD_ITM_NAME") = dt.Rows(x).Item("ROD_ITM_NAME")
                        newRow.Item("ROD_QTY") = dt.Rows(x).Item("ROD_QTY")
                        newRow.Item("ROD_UOM") = dt.Rows(x).Item("ROD_UOM")
                        newRow.Item("ROD_PCS_PER_UOM") = dt.Rows(x).Item("ROD_PCS_PER_UOM")
                        newRow.Item("ROD_TOT_PCS") = dt.Rows(x).Item("ROD_TOT_PCS")
                        'newRow.Item("ROD_LENGTH") = dt.Rows(x).Item("ROD_LENGTH")
                        'newRow.Item("ROD_WIDTH") = dt.Rows(x).Item("ROD_WIDTH")
                        'newRow.Item("ROD_HEIGHT") = dt.Rows(x).Item("ROD_HEIGHT")
                        'newRow.Item("ROD_KG") = dt.Rows(x).Item("ROD_KG")
                        'newRow.Item("ROD_CBM") = dt.Rows(x).Item("ROD_CBM")
                        'newRow.Item("ROD_LOCATION") = dt.Rows(x).Item("ROD_LOCATION")
                        'newRow.Item("ROD_SERIES_NO") = dt.Rows(x).Item("ROD_SERIES_NO")
                        newRow.Item("ROD_PALLET_NO") = dt.Rows(x).Item("ROD_PALLET_NO")
                        newRow.Item("ROD_CARTON_NO") = dt.Rows(x).Item("ROD_CARTON_NO")
                        newRow.Item("ROD_BATCH_NO") = dt.Rows(x).Item("ROD_BATCH_NO")
                        newRow.Item("ROD_REF_NO") = dt.Rows(x).Item("ROD_REF_NO")
                        newRow.Item("ROD_VND_CODE") = dt.Rows(x).Item("ROD_VND_CODE")
                        newRow.Item("ROD_EXPIRY_DATE") = dt.Rows(x).Item("ROD_EXPIRY_DATE")
                        newRow.Item("ROD_MANU_DATE") = dt.Rows(x).Item("ROD_MANU_DATE")

                        newRow.Item("itm_sku_no") = dt.Rows(x).Item("itm_sku_no")

                        newRow.Item("itm_desc") = dt.Rows(x).Item("itm_desc")

                        newRow.Item("ROD_STATUS") = "NEW"
                        newRow.Item("ROD_UOM2") = dt.Rows(x).Item("ROD_UOM2")
                        newRow.Item("ROD_QTY2") = dt.Rows(x).Item("ROD_QTY2")
                        'newRow.Item("ROD_CURR") = dt.Rows(x).Item("ROD_CURR")
                        'newRow.Item("ROD_CURR_RATE") = dt.Rows(x).Item("ROD_CURR_RATE")
                        'newRow.Item("ROD_UNIT_PRICE") = dt.Rows(x).Item("ROD_UNIT_PRICE")
                        'newRow.Item("ROD_ON_BEHALF") = dt.Rows(x).Item("ROD_ON_BEHALF")
                        newRow.Item("ROD_WH_CODE") = dt.Rows(x).Item("ROD_WH_CODE")
                        newRow.Item("ROD_LOCATION_CODE") = dt.Rows(x).Item("ROD_LOCATION_CODE")
                        'newRow.Item("ROD_HKD_EQU") = dt.Rows(x).Item("ROD_HKD_EQU")


                        If PO_TYPE.SelectedValue = "PO" Then
                            newRow.Item("mFlag") = "CN"
                        Else
                            newRow.Item("mFlag") = "NSCN"
                        End If

                        dt.Rows.InsertAt(newRow, dt.Rows.Count)

                        dt.Rows(x).Item("cb_copy") = "0"

                    End If
                Next

                dt.AcceptChanges()

                uiFun.reOrderDetails(dt, "rod_disp_seq")

                GridView1.DataSource = dt
                GridView1.DataBind()
            End If
        End If
    End Sub

    Private Sub SumPCS(ByVal PCS As Double)
        SubTotalPCS += PCS
    End Sub

    Private Function GetTotal() As Double
        Return SubTotalPCS
    End Function

    Protected Sub SendCustROMail(ByRef r_code As String, ByRef st_code As String)
        Dim app_email As String = ""
        Dim mail_title As String = "A New Replenishment Order has been created by customer."
        Dim mail_body As String = ""

        Dim selectSQL As String = "Select distinct lower(usr_email) as mailadd from wms_user, wms_user_group_alloc where " &
                                  " wms_user.usr_id = wms_user_group_alloc.usr_id " &
                                  " and wms_user_group_alloc.grp_code = 'CO_MAIL_GRP' "

        Dim mailDT As New DataTable


        Dim getValSql = "select sto_name as VALUE from wms_storer where storer_code = '" & gU.dbEncode(st_code) & "' "
        Dim lSTO_NAME = DB.getValueFromSQL(getValSql)


        mailDT = gDB.getDataTable(selectSQL)

        If mailDT.Rows.Count > 0 Then
            For i = 0 To mailDT.Rows.Count - 1
                app_email &= mailDT.Rows(i).Item("mailadd").ToString.Trim & ","
            Next
            app_email = Left(app_email, Len(app_email) - 1)
        Else
            app_email = gU.getConfig("AdminEmail")
        End If

        mail_body &= "The following Replenishment order has been created:" & vbNewLine
        mail_body &= "  RO Code: " & r_code & vbNewLine
        'mail_body &= "  Storer Name: " & st_code & vbNewLine
        mail_body &= "  Storer Name: " & lSTO_NAME & vbNewLine
        mail_body &= "  Created Date: " & Now.Date.ToString("dd/MM/yyyy HH:mm:ss")

        Dim mailog As PrgmLog

        If Not System.IO.Directory.Exists(HttpRuntime.Cache("SYSP_LOG_DIR")) Then
            System.IO.Directory.CreateDirectory(HttpRuntime.Cache("SYSP_LOG_DIR"))
        End If

        mailog = New PrgmLog(HttpRuntime.Cache("SYSP_LOG_DIR"), "emailLog" & Now.Date.ToString("ddMMyyyy") & ".txt")

        Call gU.sendEmail(app_email, mail_title, mail_body, mailog)


    End Sub

    Private Function returnBatchNO(ByVal batchPattern As String, ByVal batch_no As String, ByVal manu_date As String, exp_date As String) As String
        Dim returnBatch As String = ""

        Select Case batchPattern
            Case "MANU_DATE"
                If batch_no = "" OrElse Left(batch_no, 4) = "@B#_" Then
                    If manu_date <> "" Then
                        returnBatch = "@B#_M_" & manu_date.Replace("/", "")
                    Else
                        returnBatch = ""
                    End If
                Else
                    returnBatch = batch_no
                End If


            Case "EXP_DATE"
                If batch_no = "" OrElse Left(batch_no, 4) = "@B#_" Then
                    If exp_date <> "" Then
                        returnBatch = "@B#_E_" & exp_date.Replace("/", "")
                    Else
                        returnBatch = ""
                    End If
                Else
                    returnBatch = batch_no
                End If


            Case "EXP_MANU"
                If batch_no = "" OrElse Left(batch_no, 4) = "@B#_" Then
                    If exp_date <> "" Then
                        returnBatch = "@B#_E_" & exp_date.Replace("/", "")
                    ElseIf manu_date <> "" Then
                        returnBatch = "@B#_M_" & manu_date.Replace("/", "")
                    Else
                        returnBatch = ""
                    End If
                Else
                    returnBatch = batch_no
                End If


            Case Else
                returnBatch = batch_no
        End Select

        Return returnBatch

    End Function

    Protected Sub RO_TYPE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles PO_TYPE.SelectedIndexChanged
        Dim tempdt As DataTable

        tempdt = ViewState("dt")
        If Not tempdt Is Nothing Then tempdt.Clear()
        ViewState("n_cur_seq") = ""

        ViewState("dt") = tempdt

        GridView1.DataSource = tempdt
        GridView1.DataBind()


        If PO_TYPE.SelectedValue = "PO" Then
            'newrow.Visible = False
            selectItemBtn.Visible = True
            RO_REF_NO.CssClass = ""
        Else
            'newrow.Visible = True
            selectItemBtn.Visible = True
            'RO_REF_NO.CssClass = "REQUIRED"
        End If
    End Sub

    Protected Sub BtnCopy_Click(sender As Object, e As System.EventArgs) Handles BtnCopy.Click
        If ViewState("RO_CODE") <> "" Then
            Dim successFlag As Boolean = False
            Dim nextRO_code As String = ""
            Dim nextTrack_no As String = ""
            Dim selectInsertSQL As String = ""

            Dim gConn As SqlConnection

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction

            transaction = gConn.BeginTransaction()

            Try
                nextRO_code = DB.getDocNo("RO", gConn, transaction)
                nextTrack_no = DB.getDocNo("TRACKNO", gConn, transaction)

                selectInsertSQL = " INSERT INTO WMS_REPLENISH (IMP_CODE, STORER_CODE, RO_CODE, PRJ_CODE, RO_STATUS, RO_TYPE, RO_DATE, RO_ETD, RO_ETA, RO_ISSUED_BY, RO_RCV_BY, RO_REF_NO, " & _
                                  " RO_BATCH_NO, RO_STORER_CONT, RO_STORER_TEL, RO_STORER_FAX, RO_STORER_EMAIL, RO_DEST, RO_REM, RO_TRACK_NO, SYS_LUB, SYS_LUD, " & _
                                  " SYS_CD, SYS_CB, RO_SHIP_MODE, RO_SEAL_NO, RO_CONTAINER_NO, RO_CUST_INV_NO, RO_CTRY_ORIGIN, RO_EDI_PO_NO, RO_WH_CODE, RO_REV_NO, " & _
                                  " RO_DATE_REQ) " & _
                                  " (SELECT        IMP_CODE, STORER_CODE, '" & gU.dbEncode(nextRO_code) & "', PRJ_CODE, 'NEW', RO_TYPE, Getdate(), RO_ETD, RO_ETA, RO_ISSUED_BY, RO_RCV_BY, RO_REF_NO, " & _
                                  "   RO_BATCH_NO, RO_STORER_CONT, RO_STORER_TEL, RO_STORER_FAX, RO_STORER_EMAIL, RO_DEST, RO_REM, '" & gU.dbEncode(nextTrack_no) & "', '" & Session("usr_id") & "', Getdate(), " & _
                                  "   Getdate(), '" & Session("usr_id") & "', RO_SHIP_MODE, RO_SEAL_NO, RO_CONTAINER_NO, RO_CUST_INV_NO, RO_CTRY_ORIGIN, RO_EDI_PO_NO, RO_WH_CODE, RO_REV_NO, " & _
                                  "   RO_DATE_REQ " & _
                                  " FROM            WMS_REPLENISH where IMP_CODE='" & gU.dbEncode(Session("imp_code")) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ro_code='" & gU.dbEncode(ViewState("RO_CODE")) & "' ) "

                gDB.amendData(selectInsertSQL, gConn, transaction)


                selectInsertSQL = " INSERT INTO WMS_REPLENISH_D " &
                                  " (IMP_CODE, STORER_CODE, RO_CODE, ROD_SEQ, ROD_DISP_SEQ, ROD_CUST_CODE, ROD_PALLET_NO, ROD_CARTON_NO, ROD_BATCH_NO, ROD_REF_NO, " &
                                  " ROD_ITM_PARENT, ROD_ITM_CODE, ROD_PACK_KEY, ROD_ITM_NAME, ROD_SERIES_NO,ROD_DOC_NO, ROD_QTY, ROD_UOM, ROD_PCS_PER_UOM, ROD_TOT_PCS, " &
                                  " ROD_LENGTH, ROD_WIDTH, ROD_HEIGHT, ROD_KG, ROD_CBM, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, ROD_VND_CODE, ROD_EXPIRY_DATE, " &
                                  " ROD_MANU_DATE, ROD_POST_QTY, ROD_STATUS, ROD_UOM2, ROD_QTY2, ROD_CURR, ROD_CURR_RATE, ROD_UNIT_PRICE, ROD_ON_BEHALF, " &
                                  " ROD_WH_CODE, ROD_RCV_QTY, ROD_REJ_QTY, ROD_REJ_REASON, ROD_PD_AREA, ROD_INSP_REQ, ROD_HKD_EQU)" &
                                  "(SELECT IMP_CODE, STORER_CODE, '" & gU.dbEncode(nextRO_code) & "', ROD_SEQ, ROD_DISP_SEQ, ROD_CUST_CODE, ROD_PALLET_NO, ROD_CARTON_NO, ROD_BATCH_NO, ROD_REF_NO, " &
                                  " ROD_ITM_PARENT, ROD_ITM_CODE, ROD_PACK_KEY, ROD_ITM_NAME, ROD_SERIES_NO,ROD_DOC_NO, ROD_QTY, ROD_UOM, ROD_PCS_PER_UOM, ROD_TOT_PCS, " &
                                  " ROD_LENGTH, ROD_WIDTH, ROD_HEIGHT, ROD_KG, ROD_CBM, '" & Session("usr_id") & "', getdate(), getdate(), '" & Session("usr_id") & "', ROD_VND_CODE, NULL, " &
                                  " NULL, NULL, 'NEW', ROD_UOM2, ROD_QTY2, ROD_CURR, ROD_CURR_RATE, ROD_UNIT_PRICE, ROD_ON_BEHALF, " &
                                  " ROD_WH_CODE, NULL, NULL, '', '', '', ROD_HKD_EQU " &
                                  " FROM WMS_REPLENISH_D " &
                                  " where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ro_code='" & gU.dbEncode(ViewState("RO_CODE")) & "') "

                gDB.amendData(selectInsertSQL, gConn, transaction)

                transaction.Commit()

                successFlag = True

            Catch ex As Exception
                transaction.Rollback()
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

            If successFlag Then
                Dim rmtpost As New RemotePost

                rmtpost.Url = "ROMain.aspx"
                rmtpost.Target = "_self"
                rmtpost.Add("RO_CODE", nextRO_code)
                rmtpost.Add("STORER_CODE", STORER_CODE.SelectedValue)

                rmtpost.Post()

            End If
        End If
    End Sub

    Private Sub setPageCtrlAccess()
        exceptionEditList = New List(Of String)
        exceptionEditList.Add("BtnCopy")
        exceptionEditList.Add("btnDSCP")
    End Sub

    Protected Sub BtnGenBatchNo_Click(sender As Object, e As System.EventArgs) Handles BtnGenBatchNo.Click
        Dim checkSQL As String
        Dim updateSQL As String
        Dim checkDt As DataTable

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()

        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()

        Try

            If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
                For x As Integer = 0 To dt.Rows.Count - 1
                    If dt.Rows(x).Item("ROD_MANU_DATE").ToString <> "" Then
                        checkSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(dt.Rows(x).Item("ROD_MANU_DATE").ToString) & "'"
                        checkDt = gDB.getDataTable(checkSQL, gConn, transaction)

                        If checkDt.Rows.Count <= 0 Then

                            updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " & _
                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" & _
                                        "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                        "'" & gU.dbEncode(dt.Rows(x).Item("ROD_MANU_DATE").ToString) & "'," & _
                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                            gDB.amendData(updateSQL, gConn, transaction)

                        End If
                        dt.Rows(x).Item("rod_batch_no") = dt.Rows(x).Item("ROD_MANU_DATE").ToString
                    End If
                Next

                transaction.Commit()

                GridView1.DataSource = dt
                GridView1.DataBind()

            End If



        Catch ex As Exception
            transaction.Rollback()
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


    End Sub

    Private Function validateCPO() As Boolean
        Dim CmdText = "Select IsNull(SUM(a.ROD_QTY),0)ROD_QTY from WMS_REPLENISH_D a Inner Join WMS_REPLENISH b On a.RO_CODE = b.RO_CODE " &
                              "where a.imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                               "and a.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and b.RO_TYPE != 'CPO'  and a.RO_CODE='" & RO_CODE.Text.ToString().Split("-")(0) & "'"
        Dim dt_PO As DataTable = gDB.getDataTable(CmdText)
        If dt_PO IsNot Nothing AndAlso dt_PO.Rows.Count > 0 Then
            Dim RO_QTY = dt_PO.Rows(0)("ROD_QTY")
            CmdText = "Select IsNull(SUM(ROD_QTY),0)ROD_QTY from WMS_REPLENISH_D a Inner Join WMS_REPLENISH b On a.RO_CODE = b.RO_CODE " &
                              "where b.imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                               "and b.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and b.RO_TYPE = 'CPO' and (SELECT SUBSTRING(b.RO_CODE,0, CHARINDEX('-',b.RO_CODE)))='" & RO_CODE.Text.ToString().Split("-")(0) & "' and b.RO_CODE !='" & RO_CODE.Text.ToString() & "'"
            Dim dt_CPO As DataTable = gDB.getDataTable(CmdText)
            If dt_CPO IsNot Nothing AndAlso dt_CPO.Rows.Count > 0 Then
                Dim CPO_QTY = dt_CPO.Rows(0)("ROD_QTY")
                If PO_TYPE.SelectedValue.ToString = "CPO" Then


                    For Each row As GridViewRow In GridView1.Rows
                        CPO_QTY = CPO_QTY + Convert.ToDouble(CType(row.FindControl("rod_qty"), TextBox).Text.Trim)
                    Next
                End If
                If RO_QTY < CPO_QTY Then
                    Return False
                Else
                    Return True
                End If
            End If
        End If
        Return False
    End Function
    Protected Sub saveCPO()
        Dim gConn As SqlConnection
        gConn = gDB.getConnection()

        Dim pTransaction As SqlTransaction
        pTransaction = gConn.BeginTransaction()
        Dim CmdText As String = ""
        If validateCPO() = True Then


            If PO_TYPE.SelectedValue.ToString <> "CPO" Then
                CmdText = "Select IsNull(Count(ro_code),0)+1 as CPONo from WMS_REPLENISH a " &
                              "where a.imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                               "and a.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and a.RO_TYPE = 'CPO'  and (SELECT SUBSTRING(a.RO_CODE,0, CHARINDEX('-',a.RO_CODE)))='" & RO_CODE.Text.Trim & "'"

                Dim dt As DataTable = gDB.getDataTable(CmdText)
                Dim CPO_CODE As String = RO_CODE.Text.Trim & "-" & "CPO" & dt.Rows(0)("CPONo")
                Dim EDI_PO_NO As String = RO_EDI_PO_NO.Text.Trim & "-" & "CPO" & dt.Rows(0)("CPONo")

                Dim RowCount As Int16 = 1

                Try
                    CmdText = "Insert into WMS_REPLENISH(IMP_CODE" &
      ",STORER_CODE" &
      ",RO_CODE" &
      ",PO_TYPE" &
      ",PO_CAT" &
     " ,PRJ_CODE" &
      ",RO_STATUS" &
      ",RO_TYPE" &
      ",RO_DATE" &
      ",RO_ETD" &
      ",RO_ETA" &
      ",RO_ISSUED_BY" &
      ",RO_RCV_BY" &
      ",RO_REF_NO" &
      ",RO_BATCH_NO" &
      ",RO_STORER_CONT" &
      ",RO_STORER_TEL" &
      ",RO_STORER_FAX" &
      ",RO_STORER_EMAIL" &
      ",RO_DEST" &
      ",RO_REM" &
      ",RO_TRACK_NO" &
      ",SYS_LUB" &
      ",SYS_LUD" &
      ",SYS_CD" &
      ",SYS_CB" &
      ",RO_SHIP_MODE" &
      ",RO_SEAL_NO" &
      ",RO_CONTAINER_NO" &
      ",RO_CUST_INV_NO" &
      ",RO_CTRY_ORIGIN" &
      ",RO_EDI_PO_NO" &
      ",RO_WH_CODE" &
      ",RO_REV_NO" &
      ",RO_DATE_REQ" &
      ",RO_VND_CODE" &
      ",RO_DATA_FR" &
      ",RO_LAST_MTL" &
      ",RO_ONBEHALF_RO_NO" &
      ",SAP_MOVE_TYPE" &
      ",SAP_MAT_DOC_YEAR" &
      ",SAP_MAT_DOC_NO" &
      ",RO_INV_PACKATSITE)"
                    CmdText += "Select IMP_CODE,STORER_CODE,'" & CPO_CODE.Trim & "' ,PO_TYPE      ,PO_CAT      ,PRJ_CODE      ,RO_STATUS      ,'CPO'      ,RO_DATE      ,RO_ETD      ,RO_ETA      ,RO_ISSUED_BY      ,RO_RCV_BY      ,RO_REF_NO      ,RO_BATCH_NO      ,RO_STORER_CONT      ,RO_STORER_TEL      ,RO_STORER_FAX      ,RO_STORER_EMAIL      ,RO_DEST      ,RO_REM      ,RO_TRACK_NO      ,SYS_LUB      ,SYS_LUD      ,SYS_CD      ,SYS_CB      ,RO_SHIP_MODE      ,RO_SEAL_NO      ,RO_CONTAINER_NO      ,RO_CUST_INV_NO      ,RO_CTRY_ORIGIN      ,'" & EDI_PO_NO & "'      ,RO_WH_CODE      ,RO_REV_NO      ,RO_DATE_REQ      ,RO_VND_CODE      ,RO_DATA_FR      ,RO_LAST_MTL      ,RO_ONBEHALF_RO_NO      ,SAP_MOVE_TYPE      ,SAP_MAT_DOC_YEAR      ,SAP_MAT_DOC_NO      ,RO_INV_PACKATSITE WMS_REPLENISH FROM WMS_REPLENISH WHERE RO_CODE='" & RO_CODE.Text.Trim & "' and IMP_CODE='" & imp_code.Trim & "' and STORER_CODE='" & STORER_CODE.SelectedValue.ToString.Trim & "'"


                    Dim cmd = New SqlCommand(CmdText, gConn)
                    cmd.Transaction = pTransaction
                    cmd.ExecuteNonQuery()
                    CmdText = "Select * from WMS_REPLENISH_D " &
                              "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                               "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ro_code = '" & gU.dbEncode(RO_CODE.Text) & "' "

                    dt = gDB.getDataTable(CmdText)

                    Dim SeqNo As Int16 = 0
                    For Each row As DataRow In dt.Rows
                        CmdText = "Insert into WMS_REPLENISH_D( [IMP_CODE] ,[STORER_CODE] ,[RO_CODE] ,[ROD_SEQ] ,[ROD_DISP_SEQ] ,[ROD_CUST_CODE] ,[ROD_PALLET_NO] ,[ROD_CARTON_NO] ,[ROD_BATCH_NO] ,[ROD_REF_NO] ,[ROD_ITM_PARENT] ,[ROD_ITM_CODE] ,[ROD_PACK_KEY] ,[ROD_ITM_NAME] ,[ROD_SERIES_NO],[ROD_QTY] ,[ROD_UOM] ,[ROD_PCS_PER_UOM] ,[ROD_TOT_PCS] ,[ROD_LENGTH] ,[ROD_WIDTH] ,[ROD_HEIGHT] ,[ROD_KG] ,[ROD_CBM] ,[SYS_LUB] ,[SYS_LUD] ,[SYS_CD] ,[SYS_CB] ,[ROD_VND_CODE] ,[ROD_EXPIRY_DATE] ,[ROD_MANU_DATE] ,[ROD_POST_QTY] ,[ROD_STATUS] ,[ROD_UOM2] ,[ROD_QTY2] ,[ROD_CURR] ,[ROD_CURR_RATE] ,[ROD_UNIT_PRICE] ,[ROD_ON_BEHALF] ,[ROD_WH_CODE] ,[ROD_RCV_QTY] ,[ROD_REJ_QTY] ,[ROD_REJ_REASON] ,[ROD_PD_AREA] ,[ROD_INSP_REQ] ,[ROD_HKD_EQU] ,[ROD_DATE_REQ] ,[ROD_SCH_DELI] ,[ROD_REJ_PHOTO1] ,[ROD_DOC_TYPE] ,[ROD_DOC_NO] ,[ROD_DOC_SEQ] ,[ROD_MTL_ORG_QTY] ,[ROD_MTL_ASS_QTY] ,[ROD_MTL_LB_QTY] ,[ROD_MTL_TF_QTY] ,[ROD_MTL_MISS_QTY] ,[ROD_MTL_MISS_REASON] ,[ROD_MTL_COL_QTY] ,[ROD_MTL_TS_QTY] ,[ROD_VAL_TYPE] ,[ROD_INSP_QTY] ,[ROD_BLOCK_QTY] ,[ROD_RESTRICT_QTY] ,[ROD_UNRESTRICT_QTY] ,[ROD_PROJ_SEG] ,[ROD_VEND_SEG] ,[ROD_LOC_WH] ,[SAP_MAT_DOC_YEAR] ,[SAP_MAT_DOC_NO] ,[SAP_MAT_DOC_ITEM] ,[SAP_ITEM_NO],[ROD_LOCATION_CODE] )" &
                                            "Select [IMP_CODE] ,[STORER_CODE] ,'" & CPO_CODE.Trim & "' ,[ROD_SEQ] ,[ROD_DISP_SEQ] ,[ROD_CUST_CODE] ,[ROD_PALLET_NO] ,[ROD_CARTON_NO] ,[ROD_BATCH_NO] ,[ROD_REF_NO] ,[ROD_ITM_PARENT] ,[ROD_ITM_CODE] ,[ROD_PACK_KEY] ,[ROD_ITM_NAME] ,[ROD_SERIES_NO],'0' ,[ROD_UOM] ,[ROD_PCS_PER_UOM] ,[ROD_TOT_PCS] ,[ROD_LENGTH] ,[ROD_WIDTH] ,[ROD_HEIGHT] ,[ROD_KG] ,[ROD_CBM] ,[SYS_LUB] ,[SYS_LUD] ,[SYS_CD] ,[SYS_CB] ,[ROD_VND_CODE] ,[ROD_EXPIRY_DATE] ,[ROD_MANU_DATE] ,[ROD_POST_QTY] ,[ROD_STATUS] ,[ROD_UOM2] ,[ROD_QTY2] ,[ROD_CURR] ,[ROD_CURR_RATE] ,[ROD_UNIT_PRICE] ,[ROD_ON_BEHALF] ,[ROD_WH_CODE] ,[ROD_RCV_QTY] ,[ROD_REJ_QTY] ,[ROD_REJ_REASON] ,[ROD_PD_AREA] ,[ROD_INSP_REQ] ,[ROD_HKD_EQU] ,[ROD_DATE_REQ] ,[ROD_SCH_DELI] ,[ROD_REJ_PHOTO1] ,[ROD_DOC_TYPE] ,[ROD_DOC_NO] ,[ROD_DOC_SEQ] ,[ROD_MTL_ORG_QTY] ,[ROD_MTL_ASS_QTY] ,[ROD_MTL_LB_QTY] ,[ROD_MTL_TF_QTY] ,[ROD_MTL_MISS_QTY] ,[ROD_MTL_MISS_REASON] ,[ROD_MTL_COL_QTY] ,[ROD_MTL_TS_QTY] ,[ROD_VAL_TYPE] ,[ROD_INSP_QTY] ,[ROD_BLOCK_QTY] ,[ROD_RESTRICT_QTY] ,[ROD_UNRESTRICT_QTY] ,[ROD_PROJ_SEG] ,[ROD_VEND_SEG] ,[ROD_LOC_WH] ,[SAP_MAT_DOC_YEAR] ,[SAP_MAT_DOC_NO] ,[SAP_MAT_DOC_ITEM] ,[SAP_ITEM_NO] ,[ROD_LOCATION_CODE] From WMS_REPLENISH_D where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and ro_code = '" & gU.dbEncode(RO_CODE.Text) & "' and  ROD_SEQ='" & gU.dbEncode(row("ROD_SEQ").ToString) & "'"

                        cmd = New SqlCommand(CmdText, gConn)
                        cmd.Transaction = pTransaction
                        cmd.ExecuteNonQuery()
                        SeqNo = SeqNo + 1
                    Next

                    pTransaction.Commit()
                    gConn.Close()


                Catch ex As Exception
                    pTransaction.Rollback()
                    Dim msg = ex.Message
                    uiFun.displayMsg(Me, "1015", "", Session("gLang"))
                Finally
                    gConn.Close()
                End Try
                uiFun.displayMsg(Me, "1014", "", Session("gLang"))
                Page.Response.Redirect("../RO/ROMain.aspx?storer_code=" & gU.dbEncode(STORER_CODE.SelectedValue) & "&ro_code=" & CPO_CODE.Trim)
            Else
                uiFun.displayMsg(Me, "0", "Invalid PO TYPE", Session("gLang"))
            End If
        Else
            uiFun.displayMsg(Me, "", "CPO Qty exceeds PO Qty", Session("gLang"))
        End If

    End Sub

    Protected Sub btnGenerate_Click(sender As Object, e As EventArgs)
        saveCPO()
    End Sub

    Protected Sub rod_wh_code_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim ddlWH As DropDownList = CType(sender, DropDownList)
        Dim Row = CType(ddlWH.Parent.Parent, GridViewRow)
        Dim idx = Row.RowIndex
        If ddlWH.SelectedValue <> Nothing Then
            Dim nDropDown = CType(Row.FindControl("rod_location_code"), DropDownList)
            uiFun.load_dropdown(nDropDown, "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & ddlWH.SelectedValue & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))
            'nDropDown.SelectedValue = DataBinder.Eval(Row.DataItem, "rod_location_code").ToString.Trim
        End If
    End Sub

End Class
