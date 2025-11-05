Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text


Partial Class INBOUND_RO_ROMain_C
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
    Private DDFORMAT As String = "DD/MM/YYYY"

    Private gvCol() As String = {"CB_COPY", "ROD_DISP_SEQ", "ROD_VND_CODE", "ROD_ITM_CODE", "ITM_SKU_NO", "ROD_ITM_NAME", "ITM_DESC", "ROD_PACK_KEY", "ROD_PALLET_NO", "ROD_CARTON_NO",
                                 "ROD_BATCH_NO", "ROD_REF_NO", "ROD_EXPIRY_DATE", "ROD_MANU_DATE", "ROD_ITM_PARENT", "ROD_SERIES_NO", "ROD_QTY",
                                 "ROD_OS_QTY", "ROD_UOM", "ROD_PCS_PER_UOM", "ROD_TOT_PCS", "ROD_LENGTH", "ROD_WIDTH", "ROD_HEIGHT", "ROD_KG", "ROD_CBM"}

    Private dt As New DataTable

    Dim SubTotalPCS As Double = 0.0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DDFORMAT = gU.getConfig("DDFORMAT")

        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("IB_RO_C", Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            ViewState("STO_BATCH_FIELD_REF") = ""

            uiFun.load_dropdown(PRJ_CODE, "select PRJ_CODE, PRJ_NAME from WMS_PROJECT ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(RO_SHIP_MODE, "select colc_code, colc_eng_value from WMS_col_code where colc_tabcol='WMS_DELV_ORDER.DO_SHIP_MODE' ORDER BY colc_display_seq", "colc_code", "colc_eng_value", , Session("gSelectLabel"))

            uiFun.load_dropdown(RO_CTRY_ORIGIN, "select ctry_code, ctry_name from WMS_country ORDER BY ctry_name", "ctry_code", "ctry_name", , Session("gSelectLabel"))

            'If Session("pagemode") = "N" Then
            '    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            'End If

        End If

        If Session("pagemode") = "N" Then
            If STORER_CODE.Value = "" Then
                STORER_CODE.Value = Session("usr_pref_storer")
                ViewState("STORER_CODE") = Session("usr_pref_storer")
                If Session("usr_type") = "T" OrElse Session("usr_type") = "C" Then
                    'STORER_CODE.Enabled = False
                End If
                Call setDefStorerInfo()
            End If
            If RO_DATE.Text = "" Then
                RO_DATE.Text = Now.ToString("dd/MM/yyyy")
            End If
        End If

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Value = "Replenishment Order Maintenance"
            lbl_ImageHd.Text = "Order Items"
            lbl_RO_CODE.Text = "RO Code"
            lbl_RO_STATUS.Text = "Status"
            lbl_STORER_CODE.Text = "Storer"
            lbl_RO_DATE.Text = "Date"
            lbl_RO_RCV_BY.Text = "Received By"
            lbl_PRJ_CODE.Text = "Project Code"
            lbl_RO_TYPE.Text = "Type"
            lbl_RO_REF_NO.Text = "Storer's Ref. No."
            lbl_RO_BATCH_NO.Text = "Batch No."
            lbl_RO_ETD.Text = "ETD"
            lbl_RO_ETA.Text = "ETA"
            lbl_RO_ISSUED_BY.Text = "Issued By"
            lbl_RO_STORER_TEL.Text = "Storer's Tel."
            lbl_RO_STORER_FAX.Text = "Storer's Fax"
            lbl_RO_STORER_EMAIL.Text = "Storer's Email"
            lbl_RO_STORER_CONT.Text = "Storer's Contact"
            lbl_RO_DEST.Text = "Destination"
            lbl_RO_REM.Text = "Remarks"
            lbl_RO_TRACK_NO.Text = "Tracking No.:"
            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"

            newrow.Text = "Add"


            selectItemBtn.Text = "Select Item"
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
            lbl_STORER_CODE.Text = "貨主"
            lbl_RO_DATE.Text = "日期"
            lbl_RO_RCV_BY.Text = "收貨人"
            lbl_PRJ_CODE.Text = "項目"
            lbl_RO_TYPE.Text = "類型"
            lbl_RO_REF_NO.Text = "貨主參考編號"
            lbl_RO_BATCH_NO.Text = "批號"
            lbl_RO_ETD.Text = "預計出發時間"
            lbl_RO_ETA.Text = "預計到達時間"
            lbl_RO_ISSUED_BY.Text = "簽發者"
            lbl_RO_STORER_TEL.Text = "貨主電話"
            lbl_RO_STORER_FAX.Text = "貨主傳真"
            lbl_RO_STORER_EMAIL.Text = "貨主電郵"
            lbl_RO_STORER_CONT.Text = "貨主聯絡人"
            lbl_RO_DEST.Text = "目的地"
            lbl_RO_REM.Text = "備註"
            lbl_RO_TRACK_NO.Text = "追查編號:"
            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "儲存"
            saveBtn2.Text = "儲存"

            selectItemBtn.Text = "選擇物料"

            saveBtn1.OnClientClick = "return confirm(""確定儲存添貨指令?"");"
            saveBtn2.OnClientClick = "return confirm(""確定儲存添貨指令?"");"

            newrow.Text = "新增"

        End If
        REM **********************

        REM **********************
        REM Additional CSS
        RO_DATE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then

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

        If moduleAction = "SELECTIM" Then
            addItemtoRO()
        End If

        If dt.Rows.Count > 0 Then
            ' uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME AS STO_NAME from WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
        End If

        'Call changeLabel()

        Dim colIdx_StartWith As Integer = 1

        ar.addColDef("ROD_DISP_SEQ", "rod_disp_seq", colIdx_StartWith)
        ar.addColDef("ROD_VND_CODE", "ROD_VND_CODE", colIdx_StartWith)
        ar.addColDef("ROD_ITM_CODE", "rod_itm_code", colIdx_StartWith)
        ar.addColDef("ROD_ITM_NAME", "rod_itm_name", colIdx_StartWith)
        ar.addColDef("ITM_DESC", "itm_desc", colIdx_StartWith)
        ar.addColDef("ROD_PACK_KEY", "rod_pack_key", colIdx_StartWith)
        ar.addColDef("ROD_PALLET_NO", "rod_pallet_no", colIdx_StartWith)
        ar.addColDef("ROD_CARTON_NO", "rod_carton_no", colIdx_StartWith)
        ar.addColDef("ROD_BATCH_NO", "rod_batch_no", colIdx_StartWith)
        ar.addColDef("ROD_REF_NO", "rod_ref_no", colIdx_StartWith)
        ar.addColDef("ROD_ITM_PARENT", "rod_itm_parent", colIdx_StartWith)
        ar.addColDef("ROD_SERIES_NO", "rod_series_no", colIdx_StartWith)
        ar.addColDef("ROD_QTY", "rod_qty", colIdx_StartWith)
        ar.addColDef("ROD_OS_QTY", "ROD_OS_QTY", colIdx_StartWith)
        ar.addColDef("ROD_UOM", "rod_uom", colIdx_StartWith)
        ar.addColDef("ROD_PCS_PER_UOM", "rod_pcs_per_uom", colIdx_StartWith)
        ar.addColDef("ROD_TOT_PCS", "rod_tot_pcs", colIdx_StartWith)
        ar.addColDef("ROD_LENGTH", "rod_length", colIdx_StartWith)
        ar.addColDef("ROD_WIDTH", "rod_width", colIdx_StartWith)
        ar.addColDef("ROD_HEIGHT", "rod_height", colIdx_StartWith)
        ar.addColDef("ROD_KG", "rod_kg", colIdx_StartWith)
        ar.addColDef("ROD_CBM", "rod_cbm", colIdx_StartWith)

        'For i As Integer = 0 To GridView1.Rows.Count - 1
        '    Response.Write(i & "=" & CType(GridView1.Rows(i).FindControl("rod_itm_name"), TextBox).Text & "<BR>")
        'Next


        REM Select RO button
        If RO_STATUS.Text = "CLOSED" Then
            selectItemBtn.Attributes.Add("onclick", "ItemLookUp('" & STORER_CODE.Value & "');")
            cSBBtn.Attributes.Add("onclick", "checkSB('" & STORER_CODE.Value & "');")
        Else
            selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & storer_code.ClientID & ".value);")
            cSBBtn.Attributes.Add("onclick", "checkSB(document.myform." & storer_code.ClientID & ".value);")
        End If

        REM ****************************************************************
        If RO_STATUS.Text = "CLOSED" Then
            'ar.sec_write = "N"
            ar.sec_viewMode = "Y"


        ElseIf RO_STATUS.Text = "CONFIRMED" Then
            'ar.sec_write = "N"

        End If

        If Session("usr_pref_storer") = "" Then
            ar.sec_viewMode = "Y"
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
        'ar.hideGVForStorer(GridView1, STORER_CODE.Text, "RO", "WMS_REPLENISH_D")

        ar.setFieldCustomize(Me, "IB_RO_C", STORER_CODE.Value, "WMS_REPLENISH")

        'If GridView1.Rows.Count > 0 Then
        '    ar.setGVCustomize(GridView1, "IB_RO_C", STORER_CODE.Value, "WMS_REPLENISH_D", gvCol)
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

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        'Dim i As Integer

        If e.CommandName = "SplitItem" Then

        End If
    End Sub

    Protected Sub GridView1_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Dim gvColdt As DataTable
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim selectSQL As String = ""

        paP = New GlobalDBFunc.DBCmdPara
        selectSQL = "Select upper(FLDO_FILED_CTRL) as FLDO_FILED_CTRL from WMS_FIELD_OPTION " & _
                    "Where imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(ViewState("STORER_CODE")) & " AND FUN_CODE='IB_RO_C' AND ISNULL(FLDO_FIELD_OPTION, '') <> 'Y' AND FLDO_FIELD_TYPE is null " & _
                    " AND FLDO_TABLE_NAME='WMS_REPLENISH_D' "

        gvColdt = gDB.getDataTable(selectSQL, , , , paP)

        If gvColdt.Rows.Count > 0 Then
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
                                    "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' order by 1", "dc_date_code", "dc_date_code", , "N/A", , False)

                Dim rod_batch_no As String = DataBinder.Eval(e.Row.DataItem, "rod_batch_no").ToString.Trim

                If rod_batch_no.Length > 6 Then
                    If rod_batch_no.Substring(0, 6).Equals("@B#_E_") OrElse rod_batch_no.Substring(0, 6).Equals("@B#_M_") Then
                        CType(e.Row.FindControl("rod_batch_no"), AjaxControlToolkit.ComboBox).SelectedValue = ""
                    Else
                        CType(e.Row.FindControl("rod_batch_no"), AjaxControlToolkit.ComboBox).SelectedValue = DataBinder.Eval(e.Row.DataItem, "rod_batch_no").ToString.Trim
                    End If
                End If

                CType(e.Row.FindControl("hd_rod_batch_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "rod_batch_no").ToString.Trim

                CType(e.Row.FindControl("rod_ref_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_ref_no").ToString.Trim
                CType(e.Row.FindControl("rod_itm_parent"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_itm_parent").ToString.Trim
                CType(e.Row.FindControl("rod_itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_itm_code").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("rod_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_pack_key").ToString.Trim
                CType(e.Row.FindControl("rod_itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_itm_name").ToString.Trim
                CType(e.Row.FindControl("itm_desc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_desc").ToString.Trim
                CType(e.Row.FindControl("rod_os_qty"), Label).Text = gU.decodeEmptyCInt(DataBinder.Eval(e.Row.DataItem, "rod_qty").ToString.Trim, 0) - gU.decodeEmptyCInt(DataBinder.Eval(e.Row.DataItem, "rod_post_qty").ToString.Trim, 0)
                CType(e.Row.FindControl("dsp_rod_vnd_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_vnd_code").ToString.Trim
                CType(e.Row.FindControl("rod_vnd_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "rod_vnd_code").ToString.Trim
                Dim nDropDown As DropDownList = CType(e.Row.FindControl("rod_uom"), DropDownList)
                uiFun.load_dropdown(nDropDown, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "rod_uom").ToString.Trim

                CType(e.Row.FindControl("rod_qty"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "rod_qty").ToString.Trim)
                CType(e.Row.FindControl("rod_pcs_per_uom"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "rod_pcs_per_uom").ToString.Trim)
                CType(e.Row.FindControl("rod_tot_pcs"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "rod_tot_pcs").ToString.Trim)
                CType(e.Row.FindControl("rod_length"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_length").ToString.Trim)
                CType(e.Row.FindControl("rod_width"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_width").ToString.Trim)
                CType(e.Row.FindControl("rod_height"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_height").ToString.Trim)
                CType(e.Row.FindControl("rod_kg"), TextBox).Text = cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_kg").ToString.Trim)

                If IsDBNull(DataBinder.Eval(e.Row.DataItem, "rod_cbm")) Then
                    CType(e.Row.FindControl("rod_cbm"), TextBox).Text = "0"
                Else
                    CType(e.Row.FindControl("rod_cbm"), TextBox).Text = CDbl(DataBinder.Eval(e.Row.DataItem, "rod_cbm")).ToString("###,###,##0.0000")
                End If

                CType(e.Row.FindControl("rod_series_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "rod_series_no").ToString.Trim
                REM **********************

                CType(e.Row.FindControl("ROD_EXPIRY_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_EXPIRY_DATE").ToString.Trim
                CType(e.Row.FindControl("ROD_MANU_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ROD_MANU_DATE").ToString.Trim


                CType(e.Row.FindControl("rod_qty"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("rod_tot_pcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("rod_pcs_per_uom"), TextBox).ClientID & ".value * this.value")
                CType(e.Row.FindControl("rod_pcs_per_uom"), TextBox).Attributes.Add("onchange", "document.forms[0]." & CType(e.Row.FindControl("rod_tot_pcs"), TextBox).ClientID & ".value = document.forms[0]." & CType(e.Row.FindControl("rod_qty"), TextBox).ClientID & ".value * this.value")


                CType(e.Row.FindControl("rod_length"), TextBox).Attributes("onkeypress") = "return maskKey(event)"
                CType(e.Row.FindControl("rod_width"), TextBox).Attributes("onkeypress") = "return maskKey(event)"
                CType(e.Row.FindControl("rod_height"), TextBox).Attributes("onkeypress") = "return maskKey(event)"

                If DataBinder.Eval(e.Row.DataItem, "cb_copy").ToString.Trim = "1" Then
                    CType(e.Row.FindControl("cb_copy"), CheckBox).Checked = True
                End If

                SumPCS(cU.FormatDecimalwString(DataBinder.Eval(e.Row.DataItem, "rod_tot_pcs").ToString.Trim))

                CType(e.Row.FindControl("rod_tot_pcs"), TextBox).Attributes("onkeyup") = "sumPCStotal()"

                CType(e.Row.FindControl("rod_length"), TextBox).Attributes.Add("onkeyup", _
                         "document.forms[0]." & CType(e.Row.FindControl("rod_cbm"), TextBox).ClientID & ".value = " & _
                         "(this.value * " & _
                         "document.forms[0]." & CType(e.Row.FindControl("rod_width"), TextBox).ClientID & ".value * " & _
                         "document.forms[0]." & CType(e.Row.FindControl("rod_height"), TextBox).ClientID & ".value) / 1000000")


                CType(e.Row.FindControl("rod_width"), TextBox).Attributes.Add("onkeyup", _
                         "document.forms[0]." & CType(e.Row.FindControl("rod_cbm"), TextBox).ClientID & ".value = " & _
                         "(document.forms[0]." & CType(e.Row.FindControl("rod_length"), TextBox).ClientID & ".value * " & _
                         "this.value * " & _
                         "document.forms[0]." & CType(e.Row.FindControl("rod_height"), TextBox).ClientID & ".value) / 1000000")


                CType(e.Row.FindControl("rod_height"), TextBox).Attributes.Add("onkeyup", _
                         "document.forms[0]." & CType(e.Row.FindControl("rod_cbm"), TextBox).ClientID & ".value = " & _
                         "(document.forms[0]." & CType(e.Row.FindControl("rod_length"), TextBox).ClientID & ".value * " & _
                         "document.forms[0]." & CType(e.Row.FindControl("rod_width"), TextBox).ClientID & ".value * " & _
                         "this.value) / 1000000")

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

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

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "CN" Then

                    For k As Integer = 0 To e.Row.Cells.Count - 1
                        e.Row.Cells(k).CssClass = "GV_COPY"
                    Next
                End If

                'ar.hideGVForStorer(GridView1, STORER_CODE.Text, "RO", "WMS_REPLENISH_D", e)
            Case DataControlRowType.Footer
                CType(e.Row.FindControl("rod_sub_total_pcs"), Label).Text = GetTotal()
        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select max(to_number(rod_seq)) + 1 from wms_replenish_d " & _
                                        "where imp_code = '" & gU.dbEncode(imp_code.Trim.ToString) & "' " & _
                                        "and storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
                                        "and ro_code = '" & gU.dbEncode(RO_CODE.Text) & "' "
            REM **********************
            Dim nS_dt As New DataTable
            nS_dt = gDB.getDataTable(seq_string)
            Dim next_seq_no, temp_no As String
            Dim temp_seq_no As Integer = 1

            next_seq_no = ""

            If nS_dt.Rows.Count > 0 Then
                next_seq_no = nS_dt.Rows(0).Item(0).ToString()
            Else
                temp_no = "1"
            End If

            If next_seq_no = "" Then next_seq_no = "1"

            If ViewState("n_cur_seq") = "" Then
                ViewState("n_cur_seq") = next_seq_no
            Else
                temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                ViewState("n_cur_seq") = temp_seq_no.ToString
            End If


            dt.Rows.Add()
            rows_count = dt.Rows.Count

            REM **********************
            REM Modify Here
            dt.Rows(rows_count - 1).Item("rod_seq") = ViewState("n_cur_seq").ToString
            REM **********************
            dt.Rows(rows_count - 1).Item("mFlag") = "N"
            dt.AcceptChanges()

            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub

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

        If STORER_CODE.Value = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
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

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1

                If CType(GridView1.Rows(i).FindControl("ROD_EXPIRY_DATE"), TextBox).Text <> "" AndAlso Not gU.isValidDate(CType(GridView1.Rows(i).FindControl("ROD_EXPIRY_DATE"), TextBox).Text.Trim) Then
                    uiFun.displayMsg(Me, "", "Invalid Expiry Date!", Session("gLang"))
                    CType(GridView1.Rows(i).FindControl("ROD_EXPIRY_DATE"), TextBox).Focus()
                    Return False
                End If

                If CType(GridView1.Rows(i).FindControl("ROD_MANU_DATE"), TextBox).Text <> "" AndAlso Not gU.isValidDate(CType(GridView1.Rows(i).FindControl("ROD_MANU_DATE"), TextBox).Text.Trim) Then
                    uiFun.displayMsg(Me, "", "Invalid Manufactory Date!", Session("gLang"))
                    CType(GridView1.Rows(i).FindControl("ROD_MANU_DATE"), TextBox).Focus()
                    Return False
                End If


                If uiFun.gvValidate(Me, dt, "rod_qty", GridView1.Rows(i).Cells(9).Text, _
                                 CType(GridView1.Rows(i).FindControl("rod_qty"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "rod_pcs_per_uom", GridView1.Rows(i).Cells(11).Text, _
                                 CType(GridView1.Rows(i).FindControl("rod_pcs_per_uom"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "rod_tot_pcs", GridView1.Rows(i).Cells(12).Text, _
                                 CType(GridView1.Rows(i).FindControl("rod_tot_pcs"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "rod_length", GridView1.Rows(i).Cells(13).Text, _
                                 CType(GridView1.Rows(i).FindControl("rod_length"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "rod_length", GridView1.Rows(i).Cells(14).Text, _
                                 CType(GridView1.Rows(i).FindControl("rod_width"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "rod_height", GridView1.Rows(i).Cells(15).Text, _
                                 CType(GridView1.Rows(i).FindControl("rod_height"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "rod_kg", GridView1.Rows(i).Cells(16).Text, _
                                 CType(GridView1.Rows(i).FindControl("rod_kg"), TextBox).Text) = False Then Return False

                If uiFun.gvValidate(Me, dt, "rod_cbm", GridView1.Rows(i).Cells(17).Text, _
                                 CType(GridView1.Rows(i).FindControl("rod_cbm"), TextBox).Text) = False Then Return False
            Next
        End If

        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim nextNo As String = ""
        Dim gConn As SqlConnection

        Dim dupSQL As String
        Dim dupTbl As New DataTable

        Dim updateSQL As String = ""
        Dim rod_batch_no As String = ""

        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("RO", gConn, transaction)
                    'nextNo = RO_CODE.Text
                    REM **********************

                    If RO_TRACK_NO.Text = "" Then
                        RO_TRACK_NO.Text = DB.getDocNo("TRACKNO", gConn, transaction)
                    End If

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    dupSQL = "select 1 from wms_replenish " & _
                            "where ro_code = '" & gU.dbEncode(nextNo) & "' " & _
                            "and imp_code='" & gU.dbEncode(imp_code) & "' " & _
                            "and storer_code='" & gU.dbEncode(STORER_CODE.Value) & "' "

                    dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into wms_replenish ( " & _
                                        "ro_code, imp_code, storer_code, " & _
                                        "prj_code, ro_status, ro_type, " & _
                                        "ro_date, ro_etd, ro_eta, " & _
                                        "ro_issued_by, ro_rcv_by, ro_ref_no, " & _
                                        "ro_batch_no, ro_storer_cont, ro_storer_tel, " & _
                                        "ro_storer_fax, ro_storer_email, ro_dest, ro_rem, ro_track_no, " & _
                                        "RO_SEAL_NO,RO_CONTAINER_NO, RO_CUST_INV_NO, RO_SHIP_MODE, RO_CTRY_ORIGIN," & _
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                        "values ( " & _
                                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(imp_code)) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.Value)) & "," & _
                                        gU.convdbNVCData(gU.dbEncode(PRJ_CODE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(RO_STATUS.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_TYPE.Text)) & ", " & _
                                        gU.convdbDate(gU.dbEncode(RO_DATE.Text)) & "," & gU.convdbDate(gU.dbEncode(RO_ETD.Text)) & "," & gU.convdbDate(gU.dbEncode(RO_ETA.Text)) & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(RO_ISSUED_BY.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_RCV_BY.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_REF_NO.Text)) & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(RO_BATCH_NO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_STORER_CONT.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_STORER_TEL.Text)) & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(RO_STORER_FAX.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_STORER_EMAIL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_DEST.Text)) & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(RO_REM.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_TRACK_NO.Text)) & "," & _
                                        gU.convdbNVCData(gU.dbEncode(RO_SEAL_NO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_CONTAINER_NO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_CUST_INV_NO.Text)) & "," & gU.convdbNVCData(gU.dbEncode(RO_SHIP_MODE.SelectedValue)) & "," & _
                                        gU.convdbNVCData(gU.dbEncode(RO_CTRY_ORIGIN.SelectedValue)) & "," & _
                                        "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                        REM **********************

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                            uiFun.reOrderDetails(dt, "rod_disp_seq")

                            For Each rows As DataRow In dt.Rows

                                rod_batch_no = ""
                                rod_batch_no = returnBatchNO(ViewState("STO_BATCH_FIELD_REF"), rows.Item("rod_batch_no").ToString.Trim, gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""), gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))

                                itemSQL = ""
                                REM **********************
                                REM Modify Here
                                Select Case rows.Item("mFlag")
                                    Case "N", "CN"
                                        If rod_batch_no <> "" Then
                                            dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' AND dc_date_code='" & gU.dbEncode(rod_batch_no) & "'"
                                            dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                            If dupTbl.Rows.Count <= 0 Then

                                                updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " & _
                                                            " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" & _
                                                            "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.Value)) & ", " & _
                                                            "'" & gU.dbEncode(rod_batch_no) & "'," & _
                                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                                gDB.amendData(updateSQL, gConn, transaction)
                                            End If
                                        End If

                                        itemSQL = "insert into wms_replenish_d ( " & _
                                                    "ro_code, imp_code, storer_code, rod_seq, rod_disp_seq, " & _
                                                    "rod_pallet_no, rod_carton_no, rod_batch_no, rod_ref_no, " & _
                                                    "rod_itm_parent, rod_itm_code, rod_pack_key, rod_itm_name, " & _
                                                    "rod_qty, rod_uom, rod_pcs_per_uom, rod_tot_pcs, rod_length, " & _
                                                    "rod_width, rod_height, rod_kg, rod_cbm, rod_series_no, rod_vnd_code, ROD_EXPIRY_DATE, ROD_MANU_DATE, sys_cb, sys_cd, " & _
                                                    "sys_lub, sys_lud) " & _
                                                    "values ( " & _
                                                    gU.convdbNVCData(gU.dbEncode(nextNo)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.Value)) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_seq").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_disp_seq").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pallet_no").ToString.Trim, "000"))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_carton_no").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(rod_batch_no)) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_ref_no").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_parent").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_code").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_pack_key").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_name").ToString.Trim, ""))) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_qty").ToString.Trim, "0")) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_uom").ToString.Trim, ""))) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pcs_per_uom").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_tot_pcs").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_length").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_width").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_height").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_kg").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_cbm").ToString.Trim, "0")) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_series_no").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_vnd_code").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""))) & ", " & _
                                                    "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                End Select
                                REM **********************

                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                            Next
                        End If

                        If Session("usr_type") = "C" OrElse Session("usr_type") = "T" Then
                            SendCustROMail(nextNo, STORER_CODE.Value)
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
                    sql_string = "update wms_replenish set " & _
                                    "prj_code = " & gU.convdbNVCData(gU.dbEncode(PRJ_CODE.Text)) & ", " & _
                                    "ro_status = " & gU.convdbNVCData(gU.dbEncode(RO_STATUS.Text)) & ", " & _
                                    "ro_type = " & gU.convdbNVCData(gU.dbEncode(RO_TYPE.Text)) & ", " & _
                                    "ro_date = " & gU.convdbDate(gU.dbEncode(RO_DATE.Text)) & ", " & _
                                    "ro_etd = " & gU.convdbDate(gU.dbEncode(RO_ETD.Text)) & ", " & _
                                    "ro_eta = " & gU.convdbDate(gU.dbEncode(RO_ETA.Text)) & ", " & _
                                    "ro_issued_by = " & gU.convdbNVCData(gU.dbEncode(RO_ISSUED_BY.Text)) & ", " & _
                                    "ro_rcv_by = " & gU.convdbNVCData(gU.dbEncode(RO_RCV_BY.Text)) & ", " & _
                                    "ro_ref_no = " & gU.convdbNVCData(gU.dbEncode(RO_REF_NO.Text)) & ", " & _
                                    "ro_batch_no = " & gU.convdbNVCData(gU.dbEncode(RO_BATCH_NO.Text)) & ", " & _
                                    "ro_storer_cont = " & gU.convdbNVCData(gU.dbEncode(RO_STORER_CONT.Text)) & ", " & _
                                    "ro_storer_tel = " & gU.convdbNVCData(gU.dbEncode(RO_STORER_TEL.Text)) & ", " & _
                                    "ro_storer_fax = " & gU.convdbNVCData(gU.dbEncode(RO_STORER_FAX.Text)) & ", " & _
                                    "ro_storer_email = " & gU.convdbNVCData(gU.dbEncode(RO_STORER_EMAIL.Text)) & ", " & _
                                    "ro_dest = " & gU.convdbNVCData(gU.dbEncode(RO_DEST.Text)) & ", " & _
                                    "ro_rem = " & gU.convdbNVCData(gU.dbEncode(RO_REM.Text)) & ", " & _
                                    "ro_track_no = " & gU.convdbNVCData(gU.dbEncode(RO_TRACK_NO.Text)) & ", " & _
                                    "RO_SEAL_NO = " & gU.convdbNVCData(gU.dbEncode(RO_SEAL_NO.Text)) & ", " & _
                                    "RO_CONTAINER_NO = " & gU.convdbNVCData(gU.dbEncode(RO_CONTAINER_NO.Text)) & ", " & _
                                    "RO_CUST_INV_NO = " & gU.convdbNVCData(gU.dbEncode(RO_CUST_INV_NO.Text)) & ", " & _
                                    "RO_SHIP_MODE = " & gU.convdbNVCData(gU.dbEncode(RO_SHIP_MODE.SelectedValue)) & ", " & _
                                    "RO_CTRY_ORIGIN = " & gU.convdbNVCData(gU.dbEncode(RO_CTRY_ORIGIN.SelectedValue)) & ", " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
                                    "and ro_code = '" & gU.dbEncode(RO_CODE.Text) & "' "

                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "rod_disp_seq")

                        For Each rows As DataRow In dt.Rows
                            rod_batch_no = ""
                            If rows.Item("rod_batch_no").ToString.Trim <> "" Then
                                If rows.Item("rod_batch_no").ToString.Trim.Substring(0, 6).Equals("@B#_E_") OrElse rows.Item("rod_batch_no").ToString.Trim.Substring(0, 6).Equals("@B#_M_") Then
                                    rod_batch_no = returnBatchNO(ViewState("STO_BATCH_FIELD_REF"), "", gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""), gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))
                                Else
                                    rod_batch_no = rows.Item("rod_batch_no").ToString.Trim
                                End If

                            Else
                                rod_batch_no = returnBatchNO(ViewState("STO_BATCH_FIELD_REF"), rows.Item("rod_batch_no").ToString.Trim, gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""), gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))
                            End If


                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N", "CN"
                                    If rod_batch_no <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' AND dc_date_code='" & gU.dbEncode(rod_batch_no) & "'"
                                        dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTbl.Rows.Count <= 0 Then

                                            updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " & _
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values (" & _
                                                        "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.Value)) & ", " & _
                                                        "'" & gU.dbEncode(rod_batch_no) & "'," & _
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(updateSQL, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "insert into wms_replenish_d ( " & _
                                                    "ro_code, imp_code, storer_code, rod_seq, rod_disp_seq, " & _
                                                    "rod_pallet_no, rod_carton_no, rod_batch_no, rod_ref_no, " & _
                                                    "rod_itm_parent, rod_itm_code, rod_pack_key, rod_itm_name, " & _
                                                    "rod_qty, rod_uom, rod_pcs_per_uom, rod_tot_pcs, rod_length, " & _
                                                    "rod_width, rod_height, rod_kg, rod_cbm, rod_series_no, rod_vnd_code, ROD_EXPIRY_DATE, ROD_MANU_DATE, sys_cb, sys_cd, " & _
                                                    "sys_lub, sys_lud) " & _
                                                    "values ( " & _
                                                    gU.convdbNVCData(gU.dbEncode(RO_CODE.Text.Trim)) & ",'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.Value)) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_seq").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_disp_seq").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pallet_no").ToString.Trim, "000"))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_carton_no").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(rod_batch_no)) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_ref_no").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_parent").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_code").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_pack_key").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_name").ToString.Trim, ""))) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_qty").ToString.Trim, "0")) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_uom").ToString.Trim, ""))) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pcs_per_uom").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_tot_pcs").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_length").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_width").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_height").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_kg").ToString.Trim, "0")) & ", " & _
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_cbm").ToString.Trim, "0")) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_series_no").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_vnd_code").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))) & ", " & _
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""))) & ", " & _
                                                    "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                Case "D"
                                    itemSQL = "delete from wms_replenish_d " & _
                                            "where imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
                                            "and ro_code = '" & gU.dbEncode(RO_CODE.Text.Trim) & "' " & _
                                            "and rod_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("rod_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    If rod_batch_no <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' AND dc_date_code='" & gU.dbEncode(rod_batch_no) & "'"
                                        dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTbl.Rows.Count <= 0 Then

                                            updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " & _
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" & _
                                                        "'" & imp_code & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.Value)) & ", " & _
                                                        "'" & gU.dbEncode(rod_batch_no) & "'," & _
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"
                                            gDB.amendData(updateSQL, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "update wms_replenish_d set " & _
                                                "rod_disp_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_disp_seq").ToString.Trim, ""))) & ", " & _
                                                "rod_pallet_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pallet_no").ToString.Trim, "000"))) & ", " & _
                                                "rod_carton_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_carton_no").ToString.Trim, ""))) & ", " & _
                                                "rod_batch_no = " & gU.convdbNVCData(gU.dbEncode(rod_batch_no)) & ", " & _
                                                "rod_ref_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_ref_no").ToString.Trim, ""))) & ", " & _
                                                "rod_itm_parent = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_parent").ToString.Trim, ""))) & ", " & _
                                                "rod_itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_code").ToString.Trim, ""))) & ", " & _
                                                "rod_pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_pack_key").ToString.Trim, ""))) & ", " & _
                                                "rod_itm_name = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_itm_name").ToString.Trim, ""))) & ", " & _
                                                "rod_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_qty").ToString.Trim, "0")) & ", " & _
                                                "rod_uom = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_uom").ToString.Trim, ""))) & ", " & _
                                                "rod_pcs_per_uom = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_pcs_per_uom").ToString.Trim, "0")) & ", " & _
                                                "rod_tot_pcs = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_tot_pcs").ToString.Trim, "0")) & ", " & _
                                                "rod_length = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_length").ToString.Trim, "0")) & ", " & _
                                                "rod_width = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_width").ToString.Trim, "0")) & ", " & _
                                                "rod_height = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_height").ToString.Trim, "0")) & ", " & _
                                                "rod_kg = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_kg").ToString.Trim, "0")) & ", " & _
                                                "rod_cbm = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rod_cbm").ToString.Trim, "0")) & ", " & _
                                                "rod_series_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_series_no").ToString.Trim, ""))) & ", " & _
                                                "rod_vnd_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rod_vnd_code").ToString.Trim, ""))) & ", " & _
                                                "ROD_EXPIRY_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_EXPIRY_DATE").ToString.Trim, ""))) & ", " & _
                                                "ROD_MANU_DATE = " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ROD_MANU_DATE").ToString.Trim, ""))) & ", " & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = Getdate() " & _
                                                "where imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                                                "and storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
                                                "and ro_code = '" & gU.dbEncode(RO_CODE.Text.Trim) & "' " & _
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
                    ViewState("STORER_CODE") = STORER_CODE.Value
                    RO_CODE.ForeColor = Drawing.Color.Black
                    RO_CODE.Font.Size = 10
                    REM **********************
                End If

                If flag <> "Y" Then uiFun.displayMsg(Me, "1007", "", Session("gLang"))

                Call BindGV()

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

            ViewState("STORER_CODE") = STORER_CODE.Value
            REM **********************
        Else

            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = " SELECT wms_replenish.IMP_CODE,  wms_replenish.STORER_CODE,  wms_replenish.RO_CODE, " & _
                        "  wms_replenish.PRJ_CODE,  wms_replenish.RO_STATUS,  wms_replenish.RO_TYPE, " & _
                        "  TO_CHAR(wms_replenish.RO_DATE,'" & DDFORMAT & "') as RO_DATE,  TO_CHAR(wms_replenish.RO_ETD,'" & DDFORMAT & "') as RO_ETD,  TO_CHAR(wms_replenish.RO_ETA,'" & DDFORMAT & "') as RO_ETA , " & _
                        "  wms_replenish.RO_ISSUED_BY,  wms_replenish.RO_RCV_BY,  wms_replenish.RO_REF_NO, " & _
                        "  wms_replenish.RO_BATCH_NO,  wms_replenish.RO_STORER_CONT,  wms_replenish.RO_STORER_TEL, " & _
                        "  wms_replenish.RO_STORER_FAX,  wms_replenish.RO_STORER_EMAIL,  wms_replenish.RO_DEST, " & _
                        "  wms_replenish.RO_REM,  wms_replenish.RO_TRACK_NO,  wms_replenish.SYS_LUB, " & _
                        "  wms_replenish.SYS_LUD,  wms_replenish.SYS_CD,  wms_replenish.SYS_CB, wms_storer.STO_BATCH_FIELD_REF, " & _
                        "  wms_replenish.RO_SHIP_MODE,  wms_replenish.RO_SEAL_NO,  wms_replenish.RO_CONTAINER_NO,  wms_replenish.RO_CUST_INV_NO,wms_replenish.RO_CTRY_ORIGIN " & _
                        " FROM wms_replenish, wms_storer " & _
                        "where wms_replenish.ro_code = '" & gU.dbEncode(pk_code) & "' " & _
                        "and wms_replenish.imp_code = '" & imp_code & "' " & _
                        "and wms_replenish.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                        "and wms_storer.storer_code(+) = wms_replenish.storer_code " & _
                        "AND wms_storer.imp_code(+) = wms_replenish.imp_code "

            dt = gDB.getDataTable(SQLString)
            If dt.Rows.Count > 0 Then
                'imp_code = dt.Rows(0).Item("IMP_CODE").ToString
                ViewState("STO_BATCH_FIELD_REF") = dt.Rows(0).Item("STO_BATCH_FIELD_REF").ToString

                RO_CODE.Text = dt.Rows(0).Item("RO_CODE").ToString
                RO_CODE_HF.Value = dt.Rows(0).Item("RO_CODE").ToString
                RO_STATUS.Text = dt.Rows(0).Item("RO_STATUS").ToString
                STORER_CODE.Value = dt.Rows(0).Item("STORER_CODE").ToString
                'RO_DATE.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("RO_DATE").ToString)
                RO_DATE.Text = dt.Rows(0).Item("RO_DATE").ToString
                RO_RCV_BY.Text = dt.Rows(0).Item("RO_RCV_BY").ToString
                PRJ_CODE.SelectedValue = dt.Rows(0).Item("PRJ_CODE").ToString
                RO_TYPE.Text = dt.Rows(0).Item("RO_TYPE").ToString
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
                RO_TRACK_NO.Text = dt.Rows(0).Item("ro_track_no").ToString

                RO_CONTAINER_NO.Text = dt.Rows(0).Item("RO_CONTAINER_NO").ToString
                RO_SEAL_NO.Text = dt.Rows(0).Item("RO_SEAL_NO").ToString
                RO_CUST_INV_NO.Text = dt.Rows(0).Item("RO_CUST_INV_NO").ToString
                RO_SHIP_MODE.SelectedValue = dt.Rows(0).Item("RO_SHIP_MODE").ToString

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                RO_CTRY_ORIGIN.SelectedValue = dt.Rows(0).Item("RO_CTRY_ORIGIN").ToString

                'PRJ_NAME.Text = DB.getValueFromSQL("select PRJ_NAME from WMS_PROJECT where PRJ_CODE = '" & gU.dbEncode(dt.Rows(0).Item("PRJ_CODE").ToString) & "' ")

                'If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                '    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                'Else
                '    uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME AS STO_NAME from WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
                'End If
                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail

        SQLString = "SELECT wms_replenish_d.IMP_CODE,  wms_replenish_d.STORER_CODE, " & _
                    " wms_replenish_d.RO_CODE,  wms_replenish_d.ROD_SEQ, " & _
                    " wms_replenish_d.ROD_DISP_SEQ,  wms_replenish_d.ROD_CUST_CODE, " & _
                    " wms_replenish_d.ROD_PALLET_NO,  wms_replenish_d.ROD_CARTON_NO," & _
                    " wms_replenish_d.ROD_BATCH_NO, ''as hd_rod_batch_no,  wms_replenish_d.ROD_REF_NO," & _
                    " wms_replenish_d.ROD_ITM_PARENT,  wms_replenish_d.ROD_ITM_CODE," & _
                    " wms_replenish_d.ROD_PACK_KEY,  wms_replenish_d.ROD_ITM_NAME, wms_replenish_d.rod_post_qty, " & _
                    " wms_replenish_d.ROD_SERIES_NO,  wms_replenish_d.ROD_QTY," & _
                    " wms_replenish_d.ROD_UOM,  wms_replenish_d.ROD_PCS_PER_UOM, " & _
                    " wms_replenish_d.ROD_TOT_PCS,  wms_replenish_d.ROD_LENGTH," & _
                    " wms_replenish_d.ROD_WIDTH,  wms_replenish_d.ROD_HEIGHT," & _
                    " wms_replenish_d.ROD_KG,  wms_replenish_d.ROD_CBM," & _
                    " wms_replenish_d.ROD_VND_CODE, to_char(ROD_MANU_DATE,'" & DDFORMAT & "') as ROD_MANU_DATE, " & _
                    " to_char(ROD_EXPIRY_DATE, '" & DDFORMAT & "') as ROD_EXPIRY_DATE , " & _
                    " wms_item.itm_sku_no,  wms_item.itm_desc, " & _
                    " 'U' as mFlag, '0' as cb_copy " & _
                    "from wms_replenish_d, wms_item " & _
                    "where wms_replenish_d.ro_code = '" & gU.dbEncode(pk_code) & "' " & _
                    "and wms_replenish_d.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                    "and wms_replenish_d.imp_code = '" & imp_code & "' " & _
                    "and wms_replenish_d.imp_code = wms_item.imp_code (+) " & _
                    "and wms_replenish_d.storer_code = wms_item.storer_code (+) " & _
                    "and wms_replenish_d.rod_itm_code = wms_item.itm_code (+) " & _
                    "and wms_replenish_d.rod_pack_key = wms_item.pack_key (+) "

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

        REM **********************
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        dt.AcceptChanges()
    End Sub

    Protected Sub setDefStorerInfo()
        Dim SQLString As String = ""
        Dim ldt As New DataTable

        SQLString = "SELECT * from WMS_STORER " & _
                    "where STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
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
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(to_number(rod_seq)) + 1 from WMS_REPLENISH_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
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
            SQLString += "where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' "
            SQLString += "and wms_item.imp_code = '" & Session("IMP_CODE") & "' "
            SQLString += "and wms_item.itm_code || '_000_' || wms_item.pack_key || '_000_' || ISNULL(wms_alt_vend_item.vnd_code,'000') in ('" & Replace(itemPackList, ", ", "', '") & "') "

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

                'dt.Rows(rows_count - 1).Item("ROD_PCS_PER_UOM") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AITM_PCS_PER_PACK").ToString, CInt(RO_dt.Rows(i).Item("PACK_KEY").ToString))
                dt.Rows(rows_count - 1).Item("ROD_PCS_PER_UOM") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ITM_PCS_PER_UOM").ToString, CInt(RO_dt.Rows(i).Item("PACK_KEY").ToString))

                'dt.Rows(rows_count - 1).Item("ROD_TOT_PCS") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ITM_BALANCE").ToString, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AITM_PCS_PER_PACK").ToString, CInt(RO_dt.Rows(i).Item("PACK_KEY").ToString))
                dt.Rows(rows_count - 1).Item("ROD_TOT_PCS") = itmQty * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("ITM_PCS_PER_UOM").ToString, CInt(RO_dt.Rows(i).Item("PACK_KEY").ToString))

                dt.Rows(rows_count - 1).Item("ROD_LENGTH") = RO_dt.Rows(i).Item("AITM_LENGTH")
                dt.Rows(rows_count - 1).Item("ROD_WIDTH") = RO_dt.Rows(i).Item("AITM_WIDTH")
                dt.Rows(rows_count - 1).Item("ROD_HEIGHT") = RO_dt.Rows(i).Item("AITM_HIGHT")
                dt.Rows(rows_count - 1).Item("ROD_KG") = RO_dt.Rows(i).Item("AITM_VOL")
                'dt.Rows(rows_count - 1).Item("ROD_CBM") = RO_dt.Rows(i).Item("AITM_CBM")
                dt.Rows(rows_count - 1).Item("ROD_CBM") = gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AITM_LENGTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AITM_WIDTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(RO_dt.Rows(i).Item("AITM_HIGHT").ToString.Trim, 0) / 1000000
                dt.Rows(rows_count - 1).Item("ROD_SERIES_NO") = RO_dt.Rows(i).Item("ITM_SERIES_NO")
                dt.Rows(rows_count - 1).Item("ROD_VND_CODE") = RO_dt.Rows(i).Item("vnd_code")
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
        Call cU.newChangeGVLabel(GridView1, "No.", "編號")
        Call cU.newChangeGVLabel(GridView1, "Vendor Code", "供應商號碼")
        Call cU.newChangeGVLabel(GridView1, "Item Code", "物件號碼")
        Call cU.newChangeGVLabel(GridView1, "Stock No.", "Stock No.")
        Call cU.newChangeGVLabel(GridView1, "Item Name", "物件名稱")
        Call cU.newChangeGVLabel(GridView1, "Pack Key", "封裝內碼")
        Call cU.newChangeGVLabel(GridView1, "Pallet No.", "托盤編號")
        Call cU.newChangeGVLabel(GridView1, "Carton No.", "外箱編號")
        Call cU.newChangeGVLabel(GridView1, "Batch No.", "批號")
        Call cU.newChangeGVLabel(GridView1, "Reference", "參考編號")
        Call cU.newChangeGVLabel(GridView1, "Expiry Date", "失效日期")
        Call cU.newChangeGVLabel(GridView1, "Manufactory Date", "生產日期")
        Call cU.newChangeGVLabel(GridView1, "Parent", "Parent")
        Call cU.newChangeGVLabel(GridView1, "Series No.", "Series No.")
        Call cU.newChangeGVLabel(GridView1, "Qty", "數量")
        Call cU.newChangeGVLabel(GridView1, "OS Qty", "數量")
        Call cU.newChangeGVLabel(GridView1, "UOM", "單位")
        Call cU.newChangeGVLabel(GridView1, "Number Per UOM", "單位件數")
        Call cU.newChangeGVLabel(GridView1, "Total Number", "總件數")
        Call cU.newChangeGVLabel(GridView1, "Length(cm)", "長(cm)")
        Call cU.newChangeGVLabel(GridView1, "Width(cm)", "寬(cm)")
        Call cU.newChangeGVLabel(GridView1, "Height(cm)", "高(cm)")
        Call cU.newChangeGVLabel(GridView1, "Weight(kg)", "重量(kg)")
        Call cU.newChangeGVLabel(GridView1, "CBM", "體積(cbm)")
        Call cU.newChangeGVLabel(GridView1, "", "")
        REM **********************
    End Sub

    Protected Sub btnCopyItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopyItem.Click
        If validateAll() Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                For x As Integer = 0 To dt.Rows.Count - 1
                    If dt.Rows(x).Item("cb_copy").ToString = "1" Then
                        Dim seq_string As String = "select MAX(to_number(rod_seq)) + 1 from WMS_REPLENISH_D " & _
                                       "where IMP_CODE = '" & gU.dbEncode(imp_code.Trim) & "' " & _
                                       "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
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
                        newRow.Item("ROD_LENGTH") = dt.Rows(x).Item("ROD_LENGTH")
                        newRow.Item("ROD_WIDTH") = dt.Rows(x).Item("ROD_WIDTH")
                        newRow.Item("ROD_HEIGHT") = dt.Rows(x).Item("ROD_HEIGHT")
                        newRow.Item("ROD_KG") = dt.Rows(x).Item("ROD_KG")
                        newRow.Item("ROD_CBM") = dt.Rows(x).Item("ROD_CBM")
                        newRow.Item("ROD_SERIES_NO") = dt.Rows(x).Item("ROD_SERIES_NO")
                        newRow.Item("ROD_PALLET_NO") = dt.Rows(x).Item("ROD_PALLET_NO")
                        newRow.Item("ROD_CARTON_NO") = dt.Rows(x).Item("ROD_CARTON_NO")
                        newRow.Item("ROD_BATCH_NO") = dt.Rows(x).Item("ROD_BATCH_NO")
                        newRow.Item("HD_ROD_BATCH_NO") = dt.Rows(x).Item("HD_ROD_BATCH_NO")
                        newRow.Item("ROD_REF_NO") = dt.Rows(x).Item("ROD_REF_NO")
                        newRow.Item("ROD_VND_CODE") = dt.Rows(x).Item("ROD_VND_CODE")
                        newRow.Item("ROD_EXPIRY_DATE") = dt.Rows(x).Item("ROD_EXPIRY_DATE")
                        newRow.Item("ROD_MANU_DATE") = dt.Rows(x).Item("ROD_MANU_DATE")

                        newRow.Item("itm_sku_no") = dt.Rows(x).Item("itm_sku_no")
                        newRow.Item("itm_desc") = dt.Rows(x).Item("itm_desc")

                        newRow.Item("mFlag") = "CN"

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

        Dim selectSQL As String = "Select distinct lower(usr_email) as mailadd from wms_user, wms_user_group_alloc where " & _
                                  " wms_user.usr_id = wms_user_group_alloc.usr_id " & _
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
                If batch_no = "" Then
                    If manu_date <> "" Then
                        returnBatch = "@B#_M_" & manu_date.Replace("/", "")
                    Else
                        returnBatch = ""
                    End If
                Else
                    returnBatch = batch_no
                End If
                

            Case "EXP_DATE"
                If batch_no = "" Then
                    If exp_date <> "" Then
                        returnBatch = "@B#_E_" & exp_date.Replace("/", "")
                    Else
                        returnBatch = ""
                    End If
                Else
                    returnBatch = batch_no
                End If
                

            Case "EXP_MANU"
                If batch_no = "" Then
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


End Class
