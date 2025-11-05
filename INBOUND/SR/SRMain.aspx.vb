Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Globalization
Imports System.Data.SqlClient

Partial Class SRMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private moduleAction As String = "".Trim
    Private DDFORMAT As String = gU.getConfig("DDFORMATNO")
    Private st As New StockTrans
    Private wFun As New WMSFunc

    Private dt As New DataTable
    Private exceptionEditList As List(Of String)


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("IB_SR", Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME AS STO_NAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(RT_WH, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"))

            uiFun.load_dropdown(RT_TYPE, "Select COLC_CODE, COLC_ENG_VALUE from wms_col_code where COLC_TABCOL='WMS_STOCK_RETURN.RT_TYPE' order by COLC_DISPLAY_SEQ", "COLC_CODE", "COLC_ENG_VALUE", , Session("gSelectLabel"))
        End If

        If Session("pagemode") = "N" Then
            btnPringLbl.Visible = False
            CancelBtn.Visible = False
            BTN_C8_YN.Visible = False
            BTN_UNC8_YN.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
                'Call setDefStorerInfo()
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here

        'btnPost.OnClientClick = "javascript:getLoad();"
        btnPost.UseSubmitBehavior = False

        If Session("gLang") = "E" Then
            lheader.Text = "Stock Return Maintenance"
            lbl_ImageHd.Text = "Returned Items"
            lbl_RT_CODE.Text = "Internal Return Code WMS"
            lbl_RT_STATUS.Text = "Status"
            lbl_STORER_CODE.Text = "Organizations"
            lbl_RT_TYPE.Text = "Type"
            lbl_RT_DATE.Text = "Date"
            lbl_RT_RCV_BY.Text = "Received By"
            lbl_RT_TOT_PALLET.Text = "Total Pallet"
            lbl_RT_BY.Text = "Returned By"
            lbl_RT_BY_TEL.Text = "Serial No."
            lbl_RT_BY_EMAIL.Text = "Email"
            lbl_RT_REF_NO.Text = "Issue Control Sheet No."
            lbl_RT_WH.Text = "Subinventory"
            lbl_RT_BATCH_NO.Text = "Lot"
            lbl_RT_REM.Text = "Remarks"

            lbl_RT_CUS_CODE.Text = "MPAC Control No.:"
            lbl_RT_CONS_CODE.Text = "RMA Num:"
            lbl_RT_FAULT_REM.Text = "Fault Remark:"

            lbl_RT_C8_YN.Text = "C8 Form - Completed (Y/N)"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            saveBtn3.Text = "Save"
            saveBtn4.Text = "Save"
            CancelBtn.Text = "Cancel"
            newrow.Text = "Add"
            btnPost.Text = "Post"

            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            BTN_C8_YN.OnClientClick = "return confirm(""Are you sure to complete c8?"");"
            BTN_UNC8_YN.OnClientClick = "return confirm(""Are you sure to Un-complete c8?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            btnPost.OnClientClick = "if (confirm(""Are you sure to post this record?"")){getLoad();}else{return false;}"
            btnUnPost.OnClientClick = "if (!confirm(""Are you sure to un-post this record?\r\n(Please save your work before Un-Posting)"")) return false;"

            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[No. will be auto generated]"
            'End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "退貨維護"
            lbl_ImageHd.Text = "退貨物件"
            lbl_RT_CODE.Text = "退貨單編號"
            lbl_RT_STATUS.Text = "狀態"
            lbl_STORER_CODE.Text = "部門"
            lbl_RT_TYPE.Text = "類型"
            lbl_RT_DATE.Text = "日期"
            lbl_RT_RCV_BY.Text = "收貨者"
            lbl_RT_TOT_PALLET.Text = "貨板總量"
            lbl_RT_BY.Text = "EDI return Code"
            lbl_RT_BY_TEL.Text = "電話號碼"
            lbl_RT_BY_EMAIL.Text = "電子郵箱"
            lbl_RT_REF_NO.Text = "文件編號"
            lbl_RT_WH.Text = "子庫存"
            lbl_RT_BATCH_NO.Text = "批號"
            lbl_RT_REM.Text = "備註"

            lbl_RT_CUS_CODE.Text = "客戶編號:"
            lbl_RT_CONS_CODE.Text = "退貨單編號:"
            lbl_RT_FAULT_REM.Text = "故障備註:"

            lbl_RT_C8_YN.Text = "C8 Form - Completed (Y/N)"

            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "保存"
            saveBtn2.Text = "保存"
            saveBtn3.Text = "保存"
            saveBtn4.Text = "保存"
            CancelBtn.Text = "取消"
            newrow.Text = "新增"
            btnPost.Text = "發布"
            CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            saveBtn1.OnClientClick = "return confirm(""確定保存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""確定保存資料?"");"
            btnPost.OnClientClick = "if (confirm(""確定發布資料?"")){getLoad();}else{return false;}"
            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[号码会自动产生]"
            'End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS
        RT_TYPE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            'RT_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
            RT_C8_YN.Text = "Y"
        Else
            'RT_CODE.Enabled = False
            STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("RT_CODE") = ""

            Call BindGV()


        Else
            dt = ViewState("dt")

            If GridView1.Rows.Count > 0 Then
                For i = 0 To GridView1.Rows.Count - 1
                    'DirectCast(GridView1.Rows(i).FindControl("dsp_rtd_loc"), Label).Text = DirectCast(GridView1.Rows(i).FindControl("rtd_loc"), HiddenField).Value
                    DirectCast(GridView1.Rows(i).FindControl("dsp_rtd_vnd_code"), Label).Text = DirectCast(GridView1.Rows(i).FindControl("RTD_VND_CODE"), HiddenField).Value
                Next
            End If
        End If

        If moduleAction = "SELECTIM" Then
            addItemtoSR()
        End If

        'If RT_STATUS.Value = "POSTED" Or RT_STATUS.Value = "CANCELLED" Or RT_STATUS.Value = "PENDING" Or RT_STATUS.Value = "APPROVED" Then

        '    cSBBtn.Attributes.Add("onclick", "checkSB('" & STORER_CODE.SelectedValue & "');")
        'Else

        '    cSBBtn.Attributes.Add("onclick", "checkSB(document.myform." & STORER_CODE.ClientID & ".value);return false;")
        'End If

        setPageCtrlAccess()

        Select Case RT_STATUS.Value
            Case "NEW"

                btnApprove.Visible = False
                btnPost.Visible = False
                btnUnPost.Visible = False
            Case "PENDING"
                ar.sec_viewMode = "Y"
                exceptionEditList.Add("btnApprove")
                exceptionEditList.Add("btnUnSubmit")
                btnSubmit.Visible = False
                btnUnSubmit.Visible = True
                btnApprove.Visible = True
                btnUnPost.Visible = False
            Case "APPROVED"
                '    ar.sec_viewMode = "Y"
                exceptionEditList.Add("btnPost")
                '    btnSubmit.Visible = False
                '    btnUnSubmit.Visible = False
                '    btnApprove.Visible = False
                btnPost.Visible = True
                btnUnPost.Visible = False
            Case "CANCELLED"
                ar.sec_viewMode = "Y"
                CancelBtn.Visible = False
                BTN_C8_YN.Visible = False
                BTN_UNC8_YN.Visible = False
                btnPost.Visible = False
                btnSubmit.Visible = False
                btnUnSubmit.Visible = False
                btnUnPost.Visible = False
            Case "POSTED", "CLOSED", "PREWEIGHT"
                ar.sec_viewMode = "Y"
                btnPost.Visible = False
                If RT_STATUS.Value = "POSTED" Then
                    btnUnPost.Visible = True
                End If
                CancelBtn.Visible = False
                exceptionEditList.Add("saveBtn3")
                exceptionEditList.Add("saveBtn4")
                exceptionEditList.Add("RT_CUS_CODE")
                saveBtn3.Visible = True
                saveBtn4.Visible = True
                saveBtn1.Visible = False
                saveBtn2.Visible = False
                RT_CUS_CODE.Visible = True
        End Select

        Select Case RT_C8_YN.Text
            Case "Y"
                BTN_C8_YN.Visible = False
                BTN_UNC8_YN.Visible = True
            Case Else
                BTN_C8_YN.Visible = True
                BTN_UNC8_YN.Visible = False
        End Select

        REM ************************
        REM check right ADMIN_GP
        'ADMIN_GP
        If RT_STATUS.Value = "POSTED" Then
            Dim checkGP As String = "select usr_id from wms_user_group_alloc where grp_code = 'ADMIN_GP' and usr_id = '" & Session("usr_id") & "'"
            Dim checkGPDt As DataTable
            checkGPDt = gDB.getDataTable(checkGP)
            If checkGPDt.Rows.Count > 0 Then
                'btnUnPost.Enabled = True
                btnUnPost.Enabled = ar.hasBtnRight("BT_SR_UNPOST")
            Else
                btnUnPost.Enabled = False
            End If
        End If

        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('IB_SR','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("RT_CODE") & "','N');")

        'AddHandler btn.Click, AddressOf click_search

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)
        cSBBtn.Enabled = "true"

        'reloadDSPfield()
    End Sub

    Private Sub setPageCtrlAccess()

        exceptionEditList = New List(Of String)

        exceptionEditList.Add("CancelBtn")
        exceptionEditList.Add("BTN_C8_YN")
        exceptionEditList.Add("BTN_UNC8_YN")
        exceptionEditList.Add("btnPCancel")
        exceptionEditList.Add("pnl_RT_APPROVE_CODE")
        exceptionEditList.Add("btnPOK")
        exceptionEditList.Add("btnNew")

    End Sub

    Private Function customizectrl(ByVal ctl As Control, ByRef ctrlArrayList As ArrayList) As Boolean
        customizectrl = False
    End Function

    Private Function page_customizectrl(ByVal ctl As Control) As Boolean
        page_customizectrl = False

        If ctl.ID = "cSBBtn" Or ctl.ID = "btnUnPost" Then
            If ctl.ID = "btnUnpost" Then
                CType(ctl, Button).Enabled = ar.hasBtnRight("BT_SR_UNPOST")
            Else
                CType(ctl, Button).Enabled = True
            End If
            page_customizectrl = True
        End If
    End Function

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        'Select Case e.Row.RowType
        '    Case DataControlRowType.Header
        '        Dim oGridView As GridView = DirectCast(sender, GridView)
        '        Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

        '        REM **********************
        '        REM Use for re-create the label to change the Langauge
        '        REM Modify Here
        '        Call cU.changeGVLabel(oGridViewRow, e, "No.", "编号")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Serial No.", "序號")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "批号")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板编号")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Carton No.", "外箱编号")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Item Code", "物件号码")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封装内码")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Vendor Code", "供應商號碼")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物件名称")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Reference", "文件编号")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Expiry Date", "Expiry Date")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Manufactory Date", "Manufactory Date")
        '        Call cU.changeGVLabel(oGridViewRow, e, "RCV Qty", "收貨数量")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Loc", "位置")
        '        Call cU.changeGVLabel(oGridViewRow, e, "", "")
        '        REM **********************

        '        oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        'End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                'Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("rtd_seq"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_SEQ").ToString.Trim
                CType(e.Row.FindControl("rtd_batch_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_BATCH_NO").ToString.Trim

                'uiFun.load_ComboBox(CType(e.Row.FindControl("RTD_BATCH_NO"), AjaxControlToolkit.ComboBox), _
                '                    "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' order by 1", "dc_date_code", "dc_date_code", , "N/A", , False)
                'CType(e.Row.FindControl("RTD_BATCH_NO"), AjaxControlToolkit.ComboBox).SelectedValue = DataBinder.Eval(e.Row.DataItem, "RTD_BATCH_NO").ToString.Trim

                'Dim nDropDown As DropDownList = CType(e.Row.FindControl("RTD_UOM2"), DropDownList)
                'uiFun.load_dropdown(nDropDown, "select UOM_CODE, UOM_DESC from WMS_UOM WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "UOM_CODE", "UOM_DESC", , Session("gSelectLabel"))
                'nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "RTD_UOM2").ToString.Trim
                CType(e.Row.FindControl("RTD_UOM2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_UOM2").ToString.Trim
                CType(e.Row.FindControl("ITM_UOM"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ITM_UOM").ToString.Trim

                CType(e.Row.FindControl("rtd_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_PALLET_NO").ToString.Trim
                CType(e.Row.FindControl("rtd_carton_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_CARTON_NO").ToString.Trim
                CType(e.Row.FindControl("rtd_itm_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "RTD_ITM_CODE").ToString.Trim
                CType(e.Row.FindControl("RTD_LOC_WH"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "RTD_LOC_WH").ToString.Trim
                CType(e.Row.FindControl("SAP_MAT_DOC_NO"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "SAP_MAT_DOC_NO").ToString.Trim
                CType(e.Row.FindControl("SAP_MAT_DOC_ITEM"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "SAP_MAT_DOC_ITEM").ToString.Trim
                CType(e.Row.FindControl("RTD_STATUS"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "RTD_STATUS").ToString.Trim
                CType(e.Row.FindControl("rtd_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "RTD_PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("rtd_itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "RTD_ITM_NAME").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("rtd_ref_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_REF_NO").ToString.Trim
                CType(e.Row.FindControl("rtd_rcv_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_RCV_QTY").ToString.Trim
                CType(e.Row.FindControl("rtd_serial_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_SERIAL_NO").ToString.Trim
                CType(e.Row.FindControl("rtd_rcv_qty"), TextBox).CssClass = "REQUIRED"
                'CType(e.Row.FindControl("dsp_rtd_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "RTD_LOC").ToString.Trim
                'CType(e.Row.FindControl("rtd_loc"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "RTD_LOC").ToString.Trim
                CType(e.Row.FindControl("RTD_WH"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_WH").ToString.Trim

                Dim nDropDown As AjaxControlToolkit.ComboBox = CType(e.Row.FindControl("rtd_loc"), AjaxControlToolkit.ComboBox)
                uiFun.load_ComboBox(nDropDown, "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "RTD_WH").ToString.Trim) & "'")
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "rtd_loc").ToString.Trim

                CType(e.Row.FindControl("ITM_SERIAL_NO_YN"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ITM_SERIAL_NO_YN").ToString.Trim

                CType(e.Row.FindControl("RTD_EXPIRY_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_EXPIRY_DATE").ToString.Trim
                CType(e.Row.FindControl("RTD_MANU_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_MANU_DATE").ToString.Trim

                CType(e.Row.FindControl("dsp_rtd_vnd_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rtd_vnd_code").ToString.Trim
                CType(e.Row.FindControl("rtd_vnd_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "rtd_vnd_code").ToString.Trim

                'Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp"), Image)
                'nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                'nImage.Attributes.Add("onclick", "LocLookUp('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_rtd_loc"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("rtd_loc"), HiddenField).ClientID) & "')")
                REM **********************

                'Dim nImage2 As Image = CType(e.Row.FindControl("Image_drum_LookUp"), Image)
                'nImage2.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                'nImage2.Attributes.Add("onclick", "DrumLookUp('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("RTD_DRUM_ID"), TextBox).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("RTD_DRUM_LV"), TextBox).ClientID) & "')")

                CType(e.Row.FindControl("RTD_DRUM_ID"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_DRUM_ID").ToString.Trim
                CType(e.Row.FindControl("RTD_DRUM_LV"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_DRUM_LV").ToString.Trim

                CType(e.Row.FindControl("RTD_QTY2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_QTY2").ToString.Trim
                CType(e.Row.FindControl("RTD_KG"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "RTD_KG").ToString.Trim


                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

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
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(convert(int, RTD_SEQ)) + 1 from WMS_STOCK_RETURN_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and RT_CODE = '" & gU.dbEncode(RT_CODE.Text.Trim) & "' "
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
            dt.Rows(rows_count - 1).Item("rtd_seq") = ViewState("n_cur_seq").ToString
            REM **********************
            dt.Rows(rows_count - 1).Item("mFlag") = "N"
            dt.AcceptChanges()

            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()

        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        dt.AcceptChanges()
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""
        Dim i As Integer

        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If RT_TYPE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_RT_TYPE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_RT_TYPE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If RT_WH.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_RT_WH.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_RT_WH.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        'If String.IsNullOrWhiteSpace(RT_REF_NO2.Text) Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", "Please enter a Credit Form No.", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", "Please enter a Credit Form No.", Session("gLang"))
        '    End If

        '    Return False
        'Else

        '    Dim tempSQL As String = ""
        '    Dim creditCount As Integer = 0

        '    If Session("pagemode") <> "N" Then
        '        tempSQL = " AND RT_CODE <> '" & gU.dbEncode(ViewState("RT_CODE")) & "'"
        '    End If

        '    selectSQL = "Select Count(*) from WMS_STOCK_RETURN where WMS_STOCK_RETURN.RT_REF_NO2='" & gU.dbEncode(RT_REF_NO2.Text.Trim) & "' " & _
        '                "AND RT_STATUS <> 'CANCELLED' AND WMS_STOCK_RETURN.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
        '                "AND WMS_STOCK_RETURN.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & tempSQL

        '    creditCount = gU.decodeEmptyCInt(DB.getValueFromSQL(selectSQL), 0)

        '    If creditCount > 0 Then
        '        uiFun.displayMsg(Me, "", "This Credit Form No. has been used. Please enter another No.!", Session("gLang"))
        '        Return False
        '    End If
        'End If


        If RT_DATE.Text.Trim <> "" And Not gU.isValidDate(RT_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_RT_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_RT_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If RT_DATE.Text.Trim <> "" And gU.isValidDate(RT_DATE.Text.Trim) And gU.isValidDate(SO_DATE_HF.Value.Trim) Then
            Dim customDateTimeFormat As DateTimeFormatInfo = New DateTimeFormatInfo()
            customDateTimeFormat.DateSeparator = "/"
            customDateTimeFormat.TimeSeparator = ":"
            customDateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
            customDateTimeFormat.LongDatePattern = "dd/MM/yyyy"
            customDateTimeFormat.ShortTimePattern = "HH:mm"
            customDateTimeFormat.LongTimePattern = "HH:mm"
            customDateTimeFormat.FullDateTimePattern = "dd/MM/yyyy HH:mm"

            If Convert.ToDateTime(RT_DATE.Text.Trim, customDateTimeFormat) < Convert.ToDateTime(SO_DATE_HF.Value.Trim, customDateTimeFormat) Then
                uiFun.displayMsg(Me, "", "Return date can not be less than SO Date!", Session("gLang"))
                Return False
            End If
        End If

        If RT_BY_EMAIL.Text.Trim <> "" And Not gU.isValidEmail(RT_BY_EMAIL.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_RT_BY_EMAIL.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_RT_BY_EMAIL.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If Not gU.isDecimal(RT_TOT_PALLET.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid number, " & lbl_RT_TOT_PALLET.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的數字, " & lbl_RT_TOT_PALLET.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If RT_REM.Text.Trim.Length > 200 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Too many characters, maximum length of " & lbl_RT_REM.Text & " is 200!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "字數太多, " & lbl_RT_REM.Text & "最多只限200字!", Session("gLang"))
            End If
            Return False
        End If

        Dim itemCount As Integer = 0

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If dt.Rows(i).Item("mFlag").ToString.Trim <> "D" Then

                    If CType(GridView1.Rows(i).FindControl("rtd_rcv_qty"), TextBox).Text.Trim = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "RCV Qty cannot be empty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "收貨數量不能空白!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("rtd_rcv_qty"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid number, RCV Qty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的數字, 收貨數量!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("rtd_rcv_qty"), TextBox).Text.Trim, 0) <= 0 Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "RCV Qty need to be greater than zero !", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "收貨數量不能設零/少於零!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If CType(GridView1.Rows(i).FindControl("rtd_loc"), AjaxControlToolkit.ComboBox).SelectedValue = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Please select Location!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "位置不能空白!", Session("gLang"))
                        End If
                        Return False
                    End If
                    If CType(GridView1.Rows(i).FindControl("rtd_batch_no"), TextBox).Text = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Batch No is blank!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "Batch No is blank!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If CType(GridView1.Rows(i).FindControl("rtd_batch_no"), TextBox).Text.Trim <> "" Then
                        Dim strBatch As String = CType(GridView1.Rows(i).FindControl("rtd_batch_no"), TextBox).Text
                        Dim batchDate As String = ""
                        If strBatch IsNot Nothing AndAlso strBatch.Length >= 8 Then
                            batchDate = strBatch.Substring(0, 4) + "-" + strBatch.Substring(4, 2) + "-" + strBatch.Substring(6, 2)
                        Else
                            If Session("gLang") = "E" Then
                                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid batch number!", Session("gLang"))
                            Else
                                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid batch number!!", Session("gLang"))
                            End If
                            Return False
                        End If
                        Dim fromDateValue As DateTime
                        If DateTime.TryParse(batchDate, fromDateValue) Then

                        Else
                            If Session("gLang") = "E" Then
                                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid batch number!", Session("gLang"))
                            Else
                                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid batch number!!", Session("gLang"))
                            End If
                            Return False
                        End If
                    End If

                    If CType(GridView1.Rows(i).FindControl("rtd_uom2"), TextBox).Text = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "UOM2 is blank!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "UOM2 is blank!", Session("gLang"))
                        End If
                        Return False
                    End If
                    If CType(GridView1.Rows(i).FindControl("rtd_qty2"), TextBox).Text = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Qty2 is blank!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "Qty2 is blank!", Session("gLang"))
                        End If
                        Return False
                    End If
                    If CType(GridView1.Rows(i).FindControl("RTD_EXPIRY_DATE"), TextBox).Text = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid Expiry Date!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的失效日期!", Session("gLang"))
                        End If
                        Return False
                    End If
                    If CType(GridView1.Rows(i).FindControl("RTD_EXPIRY_DATE"), TextBox).Text <> "" AndAlso Not gU.isValidDate(CType(GridView1.Rows(i).FindControl("RTD_EXPIRY_DATE"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid Expiry Date!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的失效日期!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If CType(GridView1.Rows(i).FindControl("RTD_MANU_DATE"), TextBox).Text <> "" AndAlso Not gU.isValidDate(CType(GridView1.Rows(i).FindControl("RTD_MANU_DATE"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid Manufactory Date!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的生產日期!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("RTD_MANU_DATE"), TextBox).Text) AndAlso Not String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("RTD_EXPIRY_DATE"), TextBox).Text) Then
                        Dim customDateTimeFormat As DateTimeFormatInfo = New DateTimeFormatInfo()
                        customDateTimeFormat.DateSeparator = "/"
                        customDateTimeFormat.TimeSeparator = ":"
                        customDateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
                        customDateTimeFormat.LongDatePattern = "dd/MM/yyyy"
                        customDateTimeFormat.ShortTimePattern = "HH:mm"
                        customDateTimeFormat.LongTimePattern = "HH:mm"
                        customDateTimeFormat.FullDateTimePattern = "dd/MM/yyyy HH:mm"

                        Dim manu_date, exp_date As Date

                        manu_date = Convert.ToDateTime(CType(GridView1.Rows(i).FindControl("RTD_MANU_DATE"), TextBox).Text, customDateTimeFormat)
                        exp_date = Convert.ToDateTime(CType(GridView1.Rows(i).FindControl("RTD_EXPIRY_DATE"), TextBox).Text, customDateTimeFormat)

                        If manu_date > exp_date Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Sotck No.(" & CType(GridView1.Rows(i).FindControl("ITM_SKU_NO"), Label).Text & "): " & "Manu. Date cannot later then Expiry Date!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "Sotck No.(" & CType(GridView1.Rows(i).FindControl("ITM_SKU_NO"), Label).Text & "): " & "Manu. Date cannot later then Expiry Date!", Session("gLang"))
                            End If
                            Return False
                        End If

                    End If



                    If CType(GridView1.Rows(i).FindControl("ITM_SERIAL_NO_YN"), HiddenField).Value.Trim = "Y" AndAlso CType(GridView1.Rows(i).FindControl("rtd_serial_no"), TextBox).Text = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Serial No. is required for the item:" & CType(GridView1.Rows(i).FindControl("ITM_SKU_NO"), Label).Text, Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "Serial No. is required for the item:" & CType(GridView1.Rows(i).FindControl("ITM_SKU_NO"), Label).Text, Session("gLang"))
                        End If
                        Return False
                    End If

                    'If CType(GridView1.Rows(i).FindControl("rtd_kg"), TextBox).Text.Trim = "" Then
                    '    If Session("gLang") = "E" Then
                    '        uiFun.displayMsg(Me, "", "Item KG cannot be empty!", Session("gLang"))
                    '    Else
                    '        uiFun.displayMsg(Me, "", "收貨重量不能空白!", Session("gLang"))
                    '    End If
                    '    Return False
                    'Else
                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("rtd_kg"), TextBox).Text.Trim) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Item KG must be numeric!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "收貨重量必需為數目字!", Session("gLang"))
                        End If
                        Return False
                    End If


                    Dim ItemType As String = ""

                    ItemType = DB.getValueFromSQL("SELECT ITM_TYPE from wms_item where imp_code='" & Session("imp_code") & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and itm_code='" & CType(GridView1.Rows(i).FindControl("rtd_itm_code"), HiddenField).Value & "' and pack_key='" & gU.dbEncode(CType(GridView1.Rows(i).FindControl("rtd_pack_key"), Label).Text) & "'")


                    If CType(GridView1.Rows(i).FindControl("RTD_QTY2"), TextBox).Text = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Qty2 cannot be empty for cable item!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "收貨數量2不能空白!", Session("gLang"))
                        End If
                        Return False

                    Else
                        If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("RTD_QTY2"), TextBox).Text) Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Invalid number, Qty2!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "無效的數字, 收貨數量2!", Session("gLang"))
                            End If
                            Return False
                        Else
                            If CDbl(CType(GridView1.Rows(i).FindControl("RTD_QTY2"), TextBox).Text = 0) Then
                                If Session("gLang") = "E" Then
                                    uiFun.displayMsg(Me, "", "Qty2 cannot be zero for item!", Session("gLang"))
                                Else
                                    uiFun.displayMsg(Me, "", "收貨數量2不能設零!", Session("gLang"))
                                End If
                                Return False
                            End If
                        End If
                    End If
                    If ItemType = "CABLE" Then
                        If CType(GridView1.Rows(i).FindControl("RTD_DRUM_ID"), TextBox).Text <> "" And 1 = 2 Then
                            Dim drumCount As Integer = 0
                            drumCount = gU.decodeEmptyCInt(DB.getValueFromSQL("select Count(*) from wms_drum where imp_code='" & Session("imp_code") & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and upper(drum_id)=upper('" & gU.dbEncode(CType(GridView1.Rows(i).FindControl("RTD_DRUM_ID"), TextBox).Text) & "')"), 0)

                            If drumCount = 0 Then
                                If Session("gLang") = "E" Then
                                    uiFun.displayMsg(Me, "", "Invalid Drum ID:" & CType(GridView1.Rows(i).FindControl("RTD_DRUM_ID"), TextBox).Text & ". Drum ID Not Exists in Drum master record!", Session("gLang"))
                                Else
                                    uiFun.displayMsg(Me, "", "Invalid Drum ID:" & CType(GridView1.Rows(i).FindControl("RTD_DRUM_ID"), TextBox).Text & ". Drum ID Not Exists in Drum master record!", Session("gLang"))
                                End If
                                Return False
                            End If
                        End If

                    End If

                    itemCount += 1
                End If
            Next

        End If
        If itemCount = 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Please Select Item!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "Please Select Item!", Session("gLang"))
            End If
            Return False
        End If

        Return True

    End Function

    Protected Function save(Optional ByVal flag As String = "") As Boolean
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim gConn As SqlConnection
        Dim nextNo As String = ""

        Dim dupSQL As String = ""
        Dim dupTBL As DataTable
        Dim updateSQL As String = ""

        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("SR", gConn, transaction)
                    REM **********************

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "insert into WMS_STOCK_RETURN " &
                    "(IMP_CODE, STORER_CODE, RT_CODE, RT_TYPE, RT_STATUS, RT_DATE, RT_RCV_BY, RT_BY, " &
                     "RT_BY_TEL, RT_BY_EMAIL, RT_BATCH_NO, RT_REF_NO, RT_WH, RT_TOT_PALLET, RT_REM, " &
                     "RT_CUS_CODE, RT_CONS_CODE, RT_FAULT_REM, RT_C8_YN, " &
                     "RT_REF_NO2,RT_REF_DOC_NO," &
                     "sys_cb, sys_cd, sys_lub, sys_lud)" &
                    "values ( " &
                    "'" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " &
                    "'" & gU.dbEncode(RT_TYPE.SelectedValue) & "', '" & gU.dbEncode(RT_STATUS.Value.Trim) & "', " & gU.convdbDate(gU.dbEncode(RT_DATE.Text.Trim)) & ", " &
                    "'" & gU.dbEncode(RT_RCV_BY.Text.Trim) & "', '" & gU.dbEncode(RT_BY.Text.Trim) & "', '" & gU.dbEncode(RT_BY_TEL.Text.Trim) & "', " &
                    "'" & gU.dbEncode(RT_BY_EMAIL.Text.Trim) & "', '" & gU.dbEncode(RT_BATCH_NO.Text.Trim) & "', " &
                    gU.convdbNVCData(gU.dbEncode(RT_REF_NO.Text.Trim)) & ", '" & gU.dbEncode(RT_WH.SelectedValue) & "', " &
                    "'" & gU.dbEncode(gU.decodeNullOrEmpty(RT_TOT_PALLET.Text.Trim, "0")) & "', '" & gU.dbEncode(RT_REM.Text.Trim) & "', " &
                    "'" & gU.dbEncode(RT_CUS_CODE.Text.Trim) & "', '" & gU.dbEncode(RT_CONS_CODE.Text.Trim) & "', " &
                    "N'" & gU.dbEncode(RT_FAULT_REM.Text.Trim) & "', 'Y', " &
                    gU.convdbNVCData(gU.dbEncode(RT_REF_NO2.Text.Trim)) & "," &
                    gU.convdbNVCData(gU.dbEncode(RT_REF_DOC_NO.Text.Trim)) & "," &
                    "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "

                    '"'" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & nextNo & "', " & _
                    REM **********************
                    ViewState("RT_CODE") = nextNo

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "rtd_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    If gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, "") <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, "")) & "'"
                                        dupTBL = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTBL.Rows.Count <= 0 Then

                                            updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                        " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" &
                                                        "'" & Session("imp_code") & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                        gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, ""))) & "," &
                                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(updateSQL, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "insert into WMS_STOCK_RETURN_D " &
                                            "(IMP_CODE, STORER_CODE, RT_CODE, RTD_SEQ, RTD_PALLET_NO, RTD_CARTON_NO, RTD_BATCH_NO, RTD_REF_NO, " &
                                            "RTD_ITM_CODE,RTD_LOC_WH,SAP_MAT_DOC_NO,SAP_MAT_DOC_ITEM,RTD_STATUS, RTD_PACK_KEY, RTD_ITM_NAME, RTD_LOC, RTD_RCV_QTY, RTD_SERIAL_NO, RTD_EXPIRY_DATE, RTD_MANU_DATE,RTD_VND_CODE, " &
                                            "RTD_UOM2, RTD_QTY2, RTD_KG, RTD_DRUM_ID,RTD_WH, RTD_DRUM_LV, " &
                                            "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                            "values " &
                                            "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " &
                                             "'" & gU.dbEncode(gU.decodeNull(rows.Item("rtd_seq").ToString.Trim, "")) & "', " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(gU.decodeNull(rows.Item("rtd_pallet_no").ToString.Trim, ""), "000"))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_carton_no").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_ref_no").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_itm_code").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_LOC_WH").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("SAP_MAT_DOC_NO").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("SAP_MAT_DOC_ITEM").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_STATUS").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_pack_key").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_itm_name").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_loc").ToString.Trim, ""))) & ", " &
                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rtd_rcv_qty").ToString.Trim, "0")) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_serial_no").ToString.Trim, ""))) & ", " &
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("RTD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                             gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("RTD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_VND_CODE").ToString.Trim, "DEF_VEND"))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_UOM2").ToString.Trim, ""))) & ", " &
                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_QTY2").ToString.Trim, "0")) & ", " &
                                             gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_KG").ToString.Trim, "0")) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_DRUM_ID").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_WH").ToString.Trim, ""))) & ", " &
                                             gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_DRUM_LV").ToString.Trim, "NULL"))) & ", " &
                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                    '"(N'" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.Text) & "', '" & nextNo & "', " & _
                            End Select
                            REM **********************

                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                        Next
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode

                    'IMP_CODE, STORER_CODE, RT_CODE, RT_TYPE, RT_STATUS, RT_DATE, RT_RCV_BY, RT_BY, RT_BY_TEL, RT_BY_EMAIL, RT_BATCH_NO, RT_REF_NO, RT_WH, RT_TOT_PALLET, RT_REM
                    Dim ReturnDate As String
                    If RT_DATE.Text.Trim = System.DateTime.Now.ToString("dd/MM/yyyy") Then
                        ReturnDate = RT_DATE.Text.Trim & " " & System.DateTime.Now.ToString("HH:mm:ss")
                    Else
                        ReturnDate = RT_DATE.Text.Trim & " 23:00:00"
                    End If
                    sql_string = "update WMS_STOCK_RETURN set " &
                                    "RT_TYPE = '" & gU.dbEncode(RT_TYPE.SelectedValue) & "', " &
                                    "RT_DATE = " & gU.convdbDate(gU.dbEncode(ReturnDate)) & ", " &
                                    "RT_RCV_BY = '" & gU.dbEncode(RT_RCV_BY.Text.Trim) & "', " &
                                    "RT_BY = '" & gU.dbEncode(RT_BY.Text.Trim) & "', " &
                                    "RT_BY_TEL = '" & gU.dbEncode(RT_BY_TEL.Text.Trim) & "', " &
                                    "RT_BY_EMAIL = '" & gU.dbEncode(RT_BY_EMAIL.Text.Trim) & "', " &
                                    "RT_BATCH_NO = " & gU.convdbNVCData(gU.dbEncode(RT_BATCH_NO.Text.Trim)) & ", " &
                                    "RT_REF_NO = " & gU.convdbNVCData(gU.dbEncode(RT_REF_NO.Text.Trim)) & ", " &
                                    "RT_WH = " & gU.convdbNVCData(gU.dbEncode(RT_WH.SelectedValue)) & ", " &
                                    "RT_TOT_PALLET = " & gU.dbEncode(gU.decodeNullOrEmpty(RT_TOT_PALLET.Text.Trim, "0")) & ", " &
                                    "RT_REM =" & gU.convdbNVCData(gU.dbEncode(RT_REM.Text.Trim)) & ", " &
                                    "RT_CUS_CODE = '" & gU.dbEncode(RT_CUS_CODE.Text.Trim) & "', " &
                                    "RT_CONS_CODE = '" & gU.dbEncode(RT_CONS_CODE.Text.Trim) & "', " &
                                    "RT_FAULT_REM = " & gU.convdbNVCData(gU.dbEncode(RT_FAULT_REM.Text.Trim)) & ", " &
                                    "RT_REF_NO2 = " & gU.convdbNVCData(gU.dbEncode(RT_REF_NO2.Text.Trim)) & ", " &
                                    "RT_REF_DOC_NO = " & gU.convdbNVCData(gU.dbEncode(RT_REF_DOC_NO.Text.Trim)) & ", " &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and RT_CODE = '" & gU.dbEncode(RT_CODE.Text) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "rtd_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag").ToString.Trim
                                Case "N"
                                    If gU.decodeNullOrEmpty(rows.Item("rtd_batch_no").ToString.Trim, "") <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, "")) & "'"
                                        dupTBL = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTBL.Rows.Count <= 0 Then

                                            updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                    " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" &
                                                    "'" & Session("imp_code") & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                    "'" & gU.dbEncode(gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, "")) & "'," &
                                                     "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(updateSQL, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "insert into WMS_STOCK_RETURN_D " &
                                        "(IMP_CODE, STORER_CODE, RT_CODE, RTD_SEQ, RTD_PALLET_NO, RTD_CARTON_NO, RTD_BATCH_NO, RTD_REF_NO, " &
                                        "RTD_ITM_CODE,RTD_LOC_WH,SAP_MAT_DOC_NO,SAP_MAT_DOC_ITEM,RTD_STATUS, RTD_PACK_KEY, RTD_ITM_NAME, RTD_LOC, RTD_RCV_QTY, RTD_SERIAL_NO, RTD_EXPIRY_DATE, RTD_MANU_DATE,RTD_VND_CODE, " &
                                        "RTD_UOM2, RTD_QTY2, RTD_KG,RTD_WH, RTD_DRUM_ID, RTD_DRUM_LV, " &
                                         "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                        "values " &
                                        "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(RT_CODE.Text) & "', " &
                                         "'" & gU.dbEncode(gU.decodeNull(rows.Item("rtd_seq").ToString.Trim, "")) & "', " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(gU.decodeNull(rows.Item("rtd_pallet_no").ToString.Trim, ""), "000"))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_carton_no").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_ref_no").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_itm_code").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_LOC_WH").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("SAP_MAT_DOC_NO").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("SAP_MAT_DOC_ITEM").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_STATUS").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_pack_key").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_itm_name").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_loc").ToString.Trim, ""))) & ", " &
                                         gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rtd_rcv_qty").ToString.Trim, "0")) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_serial_no").ToString.Trim, ""))) & ", " &
                                         gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("RTD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                         gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("RTD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_VND_CODE").ToString.Trim, "DEF_VEND"))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_UOM2").ToString.Trim, ""))) & ", " &
                                         gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_QTY2").ToString.Trim, "0")) & ", " &
                                         gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_KG").ToString.Trim, "0")) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_WH").ToString.Trim, "0"))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_DRUM_ID").ToString.Trim, ""))) & ", " &
                                         gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_DRUM_LV").ToString.Trim, "NULL"))) & ", " &
                                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                Case "D"
                                    itemSQL = "delete from WMS_STOCK_RETURN_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and RT_CODE = '" & gU.dbEncode(RT_CODE.Text.Trim) & "' " &
                                        "and RTD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("rtd_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    If gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, "") <> "" Then
                                        dupSQL = "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND dc_date_code='" & gU.dbEncode(gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, "")) & "'"
                                        dupTBL = gDB.getDataTable(dupSQL, gConn, transaction)

                                        If dupTBL.Rows.Count <= 0 Then

                                            updateSQL = "Insert INTO wms_date_code (wms_date_code.IMP_CODE,  wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE, " &
                                                    " wms_date_code.SYS_CB,  wms_date_code.SYS_CD,  wms_date_code.SYS_LUB,  wms_date_code.SYS_LUD) values(" &
                                                    "'" & Session("imp_code") & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                    "'" & gU.dbEncode(gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, "")) & "'," &
                                                     "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                            gDB.amendData(updateSQL, gConn, transaction)
                                        End If
                                    End If

                                    itemSQL = "update WMS_STOCK_RETURN_D set " &
                                            "RTD_PALLET_NO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(gU.decodeNull(rows.Item("rtd_pallet_no").ToString.Trim, ""), "000"))) & ", " &
                                            "RTD_CARTON_NO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_carton_no").ToString.Trim, ""))) & ", " &
                                            "RTD_REF_NO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_ref_no").ToString.Trim, ""))) & ", " &
                                            "RTD_LOC = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_loc").ToString.Trim, ""))) & ", " &
                                            "RTD_RCV_QTY = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("rtd_rcv_qty").ToString.Trim, "0")) & ", " &
                                            "RTD_SERIAL_NO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_serial_no").ToString.Trim, ""))) & ", " &
                                            "rtd_batch_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("rtd_batch_no").ToString.Trim, ""))) & ", " &
                                            "RTD_EXPIRY_DATE =" & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("RTD_EXPIRY_DATE").ToString.Trim, ""))) & ", " &
                                            "RTD_MANU_DATE =" & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("RTD_MANU_DATE").ToString.Trim, ""))) & ", " &
                                            "RTD_UOM2=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_UOM2").ToString.Trim, ""))) & ", " &
                                            "RTD_QTY2=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_QTY2").ToString.Trim, "0")) & ", " &
                                            "RTD_KG=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_KG").ToString.Trim, "0")) & ", " &
                                            "RTD_DRUM_ID=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("RTD_DRUM_ID").ToString.Trim, ""))) & "," &
                                            "RTD_WH=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_WH").ToString.Trim, ""))) & "," &
                                            "RTD_DRUM_LV=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("RTD_DRUM_LV").ToString.Trim, "NULL"))) & ", " &
                                            "sys_lub = '" & Session("usr_id") & "', " &
                                            "sys_lud = Getdate() " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and RT_CODE = '" & gU.dbEncode(RT_CODE.Text.Trim) & "' " &
                                        "and RTD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("old_seq").ToString.Trim, "")) & "'"
                            End Select
                            REM **********************

                            'Response.Write(itemSQL)
                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                        Next
                    End If
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                For Each rows As DataRow In dt.Rows
                    If rows.Item("mFlag") = "D" Then
                        rows.Delete()
                    End If
                Next
                dt.AcceptChanges()

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    'GR_CODE.Text = nextNo
                    'Session("GR_CODE") = nextNo
                    RT_CODE.Text = nextNo
                    ViewState("RT_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    RT_CODE.ForeColor = Drawing.Color.Black
                    RT_CODE.Font.Size = 10
                    REM **********************
                End If

                If flag <> "Y" Then uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                Call BindGV()
            Catch ex As Exception
                transaction.Rollback()
                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
                Return False
            Finally
                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try

            Return True
        Else
            Return False
        End If
    End Function

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"
        Dim WhereStr As String = ""
        'Dim pk_code As String = ""
        Dim rtCode, storerCode As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("RT_CODE") <> "" Then
            'pk_code = ViewState("GR_CODE")
            rtCode = ViewState("RT_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            'pk_code = Request("GR_CODE")
            rtCode = Server.UrlDecode(Request("RT_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))

            ViewState("RT_CODE") = rtCode
            ViewState("STORER_CODE") = storerCode
        End If
        REM **********************

        REM**********************
        REM Generate Dropdown List from WMS_COL_CODE Table
        ' uiFun.load_dropdownBy_ColCode(so_currency, "WMS_QUOTATION_HD.QUO_CURRENCY", Session("gLang"))
        REM **********************

        REM**********************
        REM Generate Dropdown List from WMS_COL_CODE Table
        'uiFun.load_dropdown(so_cus_code, "select cus_code, cus_name from wms_customer", "cus_code", "cus_code")
        REM **********************

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            'GR_CODE.ForeColor = Drawing.Color.Red
            RT_STATUS.Value = "APPROVED"
            DSP_RT_STATUS.Text = "APPROVED"
            RT_DATE.Text = Now.Date.ToString("dd/MM/yyyy")
            btnPost.Visible = False
            btnApprove.Visible = False
            btnSubmit.Visible = False
            btnUnSubmit.Visible = False
            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = " SELECT wms_stock_return.IMP_CODE,  wms_stock_return.STORER_CODE,  wms_stock_return.RT_CODE, " &
                        "  wms_stock_return.RT_TYPE,  wms_stock_return.RT_STATUS,  convert(varchar, wms_stock_return.RT_DATE," & DDFORMAT & ") as RT_DATE, " &
                        "  wms_stock_return.RT_RCV_BY,  wms_stock_return.RT_CUS_CODE,  wms_stock_return.RT_C8_YN, wms_stock_return.RT_CONS_CODE, " &
                        "  wms_stock_return.RT_BY,  wms_stock_return.RT_BY_TEL,  wms_stock_return.RT_BY_EMAIL, " &
                        "  wms_stock_return.RT_BATCH_NO,  wms_stock_return.RT_REF_NO,  wms_stock_return.RT_WH, " &
                        "  wms_stock_return.RT_TOT_PALLET,  wms_stock_return.RT_REM,  wms_stock_return.RT_FAULT_REM, " &
                        "  wms_stock_return.SYS_LUB,  wms_stock_return.SYS_LUD,  wms_stock_return.SYS_CD,  wms_stock_return.SYS_CB, " &
                        " wms_stock_return.RT_REF_NO2, wms_stock_return.RT_REF_DOC_NO, convert(varchar, (select TOP(1)REQUEST_DATE from EBS_WMS_TRANS_ITX_SO_HEADER where SO_NO=wms_stock_return.RT_REF_DOC_NO AND TRIAL_TYPE='ACTUAL')," & DDFORMAT & ") as SO_DATE, wms_stock_return.RT_APPROVE_CODE " &
                        " FROM wms_stock_return " &
                        "WHERE WMS_STOCK_RETURN.RT_CODE = '" & gU.dbEncode(rtCode) & "' " &
                        "AND WMS_STOCK_RETURN.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                        "AND WMS_STOCK_RETURN.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)
            If dt.Rows.Count > 0 Then
                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                RT_CODE.Text = dt.Rows(0).Item("RT_CODE").ToString
                RT_STATUS.Value = dt.Rows(0).Item("RT_STATUS").ToString
                If dt.Rows(0).Item("RT_STATUS").ToString = "NEW" Then
                    btnSubmit.Visible = True
                End If

                DSP_RT_STATUS.Text = gDB.getColValue(dt.Rows(0).Item("RT_STATUS").ToString, "WMS_STOCK_RETURN.RT_STATUS")
                RT_CODE_HF.Value = dt.Rows(0).Item("RT_CODE").ToString
                STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                STORER_CODE_HF.Value = dt.Rows(0).Item("STORER_CODE").ToString
                RT_TYPE.SelectedValue = dt.Rows(0).Item("RT_TYPE").ToString
                RT_DATE.Text = dt.Rows(0).Item("RT_DATE").ToString
                SO_DATE_HF.Value = dt.Rows(0).Item("SO_DATE").ToString
                RT_RCV_BY.Text = dt.Rows(0).Item("RT_RCV_BY").ToString
                RT_TOT_PALLET.Text = dt.Rows(0).Item("RT_TOT_PALLET").ToString
                RT_BY.Text = dt.Rows(0).Item("RT_BY").ToString
                RT_BY_TEL.Text = dt.Rows(0).Item("RT_BY_TEL").ToString
                RT_BY_EMAIL.Text = dt.Rows(0).Item("RT_BY_EMAIL").ToString
                RT_REF_NO.Text = dt.Rows(0).Item("RT_REF_NO").ToString
                RT_WH.SelectedValue = dt.Rows(0).Item("RT_WH").ToString
                RT_BATCH_NO.Text = dt.Rows(0).Item("RT_BATCH_NO").ToString
                RT_REM.Text = dt.Rows(0).Item("RT_REM").ToString
                RT_REF_NO2.Text = dt.Rows(0).Item("RT_REF_NO2").ToString
                RT_REF_DOC_NO.Text = dt.Rows(0).Item("RT_REF_DOC_NO").ToString
                RT_APPROVE_CODE.Text = dt.Rows(0).Item("RT_APPROVE_CODE").ToString

                RT_CONS_CODE.Text = dt.Rows(0).Item("RT_CONS_CODE").ToString
                RT_FAULT_REM.Text = dt.Rows(0).Item("RT_FAULT_REM").ToString

                RT_CUS_CODE.Text = dt.Rows(0).Item("RT_CUS_CODE").ToString
                RT_C8_YN.Text = dt.Rows(0).Item("RT_C8_YN").ToString

                'RT_CUS_CODE.SelectedValue = 

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                btnGetSI.Visible = False
                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail 
        SQLString = "SELECT d.RTD_SEQ, d.RTD_SEQ as old_seq, d.RTD_PALLET_NO, d.RTD_CARTON_NO, d.RTD_BATCH_NO, d.RTD_REF_NO, d.RTD_ITM_CODE,d.RTD_LOC_WH,d.SAP_MAT_DOC_NO,d.SAP_MAT_DOC_ITEM,d.RTD_STATUS, d.RTD_PACK_KEY, d.RTD_VND_CODE," &
                    "d.RTD_ITM_NAME, d.RTD_LOC, d.RTD_RCV_QTY, d.RTD_SERIAL_NO, convert(varchar,d.RTD_MANU_DATE," & DDFORMAT & ") as RTD_MANU_DATE , convert(varchar, d.RTD_EXPIRY_DATE," & DDFORMAT & ") as RTD_EXPIRY_DATE,'0' as cb_copy, 'U' as mFlag, " &
                    " wms_item.itm_sku_no, d.RTD_UOM2, WMS_ITEM.ITM_UOM, d.RTD_QTY2, d.RTD_KG, d.RTD_DRUM_ID, d.RTD_DRUM_LV,wms_item.ITM_SERIAL_NO_YN,d.RTD_WH " &
                    " from WMS_STOCK_RETURN_D d " &
                    " INNER JOIN WMS_ITEM ON WMS_ITEM.IMP_CODE = d.IMP_CODE AND " &
                    " WMS_ITEM.STORER_CODE = d.STORER_CODE AND WMS_ITEM.ITM_CODE = d.RTD_ITM_CODE AND  " &
                    " WMS_ITEM.PACK_KEY = d.RTD_PACK_KEY " &
                    "where d.RT_CODE = '" & gU.dbEncode(rtCode) & "' " &
                    "and d.STORER_CODE = '" & gU.dbEncode(storerCode) & "' " &
                    "and d.IMP_CODE = '" & Session("IMP_CODE") & "'"

        SQLString = SQLString & " order by convert(int, d.RTD_SEQ)"
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

    Protected Sub BTN_UNC8_YN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTN_UNC8_YN.Click
        Dim cancelSql As String = "update wms_stock_return " &
                     "set rt_C8_YN = 'N', " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and rt_code = '" & gU.dbEncode(RT_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()

            RT_C8_YN.Text = "N"

            BTN_UNC8_YN.Visible = False
            BTN_C8_YN.Visible = True

            uiFun.displayMsg(Me, "", "C8 has been un-completed!", Session("gLang"))

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

    Protected Sub BTN_C8_YN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTN_C8_YN.Click
        Dim cancelSql As String = "update wms_stock_return " &
                     "set rt_C8_YN = 'Y', " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and rt_code = '" & gU.dbEncode(RT_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()

            RT_C8_YN.Text = "Y"

            BTN_C8_YN.Visible = False
            BTN_UNC8_YN.Visible = True

            uiFun.displayMsg(Me, "", "C8 has been completed!", Session("gLang"))

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

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        Dim cancelSql As String = "update wms_stock_return " &
                     "set rt_status = 'CANCELLED', " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and rt_code = '" & gU.dbEncode(RT_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()

            RT_STATUS.Value = "CANCELLED"
            DSP_RT_STATUS.Text = gDB.getColValue("CANCELLED", "WMS_STOCK_RETURN.RT_STATUS")

            ar.sec_write = "N"
            CancelBtn.Visible = False
            ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

            uiFun.displayMsg(Me, "1011", "", Session("gLang"))

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

    Protected Function saveCode(Optional ByVal flag As String = "") As Boolean
        Dim cancelSql As String = "update wms_stock_return " &
                     "set RT_CUS_CODE = '" & gU.dbEncode(RT_CUS_CODE.Text) & "', " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and rt_code = '" & gU.dbEncode(RT_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()

            uiFun.displayMsg(Me, "1007", "", Session("gLang"))

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
        saveCode = True
    End Function

    Protected Sub saveBtn3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn3.Click
        Call saveCode()
    End Sub

    Protected Sub saveBtn4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn4.Click
        Call saveCode()
    End Sub

    Protected Sub btnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPost.Click
        Dim updtSql As String
        Dim successFlag As Boolean = False
        Dim SQLString1 As String
        Dim gConn As SqlConnection
        Dim checkdt As DataTable

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        If GridView1.Rows.Count > 0 Then
            If validateAll() Then
                Call save("Y")
                Dim serial_No As String = ""
                Dim SQLStringExpDt As String = ""
                Dim checkdtExpDt As DataTable
                Try
                    For Each rows As DataRow In dt.Rows

                        'serial_No = wFun.getSerialNo(STORER_CODE.SelectedValue, rows.Item("rtd_itm_code").ToString.Trim, rows.Item("rtd_pack_key").ToString.Trim, rows.Item("RTD_SERIAL_NO").ToString.Trim)



                        st.STORER_CODE = STORER_CODE.SelectedValue
                        st.ITM_CODE = gU.decodeNull(rows.Item("rtd_itm_code").ToString.Trim, "")
                        st.PACK_KEY = gU.decodeNull(rows.Item("rtd_pack_key").ToString.Trim, "")
                        st.IO_CUST_CODE = ""
                        st.IO_AREA = ""
                        st.IO_DOC = "SR"
                        st.IO_DOC_ID = RT_CODE.Text.Trim
                        st.IO_QTY = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("rtd_rcv_qty").ToString.Trim, ""), 0)
                        st.IO_CBM = 0
                        st.IO_KG = 0
                        st.PALLET_NO = gU.decodeNull(rows.Item("rtd_pallet_no").ToString.Trim, "")
                        st.IO_WH = gU.decodeNull(rows.Item("RTD_WH").ToString.Trim, "") 'RT_WH.SelectedValue
                        st.IO_LOC = gU.decodeNull(rows.Item("rtd_loc").ToString.Trim, "")

                        SQLStringExpDt = "select Convert(varchar(10),ILOC_EXPIRY_DATE,103) ILOC_EXPIRY_DATE from WMS_ITEM_LOC_BAL where ITM_CODE='" & rows.Item("rtd_itm_code").ToString.Trim & "' and ILOC_BATCH_NO='" & rows.Item("RTD_BATCH_NO").ToString.Trim & "' and ILOC_WH='" & rows.Item("RTD_WH").ToString.Trim & "' and ILOC_LOC='" & rows.Item("rtd_loc").ToString.Trim & "' and PACK_KEY='" & rows.Item("rtd_pack_key").ToString.Trim & "' and STORER_CODE='" & STORER_CODE.SelectedValue & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE.Value.Trim) & "' "
                        checkdtExpDt = gDB.getDataTable(SQLStringExpDt, gConn, transaction)
                        If checkdtExpDt.Rows.Count > 0 Then
                            st.IO_EXPIRY_DATE = gU.decodeNull(checkdtExpDt.Rows(0).Item("ILOC_EXPIRY_DATE").ToString, "")
                        Else
                            st.IO_EXPIRY_DATE = gU.decodeNull(rows.Item("RTD_EXPIRY_DATE").ToString.Trim, "")
                        End If

                        st.IO_MANU_DATE = gU.decodeNull(rows.Item("RTD_MANU_DATE").ToString.Trim, "")
                        st.lO_BATCH_NO = gU.decodeNull(rows.Item("RTD_BATCH_NO").ToString.Trim, "")
                        st.lO_VND_CODE = gU.decodeNull(rows.Item("RTD_VND_CODE").ToString.Trim, "")

                        If rows.Item("RTD_SERIAL_NO").ToString.Trim <> "" Then
                            st.IOS_DRUM_ID = gU.decodeNull(rows.Item("RTD_DRUM_ID").ToString.Trim, "")
                            st.IOS_DRUM_LEVEL = gU.decodeNull(rows.Item("RTD_DRUM_LV").ToString.Trim, "")
                            serial_No = st.getNewSerialNo(gU.decodeNull(rows.Item("RTD_SERIAL_NO").ToString.Trim, ""), gConn, transaction)
                            st.IOS_SERIAL_NO = serial_No
                            st.IOS_QTY2 = gU.decodeEmptyCdbl(rows.Item("RTD_QTY2").ToString.Trim, 0)
                            st.IOS_UOM2 = gU.decodeNull(rows.Item("RTD_UOM2").ToString.Trim, "")

                            REM check cable
                            SQLString1 = "select ITM_CODE from WMS_ITEM WHERE ITM_SERIAL_NO_YN = 'Y' AND isnull(ITM_TYPE,'') = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "AND ITM_CODE = '" & gU.dbEncode(rows.Item("RTD_ITM_CODE").ToString.Trim) & "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("RTD_PACK_KEY").ToString.Trim) & "' "
                            checkdt = gDB.getDataTable(SQLString1)
                            If checkdt.Rows.Count > 0 Then
                                st.IOS_SL = "Y"
                            Else
                                st.IOS_SL = "N"
                            End If

                            Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("RTD_SERIAL_NO").ToString.Trim, ""), gConn, transaction)
                        End If

                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                        st.UpdateStockTrans("IN", gConn, transaction)
                        st.UpdateStockBalTrans("IN", gConn, transaction)

                        If rows.Item("RTD_SERIAL_NO").ToString.Trim <> "" Then
                            st.UpdateStockSerialTrans("IN", gConn, transaction)

                            st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                        End If

                        updtSql = "update WMS_STOCK_RETURN_D " &
                                "set RTD_SERIAL_NO = " & gU.convdbNVCData(gU.dbEncode(serial_No)) &
                                " where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and RT_CODE = '" & gU.dbEncode(RT_CODE.Text) & "' " &
                                "and RTD_SEQ='" & rows.Item("RTD_SEQ").ToString.Trim & "'"

                        gDB.amendData(updtSql, gConn, transaction)
                    Next

                    REM update related Stock Issue
                    updtSql = "update WMS_STOCK_ISSUE " &
                                "set IS_RETURN_DATE = Getdate(), IS_RETURN_DOC_NO = '" & gU.dbEncode(RT_CODE.Text) & "' " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and IS_CODE = '" & gU.dbEncode(RT_REF_DOC_NO.Text) & "' "

                    gDB.amendData(updtSql, gConn, transaction)

                    updtSql = "update WMS_STOCK_RETURN " &
                                "set RT_STATUS = 'POSTED', POSTED_DATE = GetDate() " &
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and RT_CODE = '" & gU.dbEncode(RT_CODE.Text) & "' and RT_STATUS <> 'POSTED' "

                    Dim resultCnt As Integer
                    resultCnt = gDB.amendData(updtSql, gConn, transaction)
                    If resultCnt <= 0 Then
                        transaction.Rollback()
                        successFlag = False
                    Else
                        transaction.Commit()
                        successFlag = True
                        RT_STATUS.Value = "POSTED"
                        DSP_RT_STATUS.Text = gDB.getColValue("POSTED", "WMS_STOCK_RETURN.RT_STATUS")
                        'ar.sec_viewMode = "Y"
                        ar.sec_write = "N"
                        btnPost.Visible = False
                        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
                        uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                    End If

                    If gConn IsNot Nothing Then
                        If gConn.State = ConnectionState.Open Then
                            gConn.Close()
                            gConn.Dispose()
                        End If
                    End If

                Catch ex As Exception
                    transaction.Rollback()
                    gConn.Close()
                    Throw ex
                End Try

                If successFlag Then
                    Dim rmtPost As New RemotePost
                    rmtPost.Url = "SRMAIN.aspx"
                    rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                    rmtPost.Add("RT_CODE", RT_CODE.Text)
                    rmtPost.alertMsg = "Record has been posted!"
                    rmtPost.Post()
                End If

            End If
        Else
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "No item can be posted!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "没有可供发布的物件!", Session("gLang"))
            End If
        End If

        load_ModalPopupExtender.Hide()
    End Sub

    Protected Sub RT_WH_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RT_WH.SelectedIndexChanged
        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                Dim nDropDown = CType(GridView1.Rows(i).FindControl("rtd_loc"), AjaxControlToolkit.ComboBox)
                uiFun.load_ComboBox(nDropDown, "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & gU.dbEncode(RT_WH.SelectedValue) & "'")

                'CType(GridView1.Rows(i).FindControl("dsp_rtd_loc"), Label).Text = ""
                'CType(GridView1.Rows(i).FindControl("rtd_loc"), DropDownList).SelectedValue = ""
            Next
        End If
    End Sub

    Protected Sub addItemtoSR()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(convert(int, RTD_SEQ)) + 1 from WMS_STOCK_RETURN_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and RT_CODE = '" & gU.dbEncode(RT_CODE.Text.Trim) & "' "
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

            Dim SQLString As String
            Dim tempSQL As String = ""
            Dim RO_dt As DataTable

            Dim itemPackList As String = ""
            Dim itemListarray As String()
            Dim packKeyListarray As String()
            Dim qtyListarray As String()

            Dim pref_loc As String = ""

            qtyListarray = Split(qtyList.Value, ", ")
            itemListarray = Split(itemList.Value, ", ")
            packKeyListarray = Split(packKeyList.Value, ", ")
            If itemListarray.Count = 0 Then
                If itemList.Value <> "" Then
                    itemPackList = Server.HtmlDecode(itemList.Value) & "_000_" & Server.HtmlDecode(packKeyList.Value)
                End If
            Else
                For i = 0 To itemListarray.Count - 1
                    itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemListarray(i)) & "_000_" & Server.HtmlDecode(packKeyListarray(i)))
                Next
            End If

            Dim qtyitemDict As New Dictionary(Of String, String)

            If qtyList.Value <> "" Then
                For i As Integer = 0 To itemListarray.Count - 1
                    qtyitemDict.Add(itemListarray(i) & "_000_" & packKeyListarray(i), qtyListarray(i))
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

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("RTD_SEQ") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("RTD_ITM_CODE") = RO_dt.Rows(i).Item("ITM_CODE").ToString
                dt.Rows(rows_count - 1).Item("RTD_PACK_KEY") = RO_dt.Rows(i).Item("PACK_KEY").ToString
                dt.Rows(rows_count - 1).Item("RTD_ITM_NAME") = RO_dt.Rows(i).Item("ITM_NAME").ToString
                dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = RO_dt.Rows(i).Item("ITM_SKU_NO").ToString
                dt.Rows(rows_count - 1).Item("ITM_SERIAL_NO_YN") = RO_dt.Rows(i).Item("ITM_SERIAL_NO_YN").ToString
                dt.Rows(rows_count - 1).Item("RTD_UOM2") = RO_dt.Rows(i).Item("ITM_UOM2").ToString
                dt.Rows(rows_count - 1).Item("ITM_UOM") = RO_dt.Rows(i).Item("ITM_UOM").ToString
                'RTD_UOM2

                pref_loc = ""

                If RT_WH.SelectedValue <> "" Then
                    If RT_TYPE.SelectedValue = "SCRAP" Then
                        tempSQL = "select top 1 loc from V_LOCATION where ar_scrap_area = 'Y' and wh_code='" & gU.dbEncode(RT_WH.SelectedValue) & "' order by fl_num"
                        pref_loc = DB.getValueFromSQL(tempSQL)
                    ElseIf RT_TYPE.SelectedValue = "SLCABLE" Then
                        tempSQL = "select top 1 loc from V_LOCATION where AR_SHORTLEN_CABLE_AREA = 'Y' and wh_code='" & gU.dbEncode(RT_WH.SelectedValue) & "' order by fl_num"
                        pref_loc = DB.getValueFromSQL(tempSQL)
                    Else
                        tempSQL = "select top 1 iw_pref_loc1 from wms_item_wh where wh_code='" & gU.dbEncode(RT_WH.SelectedValue) & "' and itm_code='" & gU.dbEncode(RO_dt.Rows(i).Item("ITM_CODE").ToString) & "' and pack_key='" & gU.dbEncode(RO_dt.Rows(i).Item("PACK_KEY").ToString) & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "'  order by iw_pref_loc1 asc"
                        pref_loc = DB.getValueFromSQL(tempSQL)
                    End If
                End If
                dt.Rows(rows_count - 1).Item("rtd_loc") = pref_loc


                dt.Rows(rows_count - 1).Item("RTD_VND_CODE") = RO_dt.Rows(i).Item("vnd_code")

                Dim itmQty As Integer = 0

                If qtyitemDict.ContainsKey(RO_dt.Rows(i).Item("ITM_CODE").ToString & "_000_" & RO_dt.Rows(i).Item("PACK_KEY").ToString & "_000_" & gU.decodeNullOrEmpty(RO_dt.Rows(i).Item("vnd_code").ToString, "000")) Then
                    itmQty = gU.decodeEmptyCInt(qtyitemDict(RO_dt.Rows(i).Item("ITM_CODE").ToString & "_000_" & RO_dt.Rows(i).Item("PACK_KEY").ToString & "_000_" & gU.decodeNullOrEmpty(RO_dt.Rows(i).Item("vnd_code").ToString, "000")), 0)
                End If

                dt.Rows(rows_count - 1).Item("RTD_RCV_QTY") = itmQty

                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

            Next
            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub

    Protected Sub btnGetSI_Click(sender As Object, e As System.EventArgs) Handles btnGetSI.Click
        If RT_REF_DOC_NO.Text <> "" Then
            Dim recordCount As Integer = 0
            Dim selectSQL As String = ""

            recordCount = gU.decodeEmptyCInt(DB.getValueFromSQL("select count (*) from wms_stock_issue where is_code='" & gU.dbEncode(RT_REF_DOC_NO.Text.Trim) & "' and is_status='POSTED'"), 0)

            If recordCount > 0 Then

                Dim isDT As DataTable

                selectSQL = " SELECT WMS_STOCK_ISSUE_D.IMP_CODE, WMS_STOCK_ISSUE_D.STORER_CODE, WMS_STOCK_ISSUE_D.IS_CODE, WMS_STOCK_ISSUE_D.ISD_SEQ, " &
                            " WMS_STOCK_ISSUE_D.ISD_PALLET_NO, WMS_STOCK_ISSUE_D.ISD_CARTON_NO, WMS_STOCK_ISSUE_D.ISD_BATCH_NO,  " &
                            " WMS_STOCK_ISSUE_D.ISD_REF_NO, WMS_STOCK_ISSUE_D.ISD_ITM_CODE, WMS_STOCK_ISSUE_D.ISD_PACK_KEY, WMS_STOCK_ISSUE_D.ISD_ITM_NAME,  " &
                            " WMS_STOCK_ISSUE_D.ISD_LOC, WMS_STOCK_ISSUE_D.ISD_ISSUE_QTY, WMS_STOCK_ISSUE_D.ISD_REM, WMS_STOCK_ISSUE_D.ISD_SERIAL_NO,  " &
                            " WMS_STOCK_ISSUE_D.SYS_LUB, WMS_STOCK_ISSUE_D.SYS_LUD, WMS_STOCK_ISSUE_D.SYS_CD, WMS_STOCK_ISSUE_D.SYS_CB,  " &
                            " WMS_STOCK_ISSUE_D.ISD_CUT_YN, WMS_STOCK_ISSUE_D.ISD_UOM2, WMS_STOCK_ISSUE_D.ISD_QTY2, WMS_STOCK_ISSUE_D.ISD_EXPIRY_DATE,  " &
                            " WMS_STOCK_ISSUE_D.ISD_MANU_DATE, WMS_STOCK_ISSUE_D.ISD_VND_CODE, WMS_ITEM.ITM_SKU_NO, WMS_STOCK_ISSUE_D.ISD_SERIAL_NO " &
                            " FROM WMS_STOCK_ISSUE_D INNER JOIN " &
                            " WMS_ITEM ON WMS_STOCK_ISSUE_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_ISSUE_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " &
                            " WMS_STOCK_ISSUE_D.ISD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_ISSUE_D.ISD_PACK_KEY = WMS_ITEM.PACK_KEY " &
                             "where WMS_STOCK_ISSUE_D.IS_CODE = '" & gU.dbEncode(RT_REF_DOC_NO.Text) & "' " &
                            "and WMS_STOCK_ISSUE_D.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                            "and WMS_STOCK_ISSUE_D.IMP_CODE = '" & Session("IMP_CODE") & "'"

                isDT = gDB.getDataTable(selectSQL)

                If isDT.Rows.Count > 0 Then
                    Dim newRow As DataRow
                    Dim tempDT As DataTable
                    tempDT = dt.Clone
                    tempDT.Clear()

                    For i = 0 To isDT.Rows.Count - 1

                        newRow = tempDT.NewRow

                        newRow.Item("RTD_SEQ") = i + 1
                        newRow.Item("RTD_ITM_CODE") = isDT.Rows(i).Item("ISD_ITM_CODE").ToString
                        newRow.Item("RTD_PACK_KEY") = isDT.Rows(i).Item("ISD_PACK_KEY").ToString
                        newRow.Item("RTD_ITM_NAME") = isDT.Rows(i).Item("ISD_ITM_NAME").ToString
                        newRow.Item("ITM_SKU_NO") = isDT.Rows(i).Item("ITM_SKU_NO").ToString
                        newRow.Item("RTD_PALLET_NO") = isDT.Rows(i).Item("ISD_PALLET_NO").ToString
                        newRow.Item("RTD_PALLET_NO") = isDT.Rows(i).Item("ISD_PALLET_NO").ToString
                        newRow.Item("RTD_CARTON_NO") = isDT.Rows(i).Item("ISD_CARTON_NO").ToString
                        newRow.Item("RTD_BATCH_NO") = isDT.Rows(i).Item("ISD_BATCH_NO").ToString
                        newRow.Item("RTD_REF_NO") = isDT.Rows(i).Item("ISD_REF_NO").ToString
                        newRow.Item("RTD_LOC") = isDT.Rows(i).Item("ISD_LOC").ToString
                        newRow.Item("RTD_RCV_QTY") = isDT.Rows(i).Item("ISD_ISSUE_QTY")
                        newRow.Item("RTD_EXPIRY_DATE") = isDT.Rows(i).Item("ISD_EXPIRY_DATE").ToString
                        newRow.Item("RTD_MANU_DATE") = isDT.Rows(i).Item("ISD_MANU_DATE").ToString
                        newRow.Item("RTD_VND_CODE") = isDT.Rows(i).Item("ISD_VND_CODE").ToString
                        newRow.Item("RTD_UOM2") = isDT.Rows(i).Item("ISD_UOM2").ToString
                        newRow.Item("RTD_QTY2") = isDT.Rows(i).Item("ISD_QTY2").ToString
                        newRow.Item("RTD_SERIAL_NO") = isDT.Rows(i).Item("ISD_SERIAL_NO").ToString
                        newRow.Item("mFlag") = "N"

                        tempDT.Rows.Add(newRow)
                    Next
                    tempDT.AcceptChanges()

                    dt = tempDT
                    ViewState("dt") = tempDT
                    GridView1.DataSource = tempDT
                    GridView1.DataBind()

                Else
                    uiFun.displayMsg(Me, "", "No item is found for this issue no.!", Session("gLang"))
                End If

            End If
        Else
            uiFun.displayMsg(Me, "", "Please input stock issue no. to retrieve item details!", Session("gLang"))
        End If
    End Sub

    Protected Sub btnSubmit_Click(sender As Object, e As System.EventArgs) Handles btnSubmit.Click

        If validateAll() Then
            Call save("Y")

            Dim updateSQL As String = "update wms_stock_return " &
                         "set rt_status = 'PENDING', " &
                         "RT_SBM_APP_BY='" & Session("usr_id") & "'," &
                         "RT_SBM_APP_DATE = Getdate(), " &
                         "sys_lub = '" & Session("usr_id") & "', " &
                         "sys_lud = Getdate() " &
                         "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                         "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                         "and rt_code = '" & gU.dbEncode(RT_CODE.Text) & "' "

            Dim gConn As SqlConnection

            gConn = gDB.getConnection()
            Dim transaction As SqlTransaction

            transaction = gConn.BeginTransaction()

            Try
                gDB.amendData(updateSQL, gConn, transaction)

                transaction.Commit()

                RT_STATUS.Value = "PENDING"
                DSP_RT_STATUS.Text = gDB.getColValue("PENDING", "WMS_STOCK_RETURN.RT_STATUS")

                setPageCtrlAccess()
                ar.sec_viewMode = "Y"
                exceptionEditList.Add("btnApprove")
                exceptionEditList.Add("btnUnSubmit")
                ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)

                uiFun.displayMsg(Me, "", "The order has been submitted!", Session("gLang"))
                btnSubmit.Visible = False
                btnApprove.Visible = True
                btnUnSubmit.Visible = True

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

    Protected Sub btnApprove_Click(sender As Object, e As System.EventArgs) Handles btnApprove.Click
        btnApprove_ModalPopupExtender.Show()
    End Sub

    Protected Sub btnPOK_Click(sender As Object, e As System.EventArgs) Handles btnPOK.Click
        If pnl_RT_APPROVE_CODE.Text.Trim <> "" Then
            Dim updateSQL As String = "update wms_stock_return " &
                     "set rt_status = 'APPROVED', " &
                     "RT_APPROVE_CODE=" & gU.convdbNVCData(gU.dbEncode(pnl_RT_APPROVE_CODE.Text.Trim)) & "," &
                     "RT_APPROVED_BY='" & Session("usr_id") & "'," &
                     "RT_APPROVED_DATE = Getdate(), " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and rt_code = '" & gU.dbEncode(RT_CODE.Text) & "' "

            Dim gConn As SqlConnection

            gConn = gDB.getConnection()
            Dim transaction As SqlTransaction

            transaction = gConn.BeginTransaction()

            Try
                gDB.amendData(updateSQL, gConn, transaction)

                transaction.Commit()

                RT_STATUS.Value = "APPROVED"
                DSP_RT_STATUS.Text = gDB.getColValue("APPROVED", "WMS_STOCK_RETURN.RT_STATUS")

                setPageCtrlAccess()
                ar.sec_viewMode = "Y"

                ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)
                btnPost.Enabled = True
                btnPost.Visible = True

                uiFun.displayMsg(Me, "", "The order has been approved!", Session("gLang"))

                btnSubmit.Visible = False
                btnApprove.Visible = False
                btnPost.Visible = True
                RT_APPROVE_CODE.Text = pnl_RT_APPROVE_CODE.Text.Trim
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

        Else
            uiFun.displayMsg(Me, "", "Please enter approval code to proceed!", Session("gLang"))
            btnApprove_ModalPopupExtender.Show()
        End If
    End Sub

    'Protected Sub reloadDSPfield()
    '    If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
    '        For i = 0 To GridView1.Rows.Count - 1
    '            CType(GridView1.Rows(i).FindControl("dsp_rtd_loc"), Label).Text = CType(GridView1.Rows(i).FindControl("rtd_loc"), HiddenField).Value
    '        Next
    '    End If
    'End Sub

    Protected Sub selectItemBtn_Click(sender As Object, e As System.EventArgs) Handles selectItemBtn.Click
        If RT_TYPE.SelectedValue = "SLCABLE" Then
            Session("SR_IS_Cable") = "Y"
        ElseIf RT_TYPE.SelectedValue = "SCRAP" Then
            Session("SR_IS_Scrap") = "Y"
        ElseIf RT_TYPE.SelectedValue = "NS" Then
            Session("SR_IS_NS") = "Y"
        Else
            Session("SR_IS_Scrap") = "N"
            Session("SR_IS_Cable") = "N"
            Session("SR_IS_NS") = "N"
        End If

        ScriptManager.RegisterStartupScript(LOOKUPUDP, LOOKUPUDP.GetType, "itemLookup", "ItemLookUp('" & STORER_CODE.SelectedValue & "');", True)
    End Sub

    'Protected Sub RT_REF_NO2_TextChanged(sender As Object, e As System.EventArgs) Handles RT_REF_NO2.TextChanged
    '    If Not String.IsNullOrWhiteSpace(RT_REF_NO2.Text) Then
    '        Dim selectSQL As String = ""
    '        Dim tempSQL As String = ""
    '        Dim creditCount As Integer = 0

    '        If Session("pagemode") <> "N" Then
    '            tempSQL = " AND RT_CODE <> '" & gU.dbEncode(ViewState("RT_CODE")) & "'"
    '        End If

    '        selectSQL = "Select Count(*) from WMS_STOCK_RETURN where WMS_STOCK_RETURN.RT_REF_NO2='" & gU.dbEncode(RT_REF_NO2.Text.Trim) & "' " & _
    '                    "AND RT_STATUS <> 'CANCELLED' AND WMS_STOCK_RETURN.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
    '                    "AND WMS_STOCK_RETURN.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & tempSQL

    '        creditCount = gU.decodeEmptyCInt(DB.getValueFromSQL(selectSQL), 0)

    '        If creditCount > 0 Then
    '            uiFun.displayMsgNew(cdUDP, "", "This Credit Form No. has been used. Please enter another No.!", Session("gLang"))
    '            RT_REF_NO2.BackColor = Drawing.Color.Red
    '        Else
    '            RT_REF_NO2.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFA9C")
    '        End If

    '    Else
    '        'uiFun.displayMsgNew(cdUDP, "", "Please enter Credit form No.", Session("gLang"))

    '    End If



    'End Sub

    Protected Sub btnUnSubmit_Click(sender As Object, e As System.EventArgs) Handles btnUnSubmit.Click
        If validateAll() Then
            Dim successFlag As Boolean = False

            Call save("Y")

            Dim updateSQL As String = "update wms_stock_return " &
                         "set rt_status = 'NEW', " &
                         "RT_SBM_APP_BY=NULL," &
                         "RT_SBM_APP_DATE = NULL, " &
                         "sys_lub = '" & Session("usr_id") & "', " &
                         "sys_lud = Getdate() " &
                         "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                         "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                         "and rt_code = '" & gU.dbEncode(RT_CODE.Text) & "' "

            Dim gConn As SqlConnection

            gConn = gDB.getConnection()
            Dim transaction As SqlTransaction

            transaction = gConn.BeginTransaction()

            Try
                gDB.amendData(updateSQL, gConn, transaction)

                transaction.Commit()

                'RT_STATUS.Value = "NEW"
                'DSP_RT_STATUS.Text = gDB.getColValue("NEW", "WMS_STOCK_RETURN.RT_STATUS")

                'ar = New AccessRightUtils("IB_SR", Session("usr_id"), Me)

                'setPageCtrlAccess()
                'exceptionEditList.Add("btnSubmit")
                'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)

                'uiFun.displayMsg(Me, "", "The Stock Return record has been un-submitted!", Session("gLang"))
                'btnSubmit.Visible = True
                'btnApprove.Visible = False
                'btnUnSubmit.Visible = False
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
                Dim rmtPost As New RemotePost
                rmtPost.Url = "SRMain.aspx"
                rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                rmtPost.Add("RT_CODE", ViewState("RT_CODE"))
                rmtPost.alertMsg = "The Stock Return record has been un-submitted!"
                rmtPost.Post()
            End If

        End If
    End Sub

    Protected Sub btnNew_Click(sender As Object, e As System.EventArgs) Handles btnNew.Click
        If new_action.Value = "Y" Then
            If save() Then Response.Redirect("SRMain.aspx?mode=N&menu_code=IB_SR")
        Else
            Response.Redirect("SRMain.aspx?mode=N&menu_code=IB_SR")
        End If
    End Sub

    Protected Sub btnUnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUnPost.Click
        Dim selectSql, updateSql, deleteSql As String
        Dim gConn As SqlConnection
        Dim qtyDict, cbmDict, wgtDict As Dictionary(Of String, Double)
        Dim itmKey As String
        Dim pa_dt, snDt As DataTable
        Dim SQLString1 As String
        Dim checkdt As DataTable
        Dim holddt As DataTable
        Dim coHolddt As DataTable
        Dim SQLString As String
        Dim successFlag As Boolean
        Dim dmgLoc As String
        Dim errorMsg As String = ""
        Dim hasSerial As Boolean

        SQLString1 = "SELECT RT_STATUS FROM WMS_STOCK_RETURN M " &
            "WHERE RT_STATUS <> 'POSTED' " &
            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
            "and RT_CODE = '" & gU.dbEncode(RT_CODE.Text) & "' "
        checkdt = gDB.getDataTable(SQLString1)
        If checkdt.Rows.Count > 0 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "The Stock Return status is New, please post before un-post GR!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "The Stock Return status is New, please post before un-post GR!", Session("gLang"))
            End If
            Exit Sub
        End If

        If GridView1.Rows.Count > 0 Then
            If RT_STATUS.Value = "POSTED" Then
                gConn = gDB.getConnection()

                Dim transaction As SqlTransaction
                transaction = gConn.BeginTransaction()

                Try
                    pa_dt = Session("_M_IB_GR_TMP_pa_dt")

                    snDt = Session("sn_dt")
                    Dim ifPass As Boolean
                    Dim tempEDate As String = ""
                    Dim tempMDate As String = ""
                    Dim tempUOM2 As String = ""
                    Dim tempQty2 As String = ""

                    For Each rows As DataRow In dt.Rows
                        If rows.Item("mFlag").ToString <> "D" Then
                            st.IO_SEQ = ""
                            st.STORER_CODE = STORER_CODE.SelectedValue
                            st.ITM_CODE = gU.decodeNull(rows.Item("RTD_ITM_CODE").ToString.Trim, "")
                            st.PACK_KEY = gU.decodeNull(rows.Item("RTD_PACK_KEY").ToString.Trim, "")
                            st.IO_CUST_CODE = ""
                            st.IO_WH = RT_WH.SelectedValue
                            st.IO_AREA = ""
                            st.IO_LOC = gU.decodeNull(rows.Item("RTD_LOC").ToString.Trim, "")
                            st.IO_DOC = "SR"
                            st.IO_DOC_ID = RT_CODE.Text.Trim
                            st.IO_QTY = gU.decodeNull(rows.Item("RTD_RCV_QTY").ToString.Trim, "")
                            st.PALLET_NO = gU.decodeNull(rows.Item("RTD_PALLET_NO").ToString.Trim, "")
                            st.lO_BATCH_NO = gU.decodeNull(rows.Item("RTD_BATCH_NO").ToString.Trim, "")
                            st.IO_MANU_DATE = gU.decodeNull(rows.Item("RTD_MANU_DATE").ToString.Trim, "")
                            st.lO_BATCH_NO = gU.decodeNull(rows.Item("RTD_BATCH_NO").ToString.Trim, "")
                            st.IO_CBM = 0
                            st.IO_KG = 0
                            st.IOS_DRUM_LEVEL = gU.decodeNull(rows.Item("RTD_DRUM_LV").ToString.Trim, "")
                            st.IOS_UOM2 = gU.decodeNull(rows.Item("RTD_UOM2").ToString.Trim, "")
                            st.IOS_ORG_QTY2 = gU.decodeNull(rows.Item("RTD_QTY2").ToString.Trim, "")
                            st.IOS_QTY2 = gU.decodeNull(rows.Item("RTD_QTY2").ToString.Trim, "")
                            st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                            hasSerial = False

                            REM check
                            SQLString1 = "select ITM_CODE from WMS_ITEM WHERE ITM_SERIAL_NO_YN = 'Y' AND IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "AND ITM_CODE = '" & gU.dbEncode(rows.Item("RTD_ITM_CODE").ToString.Trim) & "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("RTD_PACK_KEY").ToString.Trim) & "' "
                            checkdt = gDB.getDataTable(SQLString1)
                            If checkdt.Rows.Count > 0 Then
                                hasSerial = True
                            End If

                            If Not hasSerial Then
                                ifPass = st.UnPostStocks(StockTrans.IO_TYPE.STOCKOUT, gConn, transaction, errorMsg)

                                If Not ifPass Then
                                    Exit For
                                End If
                            Else
                                st.IOS_DRUM_ID = rows.Item("RTD_SERIAL_NO").ToString.Trim
                                st.IOS_ORG_SERIAL_NO = rows.Item("RTD_SERIAL_NO").ToString.Trim
                                st.IOS_SERIAL_NO = rows.Item("RTD_SERIAL_NO").ToString.Trim
                                st.IO_QTY = 1

                                REM check cable
                                SQLString1 = "select ITM_CODE from WMS_ITEM WHERE ITM_SERIAL_NO_YN = 'Y' AND isnull(ITM_TYPE,'') = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "AND ITM_CODE = '" & gU.dbEncode(rows.Item("RTD_ITM_CODE").ToString.Trim) & "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("RTD_PACK_KEY").ToString.Trim) & "' "
                                checkdt = gDB.getDataTable(SQLString1)
                                If checkdt.Rows.Count > 0 Then
                                    st.IOS_SL = "Y"
                                Else
                                    st.IOS_SL = "N"
                                End If

                                ifPass = st.UnPostStocks(StockTrans.IO_TYPE.STOCKOUT, gConn, transaction, errorMsg)
                                If Not ifPass Then
                                    Exit For
                                End If
                                ifPass = st.UnPostSERIAL(StockTrans.IO_TYPE.STOCKOUT, gConn, transaction, errorMsg)
                                If Not ifPass Then
                                    Exit For
                                End If
                            End If

                        End If
                    Next

                    If ifPass Then
                        updateSql = "update WMS_STOCK_RETURN " &
                                  "set RT_STATUS = 'NEW', " &
                                      "sys_lub = '" & Session("usr_id") & "', " &
                                      "sys_lud = Getdate() " &
                                  "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                  "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                  "and RT_CODE = '" & gU.dbEncode(RT_CODE.Text) & "' "

                        gDB.amendData(updateSql, gConn, transaction)

                        transaction.Commit()


                        'uiFun.displayMsgNew(updtPnlAlert, "1013", "", Session("gLang"))
                        successFlag = True
                    Else
                        transaction.Rollback()

                        successFlag = False

                    End If


                    If gConn IsNot Nothing Then
                        If gConn.State = ConnectionState.Open Then
                            gConn.Close()
                            gConn.Dispose()
                        End If
                    End If

                Catch ex As Exception
                    transaction.Rollback()
                    successFlag = False
                    gConn.Close()
                    errorMsg = ex.Message
                    'Throw ex
                End Try

                If successFlag = True Then
                    Dim rmtPost As New RemotePost
                    rmtPost.Url = "SRMain.aspx"
                    rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                    rmtPost.Add("RT_CODE", ViewState("RT_CODE"))
                    rmtPost.alertMsg = "Record has been unPosted successfully!"
                    rmtPost.Post()
                Else
                    Dim rmtPost As New RemotePost
                    rmtPost.Url = "SRMain.aspx"
                    rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                    rmtPost.Add("RT_CODE", ViewState("RT_CODE"))
                    If errorMsg <> "" Then
                        rmtPost.alertMsg = "Un-Post Failed!\r\n" & errorMsg
                    Else
                        rmtPost.alertMsg = "No item can be Un-Posted!"
                    End If
                    rmtPost.Post()
                End If

            End If
        Else
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "No item can be Un-Posted!", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "沒有可供取消發布的物件!", Session("gLang"))
            End If
        End If
    End Sub

    Private Sub reloadPage(Optional ByVal alertMsg As String = "")
        ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "RELOAD_PAGE", "reloadPage('" & gU.jsString(alertMsg) & "');", True)
    End Sub

    Protected Sub RTD_WH_TextChanged(sender As Object, e As EventArgs)
        Dim txtBox As TextBox = CType(sender, TextBox)
        Try
            If txtBox IsNot Nothing Then
                Dim strWHCOde As String = txtBox.Text.Trim
                Dim gvrow As GridViewRow = CType(sender, TextBox).NamingContainer
                Dim rowindex As Integer = CType(gvrow, GridViewRow).RowIndex
                Dim nDropDown As AjaxControlToolkit.ComboBox = CType(GridView1.Rows(rowindex).FindControl("rtd_loc"), AjaxControlToolkit.ComboBox)
                Dim selectedVal = nDropDown.SelectedValue
                uiFun.load_ComboBox(nDropDown, "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & gU.dbEncode(strWHCOde.Trim) & "'")
                If (strWHCOde = "QCFG") Then
                    nDropDown.SelectedValue = Left((selectedVal.ToString.Trim), 6) + "FG"
                ElseIf (strWHCOde = "FG01") Then
                    nDropDown.SelectedValue = Left((selectedVal.ToString.Trim), 6) + "00"
                Else
                    nDropDown.SelectedValue = selectedVal.ToString.Trim
                End If
            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub CopyItems()
        'If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
        Try
            For Each row As GridViewRow In GridView1.Rows
                Dim txtSeq As TextBox = CType(row.FindControl("rtd_seq"), TextBox)
                Dim chkCopy As CheckBox = CType(row.FindControl("cb_copy"), CheckBox)
                Dim SeqNo As Int16 = Convert.ToInt32(txtSeq.Text.Trim)
                If chkCopy.Checked Then
                    Dim tmpTbl = dt.Select("rtd_seq='" + SeqNo.ToString + "'").CopyToDataTable()
                    Dim drow As DataRow = dt.NewRow()
                    drow.ItemArray = tmpTbl.Rows(0).ItemArray
                    drow("rtd_seq") = dt.Rows.Count + 1
                    drow("old_seq") = dt.Rows.Count + 1
                    drow("mFlag") = "N"
                    drow("cb_copy") = "1"
                    dt.Rows.Add(drow)

                End If
            Next
            dt.AcceptChanges()
            uiFun.reOrderDetails(dt, "rtd_seq")
            GridView1.DataSource = dt
            GridView1.DataBind()
        Catch ex As Exception
            uiFun.displayMsgNew(updtPnlAlert, "", "Copy failed!", Session("gLang"))
        End Try
    End Sub
    Protected Sub btnCopyItem_Click(sender As Object, e As EventArgs) Handles btnCopyItem.Click

        CopyItems()
        'End If
    End Sub
End Class
