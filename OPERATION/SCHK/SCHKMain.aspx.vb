Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Partial Class OPERATION_STA_STAMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private st As New StockTrans
    Private DDFORMAT As String = "103"
    Private moduleAction As String = ""
    Private exceptionEditList As List(Of String)
    Private dt As New DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DDFORMAT = gU.getConfig("DDFORMATNO")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If

        If Not IsPostBack Then
            ViewState("pagemode") = Nothing
            ViewState("pagemode") = Request("mode")

            Session("CK_TYPE") = ""
            Session("cc_start_date") = ""
            Session("CK_IS_Cable") = ""
            Session("CK_WH") = ""
            Session("selectedSEQ") = ""
            Session("CK_PERIOD") = ""
            Session("LookupWH") = ""
            Session("SELECTED_ITMBAL") = ""

            Session("dtBin") = Nothing
            Session("dtOther") = Nothing

            If Session("usr_pref_storer") <> "" Then
                uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & Session("usr_pref_storer") & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , Session("gSelectLabel"), True)
            Else
                uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            End If
            uiFun.load_dropdown(CK_WH, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"))
            uiFun.load_dropdown(CK_TYPE, "Select COLC_CODE, COLC_ENG_VALUE from wms_col_code where COLC_TABCOL='WMS_STOCK_CHECK.CK_TYPE' order by COLC_DISPLAY_SEQ")

            uiFun.load_dropdown(CK_BY, "SELECT USR_ID, USR_FNAME + ' ' + USR_SNAME AS usr_name FROM WMS_USER where  USR_STATUS='A' and usr_type = 'S' order by usr_id", "usr_id", "usr_name")
            uiFun.load_dropdown(CK_RECHECK_BY, "SELECT USR_ID, USR_FNAME + ' ' + USR_SNAME AS usr_name FROM WMS_USER where  USR_STATUS='A' and usr_type = 'S' order by usr_id", "usr_id", "usr_name")
        End If

        If ViewState("pagemode") = "N" Then
            CancelBtn.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If
            'CK_BY.Text = Session("usr_id")

            'CK_DATE.Text = Now.Date.ToString("dd/MM/yyyy")
            ck_start_date.Text = Now.Date.ToString("dd/MM/yyyy")
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Cycle Count"
            lbl_ImageHd.Text = "Check Items"
            lbl_CK_CODE.Text = "CC Code:"
            lbl_CK_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Organizations:"
            lbl_CK_TYPE.Text = "Type:"
            lbl_CK_REF_NO.Text = "Ref. No.:"
            lbl_CK_DATE.Text = "Target Start Date:"
            lbl_CK_BY.Text = "Checked By:"
            lbl_ck_wh.Text = "Main Warehouse:"
            lbl_CK_SUB_WH.Text = "Subinventory:"
            lbl_CK_REM.Text = "Remarks:"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            newrow.Text = "Add"
            btnPost.Text = "Post"
            selectItemBtn.Text = "Select from Item Balance"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            btnPost.OnClientClick = "return confirm(""Are you sure to post this record?\r\n(Please save your work before Posting)"");"
            If ViewState("pagemode") = "N" Then
                CK_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "貨量調查維護"
            lbl_ImageHd.Text = "貨量調查詳情"

            lbl_CK_CODE.Text = "調查編號:"
            lbl_CK_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "部門:"
            lbl_CK_TYPE.Text = "類型:"
            lbl_CK_REF_NO.Text = "文件編號:"
            lbl_CK_DATE.Text = "日期:"
            lbl_CK_BY.Text = "調整者:"
            lbl_ck_wh.Text = "主倉庫:"
            lbl_CK_SUB_WH.Text = "子庫存:"
            lbl_CK_REM.Text = "備註"

            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "儲存"
            saveBtn2.Text = "儲存"
            CancelBtn.Text = "取消"
            newrow.Text = "新增"
            btnPost.Text = "發布"
            selectItemBtn.Text = "選擇物料"
            CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"
            btnPost.OnClientClick = "return confirm(""確定發布資料?"");"
            If ViewState("pagemode") = "N" Then
                CK_CODE.Text = "[號碼會自動產生]"
            End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        CK_WH.CssClass = "REQUIRED"
        CK_DATE.CssClass = "REQUIRED"
        REM **********************

        If ViewState("pagemode") = "N" Then
            'AD_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("CK_CODE") = ""

            Call BindGV()
        Else
            dt = ViewState("dt")
        End If

        If moduleAction = "SELECTIM" Then
            addItemtoSTCHK()
        ElseIf moduleAction = "SELECTIMMAST" Then
            addItemMast()
        End If

        If Not CK_IS_Cable.Checked Then
            If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
                For i = 14 To 20
                    GridView1.Columns(i).Visible = False
                Next
            End If
        End If

        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 32 To 33
                GridView1.Columns(i).Visible = False
            Next
        End If

        'cm = New CommonMenu("SCHK", lheader.text, CK_CODE.Text)
        'cm.parentDir = "../../"
        'cm.haveCheckList = "N"
        'cm.haveAttachments = "N"
        'cm.haveNotes = "N"
        'cm.haveTasks = "N"
        'cm.haveEmail = "N"
        'cm.haveHistory = "N"

        'cm.genCM(cmBar)

        'selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value,document.getElementById('CK_WH').value);")
        'selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value, document.myform." & CK_WH.ClientID & ".value);")
        setPageCtrlAccess()

        If CK_STATUS.Value = "CANCELLED" Then
            'ar.sec_write = "N"
            ar.sec_viewMode = "Y"
            CancelBtn.Visible = False


        End If

        If CK_STATUS.Value <> "NEW" Then
            btnPrintAdj.Visible = True
        Else
            btnPrintAdj.Visible = False
        End If

        Select Case CK_STATUS.Value
            Case "NEW"
                If ViewState("pagemode") <> "N" Then btnRelease.Visible = True

            Case "READY"
                'ar.sec_viewMode = "Y"
                btnRelease.Visible = False
                selectItemBtn.Visible = False
                'selectItemMastbtn.Visible = False
                'addBlank.Visible = False
                btnRechk.Visible = True
                btnCnt.Visible = True
                GridView1.Columns(0).Visible = True
                'exceptionEditList.Add("btnRechk")
            Case "CNT", "RE-CNT"
                btnCnt.Visible = True
                GridView1.Columns(0).Visible = True
                exceptionEditList.Add("btnCnt")
                btnRelease.Visible = False
                selectItemBtn.Visible = False
                'addBlank.Visible = False
                'btnRechk.Visible = True
            Case "RECHECK"
                btnCnt.Visible = True
                GridView1.Columns(0).Visible = True
                exceptionEditList.Add("btnCnt")
                selectItemBtn.Visible = False
                btnRelease.Visible = False
                'addBlank.Visible = False
                btnRechk.Visible = False
            Case Else
                btnRechk.Visible = False
                btnCnt.Visible = False
                GridView1.Columns(0).Visible = False
                btnRelease.Visible = False
                addBlank.Visible = False
                selectItemBtn.Visible = False
                selectItemMastbtn.Visible = False
                'exceptionEditList.Add("btnCnt")
                'btnRechk.Visible = True
        End Select


        If CK_STATUS.Value = "ADJUSTED" Then
            ar.sec_viewMode = "Y"
            btnGen.Visible = False
            exceptionEditList.Add("btnPrintAdj")
        End If

        If CK_TYPE.SelectedValue = "CCC" Then
            btnGen.Visible = False
            btnPrintAdj.Visible = False
        End If

        If CK_TYPE.SelectedValue = "F2L" Then
            chkITM.Visible = True
        Else
            chkITM.Visible = False
        End If

        If GridView1 IsNot Nothing AndAlso GridView1.Rows.Count > 0 Then reloadHiddenValue(GridView1)

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)
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

    Protected Sub GridView1_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Select Case e.CommandName
            Case "COUNT"
                Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)
                Dim rowNum As Integer
                Dim ckd_seq As String = ""
                Dim ckd_status As String = ""
                If Not IsNothing(gvRow) Then
                    rowNum = gvRow.RowIndex
                    ckd_seq = DirectCast(gvRow.FindControl("CKD_SEQ"), HiddenField).Value
                    ckd_status = DirectCast(gvRow.FindControl("ckd_status"), HiddenField).Value
                End If

                Call CountSingle(rowNum, ckd_seq, ckd_status)

        End Select

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        'Select Case e.Row.RowType
        '    Case DataControlRowType.Header
        '        Dim oGridView As GridView = DirectCast(sender, GridView)
        '        Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

        '        REM **********************
        '        REM Use for re-create the label to change the Langauge
        '        REM Modify Here
        '        Call cU.changeGVLabel(oGridViewRow, e, "No.", "編號")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Item Code.", "物料編號")
        '        Call cU.changeGVLabel(oGridViewRow, e, "SKU No.", "SKU No.")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內碼")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物料名稱")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "Batch No.")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Loc.", "位置")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Orig. Qty", "原來數量")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Verified Qty", "已查數量")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Variance Qty", "差異數量")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Remarks", "備註")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Expiry Date", "Expiry Date")
        '        Call cU.changeGVLabel(oGridViewRow, e, "ManuFactory Date", "ManuFactory Date")
        '        Call cU.changeGVLabel(oGridViewRow, e, "", "")
        '        REM **********************

        '        oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        'End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("mFlag"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("ckd_itm_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ckd_itm_code").ToString.Trim
                CType(e.Row.FindControl("ckd_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ckd_seq").ToString.Trim
                'If DataBinder.Eval(e.Row.DataItem, "ckd_itm_code").ToString.Trim = "" Then
                '    CType(e.Row.FindControl("itm_name_textbox"), TextBox).Visible = True
                '    CType(e.Row.FindControl("itm_name"), Label).Text = False
                'Else
                '    CType(e.Row.FindControl("itm_name_textbox"), TextBox).Visible = False
                '    CType(e.Row.FindControl("itm_name"), Label).Visible = True
                'End If

                CType(e.Row.FindControl("btnCount"), Button).CommandArgument = DataBinder.Eval(e.Row.DataItem, "ckd_seq").ToString.Trim

                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim

                CType(e.Row.FindControl("dsp_ckd_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ckd_pack_key").ToString.Trim
                CType(e.Row.FindControl("ckd_pack_key"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ckd_pack_key").ToString.Trim
                'uiFun.load_dropdown(CType(e.Row.FindControl("ckd_batch_no"), DropDownList), "select dc_date_code from wms_date_code order by 1", "dc_date_code", "dc_date_code", , Session("gSelectLabel"))
                'CType(e.Row.FindControl("ckd_batch_no"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "ckd_batch_no").ToString.Trim
                CType(e.Row.FindControl("disp_ckd_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ckd_batch_no").ToString.Trim
                CType(e.Row.FindControl("ckd_batch_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ckd_batch_no").ToString.Trim

                CType(e.Row.FindControl("itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_name").ToString.Trim
                CType(e.Row.FindControl("ckd_loc"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ckd_loc").ToString.Trim
                CType(e.Row.FindControl("dsp_ckd_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ckd_loc").ToString.Trim

                If CK_TYPE.SelectedValue = "F2L" Then
                    CType(e.Row.FindControl("dsp_ckd_org_qty"), Label).Text = ""
                    CType(e.Row.FindControl("dsp_ckd_org_qty2"), Label).Text = ""
                    CType(e.Row.FindControl("dsp_CKD_BOOK_QTY"), Label).Text = ""
                    CType(e.Row.FindControl("dsp_CKD_BOOK_QTY2"), Label).Text = ""

                Else
                    CType(e.Row.FindControl("dsp_ckd_org_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ckd_org_qty").ToString.Trim
                    CType(e.Row.FindControl("dsp_ckd_org_qty2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ckd_org_qty2").ToString.Trim
                    CType(e.Row.FindControl("dsp_CKD_BOOK_QTY"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CKD_BOOK_QTY").ToString.Trim
                    CType(e.Row.FindControl("dsp_CKD_BOOK_QTY2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CKD_BOOK_QTY2").ToString.Trim
                End If

                CType(e.Row.FindControl("ckd_org_qty"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ckd_org_qty").ToString.Trim
                CType(e.Row.FindControl("ckd_rev_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ckd_rev_qty").ToString.Trim
                CType(e.Row.FindControl("dsp_ckd_var_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ckd_var_qty").ToString.Trim

                If gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "ckd_var_qty").ToString.Trim, 0) <> 0 Then
                    CType(e.Row.FindControl("dsp_ckd_var_qty"), Label).ForeColor = Drawing.Color.Red
                Else
                    CType(e.Row.FindControl("dsp_ckd_var_qty"), Label).ForeColor = Drawing.Color.Black
                End If

                CType(e.Row.FindControl("ckd_var_qty"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ckd_var_qty").ToString.Trim
                CType(e.Row.FindControl("ckd_rem"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ckd_rem").ToString.Trim
                'CKD_ADJ_REM
                CType(e.Row.FindControl("CKD_ADJ_REM"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CKD_ADJ_REM").ToString.Trim
                CType(e.Row.FindControl("CKD_BOOK_QTY"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "CKD_BOOK_QTY").ToString.Trim
                CType(e.Row.FindControl("ckd_org_qty2"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ckd_org_qty2").ToString.Trim
                CType(e.Row.FindControl("ckd_rev_qty2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ckd_rev_qty2").ToString.Trim
                CType(e.Row.FindControl("dsp_ckd_var_qty2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ckd_var_qty2").ToString.Trim
                CType(e.Row.FindControl("ckd_var_qty2"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ckd_var_qty2").ToString.Trim

                CType(e.Row.FindControl("dsp_CKD_UOM2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CKD_UOM2").ToString.Trim
                CType(e.Row.FindControl("CKD_UOM2"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "CKD_UOM2").ToString.Trim


                If gU.decodeEmptyCdbl(DataBinder.Eval(e.Row.DataItem, "ckd_var_qty").ToString.Trim, 0) <> 0 Then
                    CType(e.Row.FindControl("dsp_ckd_var_qty2"), Label).ForeColor = Drawing.Color.Red
                Else
                    CType(e.Row.FindControl("dsp_ckd_var_qty2"), Label).ForeColor = Drawing.Color.Black
                End If


                CType(e.Row.FindControl("CKD_BOOK_QTY2"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "CKD_BOOK_QTY2").ToString.Trim

                CType(e.Row.FindControl("CKD_CC_LIST_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CKD_CC_LIST_NO").ToString.Trim
                CType(e.Row.FindControl("CKD_CC_DATE"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CKD_CC_DATE").ToString.Trim
                CType(e.Row.FindControl("CKD_STATUS"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "CKD_STATUS").ToString.Trim
                CType(e.Row.FindControl("DSP_CKD_STATUS"), Label).Text = gDB.getColValue(DataBinder.Eval(e.Row.DataItem, "CKD_STATUS").ToString.Trim, "WMS_STOCK_CHECK_D.CKD_STATUS")

                If DataBinder.Eval(e.Row.DataItem, "CKD_STATUS").ToString.Trim = "CNTD" Then
                    CType(e.Row.FindControl("btnCount"), Button).Visible = False
                End If


                CType(e.Row.FindControl("CKD_TYPE"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "CKD_TYPE").ToString.Trim
                CType(e.Row.FindControl("ckd_pallet_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ckd_pallet_no").ToString.Trim

                CType(e.Row.FindControl("ckd_rev_qty"), TextBox).Attributes.Add("onkeypress", "return maskKey(event); ")

                CType(e.Row.FindControl("CKD_ACTUAL_QTY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CKD_ACTUAL_QTY").ToString.Trim
                CType(e.Row.FindControl("CKD_ACTUAL_QTY"), TextBox).Attributes.Add("onkeypress", "return maskKey(event); ")
                CType(e.Row.FindControl("CKD_ACTUAL_QTY"), TextBox).Attributes.Add("onkeyup", "javascript:calQty(this.value, " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_CKD_BOOK_QTY"), Label).ClientID) & "', " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_ckd_var_qty"), Label).ClientID) & "', " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("CKD_BOOK_QTY"), HiddenField).ClientID) & "', " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ckd_var_qty"), HiddenField).ClientID) & "' " & _
                                                                                ");")

                CType(e.Row.FindControl("CKD_ACTUAL_QTY2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CKD_ACTUAL_QTY2").ToString.Trim
                CType(e.Row.FindControl("CKD_ACTUAL_QTY2"), TextBox).Attributes.Add("onkeypress", "return maskKey(event); ")
                CType(e.Row.FindControl("CKD_ACTUAL_QTY2"), TextBox).Attributes.Add("onkeyup", "javascript:calQty(this.value, " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_CKD_BOOK_QTY2"), Label).ClientID) & "', " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_ckd_var_qty2"), Label).ClientID) & "', " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("CKD_BOOK_QTY2"), HiddenField).ClientID) & "', " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ckd_var_qty2"), HiddenField).ClientID) & "' " & _
                                                                                ");")

                CType(e.Row.FindControl("CKD_FULL_DRUM"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "CKD_FULL_DRUM").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "CKD_DRUM_ID").ToString.Trim = "00" OrElse DataBinder.Eval(e.Row.DataItem, "CKD_DRUM_ID").ToString.Trim = "0" Then
                    CType(e.Row.FindControl("dsp_CKD_DRUM_ID"), Label).Text = "On Grd."
                Else
                    CType(e.Row.FindControl("dsp_CKD_DRUM_ID"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CKD_DRUM_ID").ToString.Trim
                End If

                CType(e.Row.FindControl("CKD_DRUM_ID"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "CKD_DRUM_ID").ToString.Trim

                CType(e.Row.FindControl("dsp_CKD_DRUM_LEVEL"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CKD_DRUM_LEVEL").ToString.Trim
                CType(e.Row.FindControl("CKD_DRUM_LEVEL"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "CKD_DRUM_LEVEL").ToString.Trim

                CType(e.Row.FindControl("dsp_CKD_SERIAL_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CKD_SERIAL_NO").ToString.Trim
                CType(e.Row.FindControl("CKD_SERIAL_NO"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "CKD_SERIAL_NO").ToString.Trim

                Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp"), Image)
                nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                nImage.Attributes.Add("onclick", "LocLookUp('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_ckd_loc"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ckd_loc"), HiddenField).ClientID) & "'," & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_wh_code"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("wh_code"), HiddenField).ClientID) & "'," & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_bn_csms_code"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("bn_csms_code"), HiddenField).ClientID) & "')")

                If DataBinder.Eval(e.Row.DataItem, "from_imast").ToString.Trim <> "Y" Then nImage.Visible = False
                CType(e.Row.FindControl("from_imast"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "from_imast").ToString.Trim


                CType(e.Row.FindControl("dsp_WH_CODE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "WH_CODE").ToString.Trim
                CType(e.Row.FindControl("WH_CODE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "WH_CODE").ToString.Trim

                CType(e.Row.FindControl("dsp_BN_CSMS_CODE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "BN_CSMS_CODE").ToString.Trim
                CType(e.Row.FindControl("BN_CSMS_CODE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "BN_CSMS_CODE").ToString.Trim

                'Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp"), Image)
                'nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                'nImage.Attributes.Add("onclick", "LocLookUp('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_ckd_loc"), Label).ClientID) & "', " & _
                '                                            "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ckd_loc"), HiddenField).ClientID) & "', " & _
                '                                            "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_ckd_org_qty"), Label).ClientID) & "', " & _
                '                                            "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ckd_org_qty"), HiddenField).ClientID) & "', " & _
                '                                            "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "ckd_itm_code").ToString.Trim) & "', " & _
                '                                            "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "ckd_pack_key").ToString.Trim) & "'    " & _
                '                                            ")")

                'nImage.Attributes.Add("onclick", "ItemLocLookUp(document.myform." & STORER_CODE.ClientID & ".value, " & _
                '                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "ckd_itm_code").ToString.Trim) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "ckd_pack_key").ToString.Trim) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_ckd_loc"), Label).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ckd_loc"), HiddenField).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ckd_pallet_no"), HiddenField).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_ckd_org_qty"), Label).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ckd_org_qty"), HiddenField).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("disp_ckd_batch_no"), Label).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ckd_batch_no"), HiddenField).ClientID) & "')")

                CType(e.Row.FindControl("CKD_WITNESS"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CKD_WITNESS").ToString.Trim
                CType(e.Row.FindControl("CKD_CHECKER"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "CKD_CHECKER").ToString.Trim

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    nButton.Text = "删除"
                End If

                If Not ar.hasBtnRight("BT_CC_RECHECK") Then
                    CType(e.Row.FindControl("CKD_ACTUAL_QTY"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("CKD_ACTUAL_QTY"), TextBox).CssClass = "READONLY"
                    CType(e.Row.FindControl("CKD_ACTUAL_QTY2"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("CKD_ACTUAL_QTY2"), TextBox).CssClass = "READONLY"
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    Call ar.hideGVRow(GridView1, e.Row)
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                    'nButton.Enabled = False
                End If

                If DataBinder.Eval(e.Row.DataItem, "CKD_AD_CODE").ToString.Trim <> "" Then
                    CType(e.Row.FindControl("CKD_AD_CODE"), HyperLink).Text = DataBinder.Eval(e.Row.DataItem, "CKD_AD_CODE").ToString.Trim
                    CType(e.Row.FindControl("CKD_AD_CODE"), HyperLink).NavigateUrl = "~/OPERATION/SADJ/SADJMain.aspx?storer_code=" & ViewState("STORER_CODE") & "&ad_code=" & DataBinder.Eval(e.Row.DataItem, "CKD_AD_CODE").ToString.Trim
                End If

                If DataBinder.Eval(e.Row.DataItem, "CKD_STATUS").ToString.Trim = "CNTD" OrElse DataBinder.Eval(e.Row.DataItem, "CKD_STATUS").ToString.Trim = "ADJ" Then
                    CType(e.Row.FindControl("CKD_ACTUAL_QTY"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("CKD_ACTUAL_QTY"), TextBox).CssClass = "READONLY"
                    CType(e.Row.FindControl("CKD_ACTUAL_QTY2"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("CKD_ACTUAL_QTY2"), TextBox).CssClass = "READONLY"

                    CType(e.Row.FindControl("ckd_rev_qty2"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("ckd_rev_qty2"), TextBox).CssClass = "READONLY"

                    CType(e.Row.FindControl("CKD_CC_LIST_NO"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("CKD_CC_LIST_NO"), TextBox).CssClass = "READONLY"
                    CType(e.Row.FindControl("ckd_rev_qty"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("ckd_rev_qty"), TextBox).CssClass = "READONLY"
                    'CType(e.Row.FindControl("ckd_rem"), TextBox).Attributes.Add("readonly", "readonly")
                    'CType(e.Row.FindControl("ckd_rem"), TextBox).CssClass = "READONLY"
                    CType(e.Row.FindControl("CKD_CC_DATE"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("CKD_CC_DATE"), TextBox).CssClass = "READONLY"

                    CType(e.Row.FindControl("CKD_TYPE"), DropDownList).Enabled = False
                    CType(e.Row.FindControl("btnCal"), ImageButton).Visible = False
                ElseIf DataBinder.Eval(e.Row.DataItem, "CKD_STATUS").ToString.Trim = "RECHECK" Then
                    CType(e.Row.FindControl("ckd_rev_qty"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("ckd_rev_qty"), TextBox).CssClass = "READONLY"
                End If

        End Select
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        dt.AcceptChanges()

        DirectCast(GridView1.Rows(e.RowIndex).FindControl("mFlag"), HiddenField).Value = "D"

        ViewState("dt") = dt
    End Sub

    Private Function validateAll(Optional ByRef flag As String = "") As Boolean
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

        If CK_WH.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_ck_wh.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_ck_wh.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If CK_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_CK_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_CK_DATE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        ElseIf Not gU.isValidDate(CK_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_CK_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_CK_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        Dim itemCount As Integer = 0

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If CType(GridView1.Rows(i).FindControl("mFlag"), HiddenField).Value <> "D" Then
                    Dim loc As String = CType(GridView1.Rows(i).FindControl("ckd_loc"), HiddenField).Value

                    'If loc = "" Then
                    '    If Session("gLang") = "E" Then
                    '        uiFun.displayMsg(Me, "", "Location Cannot Be Empty!", Session("gLang"))
                    '    Else
                    '        uiFun.displayMsg(Me, "", "位置不能空白!", Session("gLang"))
                    '    End If
                    '    Return False
                    'End If

                    If CType(GridView1.Rows(i).FindControl("ckd_rev_qty"), TextBox).Text <> "" Then
                        If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("ckd_rev_qty"), TextBox).Text) Then
                            uiFun.displayMsg(Me, "", "Checked Qty must be numeric!", Session("gLang"))
                            Return False
                        End If
                    Else
                        'If Session("gLang") = "E" Then
                        '    uiFun.displayMsg(Me, "", "Revised Qty Cannot Be Empty!", Session("gLang"))
                        'Else
                        '    uiFun.displayMsg(Me, "", "修訂數量不能空白!", Session("gLang"))
                        'End If
                        'Return False
                    End If

                    If CType(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY"), TextBox).Text <> "" Then
                        If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY"), TextBox).Text) Then
                            uiFun.displayMsg(Me, "", "ReChecked Qty must be numeric!", Session("gLang"))
                            Return False
                        End If
                    End If

                    If CType(GridView1.Rows(i).FindControl("from_imast"), HiddenField).Value = "Y" AndAlso CType(GridView1.Rows(i).FindControl("ckd_loc"), HiddenField).Value = "" Then
                        uiFun.displayMsg(Me, "", "Please select location!", Session("gLang"))
                        Return False
                    End If

                    itemCount += 1
                End If
            Next
        End If

        If itemCount = 0 Then
            uiFun.displayMsg(Me, "", "Please Select At least 1 item!", Session("gLang"))
        End If

        Return True

    End Function

    Protected Function save(Optional ByVal flag As String = "") As Boolean
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim nextNo As String = ""
        Dim gConn As SqlConnection
        Dim dupSQL As String = ""
        Dim dupTbl As New DataTable
        Dim successFlag As Boolean = False
        Dim varQty As String = ""
        Dim varQty2 As String = ""
        Dim ckd_seq As String = ""
        Dim item_name As String = ""
        Dim ccSQL As String = ""
        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                Dim isCable As String = ""

                If CK_IS_Cable.Checked Then
                    isCable = "Y"
                End If

                If ViewState("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    Dim pre_wh As String = ""

                    nextNo = DB.getDocNo("CHK", gConn, transaction)

                    If CK_TYPE.SelectedValue = "CCC" Then
                        pre_wh = "CBL"
                    Else
                        pre_wh = CK_WH.SelectedValue
                    End If

                    'nextNo = pre_wh & ck_period.SelectedValue & Now.Year.ToString & nextNo


                    'nextNo = AD_CODE.Text
                    REM **********************

                    dupSQL = "select 1 from wms_stock_check " & _
                                "where ck_code = '" & gU.dbEncode(nextNo) & "' " & _
                                "and imp_code='" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into wms_stock_check (" &
                        "ck_code, imp_code, storer_code, " &
                        "ck_status, ck_date, ck_by, CK_RECHECK_BY," &
                        "ck_type, ck_ref_no, " &
                        "ck_wh, ck_rem, " &
                        "ck_level,ck_item_level, " &
                        "CK_ACTUAL_DATE, CK_START_DATE, CK_END_DATE, CK_IS_CABLE,CK_COMPLT_DATE," &
                        "CK_SUB_WH, " &
                        "sys_cb, sys_cd, sys_lub, sys_lud)" &
                        "values ( " &
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CK_STATUS.Value.Trim)) & ", " & gU.convdbDate(gU.dbEncode(CK_DATE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(CK_BY.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(CK_RECHECK_BY.SelectedValue)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CK_TYPE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(CK_REF_NO.Text.Trim)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CK_WH.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(CK_REM.Text.Trim)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CK_LEVEL.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(CK_ITEM_LEVEL.SelectedValue)) & "," &
                        gU.convdbDate(gU.dbEncode(CK_ACTUAL_DATE.Text.Trim)) & "," & gU.convdbDate(gU.dbEncode(ck_start_date.Text.Trim)) & "," & gU.convdbDate(gU.dbEncode(ck_end_date.Text.Trim)) & "," &
                        gU.convdbNVCData(gU.dbEncode(isCable)) & "," & gU.convdbDate(gU.dbEncode(CK_COMPLT_DATE.Text.Trim)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CK_SUB_WH.SelectedValue)) & "," &
                        "'" & Session("usr_id") & "',getdate(),'" & Session("usr_id") & "',getdate()) "

                        '" & gU.convdbNVCData(gU.dbEncode(ck_period.SelectedValue)) & ",
                        ',CK_PERIOD

                        REM **********************s

                        ViewState("CK_CODE") = nextNo

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                            'uiFun.reOrderDetails(dt, "ckd_seq")
                            ViewState("dt") = dt

                            For Each rows As DataRow In dt.Rows
                                itemSQL = ""

                                REM **********************
                                REM Modify Here
                                If rows.Item("mFlag") IsNot Nothing And rows.Item("mFlag") IsNot DBNull.Value Then
                                    Select Case rows.Item("mFlag").ToString.Trim
                                        Case "N"
                                            varQty = ""
                                            varQty2 = ""
                                            'If rows.Item("ckd_rev_qty").ToString.Trim <> "" Then
                                            '    If gU.decodeEmptyCdbl(rows.Item("ckd_org_qty").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("ckd_rev_qty").ToString.Trim, 0) Then
                                            '        varQty = gU.decodeEmptyCdbl(rows.Item("ckd_org_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("ckd_rev_qty").ToString.Trim, 0)
                                            '        varQty = Math.Abs(CDbl(varQty)) * -1

                                            '    ElseIf gU.decodeEmptyCdbl(rows.Item("ckd_org_qty").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("ckd_rev_qty").ToString.Trim, 0) Then
                                            '        varQty = gU.decodeEmptyCdbl(rows.Item("ckd_rev_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("ckd_org_qty").ToString.Trim, 0)
                                            '    ElseIf gU.decodeEmptyCdbl(rows.Item("ckd_org_qty").ToString.Trim, 0) = gU.decodeEmptyCdbl(rows.Item("ckd_rev_qty").ToString.Trim, 0) Then
                                            '        varQty = 0
                                            '    End If
                                            'Else
                                            '    varQty = ""
                                            'End If

                                            If rows.Item("CKD_ACTUAL_QTY").ToString.Trim <> "" Then
                                                If gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, 0) Then
                                                    varQty = gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, 0)
                                                    varQty = Math.Abs(CDbl(varQty)) * -1

                                                ElseIf gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, 0) Then
                                                    varQty = gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY").ToString.Trim, 0)
                                                ElseIf gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY").ToString.Trim, 0) = gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, 0) Then
                                                    varQty = 0
                                                End If
                                            Else
                                                varQty = ""
                                            End If

                                            If rows.Item("CKD_ACTUAL_QTY2").ToString.Trim <> "" AndAlso rows.Item("CKD_BOOK_QTY2").ToString.Trim <> "" Then
                                                If gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY2").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, 0) Then
                                                    varQty2 = gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY2").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, 0)
                                                    varQty2 = Math.Abs(CDbl(varQty2)) * -1

                                                ElseIf gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY2").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, 0) Then
                                                    varQty2 = gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY2").ToString.Trim, 0)
                                                ElseIf gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY2").ToString.Trim, 0) = gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, 0) Then
                                                    varQty2 = 0
                                                End If
                                            Else
                                                varQty2 = ""
                                            End If

                                            Dim seqSQL As String = "select max(cast (ckd_seq as int)) + 1 from wms_stock_check_d where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                                         "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                                         "and ck_code = '" & gU.dbEncode(nextNo) & "' "
                                            ckd_seq = gU.decodeNullOrEmpty(DB.getValueFromSQL(seqSQL, gConn, transaction), "1")


                                            itemSQL = "insert into wms_stock_check_d (" &
                                                      "ck_code, imp_code, storer_code, " &
                                                      "ckd_seq, ckd_itm_code, ckd_pack_key, ckd_batch_no, " &
                                                      "ckd_loc, ckd_org_qty, ckd_rev_qty, " &
                                                      "ckd_var_qty, ckd_rem, ckd_pallet_no, " &
                                                      "ckd_expiry_date, ckd_manu_date, " &
                                                      "CKD_ACTUAL_QTY, CKD_TYPE, CKD_CC_DATE, CKD_CC_LIST_NO, " &
                                                      "CKD_FULL_DRUM, CKD_DRUM_ID, CKD_DRUM_LEVEL, CKD_SERIAL_NO, CKD_UOM2, CKD_ORG_QTY2, CKD_REV_QTY2, CKD_VAR_QTY2, CKD_ACTUAL_QTY2, CKD_BOOK_QTY, CKD_BOOK_QTY2," &
                                                      "ckd_VND_CODE, CKD_WITNESS,CKD_CHECKER,CKD_ADJ_REM,ckd_status," &
                                                      "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                      "values (" &
                                                      gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(ckd_seq, "1"))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ckd_itm_code").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ckd_pack_key").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ckd_batch_no").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ckd_loc").ToString.Trim, ""))) & ", " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ckd_org_qty").ToString.Trim, "NULL")) & ", " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ckd_rev_qty").ToString.Trim, "NULL")) & ", " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(varQty, "NULL")) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_rem").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_pallet_no").ToString.Trim, ""))) & ", " &
                                                      gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ckd_expiry_date").ToString.Trim, ""))) & ", " &
                                                      gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ckd_manu_date").ToString.Trim, ""))) & ", " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, "NULL")) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_TYPE").ToString.Trim, ""))) & ", " &
                                                      gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("CKD_CC_DATE").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_CC_LIST_NO").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_FULL_DRUM").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(CK_TYPE.SelectedValue)) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_DRUM_LEVEL").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_SERIAL_NO").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_UOM2").ToString.Trim, ""))) & ", " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_ORG_QTY2").ToString.Trim, "NULL")) & ", " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_REV_QTY2").ToString.Trim, "NULL")) & ", " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(varQty2, "NULL")) & ", " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, "NULL")) & ", " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_BOOK_QTY").ToString.Trim, "NULL")) & ", " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_BOOK_QTY2").ToString.Trim, "NULL")) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_VND_CODE").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_WITNESS").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_CHECKER").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_ADJ_REM").ToString.Trim, ""))) & ", " &
                                                      gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_STATUS").ToString.Trim, ""))) & ", " &
                                                      "'" & Session("usr_id") & "',getdate(),'" & Session("usr_id") & "',getdate()) "

                                            '"(N'" & Session("imp_code") & "', N'" & gU.dbEncode(STORER_CODE.Text) & "', N'" & nextNo & "', " & _

                                            ccSQL = "Update WMS_ITEM set ITM_CC_DATE=getdate() where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                        "and itm_code='" & gU.dbEncode(rows.Item("ckd_itm_code").ToString.Trim) & "' and pack_key='" & gU.dbEncode(rows.Item("ckd_pack_key").ToString.Trim) & "'"

                                            gDB.amendData(ccSQL, gConn, transaction)

                                    End Select
                                    REM **********************

                                    If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                                End If
                            Next
                        End If

                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Cycle Count!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "貨量調整資料重複!!", Session("gLang"))
                        End If

                        Return False
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_stock_check set " &
                                    "ck_status = " & gU.convdbNVCData(gU.dbEncode(CK_STATUS.Value)) & ", " &
                                    "ck_date = " & gU.convdbDate(gU.dbEncode(CK_DATE.Text.Trim)) & ", " &
                                    "ck_by = " & gU.convdbNVCData(gU.dbEncode(CK_BY.SelectedValue)) & ", " &
                                    "CK_RECHECK_BY = " & gU.convdbNVCData(gU.dbEncode(CK_RECHECK_BY.SelectedValue)) & ", " &
                                    "ck_type = " & gU.convdbNVCData(gU.dbEncode(CK_TYPE.SelectedValue)) & ", " &
                                    "ck_ref_no = " & gU.convdbNVCData(gU.dbEncode(CK_REF_NO.Text.Trim)) & ", " &
                                    "ck_wh = " & gU.convdbNVCData(gU.dbEncode(CK_WH.SelectedValue)) & ", " &
                                    "ck_rem = " & gU.convdbNVCData(gU.dbEncode(CK_REM.Text.Trim)) & ", " &
                                    "ck_level=" & gU.convdbNVCData(gU.dbEncode(CK_LEVEL.SelectedValue.Trim)) & "," &
                                    "ck_item_level=" & gU.convdbNVCData(gU.dbEncode(CK_ITEM_LEVEL.SelectedValue.Trim)) & ", " &
                                    "CK_ACTUAL_DATE=" & gU.convdbDate(gU.dbEncode(CK_ACTUAL_DATE.Text.Trim)) & "," &
                                    "CK_START_DATE=" & gU.convdbDate(gU.dbEncode(ck_start_date.Text.Trim)) & "," &
                                    "CK_END_DATE=" & gU.convdbDate(gU.dbEncode(ck_end_date.Text.Trim)) & "," &
                                    "CK_IS_CABLE=" & gU.convdbNVCData(isCable) & "," &
                                    "CK_COMPLT_DATE=" & gU.convdbDate(gU.dbEncode(CK_COMPLT_DATE.Text.Trim)) & "," &
                                    "CK_SUB_WH=" & gU.convdbNVCData(CK_SUB_WH.SelectedValue) & "," &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = getdate() " &
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and ck_code = '" & gU.dbEncode(CK_CODE.Text) & "' "


                    '"CK_PERIOD=" & gU.convdbNVCData(gU.dbEncode(ck_period.SelectedValue)) & "," & _
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        'uiFun.reOrderDetails(dt, "ckd_seq")
                        ViewState("dt") = dt
                        For Each rows As DataRow In dt.Rows

                            varQty = ""
                            varQty2 = ""
                            'If rows.Item("ckd_rev_qty").ToString.Trim <> "" Then
                            '    If gU.decodeEmptyCdbl(rows.Item("ckd_org_qty").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("ckd_rev_qty").ToString.Trim, 0) Then
                            '        varQty = gU.decodeEmptyCdbl(rows.Item("ckd_org_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("ckd_rev_qty").ToString.Trim, 0)
                            '        varQty = Math.Abs(CDbl(varQty)) * -1

                            '    ElseIf gU.decodeEmptyCdbl(rows.Item("ckd_org_qty").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("ckd_rev_qty").ToString.Trim, 0) Then
                            '        varQty = gU.decodeEmptyCdbl(rows.Item("ckd_rev_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("ckd_org_qty").ToString.Trim, 0)
                            '    ElseIf gU.decodeEmptyCdbl(rows.Item("ckd_org_qty").ToString.Trim, 0) = gU.decodeEmptyCdbl(rows.Item("ckd_rev_qty").ToString.Trim, 0) Then
                            '        varQty = 0
                            '    End If
                            'Else
                            '    varQty = ""
                            'End If

                            If rows.Item("CKD_ACTUAL_QTY").ToString.Trim <> "" Then
                                If gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, 0) Then
                                    varQty = gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, 0)
                                    varQty = Math.Abs(CDbl(varQty)) * -1

                                ElseIf gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, 0) Then
                                    varQty = gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY").ToString.Trim, 0)
                                ElseIf gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY").ToString.Trim, 0) = gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, 0) Then
                                    varQty = 0
                                End If
                            Else
                                varQty = ""
                            End If


                            If rows.Item("CKD_ACTUAL_QTY2").ToString.Trim <> "" AndAlso rows.Item("CKD_BOOK_QTY2").ToString.Trim <> "" Then
                                If gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY2").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, 0) Then
                                    varQty2 = gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY2").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, 0)
                                    varQty2 = Math.Abs(CDbl(varQty2)) * -1

                                ElseIf gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY2").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, 0) Then
                                    varQty2 = gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY2").ToString.Trim, 0)
                                ElseIf gU.decodeEmptyCdbl(rows.Item("CKD_BOOK_QTY2").ToString.Trim, 0) = gU.decodeEmptyCdbl(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, 0) Then
                                    varQty2 = 0
                                End If
                            Else
                                varQty2 = ""
                            End If

                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            If rows.Item("mFlag") IsNot Nothing And rows.Item("mFlag") IsNot DBNull.Value Then
                                Select Case rows.Item("mFlag")
                                    Case "N"

                                        ckd_seq = gU.decodeNullOrEmpty(DB.getValueFromSQL("select max(convert(int,ckd_seq)) + 1 from wms_stock_check_d where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                                                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                                     "and ck_code = '" & gU.dbEncode(CK_CODE.Text) & "' ", gConn, transaction), "1")

                                        itemSQL = "insert into wms_stock_check_d (" &
                                                    "ck_code, imp_code, storer_code, " &
                                                    "ckd_seq, ckd_itm_code, ckd_pack_key, ckd_batch_no, " &
                                                    "ckd_loc, ckd_org_qty, ckd_rev_qty, " &
                                                    "ckd_var_qty, ckd_rem, ckd_pallet_no, " &
                                                    "ckd_expiry_date, ckd_manu_date, " &
                                                    "CKD_ACTUAL_QTY, CKD_TYPE, CKD_CC_DATE, CKD_CC_LIST_NO, " &
                                                    "CKD_FULL_DRUM, CKD_DRUM_ID, CKD_DRUM_LEVEL, CKD_SERIAL_NO, CKD_UOM2, CKD_ORG_QTY2, CKD_REV_QTY2, CKD_VAR_QTY2, CKD_ACTUAL_QTY2, CKD_BOOK_QTY, CKD_BOOK_QTY2," &
                                                    "ckd_VND_CODE,CKD_WITNESS,CKD_CHECKER, CKD_ADJ_REM,CKD_STATUS, " &
                                                    "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                    "values (" &
                                                    gU.convdbNVCData(gU.dbEncode(CK_CODE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(ckd_seq, "1"))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_itm_code").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_pack_key").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_batch_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_loc").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ckd_org_qty").ToString.Trim, "NULL")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ckd_rev_qty").ToString.Trim, "NULL")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(varQty, "NULL")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_rem").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_pallet_no").ToString.Trim, ""))) & ", " &
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ckd_expiry_date").ToString.Trim, ""))) & ", " &
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ckd_manu_date").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, "NULL")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_TYPE").ToString.Trim, ""))) & ", " &
                                                    gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("CKD_CC_DATE").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_CC_LIST_NO").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_FULL_DRUM").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(CK_TYPE.SelectedValue)) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_DRUM_LEVEL").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_SERIAL_NO").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_UOM2").ToString.Trim, ""))) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_ORG_QTY2").ToString.Trim, "NULL")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_REV_QTY2").ToString.Trim, "NULL")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(varQty2, "NULL")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, "NULL")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_BOOK_QTY").ToString.Trim, "NULL")) & ", " &
                                                    gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_BOOK_QTY2").ToString.Trim, "NULL")) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_VND_CODE").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_WITNESS").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_CHECKER").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_ADJ_REM").ToString.Trim, ""))) & ", " &
                                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_STATUS").ToString.Trim, ""))) & ", " &
                                                    "'" & Session("usr_id") & "',getdate(),'" & Session("usr_id") & "',getdate()) "

                                        ccSQL = "Update WMS_ITEM set ITM_CC_DATE=getdate() where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                       "and itm_code='" & gU.dbEncode(rows.Item("ckd_itm_code").ToString.Trim) & "' and pack_key='" & gU.dbEncode(rows.Item("ckd_pack_key").ToString.Trim) & "'"

                                        gDB.amendData(ccSQL, gConn, transaction)

                                    Case "D"
                                        itemSQL = "delete from wms_stock_check_d " &
                                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                "and ck_code = '" & gU.dbEncode(CK_CODE.Text.Trim) & "' " &
                                                "and ckd_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("ckd_seq").ToString.Trim, "")) & "' "

                                        ccSQL = "Update WMS_ITEM set ITM_CC_DATE=NULL where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                       "and itm_code='" & gU.dbEncode(rows.Item("ckd_itm_code").ToString.Trim) & "' and pack_key='" & gU.dbEncode(rows.Item("ckd_pack_key").ToString.Trim) & "'"

                                        gDB.amendData(ccSQL, gConn, transaction)
                                    Case Else

                                        itemSQL = "update wms_stock_check_d set " &
                                                    "ckd_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_seq").ToString.Trim, ""))) & ", " &
                                                    "ckd_itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_itm_code").ToString.Trim, ""))) & ", " &
                                                    "ckd_pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_pack_key").ToString.Trim, ""))) & ", " &
                                                    "ckd_loc = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_loc").ToString.Trim, ""))) & ", " &
                                                    "ckd_org_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ckd_org_qty").ToString.Trim, "NULL")) & ", " &
                                                    "ckd_rev_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ckd_rev_qty").ToString.Trim, "NULL")) & ", " &
                                                    "ckd_var_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(varQty, "NULL")) & ", " &
                                                    "ckd_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_rem").ToString.Trim, ""))) & ", " &
                                                    "CKD_ADJ_REM = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_ADJ_REM").ToString.Trim, ""))) & ", " &
                                                    "ckd_pallet_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ckd_pallet_no").ToString.Trim, ""))) & ", " &
                                                    "ckd_expiry_date= " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ckd_expiry_date").ToString.Trim, ""))) & ", " &
                                                    "ckd_manu_date= " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("ckd_manu_date").ToString.Trim, ""))) & ", " &
                                                    "CKD_ACTUAL_QTY=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_ACTUAL_QTY").ToString.Trim, "NULL")) & ", " &
                                                    "CKD_TYPE=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_TYPE").ToString.Trim, ""))) & ", " &
                                                    "CKD_CC_LIST_NO=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_CC_LIST_NO").ToString.Trim, ""))) & ", " &
                                                    " CKD_FULL_DRUM=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_FULL_DRUM").ToString.Trim, ""))) & "," &
                                                    " CKD_REV_QTY2=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_REV_QTY2").ToString.Trim, "NULL")) & ", " &
                                                    " CKD_VAR_QTY2=" & gU.dbEncode(gU.decodeNullOrEmpty(varQty2, "NULL")) & ", " &
                                                    " CKD_ACTUAL_QTY2=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_ACTUAL_QTY2").ToString.Trim, "NULL")) & ", " &
                                                    " CKD_BOOK_QTY=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_BOOK_QTY").ToString.Trim, "NULL")) & ", " &
                                                    " CKD_BOOK_QTY2=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("CKD_BOOK_QTY2").ToString.Trim, "NULL")) & ", " &
                                                    " CKD_WITNESS=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_WITNESS").ToString.Trim, ""))) & ", " &
                                                    " CKD_CHECKER=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("CKD_CHECKER").ToString.Trim, ""))) & ", " &
                                                    "sys_lub = '" & Session("usr_id") & "', " &
                                                    "sys_lud = getdate() " &
                                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                "and ck_code = '" & gU.dbEncode(CK_CODE.Text.Trim) & "' " &
                                                "and ckd_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("old_seq").ToString.Trim, "")) & "'"
                                        '"CKD_CC_DATE=" & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("CKD_CC_DATE").ToString.Trim, ""))) & ", " & _
                                End Select
                                REM **********************

                                'Response.Write(itemSQL)
                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                            End If
                        Next
                    End If
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()
                successFlag = True
                If ViewState("pagemode") = "N" Then
                    ViewState("pagemode") = ""
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    CK_CODE.Text = nextNo
                    ViewState("CK_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    CK_CODE.ForeColor = Drawing.Color.Black
                    CK_CODE.Font.Size = 10
                    'CK_CODE.CssClass = ""
                    STORER_CODE.CssClass = ""
                    btnRelease.Visible = True
                    REM **********************
                End If

                If flag <> "Y" Then
                    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                End If


            Catch ex As Exception
                transaction.Rollback()
                Response.Write(sql_string & "<br><br>")
                Response.Write(itemSQL & "<br><br>")
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

            If successFlag Then BindGV()


        End If
        Return successFlag
    End Function

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"
        Dim WhereStr As String = ""
        Dim pk_code As String = ""
        Dim storerCode As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("CK_CODE") <> "" Then
            pk_code = ViewState("CK_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            ViewState("STORER_CODE") = Request("STORER_CODE")
            ViewState("CK_CODE") = Request("CK_CODE")

            pk_code = Server.UrlDecode(Request("CK_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))
        End If
        REM **********************

        DSP_CK_STATUS.ForeColor = Drawing.Color.Black

        If ViewState("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            CK_STATUS.Value = "NEW"
            CK_CODE.ForeColor = Drawing.Color.Red
            CK_TYPE.SelectedValue = "CC"
            IMP_CODE.Value = Session("imp_code")
            btnPrint.Visible = False
            btnPrint2.Visible = False
            btnPrint3.Visible = False
            btnRelease.Visible = False
            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = " SELECT IMP_CODE, STORER_CODE, CK_CODE, CK_STATUS, CK_TYPE, CK_REF_NO, convert(varchar,CK_DATE," & DDFORMAT & ") as ck_date, CK_BY, CK_RECHECK_BY, CK_WH, CK_REM, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, " &
                        " CK_SUB_WH, " &
                        " CK_LEVEL, CK_ITEM_LEVEL, CK_REF_AD_CODE, CK_GROUP_CODE, CK_GROUP_DESC,convert(varchar,CK_ACTUAL_DATE," & DDFORMAT & ") as CK_ACTUAL_DATE,convert(varchar,CK_START_DATE," & DDFORMAT & ") as CK_START_DATE,convert(varchar, CK_END_DATE," & DDFORMAT & ") as CK_END_DATE, CK_IS_CABLE " &
                        " FROM WMS_STOCK_CHECK " &
                        "where wms_stock_check.ck_code = '" & gU.dbEncode(pk_code) & "' " &
                        "and wms_stock_check.imp_code = '" & Session("IMP_CODE") & "' " &
                        "and wms_stock_check.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                btnPrint.Visible = True
                btnPrint2.Visible = True
                btnPrint3.Visible = True

                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                CK_CODE.Text = dt.Rows(0).Item("ck_code").ToString
                'STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                DSP_CK_STATUS.Text = gDB.getColValue(dt.Rows(0).Item("ck_STATUS").ToString, "WMS_STOCK_CHECK.CK_STATUS")
                CK_STATUS.Value = dt.Rows(0).Item("ck_STATUS").ToString
                If CK_STATUS.Value = "NEW" Then btnRelease.Visible = True

                If dt.Rows(0).Item("ck_TYPE").ToString <> "" Then CK_TYPE.SelectedValue = dt.Rows(0).Item("ck_TYPE").ToString
                CK_REF_NO.Text = dt.Rows(0).Item("ck_REF_NO").ToString
                CK_DATE.Text = dt.Rows(0).Item("ck_DATE").ToString
                CK_BY.SelectedValue = dt.Rows(0).Item("ck_BY").ToString
                CK_RECHECK_BY.SelectedValue = dt.Rows(0).Item("CK_recheck_BY").ToString
                CK_REM.Text = dt.Rows(0).Item("ck_REM").ToString

                CK_ACTUAL_DATE.Text = dt.Rows(0).Item("CK_ACTUAL_DATE").ToString
                ck_start_date.Text = dt.Rows(0).Item("CK_START_DATE").ToString
                ck_end_date.Text = dt.Rows(0).Item("CK_END_DATE").ToString
                'ck_period.SelectedValue = dt.Rows(0).Item("ck_period").ToString
                'ck_period,

                If dt.Rows(0).Item("CK_IS_CABLE").ToString.Trim = "Y" Then
                    CK_IS_Cable.Checked = True
                End If
                CK_IS_Cable.Enabled = False

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , , True)
                End If

                'CK_WH.SelectedValue = dt.Rows(0).Item("ck_WH").ToString
                uiFun.load_dropdown(CK_WH, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' and WH_MAIN_WH='" & dt.Rows(0).Item("CK_WH").ToString.Trim & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"), dt.Rows(0).Item("CK_WH").ToString.Trim, True)
                uiFun.load_dropdown(CK_SUB_WH, "Select wh_code, wh_name from wms_warehouse where imp_code='" & gU.dbEncode(Session("imp_code")) & "' AND WH_MAIN_WH='" & gU.dbEncode(dt.Rows(0).Item("CK_WH").ToString.Trim) & "' order by wh_name", "wh_code", "wh_name", , Session("gSelectLabel"), dt.Rows(0).Item("CK_SUB_WH").ToString.Trim)

                If CK_CODE.Text <> "" Then
                    'AD_CODE.ReadOnly = True
                    'AD_CODE.BorderWidth = 0
                    CK_CODE.BackColor = Drawing.Color.Transparent
                End If

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "SELECT wms_stock_check_d.IMP_CODE," &
                    " wms_stock_check_d.STORER_CODE,  wms_stock_check_d.CK_CODE, " &
                    " wms_stock_check_d.ckd_seq,  wms_stock_check_d.ckd_ITM_CODE, " &
                    " wms_stock_check_d.ckd_PACK_KEY, wms_stock_check_d.ckd_LOC,  wms_stock_check_d.ckd_REV_QTY, " &
                    " wms_stock_check_d.ckd_ORG_QTY, wms_stock_check_d.ckd_VAR_QTY,WMS_STOCK_CHECK_D.CKD_ACTUAL_QTY, " &
                    " wms_stock_check_d.ckd_REM,  wms_stock_check_d.ckd_PALLET_NO, " &
                    " wms_stock_check_d.ckd_BATCH_NO,  wms_stock_check_d.ckd_VND_CODE, " &
                    " wms_item.itm_sku_no,wms_item.itm_uom, wms_stock_check_d.CKD_STATUS, " &
                    " Convert(varchar, wms_stock_check_d.ckd_MANU_DATE,103) as ckd_MANU_DATE,  " &
                    " Convert(varchar, wms_stock_check_d.ckd_EXPIRY_DATE,103) as ckd_EXPIRY_DATE, " &
                    " WMS_STOCK_CHECK_D.CKD_CC_LIST_NO, WMS_STOCK_CHECK_D.CKD_SERIAL, WMS_STOCK_CHECK_D.CKD_TYPE, " &
                    " convert(varchar,WMS_STOCK_CHECK_D.CKD_CC_DATE," & DDFORMAT & ") as CKD_CC_DATE, WMS_STOCK_CHECK_D.CKD_FULL_DRUM, WMS_STOCK_CHECK_D.CKD_DRUM_ID, " &
                    " WMS_STOCK_CHECK_D.CKD_DRUM_LEVEL, WMS_STOCK_CHECK_D.CKD_SERIAL_NO, WMS_STOCK_CHECK_D.CKD_UOM2, " &
                    " WMS_STOCK_CHECK_D.CKD_ORG_QTY2, WMS_STOCK_CHECK_D.CKD_REV_QTY2, WMS_STOCK_CHECK_D.CKD_VAR_QTY2, WMS_STOCK_CHECK_D.CKD_ACTUAL_QTY2, " &
                    " WMS_STOCK_CHECK_D.CKD_BOOK_QTY, WMS_STOCK_CHECK_D.CKD_BOOK_QTY2, Convert(varchar,WMS_STOCK_CHECK_D.CKD_CC1_DATE," & DDFORMAT & ") as CKD_CC1_DATE," &
                    " wms_item.itm_name, 'U' as mFlag, wms_stock_check_d.ckd_seq as old_seq,wms_stock_check_d.CKD_RECHECK_BY, wms_stock_check_d.CKD_BY, wms_stock_check_d.CKD_AD_CODE,wms_stock_check_d.CKD_WITNESS,wms_stock_check_d.CKD_CHECKER,wms_stock_check_d.CKD_ADJ_REM," &
                    " CASE WHEN WMS_STOCK_CHECK_D.CKD_RECHECK_BY IS NULL THEN WMS_STOCK_CHECK_D.CKD_IN_PDA ELSE WMS_STOCK_CHECK_D.CKD_RECHK_IN_PDA END AS CKD_IN_PDA, " &
                    " WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE, " &
                    " '' as itm_name_textbox, '' as dsp_ckd_status,'' as from_imast " &
                    " from wms_stock_check_d left outer join wms_item on " &
                    " wms_stock_check_d.imp_code = wms_item.imp_code " &
                    " and wms_stock_check_d.storer_code = wms_item.storer_code " &
                    " and wms_stock_check_d.ckd_itm_code = wms_item.itm_code " &
                    " and wms_stock_check_d.ckd_PACK_KEY = wms_item.PACK_KEY " &
                    " LEFT OUTER JOIN WMS_WH_BIN on wms_stock_check_d.ckd_LOC = WMS_WH_BIN.LOC_KEY" &
                    " where wms_stock_check_d.ck_code = '" & gU.dbEncode(pk_code) & "' " &
                    " and WMS_WH_BIN.WH_CODE='" + CK_WH.SelectedValue.ToString + "' and wms_stock_check_d.storer_code = '" & gU.dbEncode(storerCode) & "' " &
                    " and wms_stock_check_d.imp_code = '" & Session("IMP_CODE") & "'"

        SQLString = SQLString & " order by Convert(decimal, wms_stock_check_d.ckd_seq)"
        REM **********************
        dt = gDB.getDataTable(SQLString)


        If dt.Rows.Count > 0 Then
            Dim selectedSEQ As String = ""

            For i = 0 To dt.Rows.Count - 1
                If dt.Rows(i).Item("ckd_ITM_CODE").ToString.Trim <> "" Then
                    selectedSEQ = gU.appendToList(selectedSEQ, dt.Rows(i).Item("STORER_CODE").ToString.Trim & "|*|" & dt.Rows(i).Item("ckd_ITM_CODE").ToString.Trim & "|*|" & dt.Rows(i).Item("ckd_PACK_KEY").ToString.Trim & "|*|" & dt.Rows(i).Item("ckd_PALLET_NO").ToString.Trim & "|*|" & dt.Rows(i).Item("ckd_BATCH_NO").ToString.Trim & "|*|" & dt.Rows(i).Item("ckd_LOC").ToString.Trim)
                End If
            Next

            Session("selectedSEQ") = selectedSEQ

            GridView1.DataSource = dt
            total_items.Text = dt.Rows.Count
        Else
            GridView1.DataSource = Nothing
        End If
        ViewState("dt") = dt
        GridView1.DataBind()

        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                GridView1.Rows(i).Visible = True
            Next
        End If

        REM **********************
    End Sub

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        Dim cancelSql As String = "update wms_stock_check " & _
                     "set ck_status = 'CANCELLED', " & _
                     "sys_lub = '" & Session("usr_id") & "', " & _
                     "sys_lud = getdate() " & _
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                     "and ck_code = '" & gU.dbEncode(CK_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction


        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            Dim selectSQL As String = "SELECT CKD_ITM_CODE, CKD_PACK_KEY FROM WMS_STOCK_CHECK_D " & _
                                      "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                      "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                      "and ck_code = '" & gU.dbEncode(CK_CODE.Text) & "' GROUP BY CKD_ITM_CODE, CKD_PACK_KEY"
            Dim tempDT As DataTable = gDB.getDataTable(selectSQL, gConn, transaction)

            Dim updateSQL As String = ""

            If tempDT IsNot Nothing AndAlso tempDT.Rows.Count > 0 Then
                For i = 0 To tempDT.Rows.Count - 1
                    updateSQL = "Update WMS_ITEM set ITM_CC_DATE=NULL where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and itm_code='" & gU.dbEncode(tempDT.Rows(i).Item("ckd_itm_code").ToString.Trim) & "' and pack_key='" & gU.dbEncode(tempDT.Rows(i).Item("ckd_pack_key").ToString.Trim) & "'"
                    gDB.amendData(updateSQL, gConn, transaction)
                Next

            End If

            transaction.Commit()

            CK_STATUS.Value = "CANCELLED"

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

    Protected Sub btnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPost.Click
        Dim updtSql As String
        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            If GridView1.Rows.Count > 0 Then
                Call save("Y")


                For Each rows As DataRow In dt.Rows
                    st.STORER_CODE = STORER_CODE.SelectedValue
                    st.ITM_CODE = gU.decodeNull(rows.Item("ckd_itm_code").ToString.Trim, "")
                    st.PACK_KEY = gU.decodeNull(rows.Item("ckd_pack_key").ToString.Trim, "")
                    st.IO_CUST_CODE = ""
                    st.IO_AREA = ""
                    st.IO_DOC = "SADJ"
                    st.IO_DOC_ID = CK_CODE.Text.Trim
                    st.IO_CBM = 0
                    st.IO_KG = 0
                    st.IO_WH = CK_WH.SelectedValue
                    st.PALLET_NO = gU.decodeNull(rows.Item("ckd_pallet_no").ToString.Trim, "")
                    st.IO_LOC = gU.decodeNull(rows.Item("ckd_loc").ToString.Trim, "")
                    st.lO_BATCH_NO = gU.decodeNull(rows.Item("ckd_batch_no").ToString.Trim, "")
                    st.IO_EXPIRY_DATE = gU.decodeNull(rows.Item("ckd_expiry_date").ToString.Trim, "")
                    st.IO_MANU_DATE = gU.decodeNull(rows.Item("ckd_manu_date").ToString.Trim, "")

                    Dim txQty As Integer = 0

                    If gU.decodeEmptyCInt(gU.decodeNull(rows.Item("ckd_rev_qty").ToString.Trim, ""), "0") > _
                        gU.decodeEmptyCInt(gU.decodeNull(rows.Item("ckd_org_qty").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCInt(gU.decodeNull(rows.Item("ckd_rev_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCInt(gU.decodeNull(rows.Item("ckd_org_qty").ToString.Trim, ""), "0")

                        st.IO_QTY = txQty
                        st.UpdateStockTrans("IN", gConn, transaction)
                        st.UpdateStockBalTrans("IN", gConn, transaction)

                    ElseIf gU.decodeEmptyCInt(gU.decodeNull(rows.Item("ckd_rev_qty").ToString.Trim, ""), "0") < _
                        gU.decodeEmptyCInt(gU.decodeNull(rows.Item("ckd_org_qty").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCInt(gU.decodeNull(rows.Item("ckd_org_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCInt(gU.decodeNull(rows.Item("ckd_rev_qty").ToString.Trim, ""), "0")

                        st.IO_QTY = txQty
                        st.UpdateStockTrans("OUT", gConn, transaction)
                        st.UpdateStockBalTrans("OUT", gConn, transaction)
                    ElseIf gU.decodeEmptyCInt(gU.decodeNull(rows.Item("ckd_rev_qty").ToString.Trim, ""), "0") = _
                        gU.decodeEmptyCInt(gU.decodeNull(rows.Item("ckd_org_qty").ToString.Trim, ""), "0") Then

                        If transaction IsNot Nothing Then
                            transaction.Rollback()
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Revised Qty Equal to Orginal Qty.\r\nYou have to adjust the Qty in order to POST!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "修訂數量和原來數量一樣。\r\你必須調整數量才可發布物件!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Next

                updtSql = "update wms_stock_check " & _
                            "set ck_status = 'POSTED', " & _
                            "sys_lub = N'" & Session("usr_id") & "', " & _
                            "sys_lud = getdate() " & _
                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            "and ck_code = '" & gU.dbEncode(CK_CODE.Text) & "' "

                gDB.amendData(updtSql)

                transaction.Commit()

                CK_STATUS.Value = "POSTED"
                ar.sec_viewMode = "Y"
                btnPost.Visible = False
                ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

                uiFun.displayMsg(Me, "1007", "", Session("gLang"))
            Else
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "No item can be posted!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "沒有可供發布的物件!", Session("gLang"))
                End If
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

    Protected Sub addItemtoSTCHK()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(ckd_seq AS int)) + 1 from wms_stock_check_d " & _
                                        "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                        "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and ck_code = '" & gU.dbEncode(CK_CODE.Text.Trim) & "' "
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
            Dim SQLStringS As String
            Dim sadj_dt As DataTable

            Dim itemPackList As String = ""
            Dim itemListarray As String()
            Dim packKeyListarray As String()
            Dim checkerListarray As String()

            Dim seqListarray As String()
            Dim seqKeyList As String = ""
            Dim seqSKeyList As String = ""
            Dim ListNoArray As String()


            itemListarray = Split(itemList.Value, ", ")
            packKeyListarray = Split(packKeyList.Value, ", ")
            seqListarray = Split(seqList.Value, ", ")
            ListNoArray = Split(ListNoList.Value, ", ")
            checkerListarray = Split(checkerList.Value, ", ")
            'If itemListarray.Count = 0 Then
            '    If itemList.Value <> "" Then
            '        itemPackList = Server.HtmlDecode(itemList.Value) & "_000_" & Server.HtmlDecode(packKeyList.Value)
            '    End If
            'Else
            '    For i = 0 To itemListarray.Count - 1
            '        itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemListarray(i)) & "_000_" & Server.HtmlDecode(packKeyListarray(i)))
            '    Next
            'End If

            Dim listNoDict As New Dictionary(Of String, String)

            If seqList.Value <> "" Then
                For i As Integer = 0 To seqListarray.Count - 1
                    If listNoDict.ContainsKey(seqListarray(i)) Then
                        listNoDict(seqListarray(i)) = CStr(gU.decodeEmptyCInt(ListNoArray(i), 1))
                    Else
                        listNoDict.Add(seqListarray(i), CStr(gU.decodeEmptyCInt(ListNoArray(i), 1)))
                    End If
                Next
            End If

            Dim checkerDict As New Dictionary(Of String, String)

            If checkerList.Value <> "" Then
                For i As Integer = 0 To seqListarray.Count - 1
                    If checkerDict.ContainsKey(seqListarray(i)) Then
                        checkerDict(seqListarray(i)) = checkerListarray(i)
                    Else
                        checkerDict.Add(seqListarray(i), checkerListarray(i))
                    End If
                Next
            End If

            If seqListarray.Count = 0 Then
                If seqList.Value <> "" Then
                    seqKeyList = Server.HtmlDecode(seqList.Value)
                End If
            Else
                For i = 0 To seqListarray.Count - 1
                    Dim tempStrArray As String()
                    tempStrArray = Split(seqListarray(i).ToString, "#_#")
                    If tempStrArray(1) <> "" Then
                        seqSKeyList = gU.appendToList(seqSKeyList, tempStrArray(1))
                    Else
                        seqKeyList = gU.appendToList(seqKeyList, tempStrArray(0))
                    End If
                Next
            End If

            'SQLString = "SELECT M.*, D.* " & _
            '"from WMS_ITEM M, WMS_ALT_VEND_ITEM D " & _
            '"where M.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
            '"and M.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
            '"AND M.ITM_CODE + '_000_' + M.PACK_KEY + '_000_' +  nvl(D.VND_CODE,'000') IN ('" & Replace(itemPackList, ", ", "', '") & "') " & _
            '"AND M.IMP_CODE = D.IMP_CODE " & _
            '"AND M.STORER_CODE = D.STORER_CODE " & _
            '"AND M.PACK_KEY = D.PACK_KEY " & _
            '"AND M.ITM_CODE = D.ITM_CODE"

            Dim addSQL As String = ""
            Dim addSSQL As String = ""
            REM non Serial
            If seqKeyList <> "" Then
                addSQL = "'" & Replace(seqKeyList, ", ", "', '") & "'"
            Else
                addSQL = "NULL"
            End If
            REM Serial
            If seqSKeyList <> "" Then
                addSSQL = "'" & Replace(seqSKeyList, ", ", "', '") & "'"
            Else
                addSSQL = "NULL"
            End If

            Dim orderbySQL As String = ""
            Dim start_year As String = ""
            If Not String.IsNullOrWhiteSpace(CK_ACTUAL_DATE.Text.Trim) Then
                start_year = Right(CK_ACTUAL_DATE.Text.Trim, 4)
            Else
                start_year = Now.Year
            End If

            'Dim PriceClass As String = ""
            'Select Case ck_period.SelectedValue
            '    Case "1YR"
            '        PriceClass = " AND loc_seq.CCLS_PERIOD='1YR' "
            '    Case "2YR"
            '        PriceClass = " AND loc_seq.CCLS_PERIOD='2YR' "
            '    Case "4YR"
            '        PriceClass = " AND loc_seq.CCLS_PERIOD='4YR' "
            '    Case Else
            '        PriceClass = " AND loc_seq.CCLS_PERIOD='ALL' "
            'End Select

            SQLString = " SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, " &
                        " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC, CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END AS ILOC_BAL_QTY, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, " &
                        " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO,WMS_ITEM.ITM_UOM, WMS_ITEM_LOC_BAL.ILOC_SEQ, WMS_ITEM_LOC_BAL_s.ILBS_SEQ, " &
                        " WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID,WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL, " &
                        " WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE, " &
                        " Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE, WMS_ITEM_LOC_BAL.VND_CODE " &
                        " FROM WMS_ITEM_LOC_BAL INNER JOIN " &
                        " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND " &
                        " WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " &
                        " LEFT OUTER JOIN WMS_WH_BIN ON WMS_ITEM_LOC_BAL.ILOC_LOC = WMS_WH_BIN.LOC_KEY And WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_BIN.WH_CODE" &
                        " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " &
                        " where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        " and wms_item.imp_code = '" & Session("IMP_CODE") & "' " &
                        " and WMS_ITEM_LOC_BAL.ILOC_SEQ in (" & addSQL & ") " &
                        " UNION " &
                        " SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, " &
                        " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC, CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END AS ILOC_BAL_QTY, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, " &
                        " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO,WMS_ITEM.ITM_UOM, WMS_ITEM_LOC_BAL.ILOC_SEQ, WMS_ITEM_LOC_BAL_s.ILBS_SEQ, " &
                        " WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID,WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL, " &
                        " WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE, " &
                        " Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE, WMS_ITEM_LOC_BAL.VND_CODE " &
                        " FROM WMS_ITEM_LOC_BAL INNER JOIN " &
                        " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND " &
                        " WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " &
                        " LEFT OUTER JOIN WMS_WH_BIN ON WMS_ITEM_LOC_BAL.ILOC_LOC = WMS_WH_BIN.LOC_KEY And WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_BIN.WH_CODE" &
                        " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " &
                        " where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        " and wms_item.imp_code = '" & Session("IMP_CODE") & "' " &
                        " and WMS_ITEM_LOC_BAL_s.ILBS_SEQ in (" & addSSQL & ") "
            '" and Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'') in (" & addSQL & ") "


            orderbySQL = " order by WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL.VND_CODE"

            SQLString = SQLString & orderbySQL

            REM **********************
            sadj_dt = gDB.getDataTable(SQLString, , , , , , , 600)
            Dim pl_list_no As String = ""
            Dim lchecker As String = ""
            Dim selectedSEQ As String = Session("selectedSEQ").ToString

            For i As Integer = 0 To sadj_dt.Rows.Count - 1
                If ViewState("n_cur_seq") = "" Then
                    ViewState("n_cur_seq") = next_seq_no
                Else
                    temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                    ViewState("n_cur_seq") = temp_seq_no.ToString
                End If

                dt.Rows.Add()

                rows_count = dt.Rows.Count

                If listNoDict.ContainsKey(sadj_dt.Rows(i).Item("ILOC_SEQ").ToString.Trim & "#_#" & sadj_dt.Rows(i).Item("ILBS_SEQ").ToString.Trim) Then
                    pl_list_no = listNoDict(sadj_dt.Rows(i).Item("ILOC_SEQ").ToString.Trim & "#_#" & sadj_dt.Rows(i).Item("ILBS_SEQ").ToString.Trim)
                End If

                If checkerDict.ContainsKey(sadj_dt.Rows(i).Item("ILOC_SEQ").ToString.Trim & "#_#" & sadj_dt.Rows(i).Item("ILBS_SEQ").ToString.Trim) Then
                    lchecker = checkerDict(sadj_dt.Rows(i).Item("ILOC_SEQ").ToString.Trim & "#_#" & sadj_dt.Rows(i).Item("ILBS_SEQ").ToString.Trim)
                End If

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("ckd_seq") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("ckd_itm_code") = sadj_dt.Rows(i).Item("itm_code").ToString.Trim
                dt.Rows(rows_count - 1).Item("itm_sku_no") = sadj_dt.Rows(i).Item("itm_sku_no").ToString.Trim
                dt.Rows(rows_count - 1).Item("itm_name") = sadj_dt.Rows(i).Item("itm_name").ToString.Trim
                dt.Rows(rows_count - 1).Item("itm_uom") = sadj_dt.Rows(i).Item("itm_uom").ToString.Trim
                dt.Rows(rows_count - 1).Item("ckd_pack_key") = sadj_dt.Rows(i).Item("pack_key").ToString.Trim
                dt.Rows(rows_count - 1).Item("ckd_PALLET_NO") = sadj_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("ckd_BATCH_NO") = sadj_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("ckd_org_qty") = sadj_dt.Rows(i).Item("ILOC_BAL_QTY")
                dt.Rows(rows_count - 1).Item("CKD_BOOK_QTY") = sadj_dt.Rows(i).Item("ILOC_BAL_QTY")
                dt.Rows(rows_count - 1).Item("ckd_LOC") = sadj_dt.Rows(i).Item("ILOC_LOC").ToString.Trim
                dt.Rows(rows_count - 1).Item("CKD_DRUM_ID") = sadj_dt.Rows(i).Item("ILBS_DRUM_ID").ToString.Trim
                dt.Rows(rows_count - 1).Item("CKD_DRUM_LEVEL") = sadj_dt.Rows(i).Item("ILBS_DRUM_LEVEL").ToString.Trim
                dt.Rows(rows_count - 1).Item("CKD_SERIAL_NO") = sadj_dt.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("CKD_UOM2") = sadj_dt.Rows(i).Item("ILBS_UOM2").ToString.Trim
                dt.Rows(rows_count - 1).Item("CKD_ORG_QTY2") = sadj_dt.Rows(i).Item("ILBS_QTY2")
                dt.Rows(rows_count - 1).Item("CKD_BOOK_QTY2") = sadj_dt.Rows(i).Item("ILBS_QTY2")

                dt.Rows(rows_count - 1).Item("ckd_VND_CODE") = sadj_dt.Rows(i).Item("VND_CODE").ToString.Trim
                dt.Rows(rows_count - 1).Item("ckd_MANU_DATE") = sadj_dt.Rows(i).Item("ILOC_MANU_DATE").ToString.Trim
                dt.Rows(rows_count - 1).Item("ckd_EXPIRY_DATE") = sadj_dt.Rows(i).Item("ILOC_EXPIRY_DATE").ToString.Trim
                dt.Rows(rows_count - 1).Item("CKD_STATUS") = "NEW"
                dt.Rows(rows_count - 1).Item("CKD_CC_LIST_NO") = pl_list_no
                dt.Rows(rows_count - 1).Item("CKD_CHECKER") = lchecker

                dt.Rows(rows_count - 1).Item("BN_CSMS_CODE") = sadj_dt.Rows(i).Item("BN_CSMS_CODE").ToString.Trim
                dt.Rows(rows_count - 1).Item("WH_CODE") = sadj_dt.Rows(i).Item("WH_CODE").ToString.Trim
                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

                If dt.Rows(i).Item("ckd_ITM_CODE").ToString.Trim <> "" Then
                    selectedSEQ = gU.appendToList(selectedSEQ, STORER_CODE.SelectedValue & "|*|" & sadj_dt.Rows(i).Item("itm_code").ToString.Trim & "|*|" & sadj_dt.Rows(i).Item("pack_key").ToString.Trim & "|*|" & sadj_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim & "|*|" & sadj_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim & "|*|" & sadj_dt.Rows(i).Item("ILOC_LOC").ToString.Trim)
                End If

            Next

            Session("selectedSEQ") = selectedSEQ

            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()

            If dt.Rows.Count > 0 Then
                CK_IS_Cable.Enabled = False

                Dim whCode As String = CK_WH.SelectedValue
                uiFun.load_dropdown(CK_WH, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' and WH_MAIN_WH='" & gU.dbEncode(whCode) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"), whCode, True)

                Dim storerCode As String = STORER_CODE.SelectedValue
                uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER where storer_code='" & gU.dbEncode(storerCode) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"), storerCode, True)

            End If

        End If
    End Sub

    Protected Sub addItemMast()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True, {"Y", "N"}) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(ckd_seq AS int)) + 1 from wms_stock_check_d " & _
                                       "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                       "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                       "and ck_code = '" & gU.dbEncode(CK_CODE.Text.Trim) & "' "
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
            Dim sadj_dt As DataTable

            Dim itemPackList As String = ""
            Dim itemListarray As String()

            Dim packKeyListarray As String()
            Dim checkerListarray As String()
            Dim ListNoArray As String()

            Dim seqListarray As String()
            Dim seqKeyList As String = ""

            itemListarray = Split(itemList.Value, ", ")
            packKeyListarray = Split(packKeyList.Value, ", ")
            seqListarray = Split(seqList.Value, ", ")
            ListNoArray = Split(ListNoList.Value, ", ")
            checkerListarray = Split(checkerList.Value, ", ")

            If itemListarray.Count = 0 Then
                If itemList.Value <> "" Then
                    itemPackList = Server.HtmlDecode(itemList.Value) & "_000_" & Server.HtmlDecode(packKeyList.Value)
                End If
            Else
                For i = 0 To itemListarray.Count - 1
                    itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemListarray(i)) & "_000_" & Server.HtmlDecode(packKeyListarray(i)))
                Next
            End If

            If seqListarray.Count = 0 Then
                If seqList.Value <> "" Then
                    seqKeyList = Server.HtmlDecode(seqList.Value)
                End If
            Else
                For i = 0 To seqListarray.Count - 1
                    seqKeyList = gU.appendToList(seqKeyList, seqListarray(i))
                Next
            End If

            Dim listNoDict As New Dictionary(Of String, String)

            If seqList.Value <> "" Then
                For i As Integer = 0 To seqListarray.Count - 1
                    If listNoDict.ContainsKey(seqListarray(i)) Then
                        listNoDict(seqListarray(i)) = CStr(gU.decodeEmptyCInt(ListNoArray(i), 1))
                    Else
                        listNoDict.Add(seqListarray(i), CStr(gU.decodeEmptyCInt(ListNoArray(i), 1)))
                    End If
                Next
            End If

            Dim checkerDict As New Dictionary(Of String, String)

            If checkerList.Value <> "" Then
                For i As Integer = 0 To seqListarray.Count - 1
                    If checkerDict.ContainsKey(seqListarray(i)) Then
                        checkerDict(seqListarray(i)) = checkerListarray(i)
                    Else
                        checkerDict.Add(seqListarray(i), checkerListarray(i))
                    End If
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
            sadj_dt = gDB.getDataTable(SQLString)

            Dim pl_list_no As String = ""
            Dim lchecker As String = ""

            For i As Integer = 0 To sadj_dt.Rows.Count - 1
                If ViewState("n_cur_seq") = "" Then
                    ViewState("n_cur_seq") = next_seq_no
                Else
                    temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                    ViewState("n_cur_seq") = temp_seq_no.ToString
                End If

                dt.Rows.Add()

                rows_count = dt.Rows.Count

                Dim itmQty As Integer = 0


                If listNoDict.ContainsKey(sadj_dt.Rows(i).Item("itm_code").ToString.Trim & "#_#" & sadj_dt.Rows(i).Item("pack_key").ToString.Trim & "#_#" & gU.decodeNullOrEmpty(sadj_dt.Rows(i).Item("vnd_code").ToString.Trim, "DEF_VEND")) Then
                    pl_list_no = listNoDict(sadj_dt.Rows(i).Item("itm_code").ToString.Trim & "#_#" & sadj_dt.Rows(i).Item("pack_key").ToString.Trim & "#_#" & gU.decodeNullOrEmpty(sadj_dt.Rows(i).Item("vnd_code").ToString.Trim, "DEF_VEND"))
                End If

                If checkerDict.ContainsKey(sadj_dt.Rows(i).Item("itm_code").ToString.Trim & "#_#" & sadj_dt.Rows(i).Item("pack_key").ToString.Trim & "#_#" & gU.decodeNullOrEmpty(sadj_dt.Rows(i).Item("vnd_code").ToString.Trim, "DEF_VEND")) Then
                    lchecker = checkerDict(sadj_dt.Rows(i).Item("itm_code").ToString.Trim & "#_#" & sadj_dt.Rows(i).Item("pack_key").ToString.Trim & "#_#" & gU.decodeNullOrEmpty(sadj_dt.Rows(i).Item("vnd_code").ToString.Trim, "DEF_VEND"))
                End If

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("ckd_seq") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("ckd_itm_code") = sadj_dt.Rows(i).Item("itm_code").ToString.Trim
                dt.Rows(rows_count - 1).Item("itm_sku_no") = sadj_dt.Rows(i).Item("itm_sku_no").ToString.Trim
                dt.Rows(rows_count - 1).Item("itm_name") = sadj_dt.Rows(i).Item("itm_name").ToString.Trim
                dt.Rows(rows_count - 1).Item("itm_uom") = sadj_dt.Rows(i).Item("itm_uom").ToString.Trim
                dt.Rows(rows_count - 1).Item("ckd_pack_key") = sadj_dt.Rows(i).Item("pack_key").ToString.Trim
                dt.Rows(rows_count - 1).Item("ckd_PALLET_NO") = "000"
                dt.Rows(rows_count - 1).Item("ckd_BATCH_NO") = ""
                dt.Rows(rows_count - 1).Item("ckd_org_qty") = 0
                dt.Rows(rows_count - 1).Item("CKD_BOOK_QTY") = 0
                dt.Rows(rows_count - 1).Item("ckd_LOC") = ""
                dt.Rows(rows_count - 1).Item("from_imast") = "Y"

                'dt.Rows(rows_count - 1).Item("CKD_DRUM_ID") = sadj_dt.Rows(i).Item("ILBS_DRUM_ID").ToString.Trim
                'dt.Rows(rows_count - 1).Item("CKD_DRUM_LEVEL") = sadj_dt.Rows(i).Item("ILBS_DRUM_LEVEL").ToString.Trim
                'dt.Rows(rows_count - 1).Item("CKD_SERIAL_NO") = sadj_dt.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim
                'dt.Rows(rows_count - 1).Item("CKD_UOM2") = sadj_dt.Rows(i).Item("ILBS_UOM2").ToString.Trim
                'dt.Rows(rows_count - 1).Item("CKD_ORG_QTY2") = sadj_dt.Rows(i).Item("ILBS_QTY2")
                'dt.Rows(rows_count - 1).Item("CKD_BOOK_QTY2") = sadj_dt.Rows(i).Item("ILBS_QTY2")

                dt.Rows(rows_count - 1).Item("ckd_VND_CODE") = sadj_dt.Rows(i).Item("VND_CODE").ToString.Trim
                dt.Rows(rows_count - 1).Item("ckd_MANU_DATE") = ""
                dt.Rows(rows_count - 1).Item("ckd_EXPIRY_DATE") = ""

                If CK_STATUS.Value = "NEW" OrElse CK_STATUS.Value = "" Then
                    dt.Rows(rows_count - 1).Item("CKD_STATUS") = "NEW"
                Else
                    dt.Rows(rows_count - 1).Item("CKD_STATUS") = "READY"
                End If


                dt.Rows(rows_count - 1).Item("CKD_CC_LIST_NO") = pl_list_no
                dt.Rows(rows_count - 1).Item("CKD_CHECKER") = lchecker
                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

            Next
            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub


    Protected Sub btnGen_Click(sender As Object, e As System.EventArgs) Handles btnGen.Click

        pnlSC_ModalPopupExtender.Show()

        If validateStockAdj() Then
            If GridView1.Rows.Count > 0 Then
                Dim selectSQL As String = "select ckd_itm_code,'' as itm_sku_no, '' as itm_name, ckd_batch_no, ckd_loc,'' as ckd_org_qty,'' as ckd_rev_qty,'' as ckd_var_qty,'' as ckd_org_qty2,'' as ckd_rev_qty2,'' as ckd_var_qty2 from wms_stock_check_d where 1=3"
                Dim tempDT As DataTable

                tempDT = gDB.getDataTable(selectSQL)

                Dim newRow As DataRow

                For i = 0 To GridView1.Rows.Count - 1
                    If DirectCast(GridView1.Rows(i).FindControl("mFlag"), HiddenField).Value <> "D" AndAlso DirectCast(GridView1.Rows(i).FindControl("ckd_itm_code"), HiddenField).Value <> "" _
                        AndAlso ((DirectCast(GridView1.Rows(i).FindControl("ckd_var_qty"), HiddenField).Value <> "" AndAlso CDbl(DirectCast(GridView1.Rows(i).FindControl("ckd_var_qty"), HiddenField).Value)) <> 0 _
                        OrElse (DirectCast(GridView1.Rows(i).FindControl("ckd_var_qty2"), HiddenField).Value <> "" AndAlso CDbl(DirectCast(GridView1.Rows(i).FindControl("ckd_var_qty2"), HiddenField).Value))) <> 0 _
                        AndAlso DirectCast(GridView1.Rows(i).FindControl("CKD_STATUS"), HiddenField).Value = "CNTD" _
                        Then

                        newRow = tempDT.NewRow

                        'newRow.Item("ckd_itm_code") = DirectCast(GridView1.Rows(i).FindControl("ckd_itm_code"), HiddenField).Value
                        newRow.Item("itm_sku_no") = DirectCast(GridView1.Rows(i).FindControl("itm_sku_no"), Label).Text
                        newRow.Item("itm_name") = DirectCast(GridView1.Rows(i).FindControl("itm_name"), Label).Text
                        newRow.Item("ckd_batch_no") = DirectCast(GridView1.Rows(i).FindControl("ckd_batch_no"), HiddenField).Value
                        newRow.Item("ckd_loc") = DirectCast(GridView1.Rows(i).FindControl("ckd_loc"), HiddenField).Value
                        newRow.Item("ckd_org_qty") = DirectCast(GridView1.Rows(i).FindControl("CKD_BOOK_QTY"), HiddenField).Value
                        newRow.Item("ckd_rev_qty") = DirectCast(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY"), TextBox).Text
                        newRow.Item("ckd_var_qty") = DirectCast(GridView1.Rows(i).FindControl("CKD_VAR_QTY"), HiddenField).Value
                        newRow.Item("ckd_org_qty2") = DirectCast(GridView1.Rows(i).FindControl("CKD_BOOK_QTY2"), HiddenField).Value
                        newRow.Item("ckd_rev_qty2") = DirectCast(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY2"), TextBox).Text
                        newRow.Item("ckd_var_qty2") = DirectCast(GridView1.Rows(i).FindControl("CKD_VAR_QTY2"), HiddenField).Value

                        tempDT.Rows.Add(newRow)
                    End If
                Next

                tempDT.AcceptChanges()

                If tempDT.Rows.Count > 0 Then
                    GVAdj.DataSource = tempDT
                    GVAdj.DataBind()

                Else
                    uiFun.displayMsg(Me, "", "No stock adjust is needed to be generated.", Session("gLang"))
                    pnlSC_ModalPopupExtender.Hide()

                End If

            Else
                uiFun.displayMsg(Me, "", "No Item Selected!", Session("gLang"))
                pnlSC_ModalPopupExtender.Hide()
            End If
        Else
            pnlSC_ModalPopupExtender.Hide()
        End If
    End Sub

    Protected Sub GenStockAdj()
        Dim selectSQL As String = ""
        Dim tempDT As DataTable
        Dim updateSQL As String = ""
        Dim sadj_no As String = ""
        Dim successFlag As Boolean = False
        If GridView1.Rows.Count > 0 Then
            If save("Y") Then
                Dim updtSql As String = ""
                Dim gConn As SqlConnection

                gConn = gDB.getConnection()
                Dim transaction As SqlTransaction

                transaction = gConn.BeginTransaction()

                Try

                    selectSQL = " Select *,Convert(varchar, ckd_EXPIRY_DATE,112) as exp_date, Convert(varchar, ckd_MANU_DATE,112) as manu_date from wms_stock_check_d where IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code='" & gU.dbEncode(ViewState("STORER_CODE")) & "' and ck_code='" & gU.dbEncode(ViewState("CK_CODE")) & "' and ((CKD_VAR_QTY <> 0 and CKD_VAR_QTY is not null) or (CKD_VAR_QTY2 <> 0 and CKD_VAR_QTY2 is not null)) and (ckd_itm_code is not null and ckd_itm_code <> '') and ckd_status='CNTD'"

                    tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                    If tempDT.Rows.Count > 0 Then
                        sadj_no = DB.getDocNo("SADJ", gConn, transaction)

                        sadj_no = Now.Year.ToString & Now.Month.ToString("d2") & sadj_no

                        updateSQL = "Insert into wms_stock_adjust( wms_stock_adjust.IMP_CODE,  wms_stock_adjust.STORER_CODE,  wms_stock_adjust.AD_CODE,  wms_stock_adjust.AD_STATUS," & _
                                    " wms_stock_adjust.AD_DATE,  wms_stock_adjust.AD_BY,  wms_stock_adjust.AD_WH, " & _
                                    " wms_stock_adjust.SYS_LUB,  wms_stock_adjust.SYS_LUD,  wms_stock_adjust.SYS_CD,  wms_stock_adjust.SYS_CB, " & _
                                    " wms_stock_adjust.AD_REF_CK_CODE,wms_stock_adjust.AD_IS_CABLE) " & _
                                    "(Select IMP_CODE, STORER_CODE, '" & sadj_no & "','NEW', getdate(), ck_by, ck_wh," & _
                                    " '" & Session("usr_id") & "',getdate(),getdate(),'" & Session("usr_id") & "',ck_code, CK_IS_CABLE from wms_stock_check where imp_code='" & gU.dbEncode(Session("IMP_CODE")) & "' AND storer_code='" & gU.dbEncode(ViewState("STORER_CODE")) & "' and ck_code='" & gU.dbEncode(ViewState("CK_CODE")) & "'" & _
                                    ")"

                        gDB.amendData(updateSQL, gConn, transaction)

                        For i = 0 To tempDT.Rows.Count - 1
                            updateSQL = "insert into wms_stock_adjust_d (IMP_CODE,  STORER_CODE,  AD_CODE,  AD_SEQ,  ADD_ITM_CODE, " & _
                                        " ADD_PACK_KEY,  ADD_LOC,  ADD_ORG_QTY,  ADD_REV_QTY,  ADD_VAR_QTY, " & _
                                        " ADD_REM,  ADD_PALLET_NO, " & _
                                        " SYS_LUB,  SYS_LUD,  SYS_CD,  SYS_CB, " & _
                                        " ADD_ORG_QTY2, ADD_REV_QTY2, ADD_VAR_QTY2, ADD_SERIAL_NO, " & _
                                        " ADD_BATCH_NO,  ADD_VND_CODE,  ADD_EXPIRY_DATE,  ADD_MANU_DATE) values (" & _
                                        "'" & gU.dbEncode(Session("IMP_CODE")) & "','" & gU.dbEncode(ViewState("STORER_CODE")) & "','" & gU.dbEncode(sadj_no) & "','" & i + 1 & "'," & _
                                        "'" & gU.dbEncode(tempDT.Rows(i).Item("CKD_ITM_CODE").ToString.Trim) & "','" & gU.dbEncode(tempDT.Rows(i).Item("CKD_PACK_KEY").ToString.Trim) & "','" & gU.dbEncode(tempDT.Rows(i).Item("CKD_LOC").ToString.Trim) & "'," & _
                                        gU.dbEncode(gU.decodeNullOrEmpty(tempDT.Rows(i).Item("CKD_BOOK_QTY").ToString.Trim, "NULL")) & "," & gU.dbEncode(gU.decodeNullOrEmpty(tempDT.Rows(i).Item("CKD_ACTUAL_QTY").ToString.Trim, "NULL")) & "," & gU.dbEncode(gU.decodeNullOrEmpty(tempDT.Rows(i).Item("CKD_VAR_QTY").ToString.Trim, "NULL")) & "," & _
                                        "'" & gU.dbEncode(tempDT.Rows(i).Item("CKD_REM").ToString.Trim) & "','" & gU.dbEncode(tempDT.Rows(i).Item("CKD_PALLET_NO").ToString.Trim) & "'," & _
                                        "'" & gU.dbEncode(Session("usr_id")) & "',getdate(),getdate(),'" & gU.dbEncode(Session("usr_id")) & "'," & _
                                        gU.dbEncode(gU.decodeNullOrEmpty(tempDT.Rows(i).Item("CKD_BOOK_QTY2").ToString.Trim, "NULL")) & "," & gU.dbEncode(gU.decodeNullOrEmpty(tempDT.Rows(i).Item("CKD_ACTUAL_QTY2").ToString.Trim, "NULL")) & "," & gU.dbEncode(gU.decodeNullOrEmpty(tempDT.Rows(i).Item("CKD_VAR_QTY2").ToString.Trim, "NULL")) & "," & _
                                        gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CKD_SERIAL_NO").ToString.Trim)) & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CKD_BATCH_NO").ToString.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CKD_VND_CODE").ToString.Trim)) & "," & _
                                        gU.convdbDate(gU.dbEncode(tempDT.Rows(i).Item("exp_date").ToString.Trim)) & "," & gU.convdbDate(gU.dbEncode(tempDT.Rows(i).Item("manu_date").ToString.Trim)) & _
                                        ") "

                            gDB.amendData(updateSQL, gConn, transaction)

                            updateSQL = "update wms_stock_check_d set ckd_status='ADJ', sys_lud=getdate(), SYS_LUB='" & gU.dbEncode(Session("usr_id")) & "', CKD_AD_CODE='" & gU.dbEncode(sadj_no) & "' " & _
                                        " where IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code='" & gU.dbEncode(ViewState("STORER_CODE")) & "' and ck_code='" & gU.dbEncode(ViewState("CK_CODE")) & "' AND ckd_seq='" & tempDT.Rows(i).Item("ckd_seq").ToString.Trim & "'"
                            gDB.amendData(updateSQL, gConn, transaction)
                        Next

                        'update WMS_STOCK_CHECK set ck_status='ADJUSTED' where CK_CODE='000016' and not exists(select 1 from WMS_STOCK_CHECK_D where  ((CKD_VAR_QTY <> 0 and CKD_VAR_QTY is not null) or (CKD_VAR_QTY2 <> 0 and CKD_VAR_QTY2 is not null)) and (ckd_itm_code is not null and ckd_itm_code <> '') and ckd_status='CNTD' and isnull(ckd_status,'N') <> 'ADJ' and ck_code='000016')

                        updateSQL = "update wms_stock_check " & _
                                    "set ck_status = 'ADJUSTED', " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = getdate() " & _
                                    "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                    "and storer_code = '" & gU.dbEncode(ViewState("STORER_CODE")) & "' " & _
                                    "and ck_code = '" & gU.dbEncode(ViewState("CK_CODE")) & "' " & _
                                    "and not exists (select 1 from WMS_STOCK_CHECK_D where ((CKD_VAR_QTY <> 0 and CKD_VAR_QTY is not null) or (CKD_VAR_QTY2 <> 0 and CKD_VAR_QTY2 is not null)) and (ckd_itm_code is not null and ckd_itm_code <> '') and ckd_status='CNTD' and isnull(ckd_status,'N') <> 'ADJ' " & _
                                    "and imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                    "and storer_code = '" & gU.dbEncode(ViewState("STORER_CODE")) & "' " & _
                                    "and ck_code = '" & gU.dbEncode(ViewState("CK_CODE")) & "' " & _
                                    ")"

                        gDB.amendData(updateSQL, gConn, transaction)

                        successFlag = True
                        transaction.Commit()
                    Else

                        uiFun.displayMsg(Me, "", "All verified item qty. is same with original item qty. No stock adjust will be generated", Session("gLang"))
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
            End If
        End If

        If successFlag Then
            uiFun.displayMsg(Me, "", "Stock Adjust Record (No." & sadj_no & ") has been generated.", Session("gLang"))
            Call BindGV()
            If CK_STATUS.Value = "ADJUSTED" Then
                ar.sec_viewMode = "Y"
                ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
            End If
        End If
    End Sub

    Protected Sub btnPnlSave_Click(sender As Object, e As System.EventArgs) Handles btnPnlSave.Click
        GenStockAdj()
    End Sub

    'Protected Sub ck_period_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ck_period.SelectedIndexChanged
    '    If Not String.IsNullOrWhiteSpace(ck_period.SelectedValue) Then
    '        Dim FR_DATE, TO_DATE As Date
    '        Dim s_date As String = ""
    '        Dim rt_Date As String = ""
    '        Dim period As String = ""

    '        Dim customDateTimeFormat As System.Globalization.DateTimeFormatInfo = New System.Globalization.DateTimeFormatInfo()
    '        customDateTimeFormat.DateSeparator = "/"
    '        customDateTimeFormat.TimeSeparator = ":"
    '        customDateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
    '        customDateTimeFormat.LongDatePattern = "dd/MM/yyyy"
    '        customDateTimeFormat.ShortTimePattern = "HH:mm"
    '        customDateTimeFormat.LongTimePattern = "HH:mm"
    '        customDateTimeFormat.FullDateTimePattern = "dd/MM/yyyy HH:mm"

    '        Dim sqlString As String = ""

    '        Select Case ck_period.SelectedValue
    '            Case "1YR"
    '                period = "1"

    '            Case "2YR"
    '                period = "2"

    '            Case "4YR"
    '                period = "4"

    '        End Select


    '        sqlString = "select convert(varchar, CCS_END_DATE," & DDFORMAT & ") as CCS_END_DATE from wms_cc_setup " & _
    '                    " Where CCS_YEAR='" & Now.Year & "' and ccs_period='" & gU.dbEncode(period) & "'"

    '        s_date = gU.decodeNullOrEmpty(DB.getValueFromSQL(sqlString), "")

    '        If s_date <> "" Then
    '            TO_DATE = Convert.ToDateTime(s_date, customDateTimeFormat)

    '            Select Case ck_period.SelectedValue
    '                Case "1YR"
    '                    FR_DATE = TO_DATE.AddYears(-1).AddDays(1)
    '                    rt_Date = FR_DATE.ToString("dd/MM/yyyy")
    '                Case "2YR"
    '                    FR_DATE = TO_DATE.AddYears(-2).AddDays(1)
    '                    rt_Date = FR_DATE.ToString("dd/MM/yyyy")
    '                Case "4YR"
    '                    FR_DATE = TO_DATE.AddYears(-4).AddDays(1)
    '                    rt_Date = FR_DATE.ToString("dd/MM/yyyy")
    '                Case Else
    '                    rt_Date = ""
    '            End Select

    '            ck_end_date.Text = s_date
    '            CK_DATE.Text = rt_Date

    '        End If

    '        'FR_DATE = Convert.ToDateTime(s_date, customDateTimeFormat)

    '        'Select Case ck_period.SelectedValue
    '        '    Case "1YR"
    '        '        TO_DATE = FR_DATE.AddYears(1).AddDays(-1)
    '        '        rt_Date = TO_DATE.ToString("dd/MM/yyyy")
    '        '    Case "2YR"
    '        '        TO_DATE = FR_DATE.AddYears(2).AddDays(-1)
    '        '        rt_Date = TO_DATE.ToString("dd/MM/yyyy")
    '        '    Case "4YR"
    '        '        TO_DATE = FR_DATE.AddYears(4).AddDays(-1)
    '        '        rt_Date = TO_DATE.ToString("dd/MM/yyyy")
    '        '    Case Else
    '        '        rt_Date = ""
    '        'End Select

    '        'ck_end_date.Text = rt_Date
    '    End If


    'End Sub

    Protected Sub selectItemBtn_Click(sender As Object, e As System.EventArgs) Handles selectItemBtn.Click

        Session("CK_TYPE") = CK_TYPE.SelectedValue
        Session("CK_START_DATE") = CK_ACTUAL_DATE.Text.Trim
        Session("CK_WH") = CK_WH.SelectedValue
        Session("LookupWH") = CK_SUB_WH.SelectedValue

        Dim priceClass As String = ""

        'Select Case ck_period.SelectedValue
        '    Case "1YR"
        '        priceClass = "HI"
        '    Case "2YR"
        '        priceClass = "MID"
        '    Case "4YR"
        '        priceClass = "LOW"
        'End Select

        'Session("CK_PERIOD") = priceClass

        If CK_IS_Cable.Checked OrElse CK_TYPE.SelectedValue = "CCC" Then
            Session("CK_IS_Cable") = "Y"
        Else
            Session("CK_IS_Cable") = ""
        End If

        If chkITM.Checked Then
            ScriptManager.RegisterStartupScript(LOOKUPDUP, LOOKUPDUP.GetType, "itemLookup", "ItemMastLookUp(document.myform." & STORER_CODE.ClientID & ".value);", True)
        Else
            ScriptManager.RegisterStartupScript(LOOKUPDUP, LOOKUPDUP.GetType, "itemLookup", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value,document.getElementById('CK_WH').value);", True)
        End If



    End Sub

    Protected Sub btnRelease_Click(sender As Object, e As System.EventArgs) Handles btnRelease.Click
        Dim pk_code As String = ViewState("CK_CODE")
        Dim successflag As Boolean = False
        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction


        transaction = gConn.BeginTransaction()

        Try
            Dim cancelSql As String = "update WMS_STOCK_CHECK " & _
                     "set CK_STATUS = 'READY', " & _
                     "sys_lub = '" & Session("usr_id") & "', " & _
                     "sys_lud = Getdate() " & _
                     "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                     "and CK_CODE = '" & gU.dbEncode(pk_code) & "' "

            gDB.amendData(cancelSql, gConn, transaction)

            Dim updateSQL As String = "update WMS_STOCK_CHECK_D " & _
                                      "set CKD_STATUS = 'READY', " & _
                                      "sys_lub = '" & Session("usr_id") & "', " & _
                                      "sys_lud = Getdate() " & _
                                      "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                      "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                      "and CK_CODE = '" & gU.dbEncode(pk_code) & "' "

            gDB.amendData(updateSQL, gConn, transaction)


            transaction.Commit()

            CK_STATUS.Value = "READY"

            setPageCtrlAccess()
            'ar.sec_viewMode = "Y"
            btnRelease.Visible = False
            btnPrint.Visible = True
            btnPrint2.Visible = True
            btnPrint3.Visible = True
            btnCnt.Visible = True
            'exceptionEditList.Add("btnRechk")
            'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)

            'uiFun.displayMsg(Me, "", "This Cycle Count is ready for PDA Checking.", Session("gLang"))
            successflag = True

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

        If successflag Then
            ScriptManager.RegisterStartupScript(udp1, udp1.GetType, "RELOAD_PAGE", "reloadPage('This Cycle Count is ready for PDA Checking.');", True)
        End If
    End Sub


    Private Sub setPageCtrlAccess()

        exceptionEditList = New List(Of String)

        exceptionEditList.Add("btnPrint")
        exceptionEditList.Add("btnPrint2")
        exceptionEditList.Add("btnPrint3")

    End Sub

    Protected Sub addBlank_Click(sender As Object, e As System.EventArgs) Handles addBlank.Click
        Dim tempDT As DataTable
        Dim next_seq_no As String
        Dim temp_seq_no As Integer = 1


        REM **********************


        tempDT = ViewState("dt")

        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

            If ViewState("n_cur_seq") = "" Then

                next_seq_no = gU.decodeNullOrEmpty(DB.getValueFromSQL("select MAX(CAST(CKD_SEQ AS int)) + 1 from WMS_STOCK_CHECK_D " & _
                                      "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                      "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                      "and CK_CODE = '" & gU.dbEncode(CK_CODE.Text.Trim) & "' "), 1)


                ViewState("n_cur_seq") = next_seq_no
            Else
                temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                ViewState("n_cur_seq") = temp_seq_no.ToString
            End If

            Dim newRow As DataRow

            newRow = tempDT.NewRow()
            newRow.Item("ckd_seq") = ViewState("n_cur_seq").ToString
            newRow.Item("mFlag") = "N"

            tempDT.Rows.Add(newRow)
            tempDT.AcceptChanges()

            ViewState("dt") = tempDT
            GridView1.DataSource = tempDT
            GridView1.DataBind()

            If Not CK_IS_Cable.Checked Then
                If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
                    For i = 20 To 28
                        GridView1.Columns(i).Visible = False
                    Next
                End If
            End If
        End If
    End Sub

    Protected Function validateStockAdj() As Boolean
        Dim tempQty As Double = 0

        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If DirectCast(GridView1.Rows(i).FindControl("ckd_itm_code"), HiddenField).Value <> "" AndAlso DirectCast(GridView1.Rows(i).FindControl("ckd_status"), HiddenField).Value = "CNTD" Then

                    If DirectCast(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY"), TextBox).Text = "" Then
                        If DirectCast(GridView1.Rows(i).FindControl("ckd_rev_qty"), TextBox).Text = "" Then
                            DirectCast(GridView1.Rows(i).FindControl("ckd_rev_qty"), TextBox).Focus()
                            uiFun.displayMsg(Me, "", "Quantity Checking has not yet performed.", Session("gLang"))
                            Return False

                        Else
                            tempQty = CDbl(DirectCast(GridView1.Rows(i).FindControl("ckd_rev_qty"), TextBox).Text) - CDbl(DirectCast(GridView1.Rows(i).FindControl("ckd_org_qty"), HiddenField).Value)

                            If tempQty <> 0 Then
                                uiFun.displayMsg(Me, "", "Re-Check has not yet been done.", Session("gLang"))
                                DirectCast(GridView1.FindControl("CKD_ACTUAL_QTY"), TextBox).Focus()
                                Return False
                            End If
                        End If
                    End If

                    If DirectCast(GridView1.Rows(i).FindControl("ckd_var_qty"), HiddenField).Value = "" Then
                        'uiFun.displayMsg(Me, "", "Re-Check has not yet been done.", Session("gLang"))
                        'DirectCast(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY"), TextBox).Focus()
                        'Return False
                    End If


                    If CK_IS_Cable.Checked Then
                        If DirectCast(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY2"), TextBox).Text = "" Then
                            If DirectCast(GridView1.Rows(i).FindControl("ckd_rev_qty2"), TextBox).Text = "" Then
                                DirectCast(GridView1.Rows(i).FindControl("ckd_rev_qty2"), TextBox).Focus()
                                uiFun.displayMsg(Me, "", "Quantity Checking has not yet performed.", Session("gLang"))
                                Return False

                            Else
                                tempQty = CDbl(DirectCast(GridView1.Rows(i).FindControl("ckd_rev_qty2"), TextBox).Text) - CDbl(DirectCast(GridView1.FindControl("ckd_org_qty2"), HiddenField).Value)

                                If tempQty <> 0 Then
                                    uiFun.displayMsg(Me, "", "Re-Check has not yet been done.", Session("gLang"))
                                    DirectCast(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY2"), TextBox).Focus()
                                    Return False
                                End If
                            End If
                        End If

                        If DirectCast(GridView1.Rows(i).FindControl("ckd_var_qty2"), HiddenField).Value = "" Then
                            '    uiFun.displayMsg(Me, "", "Re-Check has not yet been done.", Session("gLang"))
                            '    DirectCast(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY2"), TextBox).Focus()
                            '    Return False
                        End If
                    End If
                End If
            Next
        Else
            uiFun.displayMsg(Me, "", "Please Select Item before generating Stock Adjust.", Session("gLang"))
            Return False
        End If

        Return True
    End Function

    Protected Sub updateBOOK_Click(sender As Object, e As System.EventArgs) Handles updateBOOK.Click
        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            Dim SQLString As String = ""
            Dim itmCode, packKey, batchNo, LocCode, palletNo, SeriesNo, DrumID As String

            Dim paP As GlobalDBFunc.DBCmdPara
            Dim balQty As Double = 0
            Dim balQty2 As Double = 0

            For i = 0 To GridView1.Rows.Count - 1
                itmCode = DirectCast(GridView1.Rows(i).FindControl("ckd_itm_code"), HiddenField).Value
                packKey = DirectCast(GridView1.Rows(i).FindControl("ckd_pack_key"), HiddenField).Value
                batchNo = DirectCast(GridView1.Rows(i).FindControl("ckd_batch_no"), HiddenField).Value
                LocCode = DirectCast(GridView1.Rows(i).FindControl("ckd_loc"), HiddenField).Value
                palletNo = DirectCast(GridView1.Rows(i).FindControl("ckd_pallet_no"), HiddenField).Value
                SeriesNo = DirectCast(GridView1.Rows(i).FindControl("CKD_SERIAL_NO"), HiddenField).Value
                DrumID = DirectCast(GridView1.Rows(i).FindControl("CKD_DRUM_ID"), HiddenField).Value

                paP = New GlobalDBFunc.DBCmdPara
                SQLString = "Select ILOC_BAL_QTY FROM WMS_ITEM_LOC_BAL Where imp_code=" & paP.AP(Session("imp_code")) & " and storer_code=" & paP.AP(STORER_CODE.SelectedValue) & " and ITM_CODE=" & paP.AP(itmCode) &
                             " and PACK_KEY =" & paP.AP(packKey) & " and ILOC_PALLET_NO=" & paP.AP(palletNo) & " and ILOC_LOC=" & paP.AP(LocCode) & " and ISNULL(ILOC_BATCH_NO,'') = ISNULL(" & paP.AP(batchNo) & ",'')"

                balQty = gU.decodeNullOrEmpty(DB.getValueFromSQL(SQLString, , , paP), 0)

                DirectCast(GridView1.Rows(i).FindControl("CKD_BOOK_QTY"), HiddenField).Value = balQty
                DirectCast(GridView1.Rows(i).FindControl("dsp_CKD_BOOK_QTY"), Label).Text = balQty

                If CK_IS_Cable.Checked OrElse CK_TYPE.SelectedValue = "CCC" Then
                    Dim balDT As DataTable

                    paP = New GlobalDBFunc.DBCmdPara
                    SQLString = " SELECT sum(WMS_ITEM_LOC_BAL_S.ILBS_QTY2) as ILBS_QTY2, count(1) as ILOC_BAL_QTY  FROM WMS_ITEM_LOC_BAL INNER JOIN " &
                                " WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " &
                                " Where WMS_ITEM_LOC_BAL.imp_code=" & paP.AP(Session("imp_code")) & " and WMS_ITEM_LOC_BAL.storer_code=" & paP.AP(STORER_CODE.SelectedValue) & " and WMS_ITEM_LOC_BAL.ITM_CODE=" & paP.AP(itmCode) &
                                " and WMS_ITEM_LOC_BAL.PACK_KEY =" & paP.AP(packKey) & " and WMS_ITEM_LOC_BAL.ILOC_PALLET_NO=" & paP.AP(palletNo) & " and WMS_ITEM_LOC_BAL.ILOC_LOC=" & paP.AP(LocCode) & " and ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,'')=ISNULL(" & paP.AP(batchNo) & ",'') " &
                                " and WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID=" & paP.AP(DrumID)

                    Dim tempSQL As String = gDB.getCmdSql(SQLString, paP)
                    balDT = gDB.getDataTable(SQLString, , , , paP)

                    If balDT.Rows.Count > 0 Then
                        balQty = gU.decodeEmptyCdbl(balDT.Rows(0).Item("ILOC_BAL_QTY").ToString, 0)
                        balQty2 = gU.decodeEmptyCdbl(balDT.Rows(0).Item("ILBS_QTY2").ToString, 0)

                        'DirectCast(GridView1.Rows(i).FindControl("CKD_BOOK_QTY"), HiddenField).Value = "1"
                        'DirectCast(GridView1.Rows(i).FindControl("dsp_CKD_BOOK_QTY"), Label).Text = "1"

                        DirectCast(GridView1.Rows(i).FindControl("CKD_BOOK_QTY"), HiddenField).Value = balQty
                        DirectCast(GridView1.Rows(i).FindControl("dsp_CKD_BOOK_QTY"), Label).Text = balQty

                        DirectCast(GridView1.Rows(i).FindControl("CKD_BOOK_QTY2"), HiddenField).Value = balQty2
                        DirectCast(GridView1.Rows(i).FindControl("dsp_CKD_BOOK_QTY2"), Label).Text = balQty2

                    End If

                End If

            Next
            uiFun.displayMsgNew(Updatepanel3, "", "Recheck Booked Aty. has been updated.", Session("gLang"))
        Else
            uiFun.displayMsgNew(Updatepanel3, "", "Please Select Item before update.", Session("gLang"))
        End If
    End Sub


    'Protected Sub CK_TYPE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles CK_TYPE.SelectedIndexChanged
    '    If STORER_CODE.SelectedValue = "" OrElse CK_WH.SelectedValue = "" Then
    '        uiFun.displayMsgNew(Me, "", "Please Select Storer Code and Warehose.", Session("gLang"))
    '        CK_TYPE.SelectedValue = "CC"
    '    Else
    '        If CK_TYPE.SelectedValue = "CCS" Then
    '            GenerateScrapChk()
    '            selectItemBtn.Visible = False
    '        Else
    '            selectItemBtn.Visible = True

    '        End If
    '    End If
    'End Sub

    Protected Sub GenerateScrapChk()
        Dim SQLString As String = ""
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim maxBatchID As String = ""

        Dim tempDT As DataTable
        Dim newRow As DataRow
        Dim itemDT As DataTable = ViewState("dt")

        Dim next_seq_no As String
        Dim temp_seq_no As Integer = 1

        maxBatchID = DB.getValueFromSQL("Select MAX(IMP_BATCH_ID) from WMS_INTF_LOG WHERE ITF_IMP_TYPE = 'IM' and ITF_STATUS='SUCCESS' ")

        itemDT.Clear()

        If maxBatchID <> "" Then
            paP = New GlobalDBFunc.DBCmdPara
            SQLString = " SELECT WMS_CSMS_BAL.IMP_CODE, WMS_CSMS_BAL.BATCH_NO, WMS_CSMS_BAL.STORER_CODE, WMS_CSMS_BAL.ITM_CODE, WMS_CSMS_BAL.PACK_KEY, WMS_CSMS_BAL.CSBA_LOC, WMS_CSMS_BAL.CSBA_CSMS_BAL_QTY, WMS_CSMS_BAL.CSBA_WMS_BAL_QTY, WMS_CSMS_BAL.CSBA_DIFF, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, " & _
                        " WMS_ITEM.ITM_SKU_NO FROM WMS_CSMS_BAL INNER JOIN WMS_ITEM ON WMS_CSMS_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_CSMS_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                        " WMS_CSMS_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_CSMS_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                        " INNER JOIN V_LOCATION ON WMS_CSMS_BAL.IMP_CODE = V_LOCATION.IMP_CODE AND WMS_CSMS_BAL.CSBA_LOC = V_LOCATION.LOC " & _
                        " where WMS_ITEM.ITM_SCRAP_YN='Y' AND WMS_CSMS_BAL.batch_no=" & paP.AP(maxBatchID) & " AND WMS_CSMS_BAL.STORER_CODE=" & paP.AP(STORER_CODE.SelectedValue) & " and V_LOCATION.WH_CODE=" & paP.AP(CK_WH.SelectedValue)

            tempDT = gDB.getDataTable(SQLString, , , , paP)

            If tempDT.Rows.Count > 0 Then
                For i = 0 To tempDT.Rows.Count - 1

                    If ViewState("n_cur_seq") = "" Then

                        next_seq_no = gU.decodeNullOrEmpty(DB.getValueFromSQL("select MAX(CAST(CKD_SEQ AS int)) + 1 from WMS_STOCK_CHECK_D " & _
                                              "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                              "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                              "and CK_CODE = '" & gU.dbEncode(CK_CODE.Text.Trim) & "' "), 1)


                        ViewState("n_cur_seq") = next_seq_no
                    Else
                        temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                        ViewState("n_cur_seq") = temp_seq_no.ToString
                    End If


                    newRow = itemDT.NewRow

                    newRow.Item("mFlag") = "N"
                    newRow.Item("ckd_seq") = ViewState("n_cur_seq").ToString
                    newRow.Item("ckd_itm_code") = tempDT.Rows(i).Item("itm_code").ToString.Trim
                    newRow.Item("itm_sku_no") = tempDT.Rows(i).Item("itm_sku_no").ToString.Trim
                    newRow.Item("itm_name") = tempDT.Rows(i).Item("itm_name").ToString.Trim
                    newRow.Item("ckd_pack_key") = tempDT.Rows(i).Item("pack_key").ToString.Trim
                    newRow.Item("ckd_PALLET_NO") = "000"
                    'newRow.Item("ckd_BATCH_NO") = tempDT.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                    newRow.Item("ckd_org_qty") = tempDT.Rows(i).Item("CSBA_CSMS_BAL_QTY")
                    newRow.Item("ckd_rev_qty") = tempDT.Rows(i).Item("CSBA_WMS_BAL_QTY")
                    newRow.Item("ckd_var_qty") = tempDT.Rows(i).Item("CSBA_DIFF")
                    newRow.Item("CKD_BOOK_QTY") = tempDT.Rows(i).Item("CSBA_CSMS_BAL_QTY")
                    newRow.Item("ckd_LOC") = tempDT.Rows(i).Item("CSBA_LOC").ToString.Trim
                    'newRow.Item("CKD_DRUM_ID") = tempDT.Rows(i).Item("ILBS_DRUM_ID").ToString.Trim
                    'newRow.Item("CKD_DRUM_LEVEL") = tempDT.Rows(i).Item("ILBS_DRUM_LEVEL").ToString.Trim
                    'newRow.Item("CKD_SERIAL_NO") = tempDT.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim
                    'newRow.Item("CKD_UOM2") = tempDT.Rows(i).Item("ILBS_UOM2").ToString.Trim
                    'newRow.Item("CKD_ORG_QTY2") = tempDT.Rows(i).Item("ILBS_QTY2")
                    'newRow.Item("CKD_BOOK_QTY2") = tempDT.Rows(i).Item("ILBS_QTY2")

                    newRow.Item("CKD_CC_LIST_NO") = 1

                    itemDT.Rows.Add(newRow)

                Next

            Else
                uiFun.displayMsgNew(Updatepanel3, "", "No discrepancy for Scarp item is found.", Session("gLang"))
            End If


            itemDT.AcceptChanges()
            ViewState("dt") = itemDT
            GridView1.DataSource = itemDT
            GridView1.DataBind()

            If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
                For i = 20 To 28
                    GridView1.Columns(i).Visible = False
                Next
            End If
        Else
            uiFun.displayMsgNew(Updatepanel3, "", "No discrepancy for Scarp item is found.", Session("gLang"))
            Exit Sub
        End If

    End Sub

    Protected Sub btnRechk_Click(sender As Object, e As System.EventArgs) Handles btnRechk.Click
        Dim pk_code As String = ViewState("CK_CODE")
        Dim successFlag As Boolean = False
        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            Dim cancelSql As String = "update WMS_STOCK_CHECK " &
                     "set CK_STATUS='RECHECK', CK_CNT_NO=2, " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and CK_CODE = '" & gU.dbEncode(pk_code) & "' "

            gDB.amendData(cancelSql, gConn, transaction)

            Dim updateSQL As String = "update WMS_STOCK_CHECK_D " &
                                      "set CKD_STATUS = 'RECHECK', " &
                                      "sys_lub = '" & Session("usr_id") & "', " &
                                      "sys_lud = Getdate() " &
                                      "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                      "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                      "and CK_CODE = '" & gU.dbEncode(pk_code) & "' "

            gDB.amendData(updateSQL, gConn, transaction)


            transaction.Commit()

            CK_STATUS.Value = "CNT"

            ar.sec_write = "N"
            btnRechk.Visible = False
            ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

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
            rmtPost.Url = "SCHKMAIN.aspx"
            rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
            rmtPost.Add("CK_CODE", pk_code)
            rmtPost.alertMsg = "This Cycle Count is Ready for Recheck!"
            rmtPost.Post()
        End If

    End Sub

    Protected Sub btnCnt_Click(sender As Object, e As System.EventArgs) Handles btnCnt.Click
        Dim pk_code As String = ViewState("CK_CODE")
        Dim successFlag As Boolean = False
        Dim gConn As SqlConnection
        Dim updateSQL As String = ""
        Dim selectSQL As String = ""
        If validateAll() Then
            If validateCount() Then
                Call save("Y")

                gConn = gDB.getConnection()
                Dim transaction As SqlTransaction


                transaction = gConn.BeginTransaction()

                Try

                    Dim tempDT As DataTable

                    selectSQL = "Select ck_code, ckd_seq, ckd_status from WMS_STOCK_CHECK_D where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and CK_CODE = '" & gU.dbEncode(pk_code) & "' "

                    tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                    If tempDT IsNot Nothing AndAlso tempDT.Rows.Count > 0 Then
                        For i = 0 To tempDT.Rows.Count - 1
                            If gU.inList("READY, CNT", tempDT.Rows(i).Item("ckd_status").ToString.Trim) Then
                                updateSingleBookCount(pk_code, tempDT.Rows(i).Item("ckd_seq").ToString.Trim, gConn, transaction, "1")
                            End If
                        Next
                    End If

                    selectSQL = "Select * from WMS_STOCK_CHECK_D where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and CK_CODE = '" & gU.dbEncode(pk_code) & "' "
                    tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                    If Not tempDT Is Nothing AndAlso tempDT.Rows.Count > 0 Then
                        For i = 0 To tempDT.Rows.Count - 1
                            Select Case tempDT.Rows(i).Item("ckd_status").ToString.Trim
                                Case "READY", "CNT"
                                    If tempDT.Rows(i).Item("ckd_org_qty").ToString.Trim = tempDT.Rows(i).Item("ckd_rev_qty").ToString.Trim AndAlso tempDT.Rows(i).Item("ckd_org_qty2").ToString.Trim = tempDT.Rows(i).Item("ckd_rev_qty2").ToString.Trim Then

                                        updateSQL = "update WMS_STOCK_CHECK_D set CKD_STATUS='CNTD', CKD_CC1_DATE=getdate(),sys_lud = getDate(),sys_lub='" & Session("usr_id") & "', CKD_BY='" & Session("usr_id") & "' where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                    "and CK_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                                                    "and ckd_seq='" & gU.dbEncode(tempDT.Rows(i).Item("ckd_seq").ToString.Trim) & "'"

                                        gDB.amendData(updateSQL, gConn, transaction)

                                        'updateSQL = "Update WMS_ITEM set ITM_CC_DATE=getdate() where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        '            "and itm_code='" & gU.dbEncode(tempDT.Rows(i).Item("CKD_ITM_CODE").ToString.Trim) & "' and pack_key='" & gU.dbEncode(tempDT.Rows(i).Item("CKD_PACK_KEY").ToString.Trim) & "'"

                                        'gDB.amendData(updateSQL, gConn, transaction)
                                    Else
                                        updateSQL = "update WMS_STOCK_CHECK_D set CKD_STATUS='RECHECK' , CKD_CC1_DATE=getdate(), sys_lud = getDate(),sys_lub='" & Session("usr_id") & "', CKD_BY='" & Session("usr_id") & "'  where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                    "and CK_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                                                    "and ckd_seq='" & gU.dbEncode(tempDT.Rows(i).Item("ckd_seq").ToString.Trim) & "'"

                                        gDB.amendData(updateSQL, gConn, transaction)
                                    End If


                                Case "RE-CNT", "RECHECK"

                                    updateSingleBookCount(pk_code, tempDT.Rows(i).Item("ckd_seq").ToString.Trim, gConn, transaction, "2")

                                    updateSQL = "update WMS_STOCK_CHECK_D set CKD_STATUS='CNTD', CKD_CC_DATE=getdate(), sys_lud = getDate(),sys_lub='" & Session("usr_id") & "', CKD_RECHECK_BY='" & Session("usr_id") & "' where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                "and CK_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                                                "and ckd_seq='" & gU.dbEncode(tempDT.Rows(i).Item("ckd_seq").ToString.Trim) & "'"

                                    gDB.amendData(updateSQL, gConn, transaction)

                                    'updateSQL = "Update WMS_ITEM set ITM_CC_DATE=getdate() where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    '            "and itm_code='" & gU.dbEncode(tempDT.Rows(i).Item("CKD_ITM_CODE").ToString.Trim) & "' and pack_key='" & gU.dbEncode(tempDT.Rows(i).Item("CKD_PACK_KEY").ToString.Trim) & "'"

                                    'gDB.amendData(updateSQL, gConn, transaction)

                                    updateSQL = " update WMS_STOCK_CHECK_D set ckd_rem = case when (CKD_BOOK_QTY = 0 and CKD_VAR_QTY <> 0) or (CKD_BOOK_QTY2 = 0 and CKD_VAR_QTY2 <> 0) then 'Location Mismatch' " & _
                                                " when (CKD_BOOK_QTY <> 0 and CKD_VAR_QTY <> 0) or (CKD_BOOK_QTY2 <> 0 and CKD_VAR_QTY2 <> 0) then 'Quantity Mismatch'  end " & _
                                                " where ((ckd_var_qty is not null and CKD_VAR_QTY <> 0) or (ckd_var_qty2 is not null and CKD_VAR_QTY2 <> 0)) and isnull(ckd_rem,'') = '' " & _
                                                " and imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                                " and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                " and CK_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                                                " and ckd_seq='" & gU.dbEncode(tempDT.Rows(i).Item("ckd_seq").ToString.Trim) & "'"

                                    gDB.amendData(updateSQL, gConn, transaction)
                            End Select
                        Next

                        updateSQL += "merge wms_stock_check as t "
                        updateSQL += "using ( "
                        updateSQL += "select cc.imp_code, cc.storer_code, cc.ck_code, "
                        updateSQL += "case "
                        updateSQL += "when sc.os_cnt = 0 and sc.counting_cnt = 0 and sc.recheck_cnt = 0 and sc.recnt_cnt = 0 and sc.done_cnt = 0 then 'NEW' "
                        updateSQL += "when sc.os_cnt > 0 and sc.counting_cnt = 0 and sc.recheck_cnt = 0 and sc.recnt_cnt = 0 and sc.done_cnt = 0 then 'READY' "
                        updateSQL += "when sc.os_cnt = 0 and sc.counting_cnt = 0 and sc.recnt_cnt = 0 and sc.recheck_cnt > 0 then 'RECHECK' "
                        updateSQL += "when sc.os_cnt = 0 and sc.counting_cnt = 0 and sc.recnt_cnt > 0 then 'RE-CNT' "
                        updateSQL += "when sc.os_cnt = 0 and sc.counting_cnt = 0 and sc.recnt_cnt = 0 and sc.recheck_cnt = 0 and sc.done_cnt > 0 then 'CNTD' "
                        updateSQL += "else 'CNT' "
                        updateSQL += "end as ckd_status "
                        updateSQL += "from wms_stock_check cc, "
                        updateSQL += "(select imp_code, storer_code, ck_code, "
                        updateSQL += "sum(case when ckd_status = 'READY' then 1 else 0 end) as os_cnt, "
                        updateSQL += "sum(case when ckd_status = 'CNT' then 1 else 0 end) as counting_cnt, "
                        updateSQL += "sum(case when ckd_status = 'RECHECK' then 1 else 0 end) as recheck_cnt, "
                        updateSQL += "sum(case when ckd_status = 'RE-CNT' then 1 else 0 end) as recnt_cnt, "
                        updateSQL += "sum(case when ckd_status = 'CNTD' then 1 else 0 end) as done_cnt "
                        updateSQL += "from wms_stock_check_d "
                        updateSQL += "group by imp_code, storer_code, ck_code) sc "
                        updateSQL += "where cc.imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' "
                        updateSQL += "and cc.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
                        updateSQL += "and cc.ck_code = '" & gU.dbEncode(pk_code) & "' "
                        updateSQL += "and sc.imp_code = cc.imp_code "
                        updateSQL += "and sc.storer_code = cc.storer_code "
                        updateSQL += "and sc.ck_code = cc.ck_code "
                        updateSQL += ") as s "
                        updateSQL += "on ( "
                        updateSQL += "t.imp_code = s.imp_code "
                        updateSQL += "and t.storer_code = s.storer_code "
                        updateSQL += "and t.ck_code = s.ck_code "
                        updateSQL += ") "
                        updateSQL += "when matched then "
                        updateSQL += "update set "
                        updateSQL += "t.ck_status = s.ckd_status, "
                        updateSQL += "t.sys_lud = getDate(), "
                        updateSQL += "t.sys_lub = '" & Session("usr_id") & "'"
                        updateSQL += ";"

                        gDB.amendData(updateSQL, gConn, transaction)


                        transaction.Commit()

                        uiFun.displayMsg(Me, "", "The Cycle Count has been counted.", Session("gLang"))
                        successFlag = True

                    End If
                    'updateSQL = "update WMS_STOCK_CHECK_D set CKD_STATUS='CNTD' and CKD_CC1_DATE=getdate() where ckd_org_qty = ckd_rev_qty and isnull(ckd_org_qty2,0) = isnull(ckd_rev_qty2,0)  and (CKD_STATUS='CNT' or CKD_STATUS='READY') " & _
                    '            "AND imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                    '            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    '            "and CK_CODE = '" & gU.dbEncode(pk_code) & "' "
                    'gDB.amendData(updateSQL, gConn, transaction)

                    'updateSQL = "update WMS_STOCK_CHECK_D set CKD_STATUS='CNTD' where  and (CKD_STATUS='CNT' or CKD_STATUS='READY') " & _
                    '            "AND imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                    '            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    '            "and CK_CODE = '" & gU.dbEncode(pk_code) & "' "
                    'gDB.amendData(updateSQL, gConn, transaction)

                    'updateSQL = "update wms_item set wms_item.ITM_CC_DATE=getdate() where concat(wms_item.IMP_CODE,'||', wms_item.STORER_CODE,'||',wms_item.ITM_CODE,'||',wms_item.PACK_KEY) in (" & _
                    '            "select distinct concat(WMS_STOCK_CHECK_D.IMP_CODE,'||', WMS_STOCK_CHECK_D.STORER_CODE,'||', WMS_STOCK_CHECK_D.CKD_ITM_CODE,'||', WMS_STOCK_CHECK_D.CKD_PACK_KEY) from wms_stock_check_d " & _
                    '            "where WMS_STOCK_CHECK_D.imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' and WMS_STOCK_CHECK_D.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and WMS_STOCK_CHECK_D.CK_CODE = '" & gU.dbEncode(pk_code) & "') "

                    'gDB.amendData(updateSQL, gConn, transaction)

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
        End If


        If successFlag Then
            Dim rmtPost As New RemotePost
            rmtPost.Url = "SCHKMAIN.aspx"
            rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
            rmtPost.Add("CK_CODE", pk_code)
            rmtPost.alertMsg = "This Cycle Count has been counted!"
            rmtPost.Post()
        End If

    End Sub

    Protected Function validateCount() As Boolean
        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                Select Case DirectCast(GridView1.Rows(i).FindControl("CKD_STATUS"), HiddenField).Value
                    Case "READY", "CNT"
                        If String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("ckd_rev_qty"), TextBox).Text) Then
                            uiFun.displayMsgNew(Updatepanel3, "", "Please enter Checked Qty before finish counting!", Session("gLang"))
                            Return False
                            'ElseIf CDbl(CType(GridView1.Rows(i).FindControl("ckd_rev_qty"), TextBox).Text) <> CDbl(CType(GridView1.Rows(i).FindControl("ckd_org_qty"), HiddenField).Value) AndAlso String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("ckd_rem"), TextBox).Text) Then
                            '    uiFun.displayMsgNew(Updatepanel3, "", "Please enter Remarks for items with variance!", Session("gLang"))
                            '    Return False
                        End If

                        If CK_IS_Cable.Checked Then
                            If String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("ckd_rev_qty2"), TextBox).Text) Then
                                uiFun.displayMsgNew(Updatepanel3, "", "Please enter Checked Qty2 before finish counting!", Session("gLang"))
                                Return False
                                'ElseIf CDbl(CType(GridView1.Rows(i).FindControl("ckd_rev_qty2"), TextBox).Text) <> CDbl(CType(GridView1.Rows(i).FindControl("CKD_ORG_QTY2"), HiddenField).Value) AndAlso String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("ckd_rem"), TextBox).Text) Then
                                '    uiFun.displayMsgNew(Updatepanel3, "", "Please enter Remarks for items with variance!", Session("gLang"))
                                '    Return False
                            End If
                        End If

                    Case "RE-CNT", "RECHECK"
                        If String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY"), TextBox).Text) Then
                            CType(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY"), TextBox).Text = "0"
                            'uiFun.displayMsgNew(Updatepanel3, "", "Please enter ReChecked QTY before finish counting!", Session("gLang"))
                            'Return False
                            'ElseIf CDbl(CType(GridView1.Rows(i).FindControl("ckd_var_qty"), HiddenField).Value) > 0 AndAlso String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("ckd_rem"), TextBox).Text) Then
                            '    uiFun.displayMsgNew(Updatepanel3, "", "Please enter Remarks for items with variance!", Session("gLang"))
                            '    Return False
                        End If

                        If CK_IS_Cable.Checked Then
                            If String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("CKD_ACTUAL_QTY2"), TextBox).Text) Then
                                uiFun.displayMsgNew(Updatepanel3, "", "Please enter Checked Qty2 before finish counting!", Session("gLang"))
                                Return False
                                'ElseIf CDbl(CType(GridView1.Rows(i).FindControl("ckd_var_qty2"), HiddenField).Value) > 0 AndAlso String.IsNullOrWhiteSpace(CType(GridView1.Rows(i).FindControl("ckd_rem"), TextBox).Text) Then
                                '    uiFun.displayMsgNew(Updatepanel3, "", "Please enter Remarks for items with variance!", Session("gLang"))
                                '    Return False
                            End If
                        End If
                    Case ""
                        uiFun.displayMsgNew(Updatepanel3, "", "Please release the cycle count before finishing counting!", Session("gLang"))
                        Return False
                End Select
            Next
        Else
            Return False
        End If

        Return True
    End Function

    Protected Function validateCountSingle(ByVal rowNum As Integer) As Boolean
        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            Select Case DirectCast(GridView1.Rows(rowNum).FindControl("CKD_STATUS"), HiddenField).Value
                Case "READY", "CNT"
                    If String.IsNullOrWhiteSpace(CType(GridView1.Rows(rowNum).FindControl("ckd_rev_qty"), TextBox).Text) Then
                        uiFun.displayMsgNew(Updatepanel3, "", "Please enter Checked Qty before finish counting!", Session("gLang"))
                        Return False
                        'ElseIf CDbl(CType(GridView1.Rows(rowNum).FindControl("ckd_rev_qty"), TextBox).Text) <> CDbl(CType(GridView1.Rows(rowNum).FindControl("ckd_org_qty"), HiddenField).Value) AndAlso String.IsNullOrWhiteSpace(CType(GridView1.Rows(rowNum).FindControl("ckd_rem"), TextBox).Text) Then
                        '    uiFun.displayMsgNew(Updatepanel3, "", "Please enter Remarks for items with variance!", Session("gLang"))
                        '    Return False
                    End If

                    If CK_IS_Cable.Checked Then
                        If String.IsNullOrWhiteSpace(CType(GridView1.Rows(rowNum).FindControl("ckd_rev_qty2"), TextBox).Text) Then
                            uiFun.displayMsgNew(Updatepanel3, "", "Please enter Checked Qty2 before finish counting!", Session("gLang"))
                            Return False
                            'ElseIf CDbl(CType(GridView1.Rows(rowNum).FindControl("ckd_rev_qty2"), TextBox).Text) <> CDbl(CType(GridView1.Rows(rowNum).FindControl("CKD_ORG_QTY2"), HiddenField).Value) AndAlso String.IsNullOrWhiteSpace(CType(GridView1.Rows(rowNum).FindControl("ckd_rem"), TextBox).Text) Then
                            '    uiFun.displayMsgNew(Updatepanel3, "", "Please enter Remarks for items with variance!", Session("gLang"))
                            '    Return False
                        End If
                    End If



                Case "RE-CNT", "RECHECK"
                    If String.IsNullOrWhiteSpace(CType(GridView1.Rows(rowNum).FindControl("CKD_ACTUAL_QTY"), TextBox).Text) Then
                        CType(GridView1.Rows(rowNum).FindControl("CKD_ACTUAL_QTY"), TextBox).Text = "0"
                        'uiFun.displayMsgNew(Updatepanel3, "", "Please enter ReChecked QTY before finish counting!", Session("gLang"))
                        'Return False
                        'ElseIf CDbl(CType(GridView1.Rows(rowNum).FindControl("ckd_var_qty"), HiddenField).Value) > 0 AndAlso String.IsNullOrWhiteSpace(CType(GridView1.Rows(rowNum).FindControl("ckd_rem"), TextBox).Text) Then
                        '    uiFun.displayMsgNew(Updatepanel3, "", "Please enter Remarks for items with variance!", Session("gLang"))
                        '    Return False
                    End If

                    If CK_IS_Cable.Checked Then
                        If String.IsNullOrWhiteSpace(CType(GridView1.Rows(rowNum).FindControl("CKD_ACTUAL_QTY2"), TextBox).Text) Then
                            uiFun.displayMsgNew(Updatepanel3, "", "Please enter Checked Qty2 before finish counting!", Session("gLang"))
                            Return False
                            'ElseIf CDbl(CType(GridView1.Rows(rowNum).FindControl("ckd_var_qty2"), HiddenField).Value) > 0 AndAlso String.IsNullOrWhiteSpace(CType(GridView1.Rows(rowNum).FindControl("ckd_rem"), TextBox).Text) Then
                            '    uiFun.displayMsgNew(Updatepanel3, "", "Please enter Remarks for items with variance!", Session("gLang"))
                            '    Return False
                        End If
                    End If
                Case ""
                    uiFun.displayMsgNew(Updatepanel3, "", "Please release the cycle count before finishing counting!", Session("gLang"))
                    Return False
            End Select
        Else
            Return False
        End If

        Return True
    End Function

    Protected Sub status_filter_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles status_filter.SelectedIndexChanged
        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                Select Case status_filter.SelectedValue
                    Case "FC"
                        If DirectCast(GridView1.Rows(i).FindControl("CKD_STATUS"), HiddenField).Value = "READY" OrElse DirectCast(GridView1.Rows(i).FindControl("CKD_STATUS"), HiddenField).Value = "CNT" Then
                            GridView1.Rows(i).Visible = True
                        Else
                            GridView1.Rows(i).Visible = False
                        End If

                    Case "RE"
                        If DirectCast(GridView1.Rows(i).FindControl("CKD_STATUS"), HiddenField).Value = "RECHECK" OrElse DirectCast(GridView1.Rows(i).FindControl("CKD_STATUS"), HiddenField).Value = "RE-CNT" Then
                            GridView1.Rows(i).Visible = True
                        Else
                            GridView1.Rows(i).Visible = False
                        End If

                    Case "CD"
                        If DirectCast(GridView1.Rows(i).FindControl("CKD_STATUS"), HiddenField).Value = "CNTD" OrElse DirectCast(GridView1.Rows(i).FindControl("CKD_STATUS"), HiddenField).Value = "ADJ" Then
                            GridView1.Rows(i).Visible = True
                        Else
                            GridView1.Rows(i).Visible = False
                        End If

                    Case Else
                        GridView1.Rows(i).Visible = True
                End Select
            Next
        End If
    End Sub

    Private Function CountSingle(ByVal rowNum As Integer, ByVal lCKD_SEQ As String, Optional ByVal dStatus As String = "") As Boolean
        Dim pk_code As String = ViewState("CK_CODE")
        Dim successFlag As Boolean = False
        Dim gConn As SqlConnection
        Dim updateSQL As String = ""
        Dim selectSQL As String = ""

        If validateAll() Then
            If validateCountSingle(rowNum) Then
                Call save("Y")

                gConn = gDB.getConnection()
                Dim transaction As SqlTransaction
                transaction = gConn.BeginTransaction()

                Try
                    If gU.inList("READY, CNT", dStatus) Then
                        updateSingleBookCount(pk_code, lCKD_SEQ, gConn, transaction, "1")
                    End If

                    Dim tempDT As DataTable

                    selectSQL = "Select * from WMS_STOCK_CHECK_D where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and CK_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                                "and ckd_seq = '" & gU.dbEncode(lCKD_SEQ) & "'"

                    tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                    If Not tempDT Is Nothing AndAlso tempDT.Rows.Count > 0 Then

                        Select Case tempDT.Rows(0).Item("ckd_status").ToString.Trim
                            Case "READY", "CNT"
                                If tempDT.Rows(0).Item("ckd_org_qty").ToString.Trim = tempDT.Rows(0).Item("ckd_rev_qty").ToString.Trim AndAlso tempDT.Rows(0).Item("ckd_org_qty2").ToString.Trim = tempDT.Rows(0).Item("ckd_rev_qty2").ToString.Trim Then

                                    updateSQL = "update WMS_STOCK_CHECK_D set CKD_STATUS='CNTD', CKD_CC1_DATE=getdate(),CKD_CC_DATE=getdate(), sys_lud = getDate(),sys_lub='" & Session("usr_id") & "', CKD_BY='" & Session("usr_id") & "' where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                "and CK_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                                                "and ckd_seq='" & gU.dbEncode(tempDT.Rows(0).Item("ckd_seq").ToString.Trim) & "'"

                                    gDB.amendData(updateSQL, gConn, transaction)

                                    'updateSQL = "Update WMS_ITEM set ITM_CC_DATE=getdate() where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    '            "and itm_code='" & gU.dbEncode(tempDT.Rows(0).Item("CKD_ITM_CODE").ToString.Trim) & "' and pack_key='" & gU.dbEncode(tempDT.Rows(0).Item("CKD_PACK_KEY").ToString.Trim) & "'"

                                    'gDB.amendData(updateSQL, gConn, transaction)
                                Else
                                    updateSQL = "update WMS_STOCK_CHECK_D set CKD_STATUS='RECHECK' , CKD_CC1_DATE=getdate(), sys_lud = getDate(),sys_lub='" & Session("usr_id") & "' where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                "and CK_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                                                "and ckd_seq='" & gU.dbEncode(tempDT.Rows(0).Item("ckd_seq").ToString.Trim) & "'"

                                    gDB.amendData(updateSQL, gConn, transaction)
                                End If


                            Case "RE-CNT", "RECHECK"

                                updateSingleBookCount(pk_code, tempDT.Rows(0).Item("ckd_seq").ToString.Trim, gConn, transaction, "2")

                                updateSQL = "update WMS_STOCK_CHECK_D set CKD_STATUS='CNTD', CKD_CC_DATE=getdate(), sys_lud = getDate(),sys_lub='" & Session("usr_id") & "', CKD_RECHECK_BY='" & Session("usr_id") & "' where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and CK_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                                            "and ckd_seq='" & gU.dbEncode(tempDT.Rows(0).Item("ckd_seq").ToString.Trim) & "'"

                                gDB.amendData(updateSQL, gConn, transaction)

                                'updateSQL = "Update WMS_ITEM set ITM_CC_DATE=getdate() where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                '            "and itm_code='" & gU.dbEncode(tempDT.Rows(0).Item("CKD_ITM_CODE").ToString.Trim) & "' and pack_key='" & gU.dbEncode(tempDT.Rows(0).Item("CKD_PACK_KEY").ToString.Trim) & "'"

                                'gDB.amendData(updateSQL, gConn, transaction)

                                updateSQL = " update WMS_STOCK_CHECK_D set ckd_rem = case when (CKD_BOOK_QTY = 0 and CKD_VAR_QTY <> 0) or (CKD_BOOK_QTY2 = 0 and CKD_VAR_QTY2 <> 0) then 'Location Mismatch' " & _
                                            " when (CKD_BOOK_QTY <> 0 and CKD_VAR_QTY <> 0) or (CKD_BOOK_QTY2 <> 0 and CKD_VAR_QTY2 <> 0) then 'Quantity Mismatch'  end " & _
                                            " where ((ckd_var_qty is not null and CKD_VAR_QTY <> 0) or (ckd_var_qty2 is not null and CKD_VAR_QTY2 <> 0)) and isnull(ckd_rem,'') = '' " & _
                                            " and imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                            " and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            " and CK_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                                            " and ckd_seq='" & gU.dbEncode(tempDT.Rows(0).Item("ckd_seq").ToString.Trim) & "'"

                                gDB.amendData(updateSQL, gConn, transaction)

                        End Select


                        updateSQL += "merge wms_stock_check as t "
                        updateSQL += "using ( "
                        updateSQL += "select cc.imp_code, cc.storer_code, cc.ck_code, "
                        updateSQL += "case "
                        updateSQL += "when sc.os_cnt = 0 and sc.counting_cnt = 0 and sc.recheck_cnt = 0 and sc.recnt_cnt = 0 and sc.done_cnt = 0 then 'NEW' "
                        updateSQL += "when sc.os_cnt > 0 and sc.counting_cnt = 0 and sc.recheck_cnt = 0 and sc.recnt_cnt = 0 and sc.done_cnt = 0 then 'READY' "
                        updateSQL += "when sc.os_cnt = 0 and sc.counting_cnt = 0 and sc.recnt_cnt = 0 and sc.recheck_cnt > 0 then 'RECHECK' "
                        updateSQL += "when sc.os_cnt = 0 and sc.counting_cnt = 0 and sc.recnt_cnt > 0 then 'RE-CNT' "
                        updateSQL += "when sc.os_cnt = 0 and sc.counting_cnt = 0 and sc.recnt_cnt = 0 and sc.recheck_cnt = 0 and sc.done_cnt > 0 then 'CNTD' "
                        updateSQL += "else 'CNT' "
                        updateSQL += "end as ckd_status "
                        updateSQL += "from wms_stock_check cc, "
                        updateSQL += "(select imp_code, storer_code, ck_code, "
                        updateSQL += "sum(case when ckd_status = 'READY' then 1 else 0 end) as os_cnt, "
                        updateSQL += "sum(case when ckd_status = 'CNT' then 1 else 0 end) as counting_cnt, "
                        updateSQL += "sum(case when ckd_status = 'RECHECK' then 1 else 0 end) as recheck_cnt, "
                        updateSQL += "sum(case when ckd_status = 'RE-CNT' then 1 else 0 end) as recnt_cnt, "
                        updateSQL += "sum(case when ckd_status = 'CNTD' then 1 else 0 end) as done_cnt "
                        updateSQL += "from wms_stock_check_d "
                        updateSQL += "group by imp_code, storer_code, ck_code) sc "
                        updateSQL += "where cc.imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' "
                        updateSQL += "and cc.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
                        updateSQL += "and cc.ck_code = '" & gU.dbEncode(pk_code) & "' "
                        updateSQL += "and sc.imp_code = cc.imp_code "
                        updateSQL += "and sc.storer_code = cc.storer_code "
                        updateSQL += "and sc.ck_code = cc.ck_code "
                        updateSQL += ") as s "
                        updateSQL += "on ( "
                        updateSQL += "t.imp_code = s.imp_code "
                        updateSQL += "and t.storer_code = s.storer_code "
                        updateSQL += "and t.ck_code = s.ck_code "
                        updateSQL += ") "
                        updateSQL += "when matched then "
                        updateSQL += "update set "
                        updateSQL += "t.ck_status = s.ckd_status, "
                        updateSQL += "t.sys_lud = getDate(), "
                        updateSQL += "t.sys_lub = '" & Session("usr_id") & "'"
                        updateSQL += ";"

                        gDB.amendData(updateSQL, gConn, transaction)


                        transaction.Commit()

                        'uiFun.displayMsg(Me, "", "This Cycle Count item has been counted.", Session("gLang"))
                        successFlag = True

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
            End If
        End If

        If successFlag Then
            'ScriptManager.RegisterStartupScript(Updatepanel3, Updatepanel3.GetType, "RELOAD_PAGE", "reloadPage('This Cycle Count item has been counted.');", True)
            selectSQL = "SELECT wms_stock_check_d.IMP_CODE," & _
                    " wms_stock_check_d.STORER_CODE,  wms_stock_check_d.CK_CODE, " & _
                    " wms_stock_check_d.ckd_seq,  wms_stock_check_d.ckd_ITM_CODE, " & _
                    " wms_stock_check_d.ckd_PACK_KEY, wms_stock_check_d.ckd_LOC,  wms_stock_check_d.ckd_REV_QTY, " & _
                    " wms_stock_check_d.ckd_ORG_QTY, wms_stock_check_d.ckd_VAR_QTY,WMS_STOCK_CHECK_D.CKD_ACTUAL_QTY, " & _
                    " wms_stock_check_d.ckd_REM,  wms_stock_check_d.ckd_PALLET_NO, " & _
                    " wms_stock_check_d.ckd_BATCH_NO,  wms_stock_check_d.ckd_VND_CODE, " & _
                    " wms_item.itm_sku_no,wms_item.itm_uom, wms_stock_check_d.CKD_STATUS, " & _
                    " Convert(varchar, wms_stock_check_d.ckd_MANU_DATE,103) as ckd_MANU_DATE,  " & _
                    " Convert(varchar, wms_stock_check_d.ckd_EXPIRY_DATE,103) as ckd_EXPIRY_DATE, " & _
                    " WMS_STOCK_CHECK_D.CKD_CC_LIST_NO, WMS_STOCK_CHECK_D.CKD_SERIAL, WMS_STOCK_CHECK_D.CKD_TYPE, " & _
                    " convert(varchar,WMS_STOCK_CHECK_D.CKD_CC_DATE," & DDFORMAT & ") as CKD_CC_DATE, WMS_STOCK_CHECK_D.CKD_FULL_DRUM, WMS_STOCK_CHECK_D.CKD_DRUM_ID, " & _
                    " WMS_STOCK_CHECK_D.CKD_DRUM_LEVEL, WMS_STOCK_CHECK_D.CKD_SERIAL_NO, WMS_STOCK_CHECK_D.CKD_UOM2, " & _
                    " WMS_STOCK_CHECK_D.CKD_ORG_QTY2, WMS_STOCK_CHECK_D.CKD_REV_QTY2, WMS_STOCK_CHECK_D.CKD_VAR_QTY2, WMS_STOCK_CHECK_D.CKD_ACTUAL_QTY2, " & _
                    " WMS_STOCK_CHECK_D.CKD_BOOK_QTY, WMS_STOCK_CHECK_D.CKD_BOOK_QTY2, Convert(varchar,WMS_STOCK_CHECK_D.CKD_CC1_DATE," & DDFORMAT & ") as CKD_CC1_DATE," & _
                    " wms_item.itm_name, 'U' as mFlag, wms_stock_check_d.ckd_seq as old_seq,wms_stock_check_d.CKD_RECHECK_BY, wms_stock_check_d.CKD_BY, wms_stock_check_d.CKD_AD_CODE,wms_stock_check_d.CKD_WITNESS,wms_stock_check_d.CKD_CHECKER,wms_stock_check_d.CKD_ADJ_REM," & _
                    " CASE WHEN WMS_STOCK_CHECK_D.CKD_RECHECK_BY IS NULL THEN WMS_STOCK_CHECK_D.CKD_IN_PDA ELSE WMS_STOCK_CHECK_D.CKD_RECHK_IN_PDA END AS CKD_IN_PDA, " & _
                    " WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE, " & _
                    " '' as itm_name_textbox, '' as dsp_ckd_status,'' as from_imast " & _
                    " from wms_stock_check_d left outer join wms_item on " & _
                    " wms_stock_check_d.imp_code = wms_item.imp_code " & _
                    " and wms_stock_check_d.storer_code = wms_item.storer_code " & _
                    " and wms_stock_check_d.ckd_itm_code = wms_item.itm_code " & _
                    " and wms_stock_check_d.ckd_PACK_KEY = wms_item.PACK_KEY " & _
                    " LEFT OUTER JOIN WMS_WH_BIN on wms_stock_check_d.ckd_LOC = WMS_WH_BIN.LOC_KEY" & _
                    " where wms_stock_check_d.ck_code = '" & gU.dbEncode(pk_code) & "' " & _
                    " and wms_stock_check_d.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    " and wms_stock_check_d.imp_code = '" & Session("IMP_CODE") & "'" & _
                    " order by Convert(decimal, wms_stock_check_d.ckd_seq)"


            Dim tempDT2 As DataTable = gDB.getDataTable(selectSQL)

            GridView1.DataSource = tempDT2
            GridView1.DataBind()
            uiFun.displayMsgNew(updtPnlAlert, "", "This Cycle Count item has been counted.", Session("gLang"))
        End If
        Return successFlag
    End Function

    Protected Sub CK_TYPE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles CK_TYPE.SelectedIndexChanged
        If CK_TYPE.SelectedValue = "CCC" Then
            CK_IS_Cable.Checked = True
        Else
            CK_IS_Cable.Checked = False
        End If
    End Sub

    Protected Sub CK_DATE_TextChanged(sender As Object, e As System.EventArgs) Handles CK_DATE.TextChanged

        'If Not String.IsNullOrWhiteSpace(CK_DATE.Text) Then
        '    Dim FR_DATE, TO_DATE As Date
        '    Dim s_date As String = CK_DATE.Text.Trim
        '    Dim rt_Date As String = ""

        '    Dim customDateTimeFormat As System.Globalization.DateTimeFormatInfo = New System.Globalization.DateTimeFormatInfo()
        '    customDateTimeFormat.DateSeparator = "/"
        '    customDateTimeFormat.TimeSeparator = ":"
        '    customDateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
        '    customDateTimeFormat.LongDatePattern = "dd/MM/yyyy"
        '    customDateTimeFormat.ShortTimePattern = "HH:mm"
        '    customDateTimeFormat.LongTimePattern = "HH:mm"
        '    customDateTimeFormat.FullDateTimePattern = "dd/MM/yyyy HH:mm"

        '    FR_DATE = Convert.ToDateTime(s_date, customDateTimeFormat)



        '    Select Case ck_period.SelectedValue
        '        Case "1YR"
        '            TO_DATE = FR_DATE.AddYears(1).AddDays(-1)
        '            rt_Date = TO_DATE.ToString("dd/MM/yyyy")
        '        Case "2YR"
        '            TO_DATE = FR_DATE.AddYears(2).AddDays(-1)
        '            rt_Date = TO_DATE.ToString("dd/MM/yyyy")
        '        Case "4YR"
        '            TO_DATE = FR_DATE.AddYears(4).AddDays(-1)
        '            rt_Date = TO_DATE.ToString("dd/MM/yyyy")
        '        Case Else
        '            rt_Date = ""
        '    End Select

        '    ck_end_date.Text = rt_Date

        'End If
    End Sub

    Private Sub reloadHiddenValue(ByRef gv As GridView)
        For i = 0 To gv.Rows.Count - 1
            CType(gv.Rows(i).FindControl("dsp_ckd_loc"), Label).Text = CType(GridView1.Rows(i).FindControl("ckd_loc"), HiddenField).Value
            CType(gv.Rows(i).FindControl("dsp_WH_CODE"), Label).Text = CType(GridView1.Rows(i).FindControl("WH_CODE"), HiddenField).Value
            CType(gv.Rows(i).FindControl("dsp_BN_CSMS_CODE"), Label).Text = CType(GridView1.Rows(i).FindControl("BN_CSMS_CODE"), HiddenField).Value
        Next
    End Sub

    Protected Sub btnDelAll_Click(sender As Object, e As System.EventArgs) Handles btnDelAll.Click

        For i = 0 To dt.Rows.Count - 1
            dt.Rows(i).Item("mFlag") = "D"
        Next

        dt.AcceptChanges()

        GridView1.DataSource = dt
        GridView1.DataBind()

        ViewState("dt") = dt

        'uiFun.displayMsg(Me, "", "All items are deleted!", Session("gLang"))
    End Sub

    Protected Sub CK_WH_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles CK_WH.SelectedIndexChanged
        If CK_WH.SelectedValue <> "" Then
            uiFun.load_dropdown(CK_SUB_WH, "Select wh_code, wh_name from wms_warehouse where imp_code='" & gU.dbEncode(Session("imp_code")) & "' AND WH_MAIN_WH='" & gU.dbEncode(CK_WH.SelectedValue) & "' order by wh_name", "wh_code", "wh_name")
        End If

    End Sub

    Protected Sub GVBin_DataBound(sender As Object, e As System.EventArgs) Handles GVBin.DataBound
        If GVBin.DataSource IsNot Nothing Then
            Dim ddlPager As DropDownList = TryCast(GVBin.BottomPagerRow.FindControl("pager_select"), DropDownList)
            Dim tmpPagerBtn As ImageButton

            If ddlPager.Items.Count = 0 Then
                If ddlPager IsNot Nothing Then
                    For i As Integer = 1 To GVBin.PageCount
                        ddlPager.Items.Add(New ListItem(i.ToString(), i.ToString()))
                    Next
                End If
            End If

            ddlPager.SelectedValue = GVBin.PageIndex + 1

            If GVBin.PageIndex = 0 Then
                tmpPagerBtn = TryCast(GVBin.BottomPagerRow.FindControl("pager_first"), ImageButton)
                tmpPagerBtn.Enabled = False

                tmpPagerBtn = TryCast(GVBin.BottomPagerRow.FindControl("pager_previous"), ImageButton)
                tmpPagerBtn.Enabled = False
            End If

            If GVBin.PageIndex = GVBin.PageCount - 1 Then
                tmpPagerBtn = TryCast(GVBin.BottomPagerRow.FindControl("pager_next"), ImageButton)
                tmpPagerBtn.Enabled = False

                tmpPagerBtn = TryCast(GVBin.BottomPagerRow.FindControl("pager_last"), ImageButton)
                tmpPagerBtn.Enabled = False
            End If
        End If
    End Sub

    Protected Sub GVBin_PageIndexChanged(sender As Object, e As System.EventArgs) Handles GVBin.PageIndexChanged
        Dim ddlPager As DropDownList = TryCast(GVBin.BottomPagerRow.FindControl("pager_select"), DropDownList)

        If ddlPager.SelectedValue <> GridView1.PageIndex + 1 Then
            GVBin.PageIndex = ddlPager.SelectedValue - 1
            GVBin.DataSource = Session("dtBin")
            GVBin.DataBind()
        ElseIf ddlPager.SelectedValue = "1" Then
            GVBin.PageIndex = 0
            GVBin.DataSource = Session("dtBin")
            GVBin.DataBind()
        End If

        OtherMDExt.Show()
    End Sub

    Protected Sub GVBin_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GVBin.PageIndexChanging
        GVBin.PageIndex = e.NewPageIndex
        GVBin.DataSource = Session("dtBin")
        GVBin.DataBind()
        OtherMDExt.Show()
    End Sub

    Protected Sub GVBin_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVBin.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("RowIndex"), Label).Text = GVBin.PageIndex * GVBin.PageSize + e.Row.RowIndex + 1
        End Select
    End Sub

    Protected Sub GVOther_DataBound(sender As Object, e As System.EventArgs) Handles GVOther.DataBound
        If GVOther.DataSource IsNot Nothing Then
            Dim ddlPager As DropDownList = TryCast(GVOther.BottomPagerRow.FindControl("pager_selectO"), DropDownList)
            Dim tmpPagerBtn As ImageButton

            If ddlPager.Items.Count = 0 Then
                If ddlPager IsNot Nothing Then
                    For i As Integer = 1 To GVOther.PageCount
                        ddlPager.Items.Add(New ListItem(i.ToString(), i.ToString()))
                    Next
                End If
            End If

            ddlPager.SelectedValue = GVOther.PageIndex + 1

            If GVOther.PageIndex = 0 Then
                tmpPagerBtn = TryCast(GVOther.BottomPagerRow.FindControl("pager_firstO"), ImageButton)
                tmpPagerBtn.Enabled = False

                tmpPagerBtn = TryCast(GVOther.BottomPagerRow.FindControl("pager_previousO"), ImageButton)
                tmpPagerBtn.Enabled = False
            End If

            If GVOther.PageIndex = GVOther.PageCount - 1 Then
                tmpPagerBtn = TryCast(GVOther.BottomPagerRow.FindControl("pager_nextO"), ImageButton)
                tmpPagerBtn.Enabled = False

                tmpPagerBtn = TryCast(GVOther.BottomPagerRow.FindControl("pager_lastO"), ImageButton)
                tmpPagerBtn.Enabled = False
            End If
        End If
    End Sub

    Protected Sub GVOther_PageIndexChanged(sender As Object, e As System.EventArgs) Handles GVOther.PageIndexChanged
        Dim ddlPager As DropDownList = TryCast(GVOther.BottomPagerRow.FindControl("pager_selectO"), DropDownList)

        If ddlPager.SelectedValue <> GVOther.PageIndex + 1 Then
            GVOther.PageIndex = ddlPager.SelectedValue - 1
            GVOther.DataSource = Session("dtOther")
            GVOther.DataBind()
        ElseIf ddlPager.SelectedValue = "1" Then
            GVOther.PageIndex = 0
            GVOther.DataSource = Session("dtOther")
            GVOther.DataBind()
        End If

        OtherMDExt.Show()
    End Sub

    Protected Sub GVOther_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GVOther.PageIndexChanging
        GVOther.PageIndex = e.NewPageIndex
        GVOther.DataSource = Session("dtOther")
        GVOther.DataBind()
        OtherMDExt.Show()
    End Sub

    Protected Sub GVOther_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVOther.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("RowIndex"), Label).Text = GVOther.PageIndex * GVOther.PageSize + e.Row.RowIndex + 1
        End Select
    End Sub

    Protected Sub selectItemMastbtn_Click(sender As Object, e As System.EventArgs) Handles selectItemMastbtn.Click
        Session("CK_TYPE") = CK_TYPE.SelectedValue
        Session("CK_START_DATE") = CK_ACTUAL_DATE.Text.Trim
        Session("CK_WH") = CK_WH.SelectedValue
        Session("LookupWH") = CK_SUB_WH.SelectedValue

        Dim tempDT As DataTable

        tempDT = ViewState("dt")
        Dim itmStr As String = ""
        If tempDT IsNot Nothing AndAlso tempDT.Rows.Count > 0 Then
            For i = 0 To tempDT.Rows.Count - 1
                itmStr = gU.appendToList(itmStr, tempDT.Rows(i).Item("ckd_itm_code").ToString.Trim & "||" & tempDT.Rows(i).Item("ckd_pack_key").ToString.Trim)
            Next
        End If

        Session("SELECTED_ITMBAL") = itmStr

        ScriptManager.RegisterStartupScript(LOOKUPDUP, LOOKUPDUP.GetType, "itemLookup", "ItemMastLookUp(document.myform." & STORER_CODE.ClientID & ".value);", True)
    End Sub

    Protected Sub updateSingleBookCount(ByVal ck_code As String, ByVal ckd_seq As String, ByRef gConn As SqlConnection, ByRef transaction As SqlTransaction, ByVal cFlag As String)
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim SQLString As String = ""

        Dim balQty As Double = 0
        Dim balQty2 As Double = 0

        Dim tempDT As DataTable

        paP = New GlobalDBFunc.DBCmdPara
        SQLString = " SELECT WMS_ITEM_LOC_BAL.ILOC_BAL_QTY, WMS_ITEM_LOC_BAL_S.ILBS_QTY2 " & _
                    " FROM WMS_STOCK_CHECK_D LEFT OUTER JOIN " & _
                    " WMS_ITEM_LOC_BAL ON ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,'') = isnull(WMS_STOCK_CHECK_D.CKD_BATCH_NO,'') AND WMS_ITEM_LOC_BAL.ILOC_LOC = WMS_STOCK_CHECK_D.CKD_LOC AND " & _
                    " ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') = isNULL(WMS_STOCK_CHECK_D.CKD_PALLET_NO,'000') AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_STOCK_CHECK_D.CKD_PACK_KEY AND " & _
                    " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_STOCK_CHECK_D.CKD_ITM_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_STOCK_CHECK_D.STORER_CODE AND " & _
                    " WMS_ITEM_LOC_BAL.IMP_CODE = WMS_STOCK_CHECK_D.IMP_CODE " & _
                    " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL_S.ILOC_SEQ = WMS_ITEM_LOC_BAL.ILOC_SEQ and WMS_STOCK_CHECK_D.CKD_SERIAL = WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO " & _
                    " where wms_stock_check_d.ck_code = " & paP.AP(ck_code) & _
                    " and wms_stock_check_d.ckd_seq = " & paP.AP(ckd_seq) & _
                    " and wms_stock_check_d.storer_code = " & paP.AP(STORER_CODE.SelectedValue) & _
                    " and wms_stock_check_d.imp_code = '" & Session("IMP_CODE") & "'"

        tempDT = gDB.getDataTable(SQLString, gConn, transaction, , paP)


        balQty = gU.decodeEmptyCdbl(tempDT.Rows(0).Item("ILOC_BAL_QTY").ToString.Trim, 0)
        balQty2 = gU.decodeEmptyCdbl(tempDT.Rows(0).Item("ILBS_QTY2").ToString.Trim, 0)


        Dim updateField As String = ""
        Dim updateFieldQ2 As String = ""

        Select Case cFlag
            Case "1"
                updateField = "CKD_ORG_QTY"
                updateFieldQ2 = "CKD_ORG_QTY2"
            Case "2"
                updateField = "CKD_BOOK_QTY"
                updateFieldQ2 = "CKD_BOOK_QTY2"
        End Select

        'DirectCast(GridView1.Rows(i).FindControl("CKD_BOOK_QTY"), HiddenField).Value = balQty
        'DirectCast(GridView1.Rows(i).FindControl("dsp_CKD_BOOK_QTY"), Label).Text = balQty

        SQLString = "UPDATE WMS_STOCK_CHECK_D SET " & updateField & "=" & paP.AP(balQty, SqlDbType.Decimal) & "," & _
                    " CKD_VAR_QTY=" & paP.AP(balQty, SqlDbType.Decimal) & " - CKD_ACTUAL_QTY " & _
                    " where wms_stock_check_d.ck_code = " & paP.AP(ck_code) & _
                    " AND wms_stock_check_d.ckd_seq = " & paP.AP(ckd_seq) & _
                    " and wms_stock_check_d.storer_code = " & paP.AP(STORER_CODE.SelectedValue) & _
                    " and wms_stock_check_d.imp_code = '" & Session("IMP_CODE") & "'"

        gDB.amendData(SQLString, gConn, transaction, paP)

        If CK_IS_Cable.Checked OrElse CK_TYPE.SelectedValue = "CCC" Then

            SQLString = "UPDATE WMS_STOCK_CHECK_D SET " & updateFieldQ2 & "=" & paP.AP(balQty2, SqlDbType.Decimal) & "," & _
                      " CKD_VAR_QTY2=" & paP.AP(balQty2, SqlDbType.Decimal) & " - CKD_ACTUAL_QTY2 " & _
                      " where wms_stock_check_d.ck_code = " & paP.AP(ck_code) & _
                      " AND wms_stock_check_d.ckd_seq = " & paP.AP(ckd_seq) & _
                      " and wms_stock_check_d.storer_code = " & paP.AP(STORER_CODE.SelectedValue) & _
                      " and wms_stock_check_d.imp_code = '" & Session("IMP_CODE") & "'"
            gDB.amendData(SQLString, gConn, transaction, paP)
        End If

    End Sub

    Protected Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Try
            Dim connString As String = ""
            If fu_SkuUpload.HasFile Then
                If STORER_CODE.SelectedValue.ToString <> "" Then
                    If CK_WH.SelectedValue.ToString <> "" Then
                        Dim strFileType As String = Path.GetExtension(fu_SkuUpload.FileName).ToLower()
                        Dim path__1 As String = Server.MapPath("~/TempFiles/") + fu_SkuUpload.PostedFile.FileName
                        fu_SkuUpload.SaveAs(path__1)
                        'Connection String to Excel Workbook
                        If strFileType.Trim() = ".xls" Then
                            connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & path__1 & "';Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""
                        ElseIf strFileType.Trim() = ".xlsx" Then
                            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & path__1 & "';Extended Properties=""Excel 12.0;HDR=Yes;IMEX=2"""
                            'connString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""", path__1)
                        End If
                        Dim query As String = "SELECT [SKU#] FROM [Sheet1$]"
                        Dim conn As New OleDbConnection(connString)
                        If conn.State = ConnectionState.Closed Then
                            conn.Open()
                        End If
                        Dim cmd As New OleDbCommand(query, conn)
                        Dim da As New OleDbDataAdapter(cmd)
                        Dim dtTemp As New DataTable()
                        da.Fill(dtTemp)
                        Dim skuNos As String = ""
                        If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                            For Each row As DataRow In dtTemp.Rows
                                If skuNos = "" Then
                                    skuNos = "'" + row("SKU#").ToString + "'"
                                Else
                                    skuNos = skuNos + ",'" + row("SKU#").ToString + "'"
                                End If
                            Next
                        End If

                        da.Dispose()
                        conn.Close()
                        conn.Dispose()
                        If skuNos <> "" Then
                            Dim sqlString = " SELECT WMS_ITEM_LOC_BAL.PACK_KEY as ckd_pack_key, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE,WMS_ITEM_LOC_BAL.ITM_CODE as ckd_itm_code, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY,  WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC,WMS_ITEM_LOC_BAL.ILOC_LOC as ckd_loc, CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END AS ILOC_BAL_QTY, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,WMS_ITEM_LOC_BAL.ILOC_BATCH_NO as ckd_batch_no,  WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO,WMS_ITEM.ITM_UOM, WMS_ITEM_LOC_BAL.ILOC_SEQ, WMS_ITEM_LOC_BAL_s.ILBS_SEQ,  WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID,WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL,  WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE,  Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE,103) as ILOC_EXPIRY_DATE, Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE,103) as ILOC_MANU_DATE, WMS_ITEM_LOC_BAL.VND_CODE  FROM WMS_ITEM_LOC_BAL INNER JOIN  WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND  WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY  LEFT OUTER JOIN WMS_WH_BIN ON WMS_ITEM_LOC_BAL.ILOC_LOC = WMS_WH_BIN.LOC_KEY  LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ  where wms_item.storer_code = '" + STORER_CODE.SelectedValue.ToString + "'  and wms_item.imp_code = '" + IMP_CODE.Value.ToString + "' and WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0 and  WMS_ITEM_LOC_BAL.ILOC_WH = '" + CK_WH.SelectedValue.ToString + "' and WMS_WH_BIN.WH_CODE='" + CK_WH.SelectedValue.ToString + "'  and WMS_ITEM.ITM_SKU_NO in (" + skuNos + ")  "
                            Dim dtItem = gDB.getDataTable(sqlString)
                            Dim selectedSEQ As String = Session("selectedSEQ").ToString


                            Dim pl_list_no As String = ""
                            Dim lchecker As String = ""
                            Dim next_seq_no = 1




                            For i As Integer = 0 To dtItem.Rows.Count - 1

                                dt.Rows.Add()

                                Dim rows_count = dt.Rows.Count

                                'If listNoDict.ContainsKey(dtItem.Rows(i).Item("ILOC_SEQ").ToString.Trim & "#_#" & dtItem.Rows(i).Item("ILBS_SEQ").ToString.Trim) Then
                                '    pl_list_no = listNoDict(dtItem.Rows(i).Item("ILOC_SEQ").ToString.Trim & "#_#" & dtItem.Rows(i).Item("ILBS_SEQ").ToString.Trim)
                                'End If

                                'If checkerDict.ContainsKey(dtItem.Rows(i).Item("ILOC_SEQ").ToString.Trim & "#_#" & dtItem.Rows(i).Item("ILBS_SEQ").ToString.Trim) Then
                                '    lchecker = checkerDict(dtItem.Rows(i).Item("ILOC_SEQ").ToString.Trim & "#_#" & dtItem.Rows(i).Item("ILBS_SEQ").ToString.Trim)
                                'End If


                                REM **********************
                                REM Modify Here
                                dt.Rows(rows_count - 1).Item("ckd_seq") = next_seq_no.ToString()
                                dt.Rows(rows_count - 1).Item("ckd_itm_code") = dtItem.Rows(i).Item("itm_code").ToString.Trim
                                dt.Rows(rows_count - 1).Item("itm_sku_no") = dtItem.Rows(i).Item("itm_sku_no").ToString.Trim
                                dt.Rows(rows_count - 1).Item("itm_name") = dtItem.Rows(i).Item("itm_name").ToString.Trim
                                dt.Rows(rows_count - 1).Item("itm_uom") = dtItem.Rows(i).Item("itm_uom").ToString.Trim
                                dt.Rows(rows_count - 1).Item("ckd_pack_key") = dtItem.Rows(i).Item("pack_key").ToString.Trim
                                dt.Rows(rows_count - 1).Item("ckd_PALLET_NO") = dtItem.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                                dt.Rows(rows_count - 1).Item("ckd_BATCH_NO") = dtItem.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                                dt.Rows(rows_count - 1).Item("ckd_org_qty") = dtItem.Rows(i).Item("ILOC_BAL_QTY")
                                dt.Rows(rows_count - 1).Item("CKD_BOOK_QTY") = dtItem.Rows(i).Item("ILOC_BAL_QTY")
                                dt.Rows(rows_count - 1).Item("ckd_LOC") = dtItem.Rows(i).Item("ILOC_LOC").ToString.Trim
                                dt.Rows(rows_count - 1).Item("CKD_DRUM_ID") = dtItem.Rows(i).Item("ILBS_DRUM_ID").ToString.Trim
                                dt.Rows(rows_count - 1).Item("CKD_DRUM_LEVEL") = dtItem.Rows(i).Item("ILBS_DRUM_LEVEL").ToString.Trim
                                dt.Rows(rows_count - 1).Item("CKD_SERIAL_NO") = dtItem.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim
                                dt.Rows(rows_count - 1).Item("CKD_UOM2") = dtItem.Rows(i).Item("ILBS_UOM2").ToString.Trim
                                dt.Rows(rows_count - 1).Item("CKD_ORG_QTY2") = dtItem.Rows(i).Item("ILBS_QTY2")
                                dt.Rows(rows_count - 1).Item("CKD_BOOK_QTY2") = dtItem.Rows(i).Item("ILBS_QTY2")

                                dt.Rows(rows_count - 1).Item("ckd_VND_CODE") = dtItem.Rows(i).Item("VND_CODE").ToString.Trim
                                dt.Rows(rows_count - 1).Item("ckd_MANU_DATE") = dtItem.Rows(i).Item("ILOC_MANU_DATE").ToString.Trim
                                dt.Rows(rows_count - 1).Item("ckd_EXPIRY_DATE") = dtItem.Rows(i).Item("ILOC_EXPIRY_DATE").ToString.Trim
                                dt.Rows(rows_count - 1).Item("CKD_STATUS") = "NEW"
                                dt.Rows(rows_count - 1).Item("CKD_CC_LIST_NO") = pl_list_no
                                dt.Rows(rows_count - 1).Item("CKD_CHECKER") = lchecker

                                dt.Rows(rows_count - 1).Item("BN_CSMS_CODE") = dtItem.Rows(i).Item("BN_CSMS_CODE").ToString.Trim
                                dt.Rows(rows_count - 1).Item("WH_CODE") = dtItem.Rows(i).Item("WH_CODE").ToString.Trim
                                REM **********************
                                dt.Rows(rows_count - 1).Item("mFlag") = "N"

                                next_seq_no = next_seq_no + 1
                            Next

                            GridView1.DataSource = dt
                            GridView1.DataBind()
                        End If
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "Please select a Warehouse!", Session("gLang"))
                    End If
                Else
                    uiFun.displayMsgNew(updtPnlAlert, "", "Please select a storer!", Session("gLang"))
                End If

            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "Please select a file!", Session("gLang"))
            End If
        Catch ex As Exception
            uiFun.displayMsgNew(updtPnlAlert, "", "Import failed!", Session("gLang"))
        End Try

    End Sub

    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim nDataSource As DataTable
        Dim sqlString As String
        Dim ck_code As String
        ViewState("CK_CODE") = Request("CK_CODE")
        ck_code = Server.UrlDecode(Request("CK_CODE"))

        sqlString = " SELECT WMS_STOCK_CHECK.CK_CODE,CKD_SERIAL_NO as SERIAL_NO,CK_STATUS,CK_TYPE,CK_REM,CKD_ITM_CODE,WMS_ITEM.ITM_SKU_NO,WMS_ITEM.ITM_NAME,CKD_BATCH_NO,(SUBSTRING(CKD_LOC,1,2)+'.'+SUBSTRING(CKD_LOC,3,2)+'.'+SUBSTRING(CKD_LOC,5,2)+'.'+SUBSTRING(CKD_LOC,7,2)) as CKD_LOC,CKD_ORG_QTY,CKD_REV_QTY, " &
                        " CKD_VAR_QTY,CKD_REV_QTY2,CKD_VAR_QTY2,CKD_REM FROM dbo.WMS_STOCK_CHECK_D JOIN dbo.WMS_STOCK_CHECK ON WMS_STOCK_CHECK_D.CK_CODE = WMS_STOCK_CHECK.CK_CODE " &
                        " JOIN WMS_ITEM ON WMS_STOCK_CHECK_D.CKD_ITM_CODE = WMS_ITEM.ITM_CODE and WMS_STOCK_CHECK_D.STORER_CODE = WMS_ITEM.STORER_CODE where WMS_STOCK_CHECK.CK_CODE='" & gU.dbEncode(ck_code) & "'"

        nDataSource = gDB.getDataTable(sqlString)

        If nDataSource IsNot Nothing AndAlso nDataSource.Rows.Count > 0 Then
            GridTableData.DataSource = nDataSource
            GridTableData.DataBind()
        Else
            GridTableData.DataSource = Nothing
            GridTableData.DataBind()
        End If


        Try
            Response.Clear()
            Response.Buffer = True
            Response.AddHeader("content-disposition", "attachment;filename=CycleCountReport.xls")
            Response.Charset = ""
            Response.ContentType = "application/vnd.ms-excel"
            GridTableData.Visible = True
            Using sw As New System.IO.StringWriter()
                Dim hw As New HtmlTextWriter(sw)
                GridTableData.RenderControl(hw)
                Response.Output.Write(sw.ToString())
                Response.BufferOutput = True
                Response.Flush()
                Response.End()
            End Using

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(control As Control)
        ' Verifies that the control is rendered  
    End Sub
End Class