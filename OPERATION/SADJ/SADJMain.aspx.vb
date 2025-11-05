Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

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

    Private DDFORMAT As String = "DD/MM/YYYY"

    Private moduleAction As String = ""
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
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(AD_WH, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))
            uiFun.load_dropdownBy_ColCode(AD_TYPE, "WMS_STOCK_ADJUST.AD_TYPE", Session("gSelectLabel"))
        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If
            AD_BY.Text = Session("usr_id")
            If AD_DATE.Text = "" Then
                AD_DATE.Text = Now.Date.ToString("dd/MM/yyyy")
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Stock Adjustment Maintenance"
            lbl_ImageHd.Text = "Adjustment Items"
            lbl_AD_CODE.Text = "Adjust Code:"
            lbl_AD_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Organizations:"
            lbl_AD_TYPE.Text = "Type:"
            lbl_AD_REF_NO.Text = "Ref. No.:"
            lbl_AD_REF_CK_CODE.Text = "CC Code:"
            lbl_AD_DATE.Text = "Date:"
            lbl_AD_BY.Text = "Adjusted By:"
            lbl_AD_WH.Text = "Subinventory:"
            lbl_AD_REM.Text = "Remarks:"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            newrow.Text = "Add"
            btnPost.Text = "Post"
            selectItemBtn.Text = "Select Item"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            'btnPost.OnClientClick = "return confirm(""Are you sure to post this record?\r\n(Please save your work before Posting)"");"
            btnPost.OnClientClick = "if (confirm(""Are you sure to post this record?\r\n(Please save your work before Posting"")){getLoad();}else{return false;}"
            If Session("pagemode") = "N" Then
                AD_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "貨量調整維護"
            lbl_ImageHd.Text = "貨量調整詳情"

            lbl_AD_CODE.Text = "調整編號:"
            lbl_AD_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "部門:"
            lbl_AD_TYPE.Text = "類型:"
            lbl_AD_REF_NO.Text = "文件編號:"
            lbl_AD_REF_CK_CODE.Text = "CC Code:"
            lbl_AD_DATE.Text = "日期:"
            lbl_AD_BY.Text = "調整者:"
            lbl_AD_WH.Text = "子庫存:"
            lbl_AD_REM.Text = "備註"

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
            'btnPost.OnClientClick = "return confirm(""確定發布資料?"");"
            btnPost.OnClientClick = "if (confirm(""確定發布資料?"")){getLoad();}else{return false;}"
            If Session("pagemode") = "N" Then
                AD_CODE.Text = "[號碼會自動產生]"
            End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        AD_WH.CssClass = "REQUIRED"
        AD_DATE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            'AD_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("AD_CODE") = ""
            ViewState("STORER_CODE") = ""
            Call BindGV()
        Else
            dt = ViewState("dt")
        End If

        If moduleAction = "SELECTIM" Then
            addItemtoSTA()
        End If

        'cm = New CommonMenu("SADJ", lheader.text, AD_CODE.Text)
        'cm.parentDir = "../../"
        'cm.haveCheckList = "Y"
        'cm.haveAttachments = "Y"
        'cm.haveNotes = "Y"
        'cm.haveTasks = "Y"
        'cm.haveEmail = "Y"
        'cm.haveHistory = "Y"

        'cm.genCM(cmBar)

        selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value, document.myform." & AD_WH.ClientID & ".value);return false;")
        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('OP_SADJ','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("AD_CODE") & "','N');")

        If AD_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
        ElseIf AD_STATUS.Text = "POSTED" Then
            ar.sec_viewMode = "Y"
            btnPost.Visible = False
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
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
        '        Call cU.changeGVLabel(oGridViewRow, e, "Revised Qty", "修訂數量")
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

                CType(e.Row.FindControl("add_itm_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "add_itm_code").ToString.Trim

                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim

                CType(e.Row.FindControl("add_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "add_pack_key").ToString.Trim
                'uiFun.load_dropdown(CType(e.Row.FindControl("add_batch_no"), DropDownList), "select dc_date_code from wms_date_code order by 1", "dc_date_code", "dc_date_code", , Session("gSelectLabel"))
                'CType(e.Row.FindControl("add_batch_no"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "add_batch_no").ToString.Trim
                CType(e.Row.FindControl("disp_add_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "add_batch_no").ToString.Trim
                CType(e.Row.FindControl("add_batch_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "add_batch_no").ToString.Trim

                CType(e.Row.FindControl("itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_name").ToString.Trim
                CType(e.Row.FindControl("add_loc"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "add_loc").ToString.Trim
                'CType(e.Row.FindControl("dsp_add_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "add_loc").ToString.Trim

                'CType(e.Row.FindControl("dsp_add_org_qty"), Label).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "add_org_qty").ToString.Trim)
                'CType(e.Row.FindControl("add_org_qty"), HiddenField).Value = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "add_org_qty").ToString.Trim)
                'CType(e.Row.FindControl("add_rev_qty"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "add_rev_qty").ToString.Trim)
                'CType(e.Row.FindControl("dsp_add_var_qty"), Label).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "add_var_qty").ToString.Trim)
                'CType(e.Row.FindControl("add_var_qty"), HiddenField).Value = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "add_var_qty").ToString.Trim)

                CType(e.Row.FindControl("dsp_add_org_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "add_org_qty").ToString.Trim
                CType(e.Row.FindControl("add_org_qty"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "add_org_qty").ToString.Trim
                CType(e.Row.FindControl("add_rev_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "add_rev_qty").ToString.Trim
                CType(e.Row.FindControl("dsp_add_var_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "add_var_qty").ToString.Trim
                CType(e.Row.FindControl("add_var_qty"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "add_var_qty").ToString.Trim


                CType(e.Row.FindControl("add_rem"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "add_rem").ToString.Trim

                'CType(e.Row.FindControl("dsp_add_org_qty2"), Label).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "add_org_qty2").ToString.Trim)
                'CType(e.Row.FindControl("add_org_qty2"), HiddenField).Value = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "add_org_qty2").ToString.Trim)
                'CType(e.Row.FindControl("add_rev_qty2"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "add_rev_qty2").ToString.Trim)
                'CType(e.Row.FindControl("dsp_add_var_qty2"), Label).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "add_var_qty2").ToString.Trim)
                'CType(e.Row.FindControl("add_var_qty2"), HiddenField).Value = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "add_var_qty2").ToString.Trim)

                CType(e.Row.FindControl("dsp_add_org_qty2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "add_org_qty2").ToString.Trim
                CType(e.Row.FindControl("add_org_qty2"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "add_org_qty2").ToString.Trim
                CType(e.Row.FindControl("add_rev_qty2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "add_rev_qty2").ToString.Trim
                CType(e.Row.FindControl("dsp_add_var_qty2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "add_var_qty2").ToString.Trim
                CType(e.Row.FindControl("add_var_qty2"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "add_var_qty2").ToString.Trim

                CType(e.Row.FindControl("DRUM_LEVEL"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DRUM_LEVEL").ToString.Trim
                CType(e.Row.FindControl("DRUM_ID"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DRUM_ID").ToString.Trim
                CType(e.Row.FindControl("ITM_UOM"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ITM_UOM").ToString.Trim
                CType(e.Row.FindControl("dsp_ADD_SERIAL_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ITM_SERIAL_NO").ToString.Trim
                CType(e.Row.FindControl("ADD_SERIAL_NO"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ADD_SERIAL_NO").ToString.Trim
                CType(e.Row.FindControl("add_pallet_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "add_pallet_no").ToString.Trim
                'CType(e.Row.FindControl("add_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "add_pallet_no").ToString.Trim
                'CType(e.Row.FindControl("dsp_add_pallet_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "add_pallet_no").ToString.Trim


                'CType(e.Row.FindControl("add_expiry_date"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "add_expiry_date").ToString.Trim
                'CType(e.Row.FindControl("dsp_add_expiry_date"), Label).Text = DataBinder.Eval(e.Row.DataItem, "add_expiry_date").ToString.Trim
                CType(e.Row.FindControl("add_manu_date"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "add_manu_date").ToString.Trim
                'CType(e.Row.FindControl("dsp_add_manu_date"), Label).Text = DataBinder.Eval(e.Row.DataItem, "add_manu_date").ToString.Trim

                CType(e.Row.FindControl("add_rev_qty"), TextBox).Attributes.Add("onkeypress", "return maskKey(event); ")
                CType(e.Row.FindControl("add_rev_qty"), TextBox).Attributes.Add("onkeyup", "javascript:calQty(this.value, " &
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_add_org_qty"), Label).ClientID) & "', " &
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_add_var_qty"), Label).ClientID) & "', " &
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("add_org_qty"), HiddenField).ClientID) & "', " &
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("add_var_qty"), HiddenField).ClientID) & "' " &
                                                                                ");")


                CType(e.Row.FindControl("add_rev_qty2"), TextBox).Attributes.Add("onkeypress", "return maskKey(event); ")
                CType(e.Row.FindControl("add_rev_qty2"), TextBox).Attributes.Add("onkeyup", "javascript:calQty(this.value, " &
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_add_org_qty2"), Label).ClientID) & "', " &
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_add_var_qty2"), Label).ClientID) & "', " &
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("add_org_qty2"), HiddenField).ClientID) & "', " &
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("add_var_qty2"), HiddenField).ClientID) & "' " &
                                                                                ");")

                'Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp"), Image)
                'nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                'nImage.Attributes.Add("onclick", "LocLookUp('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_add_loc"), Label).ClientID) & "', " & _
                '                                            "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("add_loc"), HiddenField).ClientID) & "', " & _
                '                                            "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_add_org_qty"), Label).ClientID) & "', " & _
                '                                            "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("add_org_qty"), HiddenField).ClientID) & "', " & _
                '                                            "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "add_itm_code").ToString.Trim) & "', " & _
                '                                            "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "add_pack_key").ToString.Trim) & "'    " & _
                '                                            ")")

                'nImage.Attributes.Add("onclick", "ItemLocLookUp(document.myform." & STORER_CODE.ClientID & ".value, " & _
                '                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "add_itm_code").ToString.Trim) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "add_pack_key").ToString.Trim) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_add_loc"), Label).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("add_loc"), HiddenField).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_add_pallet_no"), Label).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("add_pallet_no"), HiddenField).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_add_org_qty"), Label).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("add_org_qty"), HiddenField).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("disp_add_batch_no"), Label).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("add_batch_no"), HiddenField).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_ADD_EXPIRY_DATE"), Label).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ADD_EXPIRY_DATE"), HiddenField).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_ADD_MANU_DATE"), Label).ClientID) & "', " & _
                '                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("ADD_MANU_DATE"), HiddenField).ClientID) & "' " & _
                '                      ")")


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
            Dim seq_string As String = "select MAX(CAST(ad_SEQ AS int)) + 1 from WMS_STOCK_ISSUE_D " &
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and AD_CODE = '" & gU.dbEncode(AD_CODE.Text.Trim) & "' "
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
            dt.Rows(rows_count - 1).Item("ad_seq") = ViewState("n_cur_seq").ToString
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

        If AD_WH.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_AD_WH.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_AD_WH.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If AD_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_AD_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_AD_DATE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        ElseIf Not gU.isValidDate(AD_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_AD_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_AD_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                Dim loc As String = CType(GridView1.Rows(i).FindControl("add_loc"), TextBox).Text

                If loc = "" Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Location Cannot Be Empty!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "位置不能空白!", Session("gLang"))
                    End If
                    Return False
                End If

                If CType(GridView1.Rows(i).FindControl("add_rev_qty"), TextBox).Text <> "" Then
                    If uiFun.gvValidate(Me, dt, "add_rev_qty", "Revised Qty",
                                     CType(GridView1.Rows(i).FindControl("add_rev_qty"), TextBox).Text) = False Then Return False
                Else
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Revised Qty Cannot Be Empty!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "修訂數量不能空白!", Session("gLang"))
                    End If
                    Return False
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
        Dim gConn As SqlConnection
        Dim dupSQL As String = ""
        Dim dupTbl As New DataTable

        Dim varQty As Double = 0
        Dim varQty2 As Double = 0

        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("SADJ", gConn, transaction)
                    'nextNo = AD_CODE.Text
                    REM **********************

                    dupSQL = "select 1 from wms_stock_adjust " &
                                "where ad_code = '" & gU.dbEncode(nextNo) & "' " &
                                "and imp_code='" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into wms_stock_adjust (" &
                        "ad_code, imp_code, storer_code, " &
                        "ad_status, ad_date, ad_by, " &
                        "ad_type, ad_ref_no, " &
                        "ad_wh, ad_rem, " &
                         "sys_cb, sys_cd, sys_lub, sys_lud)" &
                        "values ( " &
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," &
                        gU.convdbNVCData(gU.dbEncode(AD_STATUS.Text.Trim)) & ", " & gU.convdbDate(gU.dbEncode(AD_DATE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(AD_BY.Text.Trim)) & "," &
                        gU.convdbNVCData(gU.dbEncode(AD_TYPE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(AD_REF_NO.Text.Trim)) & "," &
                        gU.convdbNVCData(gU.dbEncode(AD_WH.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(AD_REM.Text.Trim)) & "," &
                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                        REM **********************s

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                            uiFun.reOrderDetails(dt, "ad_seq")

                            For Each rows As DataRow In dt.Rows
                                itemSQL = ""

                                REM **********************
                                REM Modify Here
                                Select Case rows.Item("mFlag")
                                    Case "N"

                                        varQty = 0
                                        If gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0) Then
                                            varQty = gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0)
                                            varQty = Math.Abs(varQty) * -1

                                        ElseIf gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0) Then
                                            varQty = gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0)
                                        End If

                                        varQty2 = 0
                                        If gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0) Then
                                            varQty2 = gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0)
                                            varQty2 = Math.Abs(varQty2) * -1

                                        ElseIf gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0) Then
                                            varQty2 = gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0)
                                        End If

                                        itemSQL = "insert into wms_stock_adjust_d (" &
                                                "ad_code, imp_code, storer_code, " &
                                                "ad_seq, add_itm_code, add_pack_key, add_batch_no, " &
                                                "add_loc, add_org_qty, add_rev_qty, " &
                                                "add_var_qty, add_rem, add_pallet_no, " &
                                                "add_expiry_date, add_manu_date, " &
                                                "ADD_ORG_QTY2, ADD_REV_QTY2, ADD_VAR_QTY2, ADD_SERIAL_NO," &
                                                "sys_cb, sys_cd, sys_lub, sys_lud,DRUM_ID,DRUM_LEVEL) " &
                                                "values (" &
                                                gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ad_seq").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_itm_code").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_pack_key").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_batch_no").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_loc").ToString.Trim, ""))) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_org_qty").ToString.Trim, "NULL")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_rev_qty").ToString.Trim, "NULL")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(varQty, "NULL")) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_rem").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, ""))) & ", " &
                                                gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("add_expiry_date").ToString.Trim, ""))) & ", " &
                                                gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("add_manu_date").ToString.Trim, ""))) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_org_qty2").ToString.Trim, "NULL")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_rev_qty2").ToString.Trim, "NULL")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(varQty2, "NULL")) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, ""))) & ", " &
                                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate(),'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUM_ID").ToString.Trim, "NULL")) & "','" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUM_LEVEL").ToString.Trim, "NULL")) & "') "

                                        '"(N'" & Session("imp_code") & "', N'" & gU.dbEncode(STORER_CODE.Text) & "', N'" & nextNo & "', " & _
                                End Select
                                REM **********************

                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                            Next
                        End If
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Stock Adjustment!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "貨量調整資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_stock_adjust set " &
                                    "ad_status = " & gU.convdbNVCData(gU.dbEncode(AD_STATUS.Text)) & ", " &
                                    "ad_date = " & gU.convdbDate(gU.dbEncode(AD_DATE.Text.Trim)) & ", " &
                                    "ad_by = " & gU.convdbNVCData(gU.dbEncode(AD_BY.Text.Trim)) & ", " &
                                    "ad_type = " & gU.convdbNVCData(gU.dbEncode(AD_TYPE.SelectedValue)) & ", " &
                                    "ad_ref_no = " & gU.convdbNVCData(gU.dbEncode(AD_REF_NO.Text.Trim)) & ", " &
                                    "ad_wh = " & gU.convdbNVCData(gU.dbEncode(AD_WH.SelectedValue)) & ", " &
                                    "ad_rem = " & gU.convdbNVCData(gU.dbEncode(AD_REM.Text.Trim)) & ", " &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and ad_code = '" & gU.dbEncode(AD_CODE.Text) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "ad_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"

                                    varQty = 0

                                    If gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0) Then
                                        varQty = gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0)
                                        varQty = Math.Abs(varQty) * -1

                                    ElseIf gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0) Then
                                        varQty = gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0)
                                    End If

                                    varQty2 = 0
                                    If gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0) Then
                                        varQty2 = gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0)
                                        varQty2 = Math.Abs(varQty2) * -1

                                    ElseIf gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0) Then
                                        varQty2 = gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0)
                                    End If

                                    itemSQL = "insert into wms_stock_adjust_d (" &
                                                "ad_code, imp_code, storer_code, " &
                                                "ad_seq, add_itm_code, add_pack_key, add_batch_no, " &
                                                "add_loc, add_org_qty, add_rev_qty, " &
                                                "add_var_qty, add_rem, add_pallet_no, " &
                                                "add_expiry_date, add_manu_date, " &
                                                "ADD_ORG_QTY2, ADD_REV_QTY2, ADD_VAR_QTY2, ADD_SERIAL_NO," &
                                                "sys_cb, sys_cd, sys_lub, sys_lud,DRUM_ID,DRUM_LEVEL) " &
                                                "values (" &
                                                gU.convdbNVCData(gU.dbEncode(AD_CODE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ad_seq").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_itm_code").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_pack_key").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_batch_no").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_loc").ToString.Trim, ""))) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_org_qty").ToString.Trim, "0")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_rev_qty").ToString.Trim, "0")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(varQty, "0")) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_rem").ToString.Trim, ""))) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, ""))) & ", " &
                                                gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("add_expiry_date").ToString.Trim, ""))) & ", " &
                                                gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("add_manu_date").ToString.Trim, ""))) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_org_qty2").ToString.Trim, "NULL")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_rev_qty2").ToString.Trim, "NULL")) & ", " &
                                                gU.dbEncode(gU.decodeNullOrEmpty(varQty2, "NULL")) & ", " &
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, ""))) & ", " &
                                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate(),'" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUM_ID").ToString.Trim, "NULL")) & "','" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUM_LEVEL").ToString.Trim, "NULL")) & "') "

                                Case "D"
                                    itemSQL = "delete from wms_stock_adjust_d " &
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "and ad_code = '" & gU.dbEncode(AD_CODE.Text.Trim) & "' " &
                                            "and ad_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("ad_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    varQty = 0

                                    If gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0) Then
                                        varQty = gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0)
                                        varQty = Math.Abs(varQty) * -1

                                    ElseIf gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0) Then
                                        varQty = gU.decodeEmptyCdbl(rows.Item("add_rev_qty").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_org_qty").ToString.Trim, 0)
                                    End If

                                    varQty2 = 0
                                    If gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0) > gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0) Then
                                        varQty2 = gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0)
                                        varQty2 = Math.Abs(varQty2) * -1

                                    ElseIf gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0) < gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0) Then
                                        varQty2 = gU.decodeEmptyCdbl(rows.Item("add_rev_qty2").ToString.Trim, 0) - gU.decodeEmptyCdbl(rows.Item("add_org_qty2").ToString.Trim, 0)
                                    End If

                                    itemSQL = "update wms_stock_adjust_d set " &
                                              "ad_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ad_seq").ToString.Trim, ""))) & ", " &
                                              "add_itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_itm_code").ToString.Trim, ""))) & ", " &
                                              "add_pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_pack_key").ToString.Trim, ""))) & ", " &
                                              "add_loc = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_loc").ToString.Trim, ""))) & ", " &
                                              "add_org_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_org_qty").ToString.Trim, "NULL")) & ", " &
                                              "add_rev_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_rev_qty").ToString.Trim, "NULL")) & ", " &
                                              "add_var_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(varQty, "NULL")) & ", " &
                                              "add_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_rem").ToString.Trim, ""))) & ", " &
                                              "add_pallet_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, ""))) & ", " &
                                              "add_expiry_date= " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("add_expiry_date").ToString.Trim, ""))) & ", " &
                                              "add_manu_date= " & gU.convdbDate(gU.dbEncode(gU.decodeNull(rows.Item("add_manu_date").ToString.Trim, ""))) & ", " &
                                              "add_org_qty2 = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_org_qty2").ToString.Trim, "NULL")) & ", " &
                                              "add_rev_qty2 = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("add_rev_qty2").ToString.Trim, "NULL")) & ", " &
                                              "add_var_qty2 = " & gU.dbEncode(gU.decodeNullOrEmpty(varQty2, "NULL")) & ", " &
                                              "DRUM_ID = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUM_ID").ToString.Trim, "NULL")) & ", " &
                                              "DRUM_LEVEL = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("DRUM_LEVEL").ToString.Trim, "NULL")) & ", " &
                                              "sys_lub = '" & Session("usr_id") & "', " &
                                              "sys_lud = Getdate() " &
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                            "and ad_code = '" & gU.dbEncode(AD_CODE.Text.Trim) & "' " &
                                            "and ad_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("old_seq").ToString.Trim, "")) & "'"

                            End Select
                            REM **********************

                            'Response.Write(itemSQL)
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
                    AD_CODE.Text = nextNo
                    ViewState("AD_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    AD_CODE.ForeColor = Drawing.Color.Black
                    AD_CODE.Font.Size = 10
                    'AD_CODE.CssClass = ""
                    STORER_CODE.CssClass = ""

                    REM **********************
                End If

                If flag <> "Y" Then
                    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                End If

                Call BindGV()
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
        End If
    End Sub

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
        If ViewState("AD_CODE") <> "" Then
            pk_code = ViewState("AD_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("AD_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))

            ViewState("AD_CODE") = pk_code
            ViewState("STORER_CODE") = storerCode
        End If
        REM **********************

        AD_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            AD_STATUS.Text = "NEW"
            AD_CODE.ForeColor = Drawing.Color.Red

            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = "select CONVERT(VARCHAR(30), wms_stock_adjust.AD_DATE, " & DDFORMAT & ") as AD_DATE, wms_stock_adjust.* from wms_stock_adjust " &
                        "where wms_stock_adjust.ad_code = '" & gU.dbEncode(pk_code) & "' " &
                        "and wms_stock_adjust.imp_code = '" & Session("IMP_CODE") & "' " &
                        "and wms_stock_adjust.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                AD_CODE.Text = dt.Rows(0).Item("AD_code").ToString
                'STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString

                AD_STATUS.Text = dt.Rows(0).Item("AD_STATUS").ToString
                AD_TYPE.SelectedValue = dt.Rows(0).Item("AD_TYPE").ToString
                AD_REF_NO.Text = dt.Rows(0).Item("AD_REF_NO").ToString
                AD_REF_CK_CODE.Text = dt.Rows(0).Item("AD_REF_CK_CODE").ToString
                AD_DATE.Text = dt.Rows(0).Item("AD_DATE").ToString
                AD_BY.Text = dt.Rows(0).Item("AD_BY").ToString
                AD_WH.SelectedValue = dt.Rows(0).Item("AD_WH").ToString
                AD_REM.Text = dt.Rows(0).Item("AD_REM").ToString

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , , True)
                End If

                If AD_CODE.Text <> "" Then
                    'AD_CODE.ReadOnly = True
                    'AD_CODE.BorderWidth = 0
                    AD_CODE.BackColor = Drawing.Color.Transparent
                End If

                If AD_BY.Text <> "" Then
                    AD_BY.ReadOnly = True
                    AD_BY.BorderWidth = 0
                    AD_BY.BackColor = Drawing.Color.Transparent
                End If

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "SELECT wms_stock_adjust_d.IMP_CODE," &
                    " wms_stock_adjust_d.STORER_CODE,  wms_stock_adjust_d.AD_CODE,wms_stock_adjust_d.DRUM_ID,wms_stock_adjust_d.DRUM_LEVEL,wms_item.ITM_UOM,wms_item.ITM_SERIAL_NO, " &
                    " wms_stock_adjust_d.AD_SEQ,  wms_stock_adjust_d.ADD_ITM_CODE, " &
                    " wms_stock_adjust_d.ADD_PACK_KEY, wms_stock_adjust_d.ADD_LOC,  wms_stock_adjust_d.ADD_REV_QTY, " &
                    " wms_stock_adjust_d.ADD_ORG_QTY, wms_stock_adjust_d.ADD_VAR_QTY, " &
                    " wms_stock_adjust_d.ADD_REM,  wms_stock_adjust_d.ADD_PALLET_NO, " &
                    " wms_stock_adjust_d.ADD_BATCH_NO,  wms_stock_adjust_d.ADD_VND_CODE, " &
                    " wms_item.itm_sku_no, " &
                    " CONVERT(VARCHAR(30), wms_stock_adjust_d.ADD_MANU_DATE," & DDFORMAT & ") as ADD_MANU_DATE,  " &
                    " CONVERT(VARCHAR(30), wms_stock_adjust_d.ADD_EXPIRY_DATE," & DDFORMAT & ") as ADD_EXPIRY_DATE, " &
                    " ADD_ORG_QTY2, ADD_REV_QTY2, ADD_VAR_QTY2, ADD_SERIAL_NO, " &
                    " wms_item.itm_name, 'U' as mFlag, wms_stock_adjust_d.ad_seq as old_seq " &
                    " from wms_stock_adjust_d left outer join wms_item on " &
                    " wms_stock_adjust_d.imp_code = wms_item.imp_code " &
                    " and wms_stock_adjust_d.storer_code = wms_item.storer_code " &
                    " and wms_stock_adjust_d.add_itm_code = wms_item.itm_code " &
                    " and wms_stock_adjust_d.ADD_PACK_KEY = wms_item.PACK_KEY " &
                    " where wms_stock_adjust_d.ad_code = '" & gU.dbEncode(pk_code) & "' " &
                    " and wms_stock_adjust_d.storer_code = '" & gU.dbEncode(storerCode) & "' " &
                    " and wms_stock_adjust_d.imp_code = '" & Session("IMP_CODE") & "'"

        SQLString = SQLString & " order by CONVERT(int, wms_stock_adjust_d.ad_seq)"
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

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        Dim cancelSql As String = "update wms_stock_adjust " &
                     "set ad_status = 'CANCELLED', " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and ad_code = '" & gU.dbEncode(AD_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction


        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql)

            transaction.Commit()

            AD_STATUS.Text = "CANCELLED"

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
        Dim qtyTbl As New DataTable
        Dim imTbl As New DataTable
        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            If GridView1.Rows.Count > 0 Then
                Call save("Y")

                Dim StkChkDt As DataTable
                Dim SQLStkChk As String = "select * from WMS_STOCK_ADJUST_D where AD_CODE='" & AD_CODE.Text.Trim & "' and (select count(1) from [WMS_STOCK_ADJUST_D] where ADD_REV_QTY<0 and  AD_CODE='" & AD_CODE.Text.Trim & "')=0"
                StkChkDt = gDB.getDataTable(SQLStkChk, gConn, transaction)

                If 1 = 1 Then 'StkChkDt IsNot Nothing AndAlso StkChkDt.Rows.Count > 0


                    For Each rows As DataRow In dt.Rows
                        st.STORER_CODE = STORER_CODE.SelectedValue
                        st.ITM_CODE = gU.decodeNull(rows.Item("add_itm_code").ToString.Trim, "")
                        st.PACK_KEY = gU.decodeNull(rows.Item("add_pack_key").ToString.Trim, "")
                        st.IO_CUST_CODE = ""
                        st.IO_AREA = ""
                        st.IO_DOC = "SADJ"
                        st.IO_DOC_ID = AD_CODE.Text.Trim
                        st.IO_CBM = 0
                        st.IO_KG = 0
                        If AD_TYPE.SelectedValue = "EBS" Then
                            st.IO_WH = gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, "")
                        Else
                            st.IO_WH = AD_WH.SelectedValue
                        End If
                        st.PALLET_NO = gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, "")
                        st.IO_LOC = gU.decodeNull(rows.Item("add_loc").ToString.Trim, "")
                        st.lO_BATCH_NO = gU.decodeNull(rows.Item("add_batch_no").ToString.Trim, "")
                        st.IO_EXPIRY_DATE = gU.decodeNull(rows.Item("add_expiry_date").ToString.Trim, "")
                        st.IO_MANU_DATE = gU.decodeNull(rows.Item("add_manu_date").ToString.Trim, "")

                        Dim txQty As Double = 0

                        If gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") >
                            gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                            txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0")

                            st.IO_QTY = txQty
                            st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                            st.UpdateStockTrans("IN", gConn, transaction)
                            st.UpdateStockBalTrans("IN", gConn, transaction)

                            If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                                st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                                st.IOS_QTY2 = 1
                                st.UpdateStockSerialTrans("IN", gConn, transaction)
                                st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                            End If
                        ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") <
                            gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                            txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0")

                            st.IO_QTY = txQty
                            st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                            st.UpdateStockTrans("OUT", gConn, transaction)
                            st.UpdateStockBalTrans("OUT", gConn, transaction)

                            If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                                st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                                st.IOS_QTY2 = 1
                                st.UpdateStockSerialTrans("OUT", gConn, transaction)
                                st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                            End If
                        ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") >
                            gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") Then

                            txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0")


                            Dim SQLString As String
                            Dim drumDt As DataTable
                            SQLString = "SELECT ILBS_DRUM_ID, ILBS_DRUM_LEVEL FROM WMS_ITEM_LOC_BAL_S WHERE " &
                                "IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue.Trim) & "' " &
                                "AND ITM_CODE = '" & gU.decodeNull(rows.Item("add_itm_code").ToString.Trim, "") & "' " &
                                "AND PACK_KEY = '" & gU.decodeNull(rows.Item("add_pack_key").ToString.Trim, "") & "' " &
                                "AND ILBS_SERIAL_NO = '" & gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") & "' AND ILBS_QTY2 > 0 "
                            REM **********************

                            drumDt = gDB.getDataTable(SQLString, gConn, transaction)

                            If drumDt.Rows.Count > 0 Then
                                st.IOS_DRUM_ID = drumDt.Rows(0).Item("ILBS_DRUM_ID").ToString
                                st.IOS_DRUM_LEVEL = drumDt.Rows(0).Item("ILBS_DRUM_LEVEL")
                            End If

                            st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                            st.IOS_QTY2 = txQty
                            st.IO_QTY = 1
                            st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                            st.UpdateStockTrans("IN", gConn, transaction)
                            st.UpdateStockBalTrans("IN", gConn, transaction)
                            If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                                st.IOS_SL = "Y"
                                st.UpdateStockSerialTrans("IN", gConn, transaction)
                                st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                            End If

                        ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") <
                            gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") Then

                            txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0")
                            st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                            st.IOS_QTY2 = txQty
                            st.IO_QTY = 1
                            Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, ""), gConn, transaction)
                            st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                            st.UpdateStockTrans("OUT", gConn, transaction)
                            st.UpdateStockBalTrans("OUT", gConn, transaction)
                            If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                                st.IOS_SL = "Y"
                                st.UpdateStockSerialTrans("OUT", gConn, transaction)
                                st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                            End If
                        ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") =
                            gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                            'If gU.decodeNull(rows.Item("add_manu_date").ToString.Trim, "") <> "" Then

                            txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0")

                                st.IO_QTY = txQty
                                st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                                st.UpdateStockTrans("IN", gConn, transaction)
                                st.UpdateStockBalTrans("IN", gConn, transaction)

                                If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                                    st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                                    st.IOS_QTY2 = 1
                                    st.UpdateStockSerialTrans("IN", gConn, transaction)
                                    st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                                End If
                            'Else
                            '    If transaction IsNot Nothing Then
                            '        transaction.Rollback()
                            '    End If

                            '    If Session("gLang") = "E" Then
                            '        uiFun.displayMsg(Me, "", "Revised Qty Equal to Orginal Qty.\r\nYou have to adjust the Qty in order to POST!", Session("gLang"))
                            '    Else
                            '        uiFun.displayMsg(Me, "", "修訂數量和原來數量一樣。\r\你必須調整數量才可發布物件!", Session("gLang"))
                            '    End If
                            '    Exit Sub
                            'End If

                        End If
                    Next

                    updtSql = "update wms_stock_adjust " &
                                "set ad_status = 'POSTED', " &
                                "sys_lub = N'" & Session("usr_id") & "', " &
                                "sys_lud = Getdate() " &
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                "and ad_code = '" & gU.dbEncode(AD_CODE.Text) & "' "

                    gDB.amendData(updtSql)

                    transaction.Commit()

                    AD_STATUS.Text = "POSTED"
                    ar.sec_viewMode = "Y"
                    btnPost.Visible = False
                    ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

                    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                Else
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "-ve stock found!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "-ve stock found!", Session("gLang"))
                    End If
                End If


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

            load_ModalPopupExtender.Hide()
        End Try
    End Sub

    Protected Sub addItemtoSTA()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(ad_SEQ AS int)) + 1 from wms_stock_adjust_d " &
                                        "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                        "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                        "and ad_code = '" & gU.dbEncode(AD_CODE.Text.Trim) & "' "
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
            Dim sadj_dt As DataTable

            Dim itemPackList As String = ""
            Dim itemListarray As String()
            Dim packKeyListarray As String()
            Dim seqListArray As String()

            Dim seqKeyList As String = ""

            itemListarray = Split(itemList.Value, ", ")
            packKeyListarray = Split(packKeyList.Value, ", ")
            seqListArray = Split(seqList.Value, ", ")

            'If itemListarray.Count = 0 Then
            '    If itemList.Value <> "" Then
            '        itemPackList = Server.HtmlDecode(itemList.Value) & "_000_" & Server.HtmlDecode(packKeyList.Value)
            '    End If
            'Else
            '    For i = 0 To itemListarray.Count - 1
            '        itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemListarray(i)) & "_000_" & Server.HtmlDecode(packKeyListarray(i)))
            '    Next
            'End If
            If seqListArray.Count = 0 Then
                If seqList.Value <> "" Then
                    seqKeyList = Server.HtmlDecode(seqList.Value)
                End If
            Else
                For i = 0 To seqListArray.Count - 1
                    seqKeyList = gU.appendToList(seqKeyList, seqListArray(i))
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

            If seqKeyList <> "" Then
                addSQL = "'" & Replace(seqKeyList, ", ", "', '") & "'"
            Else
                addSQL = "NULL"
            End If

            SQLString = " SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, " &
                        " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC, CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END AS ILOC_BAL_QTY, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, " &
                        " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL.ILOC_SEQ, WMS_ITEM_LOC_BAL_s.ILBS_SEQ, " &
                        " WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID,WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL, " &
                        " Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE, WMS_ITEM_LOC_BAL.VND_CODE " &
                        " FROM WMS_ITEM_LOC_BAL INNER JOIN " &
                        " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND " &
                        " WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " &
                        " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " &
                        " where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        " and wms_item.imp_code = '" & Session("IMP_CODE") & "' " &
                        " and Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'') in (" & addSQL & ") "

            SQLString = SQLString & " order by wms_item.ITM_CODE, wms_item.PACK_KEY, VND_CODE"
            REM **********************
            sadj_dt = gDB.getDataTable(SQLString)

            For i As Integer = 0 To sadj_dt.Rows.Count - 1
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
                dt.Rows(rows_count - 1).Item("ad_seq") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("add_itm_code") = sadj_dt.Rows(i).Item("itm_code")
                dt.Rows(rows_count - 1).Item("itm_sku_no") = sadj_dt.Rows(i).Item("itm_sku_no")
                dt.Rows(rows_count - 1).Item("itm_name") = sadj_dt.Rows(i).Item("itm_name")
                dt.Rows(rows_count - 1).Item("add_pack_key") = sadj_dt.Rows(i).Item("pack_key")
                'dt.Rows(rows_count - 1).Item("add_org_qty") = sadj_dt.Rows(i).Item("itm_balance")

                dt.Rows(rows_count - 1).Item("ADD_PALLET_NO") = sadj_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("ADD_BATCH_NO") = sadj_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("ADD_ORG_QTY") = sadj_dt.Rows(i).Item("ILOC_BAL_QTY")
                dt.Rows(rows_count - 1).Item("ADD_LOC") = sadj_dt.Rows(i).Item("ILOC_LOC").ToString.Trim
                dt.Rows(rows_count - 1).Item("ADD_SERIAL_NO") = sadj_dt.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("ADD_ORG_QTY2") = sadj_dt.Rows(i).Item("ILBS_QTY2")

                dt.Rows(rows_count - 1).Item("ADD_VND_CODE") = sadj_dt.Rows(i).Item("VND_CODE").ToString.Trim
                dt.Rows(rows_count - 1).Item("ADD_MANU_DATE") = sadj_dt.Rows(i).Item("ILOC_MANU_DATE").ToString.Trim
                dt.Rows(rows_count - 1).Item("ADD_EXPIRY_DATE") = sadj_dt.Rows(i).Item("ILOC_EXPIRY_DATE")

                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

            Next
            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub


End Class
