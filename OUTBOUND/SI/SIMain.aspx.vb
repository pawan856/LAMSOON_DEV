Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OUTBOUND_SI_SIMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private st As New StockTrans
    Private moduleAction As String = ""
    Private checkBoxValue() As String = {"Y", ""}
    Private DDFORMAT As String = "DD/MM/YYYY"

    Private dt As New DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DDFORMAT = gU.getConfig("DDFORMATNO")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_SI", Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(IS_PROJECT_NO, "select PRJ_CODE, PRJ_NAME from WMS_PROJECT ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(IS_WH, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))
        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
                Call setDefStorerInfo()
            End If
            IS_ISSUED_BY.Text = Session("usr_id")
            If IS_DATE.Text = "" Then
                IS_DATE.Text = Now.Date.ToString("dd/MM/yyyy")
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Stock Issue Maintenance"
            lbl_ImageHd.Text = "Issued Items"
            lbl_IS_CODE.Text = "Issue Code:"
            lbl_IS_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Storer:"
            lbl_IS_TYPE.Text = "Type:"
            lbl_IS_DATE.Text = "Date:"
            lbl_IS_PLAN_DATE.Text = "Planned Return Date:"
            lbl_IS_RETURN_DATE.Text = "Actual Returned Date"
            lbl_IS_RETURN_DOC_NO.Text = "Stock Return No."
            lbl_IS_ISSUED_BY.Text = "Issued By:"
            lbl_IS_TOT_PALLET.Text = "Total Pallet:"
            lbl_IS_PROJECT_NO.Text = "Project:"
            lbl_IS_REQ_BY.Text = "Requested By:"
            lbl_IS_REQ_TEL.Text = "Tel:"
            lbl_IS_REQ_EMAIL.Text = "Email:"
            lbl_IS_REF_NO.Text = "Ref. No.:"
            lbl_IS_WH.Text = "Warehouse:"
            lbl_IS_BATCH_NO.Text = "Batch No.:"

            lbl_IS_REM.Text = "Remarks"
            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            newrow.Text = "Add"
            btnPost.Text = "Post"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            'btnPost.OnClientClick = "return confirm(""Are you sure to post this record?\r\n(Please save your work before Posting)"");"
            btnPost.OnClientClick = "if (confirm(""Are you sure to post this record?\r\n(Please save your work before Posting)"")){getLoad();}else{return false;}"
            If Session("pagemode") = "N" Then
                IS_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "出貨維護"
            lbl_ImageHd.Text = "出貨詳情"
            lbl_IS_CODE.Text = "出貨號碼:"
            lbl_IS_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "貨主:"
            lbl_IS_TYPE.Text = "類型:"
            lbl_IS_DATE.Text = "日期:"
            lbl_IS_ISSUED_BY.Text = "簽發者:"
            lbl_IS_TOT_PALLET.Text = "總貨板數:"
            lbl_IS_PROJECT_NO.Text = "項目:"
            lbl_IS_REQ_BY.Text = "要求者:"
            lbl_IS_REQ_TEL.Text = "電話號碼:"
            lbl_IS_REQ_EMAIL.Text = "電子郵箱:"
            lbl_IS_REF_NO.Text = "文件編號:"
            lbl_IS_WH.Text = "倉庫:"
            lbl_IS_BATCH_NO.Text = "批號:"
            lbl_IS_REM.Text = "備註:"

            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "儲存"
            saveBtn2.Text = "儲存"
            CancelBtn.Text = "取消"
            newrow.Text = "新增"
            btnPost.Text = "發布"
            CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"
            'btnPost.OnClientClick = "return confirm(""確定發布資料?"");"
            btnPost.OnClientClick = "if (confirm(""確定發布資料?"")){getLoad();}else{return false;}"
            If Session("pagemode") = "N" Then
                IS_CODE.Text = "[号码会自动产生]"
            End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS
        IS_TYPE.CssClass = "REQUIRED"
        IS_WH.CssClass = "REQUIRED"
        IS_DATE.CssClass = "REQUIRED"
        IS_PLAN_DATE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            'IS_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("IS_CODE") = ""

            Call BindGV()
        Else
            dt = ViewState("dt")
        End If

        'cm = New CommonMenu("SI", lheader.text, IS_CODE.Text)
        'cm.parentDir = "../../"
        'cm.haveCheckList = "Y"
        'cm.haveAttachments = "Y"
        'cm.haveNotes = "Y"
        'cm.haveTasks = "Y"
        'cm.haveEmail = "Y"
        'cm.haveHistory = "Y"

        'cm.genCM(cmBar)

        If moduleAction = "SELECTIM" Then
            addItemtoSI()
        End If

        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('OB_SI','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("IS_CODE") & "','N');")
        'selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value);")

        If IS_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
        ElseIf IS_STATUS.Text = "POSTED" Then
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
        '        Call cU.changeGVLabel(oGridViewRow, e, "Serial No.", "序號")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "批號")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板编號")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Carton No.", "外箱編號")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Item Code", "物件號碼")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內碼")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物件名稱")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Ref. No.", "文件編號")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Qty", "數量")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Loc", "位置")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Remarks", "備註")
        '        Call cU.changeGVLabel(oGridViewRow, e, "", "")
        '        REM **********************

        '        oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        'End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("isd_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ISD_BATCH_NO").ToString.Trim
                CType(e.Row.FindControl("isd_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ISD_PALLET_NO").ToString.Trim
                CType(e.Row.FindControl("isd_carton_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ISD_CARTON_NO").ToString.Trim
                CType(e.Row.FindControl("isd_itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ISD_ITM_CODE").ToString.Trim
                CType(e.Row.FindControl("itm_SKU_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("isd_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ISD_PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("isd_itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ISD_ITM_NAME").ToString.Trim
                CType(e.Row.FindControl("isd_ref_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ISD_REF_NO").ToString.Trim
                CType(e.Row.FindControl("isd_issue_qty"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "ISD_ISSUE_QTY").ToString.Trim)
                'CType(e.Row.FindControl("dsp_isd_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ISD_LOC").ToString.Trim
                'CType(e.Row.FindControl("isd_loc"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ISD_LOC").ToString.Trim

                Dim nDropDown As DropDownList = CType(e.Row.FindControl("isd_loc"), DropDownList)
                uiFun.load_dropdown(nDropDown, "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & gU.dbEncode(IS_WH.SelectedValue) & "'")
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "isd_loc").ToString.Trim

                CType(e.Row.FindControl("isd_rem"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ISD_REM").ToString.Trim
                CType(e.Row.FindControl("ISD_UOM2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ISD_UOM2").ToString.Trim
                CType(e.Row.FindControl("ISD_SL"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ISD_SL").ToString.Trim


                If DataBinder.Eval(e.Row.DataItem, "ISD_CUT_YN").ToString.Trim = "Y" Then
                    CType(e.Row.FindControl("ISD_CUT_YN"), CheckBox).Checked = True
                Else
                    CType(e.Row.FindControl("ISD_CUT_YN"), CheckBox).Checked = False
                End If

                CType(e.Row.FindControl("ISD_QTY2"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "ISD_QTY2").ToString.Trim)

                CType(e.Row.FindControl("ISD_SERIAL_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ISD_SERIAL_NO").ToString.Trim

                'Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp"), Image)
                'nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                'nImage.Attributes.Add("onclick", "LocLookUp('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_isd_loc"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("isd_loc"), HiddenField).ClientID) & "')")


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
        If cU.gfBuildDataTableforGridView(dt, GridView1, True, checkBoxValue) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CONVERT(int, ISD_SEQ)) + 1 from WMS_STOCK_ISSUE_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and IS_CODE = '" & gU.dbEncode(IS_CODE.Text.Trim) & "' "
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
            dt.Rows(rows_count - 1).Item("isd_seq") = ViewState("n_cur_seq").ToString
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

        If IS_TYPE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_IS_TYPE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_IS_TYPE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If IS_WH.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_IS_WH.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_IS_WH.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If IS_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_IS_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_IS_DATE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        ElseIf Not gU.isValidDate(IS_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_IS_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_IS_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If IS_PLAN_DATE.Text.Trim = "" And IS_TYPE.SelectedValue <> "WRITEOFF" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_IS_PLAN_DATE.Text & "cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_IS_PLAN_DATE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        Else
            If IS_PLAN_DATE.Text.Trim <> "" Then
                If Not gU.isValidDate(IS_PLAN_DATE.Text.Trim) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Invalid date, " & lbl_IS_PLAN_DATE.Text & "!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "無效的日期, " & lbl_IS_PLAN_DATE.Text & "!", Session("gLang"))
                    End If
                    Return False
                End If
            End If
        End If


        If IS_RETURN_DATE.Text.Trim <> "" Then
            If Not gU.isValidDate(IS_RETURN_DATE.Text.Trim) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Invalid date, " & lbl_IS_RETURN_DATE.Text & "!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "無效的日期, " & lbl_IS_RETURN_DATE.Text & "!", Session("gLang"))
                End If
                Return False
            End If
        End If

        If IS_REQ_EMAIL.Text.Trim <> "" And Not gU.isValidEmail(IS_REQ_EMAIL.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_IS_REQ_EMAIL.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_IS_REQ_EMAIL.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If Not gU.isDecimal(IS_TOT_PALLET.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid number, " & lbl_IS_TOT_PALLET.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的數字, " & lbl_IS_TOT_PALLET.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        Dim itemCount As Integer = 0

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If dt.Rows(i).Item("mFlag").ToString.Trim <> "D" Then

                    If CType(GridView1.Rows(i).FindControl("isd_issue_qty"), TextBox).Text <> "" Then
                        If uiFun.gvValidate(Me, dt, "isd_issue_qty", "Qty", _
                                         CType(GridView1.Rows(i).FindControl("isd_issue_qty"), TextBox).Text) = False Then Return False

                        If CInt(CType(GridView1.Rows(i).FindControl("isd_issue_qty"), TextBox).Text) <= 0 Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Issued Item Qty Cannot Be Zero!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "物料數量不能零!", Session("gLang"))
                            End If
                            Return False
                        End If
                    Else
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Issued Item Qty Cannot Be Empty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "物料數量不能空白!", Session("gLang"))
                        End If

                        Return False
                    End If

                    If CType(GridView1.Rows(i).FindControl("ISD_QTY2"), TextBox).Text <> "" Then
                        If uiFun.gvValidate(Me, dt, "ISD_QTY2", "Qty", _
                                         CType(GridView1.Rows(i).FindControl("ISD_QTY2"), TextBox).Text) = False Then Return False
                    End If

                    If CType(GridView1.Rows(i).FindControl("isd_loc"), DropDownList).SelectedValue = "" Then

                        uiFun.displayMsg(Me, "", "Location Cannot Be Empty!", Session("gLang"))
                        Return False
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

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim nextNo As String = ""
        Dim gConn As SqlConnection
        'Dim nextNo As String = ""
        Dim cutYN As String = ""
        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                Dim isCable As String = ""

                If IS_CABLE_YN.Checked Then isCable = "Y"

                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("SI", gConn, transaction)
                    'nextNo = IS_CODE.Text
                    REM **********************

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "insert into wms_stock_issue (" & _
                    "is_code, imp_code, storer_code, " & _
                    "is_type, is_status, is_date, IS_PLAN_DATE, IS_RETURN_DATE, IS_RETURN_DOC_NO, " & _
                    "is_project_no, is_issued_by, is_req_by, " & _
                    "is_req_tel, is_req_email, is_batch_no, " & _
                    "is_ref_no, is_wh, is_tot_pallet, is_rem, " & _
                    "is_cable_yn, " & _
                     "sys_cb, sys_cd, sys_lub, sys_lud)" & _
                    "values ( " & _
                    gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & _
                    gU.convdbNVCData(gU.dbEncode(IS_TYPE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(IS_STATUS.Text.Trim)) & "," & gU.convdbDate(gU.dbEncode(IS_DATE.Text.Trim)) & "," & gU.convdbDate(gU.dbEncode(IS_PLAN_DATE.Text.Trim)) & "," & gU.convdbDate(gU.dbEncode(IS_RETURN_DATE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(IS_RETURN_DOC_NO.Text.Trim)) & "," & _
                    gU.convdbNVCData(gU.dbEncode(IS_PROJECT_NO.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(IS_ISSUED_BY.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(IS_REQ_BY.Text.Trim)) & "," & _
                    gU.convdbNVCData(gU.dbEncode(IS_REQ_TEL.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(IS_REQ_EMAIL.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(IS_BATCH_NO.Text.Trim)) & "," & _
                    gU.convdbNVCData(gU.dbEncode(IS_REF_NO.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(IS_WH.SelectedValue)) & "," & gU.dbEncode(gU.decodeNullOrEmpty(IS_TOT_PALLET.Text, "0")) & "," & _
                    gU.convdbNVCData(gU.dbEncode(IS_REM.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(isCable)) & "," & _
                     "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                    '"'" & Session("imp_code") & "', N'" & gU.dbEncode(STORER_CODE.SelectedValue) & "', N'" & nextNo & "', " & _
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True, checkBoxValue) Then

                        uiFun.reOrderDetails(dt, "isd_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    itemSQL = "insert into wms_stock_issue_d (" & _
                                            "is_code, imp_code, storer_code, " & _
                                            "isd_seq, isd_pallet_no, isd_carton_no, " & _
                                            "isd_batch_no, isd_ref_no, isd_itm_code, " & _
                                            "isd_pack_key, isd_itm_name, isd_loc, " & _
                                            "isd_issue_qty, isd_rem, isd_serial_no, " & _
                                            "ISD_CUT_YN,ISD_UOM2,ISD_SL,ISD_QTY2,ISD_EXPIRY_DATE,ISD_MANU_DATE,ISD_VND_CODE," & _
                                            "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                            "values (" & _
                                            gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_seq").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_pallet_no").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_carton_no").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_batch_no").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_ref_no").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_itm_code").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_pack_key").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_itm_name").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_loc").ToString.Trim, ""))) & ", " & _
                                            gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("isd_issue_qty").ToString.Trim, "0")) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_rem").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_serial_no").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_CUT_YN").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_UOM2").ToString.Trim, ""))) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_SL").ToString.Trim, ""))) & ", " & _
                                            gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ISD_QTY2").ToString.Trim, "NULL")) & ", " & _
                                            gU.convdbDate(gU.dbEncode(rows.Item("ISD_EXPIRY_DATE").ToString.Trim)) & "," & _
                                            gU.convdbDate(gU.dbEncode(rows.Item("ISD_MANU_DATE").ToString.Trim)) & "," & _
                                            gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_VND_CODE").ToString.Trim, ""))) & ", " & _
                                            "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                    '"(N'" & Session("imp_code") & "', N'" & gU.dbEncode(STORER_CODE.Text) & "', N'" & nextNo & "', " & _
                            End Select
                            REM **********************

                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                        Next
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_stock_issue set " & _
                                    "is_status = " & gU.convdbNVCData(gU.dbEncode(IS_STATUS.Text)) & ", " & _
                                    "is_type = " & gU.convdbNVCData(gU.dbEncode(IS_TYPE.SelectedValue)) & ", " & _
                                    "is_date = " & gU.convdbDate(gU.dbEncode(IS_DATE.Text.Trim)) & ", " & _
                                    "IS_PLAN_DATE = " & gU.convdbDate(gU.dbEncode(IS_PLAN_DATE.Text.Trim)) & ", " & _
                                    "IS_RETURN_DATE = " & gU.convdbDate(gU.dbEncode(IS_RETURN_DATE.Text.Trim)) & ", " & _
                                    "IS_RETURN_DOC_NO = " & gU.convdbNVCData(gU.dbEncode(IS_RETURN_DOC_NO.Text.Trim)) & ", " & _
                                    "is_project_no = " & gU.convdbNVCData(gU.dbEncode(IS_PROJECT_NO.SelectedValue)) & ", " & _
                                    "is_issued_by = " & gU.convdbNVCData(gU.dbEncode(IS_ISSUED_BY.Text.Trim)) & ", " & _
                                    "is_req_by = " & gU.convdbNVCData(gU.dbEncode(IS_REQ_BY.Text.Trim)) & ", " & _
                                    "is_req_tel = " & gU.convdbNVCData(gU.dbEncode(IS_REQ_TEL.Text.Trim)) & ", " & _
                                    "is_req_email = " & gU.convdbNVCData(gU.dbEncode(IS_REQ_EMAIL.Text.Trim)) & ", " & _
                                    "is_batch_no = " & gU.convdbNVCData(gU.dbEncode(IS_BATCH_NO.Text.Trim)) & ", " & _
                                    "is_ref_no = " & gU.convdbNVCData(gU.dbEncode(IS_REF_NO.Text.Trim)) & ", " & _
                                    "is_wh = " & gU.convdbNVCData(gU.dbEncode(IS_WH.SelectedValue)) & ", " & _
                                    "is_tot_pallet = " & gU.dbEncode(gU.decodeNullOrEmpty(IS_TOT_PALLET.Text.Trim, "0")) & ", " & _
                                    "is_rem = " & gU.convdbNVCData(gU.dbEncode(IS_REM.Text.Trim)) & ", " & _
                                    "is_cable_yn = " & gU.convdbNVCData(gU.dbEncode(isCable)) & ", " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and is_code = '" & gU.dbEncode(IS_CODE.Text) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True, checkBoxValue) Then

                        uiFun.reOrderDetails(dt, "isd_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    itemSQL = "insert into wms_stock_issue_d (" & _
                                           "is_code, imp_code, storer_code, " & _
                                           "isd_seq, isd_pallet_no, isd_carton_no, " & _
                                           "isd_batch_no, isd_ref_no, isd_itm_code, " & _
                                           "isd_pack_key, isd_itm_name, isd_loc, " & _
                                           "isd_issue_qty, isd_rem, isd_serial_no, " & _
                                           "ISD_CUT_YN,ISD_UOM2,ISD_SL,ISD_QTY2,ISD_EXPIRY_DATE,ISD_MANU_DATE,ISD_VND_CODE," & _
                                           "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                           "values (" & _
                                           gU.convdbNVCData(gU.dbEncode(IS_CODE.Text)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_seq").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_pallet_no").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_carton_no").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_batch_no").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_ref_no").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_itm_code").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_pack_key").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_itm_name").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_loc").ToString.Trim, ""))) & ", " & _
                                           gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("isd_issue_qty").ToString.Trim, "0")) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_rem").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_serial_no").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_CUT_YN").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_UOM2").ToString.Trim, ""))) & ", " & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_SL").ToString.Trim, ""))) & ", " & _
                                           gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ISD_QTY2").ToString.Trim, "NULL")) & ", " & _
                                           gU.convdbDate(gU.dbEncode(rows.Item("ISD_EXPIRY_DATE").ToString.Trim)) & "," & _
                                           gU.convdbDate(gU.dbEncode(rows.Item("ISD_MANU_DATE").ToString.Trim)) & "," & _
                                           gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_VND_CODE").ToString.Trim, ""))) & ", " & _
                                           "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                Case "D"
                                    itemSQL = "delete from wms_stock_issue_d " & _
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and is_code = '" & gU.dbEncode(IS_CODE.Text.Trim) & "' " & _
                                            "and isd_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("isd_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    itemSQL = "update wms_stock_issue_d set " & _
                                                "isd_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_seq").ToString.Trim, ""))) & ", " & _
                                                "isd_pallet_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_pallet_no").ToString.Trim, ""))) & ", " & _
                                                "isd_carton_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_carton_no").ToString.Trim, ""))) & ", " & _
                                                "isd_batch_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_batch_no").ToString.Trim, ""))) & ", " & _
                                                "isd_ref_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_ref_no").ToString.Trim, ""))) & ", " & _
                                                "isd_itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_itm_code").ToString.Trim, ""))) & ", " & _
                                                "isd_pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_pack_key").ToString.Trim, ""))) & ", " & _
                                                "isd_itm_name = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_itm_name").ToString.Trim, ""))) & ", " & _
                                                "isd_loc = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_loc").ToString.Trim, ""))) & ", " & _
                                                "isd_issue_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("isd_issue_qty").ToString.Trim, "0")) & ", " & _
                                                "isd_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_rem").ToString.Trim, ""))) & ", " & _
                                                "isd_serial_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("isd_serial_no").ToString.Trim, ""))) & ", " & _
                                                "ISD_CUT_YN=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_CUT_YN").ToString.Trim, ""))) & "," & _
                                                "ISD_UOM2=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_UOM2").ToString.Trim, ""))) & "," & _
                                                "ISD_SL=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("ISD_SL").ToString.Trim, ""))) & "," & _
                                                "ISD_QTY2=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("ISD_QTY2").ToString.Trim, "0")) & "," & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = Getdate() " & _
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and is_code = '" & gU.dbEncode(IS_CODE.Text.Trim) & "' " & _
                                            "and isd_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("old_seq").ToString.Trim, "")) & "'"
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
                    IS_CODE.Text = nextNo
                    ViewState("IS_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    IS_CODE.ForeColor = Drawing.Color.Black
                    IS_CODE.Font.Size = 10
                    IS_CODE.CssClass = ""
                    STORER_CODE.CssClass = ""

                    REM **********************
                End If

                If flag <> "Y" Then uiFun.displayMsg(Me, "1007", "", Session("gLang"))
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
        If ViewState("IS_CODE") <> "" Then
            pk_code = ViewState("IS_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("IS_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))
        End If
        REM **********************

        uiFun.load_dropdownBy_ColCode(IS_TYPE, "WMS_STOCK_ISSUE.IS_TYPE", Session("gLang"))

        IS_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            IS_CODE.ForeColor = Drawing.Color.Red
            IS_STATUS.Text = "NEW"
            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = "SELECT wms_stock_issue.IMP_CODE,  wms_stock_issue.STORER_CODE,  wms_stock_issue.IS_CODE, " & _
                        "  wms_stock_issue.IS_TYPE,  wms_stock_issue.IS_STATUS,  CONVERT(nvarchar(30), wms_stock_issue.IS_PLAN_DATE," & DDFORMAT & ") as IS_PLAN_DATE, CONVERT(nvarchar(30), wms_stock_issue.IS_RETURN_DATE," & DDFORMAT & ") as IS_RETURN_DATE, IS_RETURN_DOC_NO, CONVERT(nvarchar(30), wms_stock_issue.IS_DATE," & DDFORMAT & ") as IS_DATE, " & _
                        "  wms_stock_issue.IS_PROJECT_NO,  wms_stock_issue.IS_ISSUED_BY,  wms_stock_issue.IS_REQ_BY, " & _
                        "  wms_stock_issue.IS_REQ_TEL,  wms_stock_issue.IS_REQ_EMAIL,  wms_stock_issue.IS_BATCH_NO, " & _
                        "  wms_stock_issue.IS_REF_NO,  wms_stock_issue.IS_WH,  wms_stock_issue.IS_TOT_PALLET, wms_stock_issue.IS_CABLE_YN," & _
                        "  wms_stock_issue.IS_REM,  wms_stock_issue.SYS_LUB,  wms_stock_issue.SYS_LUD, " & _
                        "  wms_stock_issue.SYS_CD,  wms_stock_issue.SYS_CB " & _
                        "FROM wms_stock_issue " & _
                        "where wms_stock_issue.is_code = '" & gU.dbEncode(pk_code) & "' " & _
                        "and wms_stock_issue.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        "and wms_stock_issue.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then
                ViewState("IS_CODE") = dt.Rows(0).Item("is_code").ToString
                ViewState("STORER_CODE") = dt.Rows(0).Item("STORER_CODE").ToString

                If dt.Rows(0).Item("IS_CABLE_YN").ToString.Trim = "Y" Then
                    IS_CABLE_YN.Checked = True
                Else
                    IS_CABLE_YN.Checked = False
                End If

                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                IS_CODE.Text = dt.Rows(0).Item("is_code").ToString
                STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                IS_STATUS.Text = dt.Rows(0).Item("is_status").ToString
                IS_TYPE.SelectedValue = dt.Rows(0).Item("is_type").ToString
                'IS_DATE.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("is_date").ToString)
                IS_DATE.Text = dt.Rows(0).Item("is_date").ToString
                IS_PLAN_DATE.Text = dt.Rows(0).Item("IS_PLAN_DATE").ToString
                IS_RETURN_DATE.Text = dt.Rows(0).Item("IS_RETURN_DATE").ToString
                IS_RETURN_DOC_NO.Text = dt.Rows(0).Item("IS_RETURN_DOC_NO").ToString
                IS_ISSUED_BY.Text = dt.Rows(0).Item("is_issued_by").ToString
                IS_REQ_BY.Text = dt.Rows(0).Item("is_req_by").ToString
                IS_REQ_TEL.Text = dt.Rows(0).Item("is_req_tel").ToString
                IS_REQ_EMAIL.Text = dt.Rows(0).Item("is_req_email").ToString
                IS_BATCH_NO.Text = dt.Rows(0).Item("is_batch_no").ToString
                IS_REF_NO.Text = dt.Rows(0).Item("is_ref_no").ToString
                'IS_WH.SelectedValue = dt.Rows(0).Item("is_wh").ToString
                IS_TOT_PALLET.Text = cU.FormatIntegerString(dt.Rows(0).Item("is_tot_pallet").ToString)
                IS_REM.Text = dt.Rows(0).Item("is_rem").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , , True)
                End If

                If dt.Rows(0).Item("IS_PROJECT_NO").ToString = "" Then
                    uiFun.load_dropdown(IS_PROJECT_NO, "SELECT PRJ_CODE, PRJ_NAME FROM WMS_PROJECT ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(IS_PROJECT_NO, "SELECT PRJ_CODE, PRJ_NAME FROM WMS_PROJECT WHERE PRJ_CODE = '" & gU.dbEncode(dt.Rows(0).Item("IS_PROJECT_NO").ToString) & "' ORDER BY 2", "PRJ_CODE", "PRJ_NAME", , , , True)
                End If

                uiFun.load_dropdown(IS_WH, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' and wh_code='" & dt.Rows(0).Item("is_wh").ToString.Trim & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"), dt.Rows(0).Item("is_wh").ToString.Trim, True)

                If IS_CODE.Text <> "" Then
                    'IS_CODE.ReadOnly = True
                    'IS_CODE.BorderWidth = 0
                    'IS_CODE.BackColor = Drawing.Color.Transparent
                End If

                If IS_ISSUED_BY.Text <> "" Then
                    IS_ISSUED_BY.ReadOnly = True
                    IS_ISSUED_BY.BorderWidth = 0
                    IS_ISSUED_BY.BackColor = Drawing.Color.Transparent
                End If

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "SELECT 'U' as mFlag,wms_stock_issue_d.isd_seq as old_seq, wms_item.itm_sku_no,  " & _
                    " WMS_STOCK_ISSUE_D.IMP_CODE, WMS_STOCK_ISSUE_D.STORER_CODE, WMS_STOCK_ISSUE_D.IS_CODE, WMS_STOCK_ISSUE_D.ISD_SEQ, " & _
                    " WMS_STOCK_ISSUE_D.ISD_PALLET_NO, WMS_STOCK_ISSUE_D.ISD_CARTON_NO, WMS_STOCK_ISSUE_D.ISD_BATCH_NO,  " & _
                    " WMS_STOCK_ISSUE_D.ISD_REF_NO, WMS_STOCK_ISSUE_D.ISD_ITM_CODE, WMS_STOCK_ISSUE_D.ISD_PACK_KEY, WMS_STOCK_ISSUE_D.ISD_ITM_NAME,  " & _
                    " WMS_STOCK_ISSUE_D.ISD_LOC, WMS_STOCK_ISSUE_D.ISD_ISSUE_QTY, WMS_STOCK_ISSUE_D.ISD_REM, WMS_STOCK_ISSUE_D.ISD_SERIAL_NO,  " & _
                    " WMS_STOCK_ISSUE_D.SYS_LUB, WMS_STOCK_ISSUE_D.SYS_LUD, WMS_STOCK_ISSUE_D.SYS_CD, WMS_STOCK_ISSUE_D.SYS_CB,  " & _
                    " WMS_STOCK_ISSUE_D.ISD_CUT_YN, WMS_STOCK_ISSUE_D.ISD_UOM2, WMS_STOCK_ISSUE_D.ISD_SL, WMS_STOCK_ISSUE_D.ISD_QTY2, Convert(varchar,WMS_STOCK_ISSUE_D.ISD_EXPIRY_DATE," & DDFORMAT & ") as ISD_EXPIRY_DATE,  " & _
                    " Convert(varchar,WMS_STOCK_ISSUE_D.ISD_MANU_DATE," & DDFORMAT & ") as ISD_MANU_DATE, WMS_STOCK_ISSUE_D.ISD_VND_CODE " & _
                    "from wms_stock_issue_d INNER JOIN " & _
                    " WMS_ITEM ON WMS_STOCK_ISSUE_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_ISSUE_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                    " WMS_STOCK_ISSUE_D.ISD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_ISSUE_D.ISD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                    "where wms_stock_issue_d.is_code = '" & gU.dbEncode(pk_code) & "' " & _
                    "and wms_stock_issue_d.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                    "and wms_stock_issue_d.imp_code = '" & Session("IMP_CODE") & "'"

        SQLString = SQLString & " order by CONVERT(int, wms_stock_issue_d.isd_seq)"
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
        IS_STATUS.Text = "CANCELLED"
        Call save()
        ar.sec_write = "N"
        CancelBtn.Visible = False
        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
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

                    Dim SrchStr As String = ""
                    SrchStr += "select ILOC_BAL_QTY from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' "
                    SrchStr += "AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
                    SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("isd_itm_code").ToString.Trim, "")) & "' "
                    SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("isd_pack_key").ToString.Trim, "")) & "' "
                    SrchStr += "AND ILOC_LOC = '" & gU.dbEncode(gU.decodeNull(rows.Item("isd_loc").ToString.Trim, "")) & "' "
                    SrchStr += "AND ILOC_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(gU.decodeNull(rows.Item("isd_pallet_no").ToString.Trim, ""), "000")) & "' "

                    If rows.Item("ISD_BATCH_NO").ToString.Trim <> "" Then
                        SrchStr += "AND ILOC_BATCH_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("ISD_BATCH_NO").ToString.Trim, "")) & "' "
                    Else
                        SrchStr += "AND ISNULL(ILOC_BATCH_NO,'') = '' "
                    End If

                    Dim qtydt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)

                    If qtydt.Rows.Count > 0 Then
                        Dim stQTY As Integer = gU.decodeEmptyCInt(qtydt.Rows(0).Item("ILOC_BAL_QTY").ToString, 0)
                        Dim plQTY As Integer = gU.decodeEmptyCInt(rows.Item("isd_issue_qty").ToString.Trim, 0)

                        If stQTY < plQTY Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsgNew(Me, "", "Item " & gU.decodeNull(rows.Item("isd_itm_code").ToString.Trim, "") & " | " &
                                                    gU.decodeNull(rows.Item("isd_pack_key").ToString.Trim, "") & " | " &
                                                    gU.decodeNull(rows.Item("isd_pallet_no").ToString.Trim, "") & " | " &
                                                    gU.decodeNull(rows.Item("ISD_BATCH_NO").ToString.Trim, "") & " " &
                                                    "don't have enough stock to delivery!", Session("gLang"))
                            Else
                                uiFun.displayMsgNew(Me, "", "物料" & gU.decodeNull(rows.Item("isd_itm_code").ToString.Trim, "") & " | " &
                                                    gU.decodeNull(rows.Item("isd_pack_key").ToString.Trim, "") & " | " &
                                                    gU.decodeNull(rows.Item("isd_pallet_no").ToString.Trim, "") & " | " &
                                                    gU.decodeNull(rows.Item("ISD_BATCH_NO").ToString.Trim, "") & " " &
                                                    "沒有足夠貨存出貨!", Session("gLang"))
                            End If

                            transaction.Rollback()
                            gConn.Close()
                            gConn.Dispose()
                            Exit Sub
                        End If
                    Else
                        If Session("gLang") = "E" Then
                            uiFun.displayMsgNew(Me, "", "Item " & gU.decodeNull(rows.Item("isd_itm_code").ToString.Trim, "") & " | " &
                                                gU.decodeNull(rows.Item("isd_pack_key").ToString.Trim, "") & " | " &
                                                gU.decodeNull(rows.Item("isd_pallet_no").ToString.Trim, "") & " | " &
                                                gU.decodeNull(rows.Item("ISD_BATCH_NO").ToString.Trim, "") & " " &
                                                "don't have stock to delivery!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(Me, "", "物料" & gU.decodeNull(rows.Item("isd_itm_code").ToString.Trim, "") & " | " &
                                                gU.decodeNull(rows.Item("isd_pack_key").ToString.Trim, "") & " | " &
                                                gU.decodeNull(rows.Item("isd_pallet_no").ToString.Trim, "") & " | " &
                                                gU.decodeNull(rows.Item("ISD_BATCH_NO").ToString.Trim, "") & " " &
                                                "沒有此貨存出貨!", Session("gLang"))
                        End If

                        transaction.Rollback()
                        gConn.Close()
                        gConn.Dispose()
                        Exit Sub
                    End If


                    st.STORER_CODE = STORER_CODE.SelectedValue
                    st.ITM_CODE = gU.decodeNull(rows.Item("isd_itm_code").ToString.Trim, "")
                    st.PACK_KEY = gU.decodeNull(rows.Item("isd_pack_key").ToString.Trim, "")
                    st.IO_CUST_CODE = ""
                    st.IO_WH = IS_WH.SelectedValue
                    st.IO_AREA = ""
                    st.IO_LOC = gU.decodeNull(rows.Item("isd_loc").ToString.Trim, "")
                    st.IO_DOC = "SI"
                    st.IO_DOC_ID = IS_CODE.Text.Trim
                    st.IO_QTY = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("isd_issue_qty").ToString.Trim, ""), "0")
                    st.PALLET_NO = gU.decodeNull(rows.Item("isd_pallet_no").ToString.Trim, "")
                    st.lO_BATCH_NO = gU.decodeNull(rows.Item("ISD_BATCH_NO").ToString.Trim, "")
                    st.IO_CBM = 0
                    st.IO_KG = 0
                    st.IO_EXPIRY_DATE = gU.decodeNull(rows.Item("ISD_EXPIRY_DATE").ToString.Trim, "")
                    st.IO_MANU_DATE = gU.decodeNull(rows.Item("ISD_MANU_DATE").ToString.Trim, "")
                    st.lO_VND_CODE = gU.decodeNull(rows.Item("ISD_VND_CODE").ToString.Trim, "")

                    If rows.Item("ISD_SERIAL_NO").ToString.Trim <> "" Then
                        st.IOS_SERIAL_NO = rows.Item("ISD_SERIAL_NO").ToString.Trim
                        st.IOS_QTY2 = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("ISD_QTY2").ToString.Trim, ""), "0")
                        st.IOS_UOM2 = gU.decodeNull(rows.Item("ISD_UOM2").ToString.Trim, "")
                        Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("ISD_SERIAL_NO").ToString.Trim, ""), gConn, transaction)
                    End If

                    st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                    st.UpdateStockTrans("OUT", gConn, transaction)
                    st.UpdateStockBalTrans("OUT", gConn, transaction)

                    If rows.Item("ISD_SERIAL_NO").ToString.Trim <> "" Then
                        st.UpdateStockSerialTrans("OUT", gConn, transaction)

                        st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                    End If
                Next

                updtSql = "update wms_stock_issue " &
                            "set is_status = 'POSTED', " &
                            "sys_lub = '" & Session("usr_id") & "', " &
                            "sys_lud = Getdate() " &
                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " &
                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                            "and is_code = '" & gU.dbEncode(IS_CODE.Text) & "' "

                gDB.amendData(updtSql)

                transaction.Commit()

                IS_STATUS.Text = "POSTED"
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

            load_ModalPopupExtender.Hide()
        End Try
    End Sub

    Protected Sub IS_WH_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles IS_WH.SelectedIndexChanged
        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                Dim nDropDown = CType(GridView1.Rows(i).FindControl("isd_loc"), DropDownList)
                uiFun.load_dropdown(nDropDown, "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & gU.dbEncode(IS_WH.SelectedValue) & "'")

                'CType(GridView1.Rows(i).FindControl("dsp_rtd_loc"), Label).Text = ""
                'CType(GridView1.Rows(i).FindControl("rtd_loc"), DropDownList).SelectedValue = ""
            Next
        End If
    End Sub

    Protected Sub setDefStorerInfo()
        Dim SQLString As String = ""
        Dim ldt As New DataTable

        SQLString = "SELECT * from WMS_STORER " &
                    "where STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                    "and IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "'"

        REM **********************
        ldt = gDB.getDataTable(SQLString)
        If ldt.Rows.Count > 0 Then
            IS_REQ_TEL.Text = ldt.Rows(0).Item("STO_CONT_TEL_ORD").ToString()
            IS_REQ_EMAIL.Text = ldt.Rows(0).Item("STO_CONT_EMAIL_ORD").ToString()
        End If
        ldt = Nothing
    End Sub

    Protected Sub STORER_CODE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles STORER_CODE.SelectedIndexChanged
        Call setDefStorerInfo()
    End Sub

    Protected Sub addItemtoSI()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True, checkBoxValue) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CONVERT(int, ISD_SEQ)) + 1 from WMS_STOCK_ISSUE_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and IS_CODE = '" & gU.dbEncode(IS_CODE.Text.Trim) & "' "
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
            Dim addSQL As String = ""
            Dim RO_dt As DataTable

            Dim itemPackList As String = ""
            Dim itemListarray As String()
            Dim packKeyListarray As String()
            Dim seqKeyList As String = ""
            Dim seqListArray As String()

            itemListarray = Split(itemList.Value, ", ")
            packKeyListarray = Split(packKeyList.Value, ", ")
            seqListArray = Split(seqList.Value, ", ")

            If itemListarray.Count = 0 Then
                If itemList.Value <> "" Then
                    itemPackList = Server.HtmlDecode(itemList.Value) & "_000_" & Server.HtmlDecode(packKeyList.Value)
                End If
            Else
                For i = 0 To itemListarray.Count - 1
                    itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemListarray(i)) & "_000_" & Server.HtmlDecode(packKeyListarray(i)))
                Next
            End If

            If seqListArray.Count = 0 Then
                If seqList.Value <> "" Then
                    seqKeyList = Server.HtmlDecode(seqList.Value)
                End If
            Else
                For i = 0 To seqListArray.Count - 1
                    seqKeyList = gU.appendToList(seqKeyList, seqListArray(i))
                Next
            End If

            If seqKeyList <> "" Then
                addSQL = "'" & Replace(seqKeyList, ", ", "', '") & "'"
            Else
                addSQL = "NULL"
            End If

            SQLString = " SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, " & _
                        " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC, CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END AS ILOC_BAL_QTY, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, " & _
                        " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL.ILOC_SEQ, WMS_ITEM_LOC_BAL_s.ILBS_SEQ, " & _
                        " WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, WMS_ITEM_LOC_BAL_S.ILBS_SL, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID,WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL, " & _
                        " Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE, WMS_ITEM_LOC_BAL.VND_CODE " & _
                        " FROM WMS_ITEM_LOC_BAL INNER JOIN " & _
                        " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND " & _
                        " WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " & _
                        " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                        " where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                        " and wms_item.imp_code = '" & Session("IMP_CODE") & "' " & _
                        " and Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'') in (" & addSQL & ") "

            SQLString = SQLString & " order by wms_item.ITM_CODE, wms_item.PACK_KEY, VND_CODE"

            REM **********************
            RO_dt = gDB.getDataTable(SQLString)

            If RO_dt.Rows.Count > 0 Then
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
                    dt.Rows(rows_count - 1).Item("ISD_SEQ") = ViewState("n_cur_seq").ToString
                    dt.Rows(rows_count - 1).Item("itm_sku_no") = RO_dt.Rows(i).Item("itm_sku_no")
                    dt.Rows(rows_count - 1).Item("ISD_ITM_CODE") = RO_dt.Rows(i).Item("ITM_CODE")
                    dt.Rows(rows_count - 1).Item("ISD_PACK_KEY") = RO_dt.Rows(i).Item("PACK_KEY")
                    dt.Rows(rows_count - 1).Item("ISD_ITM_NAME") = RO_dt.Rows(i).Item("ITM_NAME")
                    dt.Rows(rows_count - 1).Item("ISD_PALLET_NO") = RO_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                    dt.Rows(rows_count - 1).Item("ISD_BATCH_NO") = RO_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                    dt.Rows(rows_count - 1).Item("ISD_LOC") = RO_dt.Rows(i).Item("ILOC_LOC").ToString.Trim
                    'dt.Rows(rows_count - 1).Item("ISD_ISSUE_QTY") = RO_dt.Rows(i).Item("ILOC_BAL_QTY")
                    If RO_dt.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim <> "" Then
                        dt.Rows(rows_count - 1).Item("ISD_ISSUE_QTY") = 1
                    End If
                    dt.Rows(rows_count - 1).Item("ISD_SERIAL_NO") = RO_dt.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim

                    dt.Rows(rows_count - 1).Item("ISD_UOM2") = RO_dt.Rows(i).Item("ILBS_UOM2")
                    dt.Rows(rows_count - 1).Item("ISD_SL") = RO_dt.Rows(i).Item("ILBS_SL")

                    dt.Rows(rows_count - 1).Item("ISD_QTY2") = RO_dt.Rows(i).Item("ILBS_QTY2")

                    dt.Rows(rows_count - 1).Item("ISD_VND_CODE") = RO_dt.Rows(i).Item("VND_CODE").ToString.Trim
                    dt.Rows(rows_count - 1).Item("ISD_MANU_DATE") = RO_dt.Rows(i).Item("ILOC_MANU_DATE").ToString.Trim
                    dt.Rows(rows_count - 1).Item("ISD_EXPIRY_DATE") = RO_dt.Rows(i).Item("ILOC_EXPIRY_DATE")
                    REM **********************
                    dt.Rows(rows_count - 1).Item("mFlag") = "N"

                Next

                dt.AcceptChanges()
                ViewState("dt") = dt
                GridView1.DataSource = dt
                GridView1.DataBind()

                Dim tempSelValue As String = ""
                tempSelValue = IS_WH.SelectedValue
                uiFun.load_dropdown(IS_WH, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' and wh_code='" & tempSelValue & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"), tempSelValue, True)

                tempSelValue = STORER_CODE.SelectedValue
                uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(tempSelValue) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , tempSelValue, True)

            End If

        End If
    End Sub

    Protected Sub selectItemBtn_Click(sender As Object, e As System.EventArgs) Handles selectItemBtn.Click
        If IS_CABLE_YN.Checked Then
            Session("IS_Cable") = "Y"
        Else
            Session("IS_Cable") = ""
        End If

        ScriptManager.RegisterStartupScript(LOOKUPUDP, LOOKUPUDP.GetType, "itemLookup", "ItemLookUp(document.getElementById('" & STORER_CODE.ClientID & "').value,document.getElementById('" & IS_WH.ClientID & "').value);", True)
    End Sub
End Class
