Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OPERATION_SRV_SRVMain
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
    Private LocDict As Dictionary(Of String, LocItems)

    Structure LocItems
        Dim loc_code As String
        Dim totalCBM As Double
        Dim resrvCBM As Double
    End Structure

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
            uiFun.load_dropdown(LOBH_WH, "select WH_CODE,isnull(wh_name,wh_code) as wh_name from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "wh_name", , Session("gSelectLabel"))

        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If
            LOBH_BOOK_BY.Text = Session("usr_id")
            If LOBH_START_DATE.Text = "" Then
                LOBH_START_DATE.Text = Now.Date.ToString("dd/MM/yyyy")
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Location Reservation Maintenance"
            lbl_ImageHd.Text = "Location Details"
            lbl_LOBH_CODE.Text = "Reservation No:"
            lbl_LOBH_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Storer:"
            lbl_LOBH_BOOK_BY.Text = "Reserved By:"
            lbl_LOBH_WH.Text = "Warehouse:"
            lbl_LOBH_REMARKS.Text = "Remarks:"

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
            btnPost.OnClientClick = "return confirm(""Are you sure to post this record?\r\n(Please save your work before Posting)"");"
            If Session("pagemode") = "N" Then
                LOBH_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "貨量調整維護"
            lbl_ImageHd.Text = "貨量調整詳情"

            lbl_LOBH_CODE.Text = "調整編號:"
            lbl_LOBH_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "貨主:"
            lbl_LOBH_BOOK_BY.Text = "調整者:"
            lbl_LOBH_WH.Text = "倉庫:"
            lbl_LOBH_REMARKS.Text = "備註"

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
            If Session("pagemode") = "N" Then
                LOBH_CODE.Text = "[號碼會自動產生]"
            End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        LOBH_WH.CssClass = "REQUIRED"

        REM **********************

        If Session("pagemode") = "N" Then
            'LOBH_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("LOBH_CODE") = ""
            ViewState("STORER_CODE") = ""
            Call BindGV()
        Else
            dt = ViewState("dt")
        End If

        If moduleAction = "SELECTIM" Then
            addItemtoRSV()
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

        Image_RO_LookUp.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value, document.myform." & LOBH_WH.ClientID & ".value);")
        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('OP_SRSRV','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("LOBH_CODE") & "','N');")

        If LOBH_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
        ElseIf LOBH_STATUS.Text = "POSTED" Then
            ar.sec_viewMode = "Y"
            btnPost.Visible = False
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

        If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                DirectCast(GridView1.Rows(i).FindControl("dsp_LOBD_LOC"), Label).Text = DirectCast(GridView1.Rows(i).FindControl("LOBD_LOC"), HiddenField).Value
                DirectCast(GridView1.Rows(i).FindControl("dsp_LOBD_CBM_PERC"), Label).Text = DirectCast(GridView1.Rows(i).FindControl("LOBD_CBM_PERC"), HiddenField).Value
            Next
        End If

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

                CType(e.Row.FindControl("mFlag"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim
                CType(e.Row.FindControl("LOBD_ITM_CODE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "LOBD_ITM_CODE").ToString.Trim
                CType(e.Row.FindControl("PACK_KEY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("LOBD_SKU_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "LOBD_SKU_NO").ToString.Trim
                CType(e.Row.FindControl("LOBD_ITM_DESC"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "LOBD_ITM_DESC").ToString.Trim
                CType(e.Row.FindControl("LOBD_CBM"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "LOBD_CBM").ToString.Trim

                CType(e.Row.FindControl("dsp_LOBD_LOC"), Label).Text = DataBinder.Eval(e.Row.DataItem, "LOBD_LOC").ToString.Trim
                CType(e.Row.FindControl("LOBD_LOC"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "LOBD_LOC").ToString.Trim

                CType(e.Row.FindControl("dsp_LOBD_CBM_PERC"), Label).Text = DataBinder.Eval(e.Row.DataItem, "LOBD_CBM_PERC").ToString.Trim
                CType(e.Row.FindControl("LOBD_CBM_PERC"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "LOBD_CBM_PERC").ToString.Trim
                CType(e.Row.FindControl("BN_UTILIZATION_TYPE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "BN_UTILIZATION_TYPE").ToString.Trim
                CType(e.Row.FindControl("BN_CBM"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "BN_CBM").ToString.Trim
                CType(e.Row.FindControl("BN_AREA"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "BN_AREA").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "LOBD_ITM_CODE").ToString.Trim <> "" Then
                    CType(e.Row.FindControl("PACK_KEY"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("PACK_KEY"), TextBox).CssClass = "READONLY"
                    CType(e.Row.FindControl("LOBD_SKU_NO"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("LOBD_SKU_NO"), TextBox).CssClass = "READONLY"
                    CType(e.Row.FindControl("LOBD_ITM_DESC"), TextBox).Attributes.Add("readonly", "readonly")
                    CType(e.Row.FindControl("LOBD_ITM_DESC"), TextBox).CssClass = "READONLY"
                End If


                CType(e.Row.FindControl("LOBD_STATUS"), Label).Text = DataBinder.Eval(e.Row.DataItem, "LOBD_STATUS").ToString.Trim

                CType(e.Row.FindControl("LOBD_CBM"), TextBox).Attributes.Add("onkeypress", "return maskKeyAndCheck(event, " & _
                                                                                "document.myform." & HttpUtility.HtmlEncode(CType(e.Row.FindControl("LOBD_LOC"), HiddenField).ClientID) & ".value); ")

                CType(e.Row.FindControl("LOBD_CBM"), TextBox).Attributes.Add("onkeyup", "javascript:calCBM(this.value, " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("BN_CBM"), HiddenField).ClientID) & "', " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("BN_AREA"), HiddenField).ClientID) & "', " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("BN_UTILIZATION_TYPE"), HiddenField).ClientID) & "', " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_LOBD_CBM_PERC"), Label).ClientID) & "', " & _
                                                                                "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("LOBD_CBM_PERC"), HiddenField).ClientID) & "' " & _
                                                                                ");")

                'CType(e.Row.FindControl("LOBD_CBM"), TextBox).Attributes.Add("onkeyup", "javascript:alert('hi');")

                Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp"), Image)
                nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")


                nImage.Attributes.Add("onclick", "LocLookUp('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_LOBD_LOC"), Label).ClientID) & "', " & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("LOBD_LOC"), HiddenField).ClientID) & "'," & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("BN_CBM"), HiddenField).ClientID) & "'," & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("BN_AREA"), HiddenField).ClientID) & "'," & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("BN_UTILIZATION_TYPE"), HiddenField).ClientID) & "'," & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("LOBD_CBM"), TextBox).ClientID) & "'," & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("LOBD_CBM_PERC"), HiddenField).ClientID) & "'," & _
                                      "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_LOBD_CBM_PERC"), Label).ClientID) & "'" & _
                                      ");")


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
        Dim tempDT As DataTable

        tempDT = ViewState("dt")

        If cU.gfBuildDataTableforGridView(tempDT, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(LOBD_SEQ AS int)) + 1 from WMS_LOC_BOOK_DTL " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and LOBH_CODE = '" & gU.dbEncode(ViewState("LOBH_CODE")) & "' "
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

            Dim newRow As DataRow

            newRow = tempDT.NewRow

            REM **********************
            REM Modify Here

            REM **********************

            newRow.Item("LOBD_SEQ") = ViewState("n_cur_seq").ToString
            newRow.Item("mFlag") = "N"
            newRow.Item("LOBD_STATUS") = "FREE"

            tempDT.Rows.Add(newRow)

            tempDT.AcceptChanges()

            ViewState("dt") = tempDT
            GridView1.DataSource = tempDT
            GridView1.DataBind()

        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        dt.AcceptChanges()
        ViewState("dt") = dt

        DirectCast(GridView1.Rows(e.RowIndex).FindControl("mFlag"), HiddenField).Value = "D"

        'GridView1.DataSource = dt
        'GridView1.DataBind()

    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""
        Dim i As Integer

        Dim iLocitem As LocItems

        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If LOBH_WH.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_LOBH_WH.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_LOBH_WH.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        'If LOBH_DATE.Text.Trim = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_LOBH_DATE.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_LOBH_DATE.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'ElseIf Not gU.isValidDate(LOBH_DATE.Text.Trim) Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", "Invalid date, " & lbl_LOBH_DATE.Text & "!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", "無效的日期, " & lbl_LOBH_DATE.Text & "!", Session("gLang"))
        '    End If
        '    Return False
        'End If
        Dim bn_cbm As String = ""
        Dim bn_area As String = ""

        LocDict = New Dictionary(Of String, LocItems)

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If CType(GridView1.Rows(i).FindControl("mFlag"), HiddenField).Value <> "D" Then
                    Dim loc As String = CType(GridView1.Rows(i).FindControl("LOBD_LOC"), HiddenField).Value

                    If loc = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Location Cannot Be Empty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "位置不能空白!", Session("gLang"))
                        End If
                        Return False
                    Else
                        bn_cbm = DirectCast(GridView1.Rows(i).FindControl("BN_CBM"), HiddenField).Value
                        bn_area = DirectCast(GridView1.Rows(i).FindControl("BN_AREA"), HiddenField).Value
                        'for demo
                        'If bn_cbm = "" AndAlso bn_area = "" Then
                        '    uiFun.displayMsg(Me, "", "CBM value has not been defined in location(" & loc & ")!", Session("gLang"))
                        '    Return False
                        'End If

                        'If Not gU.isDecimal(bn_cbm) AndAlso Not gU.isDecimal(bn_area) Then
                        '    uiFun.displayMsg(Me, "", "CBM value is invalid in location(" & loc & ")!", Session("gLang"))
                        '    Return False
                        'End If

                        'If CDbl(bn_cbm) = 0 AndAlso CDbl(bn_area) = 0 Then
                        '    uiFun.displayMsg(Me, "", "CBM value defined cannot be 0 in location(" & loc & ")!", Session("gLang"))
                        '    Return False
                        'End If
                    End If

                    If CType(GridView1.Rows(i).FindControl("LOBD_CBM"), TextBox).Text <> "" Then
                        If uiFun.gvValidate(Me, ViewState("dt"), "LOBD_CBM", "CBM", _
                                         CType(GridView1.Rows(i).FindControl("LOBD_CBM"), TextBox).Text) = False Then Return False
                    Else
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "CBM Cannot Be Empty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "CBM不能空白!", Session("gLang"))
                        End If
                        Return False
                    End If
                    'for demo
                    'If CType(GridView1.Rows(i).FindControl("LOBD_CBM_PERC"), HiddenField).Value <> "" AndAlso gU.isDecimal(CType(GridView1.Rows(i).FindControl("LOBD_CBM_PERC"), HiddenField).Value) Then
                    '    If CDbl(CType(GridView1.Rows(i).FindControl("LOBD_CBM_PERC"), HiddenField).Value) > 100 Then
                    '        uiFun.displayMsg(Me, "", "CBM reserve value cannot exceed the bin capacity!", Session("gLang"))
                    '        Return False

                    '    End If
                    'End If

                    'If Not LocDict.ContainsKey(CType(GridView1.Rows(i).FindControl("LOBD_LOC"), HiddenField).Value) Then
                    '    iLocitem = New LocItems
                    '    iLocitem.loc_code = CType(GridView1.Rows(i).FindControl("LOBD_LOC"), HiddenField).Value
                    '    Select Case CType(GridView1.Rows(i).FindControl("BN_UTILIZATION_TYPE"), HiddenField).Value
                    '        Case "CBM"
                    '            iLocitem.totalCBM = CDbl(CType(GridView1.Rows(i).FindControl("BN_CBM"), HiddenField).Value)
                    '        Case "AREA"
                    '            iLocitem.totalCBM = CDbl(CType(GridView1.Rows(i).FindControl("BN_AREA"), HiddenField).Value)
                    '    End Select

                    '    iLocitem.resrvCBM = CDbl(CType(GridView1.Rows(i).FindControl("LOBD_CBM"), TextBox).Text)

                    '    LocDict.Add(CType(GridView1.Rows(i).FindControl("LOBD_LOC"), HiddenField).Value, iLocitem)
                    'Else
                    '    iLocitem = New LocItems
                    '    iLocitem = LocDict(loc)
                    '    iLocitem.resrvCBM += CDbl(CType(GridView1.Rows(i).FindControl("LOBD_CBM"), TextBox).Text)

                    '    LocDict(loc) = iLocitem
                    'End If

                End If
            Next

            'for demo
            'If Not LocDict Is Nothing AndAlso LocDict.Keys.Count > 0 Then
            '    For i = 0 To LocDict.Keys.Count - 1
            '        If LocDict(LocDict.Keys(i)).resrvCBM > LocDict(LocDict.Keys(i)).totalCBM Then
            '            uiFun.displayMsg(Me, "", "The total CBM reservation value for location(" & LocDict.Keys(i) & ") exceed the bin capacity!", Session("gLang"))
            '            Return False
            '        End If
            '    Next
            'End If

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
        Dim seqSQL As String = ""
        Dim LOBD_SEQ As String

        Dim successFlag As Boolean = False

        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("LOCB", gConn, transaction)
                    'nextNo = AD_CODE.Text
                    REM **********************

                    dupSQL = "select 1 from WMS_LOC_BOOK_HD " & _
                                "where LOBH_CODE = '" & gU.dbEncode(nextNo) & "' " & _
                                "and imp_code='" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    If dupTbl.Rows.Count = 0 Then
                        sql_string = " INSERT INTO WMS_LOC_BOOK_HD " & _
                                     " (IMP_CODE, STORER_CODE, LOBH_CODE, LOBH_STATUS, LOBH_BOOK_BY, LOBH_START_DATE, LOBH_END_DATE, LOBH_WH, LOBH_FD_DATE, " & _
                                     " LOBH_AD_DATE, LOBH_RO_CODE, LOBH_PO_CODE, LOBH_REMARKS, SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) " & _
                                     " VALUES (" & _
                                     gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & gU.convdbNVCData(gU.dbEncode(nextNo)) & ",'NEW'," & _
                                     gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_BOOK_BY.Text.Trim, ""))) & ", " & gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_START_DATE.Text.Trim, ""))) & ", " & _
                                     gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_END_DATE.Text.Trim, ""))) & ", " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_WH.SelectedValue.Trim, ""))) & ", " & _
                                     gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_FD_DATE.Text.Trim, ""))) & ", " & gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_AD_DATE.Text.Trim, ""))) & ", " & _
                                     gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_RO_CODE.Text.Trim, ""))) & ", " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_PO_CODE.Text.Trim, ""))) & ", " & _
                                     gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_REMARKS.Text.Trim, ""))) & ", " & _
                                     "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "


                        REM **********************

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                            uiFun.reOrderDetails(dt, "LOBD_SEQ")

                            For Each rows As DataRow In dt.Rows
                                itemSQL = ""

                                seqSQL = "select max(cast (LOBD_SEQ as int)) + 1 from WMS_LOC_BOOK_DTL where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                             "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                             "and LOBH_CODE = '" & gU.dbEncode(nextNo) & "' "

                                LOBD_SEQ = gU.decodeNullOrEmpty(DB.getValueFromSQL(seqSQL, gConn, transaction), "1")

                                REM **********************
                                REM Modify Here
                                Select Case rows.Item("mFlag")
                                    Case "N"

                                        itemSQL = " INSERT INTO WMS_LOC_BOOK_DTL " & _
                                                  " (IMP_CODE, STORER_CODE, LOBH_CODE, LOBD_SEQ, LOBD_ITM_CODE, LOBD_SKU_NO, PACK_KEY, LOBD_ITM_DESC, LOBD_LOC, LOBD_CBM,  " & _
                                                  " LOBD_CBM_PERC, LOBD_STATUS, SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) " & _
                                                  " VALUES        (" & _
                                                  gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & _
                                                  gU.convdbNVCData(gU.dbEncode(LOBD_SEQ)) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_ITM_CODE").ToString.Trim, ""))) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_SKU_NO").ToString.Trim, ""))) & "," & _
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PACK_KEY").ToString.Trim, ""))) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_ITM_DESC").ToString.Trim, ""))) & "," & _
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_LOC").ToString.Trim, ""))) & "," & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_CBM").ToString.Trim, "NULL")) & "," & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_CBM_PERC").ToString.Trim, "NULL")) & "," & _
                                                  gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_STATUS").ToString.Trim, "NEW"))) & "," & _
                                                  "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                    Case Else

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
                    sql_string = " UPDATE WMS_LOC_BOOK_HD " & _
                                 " SET LOBH_START_DATE =" & gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_START_DATE.Text.Trim, ""))) & ",  " & _
                                 " 	LOBH_END_DATE =" & gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_END_DATE.Text.Trim, ""))) & ",  " & _
                                 " 	LOBH_WH =" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_WH.SelectedValue.Trim, ""))) & ",  " & _
                                 " 	LOBH_FD_DATE =" & gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_FD_DATE.Text.Trim, ""))) & ",  " & _
                                 "  LOBH_AD_DATE =" & gU.convdbDate(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_AD_DATE.Text.Trim, ""))) & ",  " & _
                                 " 	LOBH_REMARKS =" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_REMARKS.Text.Trim, ""))) & ", " & _
                                 " 	SYS_LUB= '" & Session("usr_id") & "', SYS_LUD=getdate() " & _
                                 "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                 "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                 "and LOBH_CODE = '" & gU.dbEncode(LOBH_CODE.Text) & "' "


                    '" 	LOBH_RO_CODE =" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_RO_CODE.Text.Trim, ""))) & ",  " & _
                    '" 	LOBH_PO_CODE =" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(LOBH_PO_CODE.Text.Trim, ""))) & ",  " & _
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "LOBD_SEQ")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"

                                    seqSQL = "select max(cast (LOBD_SEQ as int)) + 1 from WMS_LOC_BOOK_DTL where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                             "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                             "and LOBH_CODE = '" & gU.dbEncode(LOBH_CODE.Text.Trim) & "' "

                                    LOBD_SEQ = gU.decodeNullOrEmpty(DB.getValueFromSQL(seqSQL, gConn, transaction), "1")

                                    itemSQL = " INSERT INTO WMS_LOC_BOOK_DTL " & _
                                              " (IMP_CODE, STORER_CODE, LOBH_CODE, LOBD_SEQ, LOBD_ITM_CODE, LOBD_SKU_NO, PACK_KEY, LOBD_ITM_DESC, LOBD_LOC, LOBD_CBM,  " & _
                                              " LOBD_CBM_PERC, LOBD_STATUS, SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) " & _
                                              " VALUES        (" & _
                                              gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & gU.convdbNVCData(gU.dbEncode(LOBH_CODE.Text.Trim)) & "," & _
                                              gU.convdbNVCData(gU.dbEncode(LOBD_SEQ)) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_ITM_CODE").ToString.Trim, ""))) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_SKU_NO").ToString.Trim, ""))) & "," & _
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PACK_KEY").ToString.Trim, ""))) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_ITM_DESC").ToString.Trim, ""))) & "," & _
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_LOC").ToString.Trim, ""))) & "," & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_CBM").ToString.Trim, "NULL")) & "," & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_CBM_PERC").ToString.Trim, "NULL")) & "," & _
                                              gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_STATUS").ToString.Trim, "NEW"))) & "," & _
                                              "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                Case "D"
                                    itemSQL = "delete from WMS_LOC_BOOK_DTL " & _
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and LOBH_CODE = '" & gU.dbEncode(LOBH_CODE.Text.Trim) & "' " & _
                                            "and LOBD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("LOBD_SEQ").ToString.Trim, "")) & "' "
                                Case Else

                                    itemSQL = " UPDATE  WMS_LOC_BOOK_DTL SET " & _
                                                " LOBD_SKU_NO =" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_SKU_NO").ToString.Trim, ""))) & ", " & _
                                                " PACK_KEY =" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PACK_KEY").ToString.Trim, ""))) & ", " & _
                                                " LOBD_ITM_DESC =" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_ITM_DESC").ToString.Trim, ""))) & ", " & _
                                                " LOBD_LOC =" & gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_LOC").ToString.Trim, ""))) & ", " & _
                                                " LOBD_CBM =" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_CBM").ToString.Trim, "NULL")) & ", " & _
                                                " LOBD_CBM_PERC =" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("LOBD_CBM_PERC").ToString.Trim, "NULL")) & ", " & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = Getdate() " & _
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and LOBH_CODE = '" & gU.dbEncode(LOBH_CODE.Text.Trim) & "' " & _
                                            "and LOBD_SEQ = '" & gU.dbEncode(gU.decodeNull(rows.Item("LOBD_SEQ").ToString.Trim, "")) & "'"
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
                    LOBH_CODE.Text = nextNo
                    ViewState("LOBH_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    LOBH_CODE.ForeColor = Drawing.Color.Black
                    LOBH_CODE.Font.Size = 10
                    'AD_CODE.CssClass = ""
                    STORER_CODE.CssClass = ""

                    REM **********************
                End If

                If flag <> "Y" Then
                    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                End If

                successFlag = True
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

        If successFlag Then Call BindGV()

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
        If ViewState("LOBH_CODE") <> "" Then
            pk_code = ViewState("LOBH_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("LOBH_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))

            ViewState("LOBH_CODE") = pk_code
            ViewState("STORER_CODE") = storerCode
        End If
        REM **********************

        LOBH_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            LOBH_STATUS.Text = "NEW"
            LOBH_CODE.ForeColor = Drawing.Color.Red

            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = " SELECT WMS_LOC_BOOK_HD.IMP_CODE, WMS_LOC_BOOK_HD.STORER_CODE, WMS_LOC_BOOK_HD.LOBH_CODE, WMS_LOC_BOOK_HD.LOBH_STATUS, WMS_LOC_BOOK_HD.LOBH_BOOK_BY, convert(varchar, WMS_LOC_BOOK_HD.LOBH_START_DATE,103) as LOBH_START_DATE,convert(varchar,WMS_LOC_BOOK_HD.LOBH_END_DATE,103 ) as LOBH_END_DATE, WMS_LOC_BOOK_HD.LOBH_WH, " & _
                        " convert(varchar, WMS_LOC_BOOK_HD.LOBH_FD_DATE,103) as LOBH_FD_DATE, convert(varchar,WMS_LOC_BOOK_HD.LOBH_AD_DATE,103) as LOBH_AD_DATE, WMS_LOC_BOOK_HD.LOBH_RO_CODE , WMS_LOC_BOOK_HD.LOBH_PO_CODE, WMS_LOC_BOOK_HD.LOBH_REMARKS, " & _
                        " WMS_LOC_BOOK_HD.SYS_LUB, WMS_LOC_BOOK_HD.SYS_LUD, WMS_LOC_BOOK_HD.SYS_CD, WMS_LOC_BOOK_HD.SYS_CB " & _
                        " FROM WMS_LOC_BOOK_HD " & _
                        " where WMS_LOC_BOOK_HD.LOBH_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                        " and WMS_LOC_BOOK_HD.imp_code = '" & Session("IMP_CODE") & "' " & _
                        " and WMS_LOC_BOOK_HD.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                LOBH_CODE.Text = dt.Rows(0).Item("LOBH_code").ToString
                STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                LOBH_STATUS.Text = dt.Rows(0).Item("LOBH_STATUS").ToString

                LOBH_START_DATE.Text = dt.Rows(0).Item("LOBH_START_DATE").ToString
                LOBH_END_DATE.Text = dt.Rows(0).Item("LOBH_END_DATE").ToString
                LOBH_FD_DATE.Text = dt.Rows(0).Item("LOBH_FD_DATE").ToString
                LOBH_AD_DATE.Text = dt.Rows(0).Item("LOBH_AD_DATE").ToString
                LOBH_BOOK_BY.Text = dt.Rows(0).Item("LOBH_BOOK_BY").ToString
                LOBH_WH.SelectedValue = dt.Rows(0).Item("LOBH_WH").ToString
                LOBH_REMARKS.Text = dt.Rows(0).Item("LOBH_REMARKS").ToString.Trim

                LOBH_RO_CODE.Text = dt.Rows(0).Item("LOBH_RO_CODE").ToString
                LOBH_PO_CODE.Text = dt.Rows(0).Item("LOBH_PO_CODE").ToString

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , , True)
                End If

                'If dt.Rows(0).Item("LOBH_WH").ToString = "" Then
                '    uiFun.load_dropdown(LOBH_WH, "select WH_CODE,isnull(wh_name,wh_code) as wh_name from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "wh_name", , Session("gSelectLabel"))
                'Else
                '    uiFun.load_dropdown(LOBH_WH, "select WH_CODE,isnull(wh_name,wh_code) as wh_name from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' and wh_code='" & gU.dbEncode(dt.Rows(0).Item("LOBH_WH").ToString) & "' ORDER BY 1", "WH_CODE", "wh_name", , Session("gSelectLabel"), dt.Rows(0).Item("LOBH_WH").ToString)
                'End If

                If LOBH_CODE.Text <> "" Then
                    'AD_CODE.ReadOnly = True
                    'AD_CODE.BorderWidth = 0
                    LOBH_CODE.BackColor = Drawing.Color.Transparent
                End If

                Image_RO_LookUp.Visible = False

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = " SELECT 'U' as mFlag, WMS_LOC_BOOK_DTL.LOBH_CODE, WMS_LOC_BOOK_DTL.LOBD_SEQ, WMS_LOC_BOOK_DTL.LOBD_ITM_CODE, WMS_LOC_BOOK_DTL.LOBD_SKU_NO, WMS_LOC_BOOK_DTL.PACK_KEY, WMS_LOC_BOOK_DTL.LOBD_ITM_DESC, WMS_LOC_BOOK_DTL.LOBD_LOC, LOBD_CBM, " & _
                    " WMS_LOC_BOOK_DTL.LOBD_CBM_PERC, WMS_LOC_BOOK_DTL.LOBD_STATUS, " & _
                    " WMS_WH_BIN.BN_LENGTH, WMS_WH_BIN.BN_WIDTH, WMS_WH_BIN.BN_DEPTH, WMS_WH_BIN.BN_UTILIZATION_TYPE, " & _
                    " (isnull(WMS_WH_BIN.BN_LENGTH,0) * isnull(WMS_WH_BIN.BN_WIDTH,0) * isnull(WMS_WH_BIN.BN_DEPTH,0) / 1000000) as BN_CBM," & _
                    " isnull(WMS_WH_BIN.BN_LENGTH,0) * isnull(WMS_WH_BIN.BN_WIDTH,0) as BN_AREA" & _
                    " FROM WMS_LOC_BOOK_DTL " & _
                    " INNER JOIN V_LOCATION ON WMS_LOC_BOOK_DTL.IMP_CODE = V_LOCATION.IMP_CODE AND WMS_LOC_BOOK_DTL.LOBD_LOC = V_LOCATION.LOC INNER JOIN " & _
                    " WMS_WH_BIN ON V_LOCATION.IMP_CODE = WMS_WH_BIN.IMP_CODE AND V_LOCATION.WH_CODE = WMS_WH_BIN.WH_CODE AND  " & _
                    " V_LOCATION.FL_NUM = WMS_WH_BIN.FL_NUM AND V_LOCATION.AR_CODE = WMS_WH_BIN.AR_CODE AND  " & _
                    " V_LOCATION.RK_CODE = WMS_WH_BIN.RK_CODE AND V_LOCATION.BN_CODE = WMS_WH_BIN.BN_CODE " & _
                    " where WMS_LOC_BOOK_DTL.LOBH_CODE = '" & gU.dbEncode(pk_code) & "' " & _
                    " and WMS_LOC_BOOK_DTL.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                    " and WMS_LOC_BOOK_DTL.imp_code = '" & Session("IMP_CODE") & "'"

        SQLString = SQLString & " order by CONVERT(int, WMS_LOC_BOOK_DTL.LOBD_SEQ)"
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
        Dim cancelSql As String = "update WMS_LOC_BOOK_HD " & _
                     "set LOBH_STATUS = 'CANCELLED', " & _
                     "sys_lub = '" & Session("usr_id") & "', " & _
                     "sys_lud = Getdate() " & _
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                     "and LOBH_CODE = '" & gU.dbEncode(LOBH_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction


        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql)

            transaction.Commit()

            LOBH_STATUS.Text = "CANCELLED"

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

                For Each rows As DataRow In dt.Rows
                    st.STORER_CODE = STORER_CODE.SelectedValue
                    st.ITM_CODE = gU.decodeNull(rows.Item("add_itm_code").ToString.Trim, "")
                    st.PACK_KEY = gU.decodeNull(rows.Item("add_pack_key").ToString.Trim, "")
                    st.IO_CUST_CODE = ""
                    st.IO_AREA = ""
                    st.IO_DOC = "SADJ"
                    st.IO_DOC_ID = LOBH_CODE.Text.Trim
                    st.IO_CBM = 0
                    st.IO_KG = 0
                    st.IO_WH = LOBH_WH.SelectedValue
                    st.PALLET_NO = gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, "")
                    st.IO_LOC = gU.decodeNull(rows.Item("add_loc").ToString.Trim, "")
                    st.lO_BATCH_NO = gU.decodeNull(rows.Item("add_batch_no").ToString.Trim, "")
                    st.IO_EXPIRY_DATE = gU.decodeNull(rows.Item("add_expiry_date").ToString.Trim, "")
                    st.IO_MANU_DATE = gU.decodeNull(rows.Item("add_manu_date").ToString.Trim, "")

                    Dim txQty As Integer = 0

                    If gU.decodeEmptyCInt(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") > _
                        gU.decodeEmptyCInt(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCInt(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCInt(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0")

                        st.IO_QTY = txQty
                        st.UpdateStockTrans("IN", gConn, transaction)
                        st.UpdateStockBalTrans("IN", gConn, transaction)

                    ElseIf gU.decodeEmptyCInt(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") < _
                        gU.decodeEmptyCInt(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCInt(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCInt(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0")

                        st.IO_QTY = txQty
                        st.UpdateStockTrans("OUT", gConn, transaction)
                        st.UpdateStockBalTrans("OUT", gConn, transaction)
                    ElseIf gU.decodeEmptyCInt(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") = _
                        gU.decodeEmptyCInt(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

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

                updtSql = "update wms_stock_adjust " & _
                            "set ad_status = 'POSTED', " & _
                            "sys_lub = N'" & Session("usr_id") & "', " & _
                            "sys_lud = Getdate() " & _
                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            "and ad_code = '" & gU.dbEncode(LOBH_CODE.Text) & "' "

                gDB.amendData(updtSql)

                transaction.Commit()

                LOBH_STATUS.Text = "POSTED"
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

    Protected Sub addItemtoRSV()

        Dim selectedRO As String = itemList.Value

        If itemList.Value <> "" Then
            Dim tempDT As DataTable
            Dim selectSQL As String = ""
            Dim itemDT As DataTable
            Dim newRow As DataRow

            selectSQL = " SELECT WMS_REPLENISH.RO_CODE, WMS_REPLENISH.RO_EDI_PO_NO, WMS_ITEM.ITM_CODE, WMS_ITEM.PACK_KEY, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC,WMS_ITEM.ITM_SKU_NO, " & _
                        " SUM(case when wms_item.ITM_STACKABLE_YN = 'Y' then V_ALT_VEND_ITEM.CARTON_CBM else (V_ALT_VEND_ITEM.AITM_LENGTH * V_ALT_VEND_ITEM.AITM_WIDTH) end) as CBM_VAL " & _
                        " FROM  WMS_REPLENISH INNER JOIN " & _
                        " WMS_REPLENISH_D ON WMS_REPLENISH.IMP_CODE = WMS_REPLENISH_D.IMP_CODE AND  " & _
                        " WMS_REPLENISH.STORER_CODE = WMS_REPLENISH_D.STORER_CODE AND WMS_REPLENISH.RO_CODE = WMS_REPLENISH_D.RO_CODE  " & _
                        " INNER JOIN WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_REPLENISH_D.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_REPLENISH_D.STORER_CODE AND  " & _
                        " WMS_ITEM.ITM_CODE = WMS_REPLENISH_D.ROD_ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_REPLENISH_D.ROD_PACK_KEY " & _
                        " left outer JOIN  " & _
                        " V_ALT_VEND_ITEM ON WMS_ITEM.IMP_CODE = V_ALT_VEND_ITEM.IMP_CODE AND WMS_ITEM.STORER_CODE = V_ALT_VEND_ITEM.STORER_CODE AND  " & _
                        " WMS_ITEM.ITM_CODE = V_ALT_VEND_ITEM.ITM_CODE AND WMS_ITEM.PACK_KEY = V_ALT_VEND_ITEM.PACK_KEY " & _
                        " WHERE  WMS_REPLENISH.IMP_CODE='" & gU.dbEncode(Session("imp_code")) & "'  AND WMS_REPLENISH.STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND WMS_REPLENISH.RO_CODE='" & gU.dbEncode(selectedRO) & "'" & _
                        " group by WMS_REPLENISH.RO_CODE, WMS_REPLENISH.RO_EDI_PO_NO, WMS_ITEM.ITM_CODE, WMS_ITEM.PACK_KEY, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC,WMS_ITEM.ITM_SKU_NO "

            tempDT = gDB.getDataTable(selectSQL)

            If tempDT.Rows.Count > 0 Then

                LOBH_RO_CODE.Text = selectedRO
                LOBH_PO_CODE.Text = tempDT.Rows(0).Item("RO_EDI_PO_NO").ToString.Trim

                itemDT = ViewState("dt")

                itemDT.Clear()

                For i = 0 To tempDT.Rows.Count - 1
                    newRow = itemDT.NewRow
                    newRow.Item("mFlag") = "N"
                    newRow.Item("LOBD_SEQ") = i + 1
                    newRow.Item("LOBD_ITM_CODE") = tempDT.Rows(i).Item("ITM_CODE").ToString.Trim
                    newRow.Item("PACK_KEY") = tempDT.Rows(i).Item("PACK_KEY").ToString.Trim
                    newRow.Item("LOBD_SKU_NO") = tempDT.Rows(i).Item("ITM_SKU_NO").ToString.Trim
                    newRow.Item("LOBD_ITM_DESC") = tempDT.Rows(i).Item("ITM_NAME").ToString.Trim
                    newRow.Item("LOBD_CBM") = tempDT.Rows(i).Item("CBM_VAL")
                    newRow.Item("LOBD_STATUS") = "FREE"

                    itemDT.Rows.Add(newRow)
                Next


                itemDT.AcceptChanges()

                ViewState("n_cur_seq") = CStr(itemDT.Rows.Count)

                dt = itemDT
                ViewState("dt") = itemDT

                GridView1.DataSource = itemDT
                GridView1.DataBind()

                Dim tempsto_code = STORER_CODE.SelectedValue
                uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(tempsto_code) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , , True)

            End If

        End If




    End Sub

    Protected Sub LOBH_WH_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LOBH_WH.SelectedIndexChanged
        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                DirectCast(GridView1.Rows(i).FindControl("dsp_LOBD_LOC"), Label).Text = ""
                DirectCast(GridView1.Rows(i).FindControl("dsp_LOBD_CBM_PERC"), Label).Text = ""

                DirectCast(GridView1.Rows(i).FindControl("LOBD_CBM_PERC"), HiddenField).Value = ""
                DirectCast(GridView1.Rows(i).FindControl("LOBD_LOC"), HiddenField).Value = ""
                DirectCast(GridView1.Rows(i).FindControl("BN_UTILIZATION_TYPE"), HiddenField).Value = ""
                DirectCast(GridView1.Rows(i).FindControl("BN_CBM"), HiddenField).Value = ""
                DirectCast(GridView1.Rows(i).FindControl("BN_AREA"), HiddenField).Value = ""

            Next
        End If
    End Sub
End Class
