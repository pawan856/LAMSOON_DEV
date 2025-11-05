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
        ar = New AccessRightUtils("IB_SRC", Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME AS STO_NAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(SRC_WH, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"))

          
        End If

        If Session("pagemode") = "N" Then

            CancelBtn.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
                'Call setDefStorerInfo()
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here

        If Session("gLang") = "E" Then
            lheader.Text = "Stock Return Close-out Maintenance"
            lbl_ImageHd.Text = "SR Close-out Item Details"
            lbl_SRC_CODE.Text = "Return Code"
            lbl_SRC_STATUS.Text = "Status"
            lbl_STORER_CODE.Text = "Storer"
           

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            newrow.Text = "Add"

            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"

            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[No. will be auto generated]"
            'End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "退貨維護"
            lbl_ImageHd.Text = "退貨物件"
            lbl_SRC_CODE.Text = "退貨號碼"
            lbl_SRC_STATUS.Text = "狀態"
            lbl_STORER_CODE.Text = "儲存庫"
           


            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "保存"
            saveBtn2.Text = "保存"
            CancelBtn.Text = "取消"
            newrow.Text = "新增"

            CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            saveBtn1.OnClientClick = "return confirm(""確定保存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""確定保存資料?"");"

            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[号码会自动产生]"
            'End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS
        REM **********************

        If Session("pagemode") = "N" Then
            'SRC_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'SRC_CODE.Enabled = False
            STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("SRC_CODE") = ""

            Session("SR_MWH") = ""

            Call BindGV()
        Else
            dt = ViewState("dt")

        
        End If

        If moduleAction = "SELECTIM" Then
            addItemtoSR()
        ElseIf moduleAction = "SELECTWIT" Then
            addWIT()
        End If

        
        setPageCtrlAccess()

        Select Case SRC_STATUS.Value
            Case "NEW"
                btnClose.Visible = False
                btnUnClose.Visible = False
                btnPreweight.Visible = True
                btnUnPreweight.Visible = False
                CancelBtn.Visible = True
            Case "CANCELLED"
                ar.sec_viewMode = "Y"
                CancelBtn.Visible = False
                btnClose.Visible = False
                btnUnClose.Visible = False
                btnPreweight.Visible = False
                btnUnPreweight.Visible = False

            Case "CLOSED"
                ar.sec_viewMode = "Y"
                btnClose.Visible = False
                btnUnClose.Visible = True
                btnPreweight.Visible = False
                btnUnPreweight.Visible = False
                exceptionEditList.Add("btnUnClose")
                btnUnClose.Visible = ar.hasBtnRight("BT_SRC_UNCLOSE")
                CancelBtn.Visible = False
            Case "PREWEIGHT"
                btnUnClose.Visible = False
                btnPreweight.Visible = False
                exceptionEditList.Add("btnClose")
                exceptionEditList.Add("btnUnPreweight")
                btnClose.Visible = True
                btnUnPreweight.Visible = True
                CancelBtn.Visible = False
        End Select

        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('IB_SRC','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("SRC_CODE") & "','N');")

        'AddHandler btn.Click, AddressOf click_search

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)


        updateCount()
        reloadDSPfield()
    End Sub

    Private Sub setPageCtrlAccess()

        exceptionEditList = New List(Of String)

        exceptionEditList.Add("btnPCancel")
        exceptionEditList.Add("pnl_SRC_APPROVE_CODE")
        exceptionEditList.Add("btnPOK")
        exceptionEditList.Add("btnNew")

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

                CType(e.Row.FindControl("SEQ_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "SEQ_NO").ToString.Trim
                CType(e.Row.FindControl("RT_CODE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "RT_CODE").ToString.Trim
                CType(e.Row.FindControl("RT_STATUS"), Label).Text = DataBinder.Eval(e.Row.DataItem, "RT_STATUS").ToString.Trim

                Select Case DataBinder.Eval(e.Row.DataItem, "RT_STATUS").ToString.Trim
                    Case "NEW", "PENDING", "APPROVED"
                        e.Row.Cells(5).BackColor = Drawing.Color.Red
                    Case Else
                        e.Row.Cells(6).BackColor = System.Drawing.ColorTranslator.FromHtml("#F0F0F0")
                End Select

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to remove this SR?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    nButton.Text = "删除"
                End If

                If SRC_STATUS.Value = "NEW" Then
                    nButton.Visible = True
                Else
                    nButton.Visible = False
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    e.Row.Enabled = False
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
            Dim seq_string As String = "select MAX(convert(int, SEQ_NO)) + 1 from wms_sr_co_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and SRC_CODE = '" & gU.dbEncode(SRC_CODE.Text.Trim) & "' "
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
            dt.Rows(rows_count - 1).Item("SEQ_NO") = ViewState("n_cur_seq").ToString
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
        ViewState("dt") = dt
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


        If SRC_WH.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_SRC_WH.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_SRC_WH.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If



        If SRC_DATE.Text.Trim <> "" And Not gU.isValidDate(SRC_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_SRC_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_SRC_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If SRC_REM.Text.Trim.Length > 200 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Too many characters, maximum length of " & lbl_SRC_REM.Text & " is 200!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "字數太多, " & lbl_SRC_REM.Text & "最多只限200字!", Session("gLang"))
            End If
            Return False
        End If

        Dim itemCount As Integer = 0

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If dt.Rows(i).Item("mFlag").ToString.Trim <> "D" Then


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
        Dim maxSeq As Integer = 0
        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("SRC", gConn, transaction)
                    REM **********************

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = " INSERT INTO WMS_SR_CO " & _
                                 " (IMP_CODE, STORER_CODE, SRC_CODE, SRC_STATUS, SRC_DATE, SRC_WIT_NO, SRC_WH, SRC_REM, " & _
                                 " SRC_WIT_TOTAL_QTY, SRC_WIT_TOTAL_KG, " & _
                                 " SYS_LUB, SYS_LUD, SYS_CB, SYS_CD)  VALUES " & _
                                 "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', 'NEW', " & _
                                 gU.convdbDate(gU.dbEncode(SRC_DATE.Text.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(SRC_WIT_NO.Text.Trim)) & "," & _
                                 gU.convdbNVCData(gU.dbEncode(SRC_WH.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(SRC_REM.Text.Trim)) & "," & _
                                 gU.convdbNumData(gU.dbEncode(SRC_WIT_TOTAL_QTY.Text.Trim)) & "," & gU.convdbNumData(gU.dbEncode(SRC_WIT_TOTAL_KG.Text.Trim)) & ", " & _
                                 "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"

                    REM **********************
                    ViewState("SRC_CODE") = nextNo

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "SEQ_NO")


                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    maxSeq += 1

                                    itemSQL = "insert into wms_sr_co_D " & _
                                            "(IMP_CODE, STORER_CODE, SRC_CODE, SEQ_NO, RT_CODE, RT_STATUS_BEF, SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) " & _
                                            "values " & _
                                            "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(nextNo) & "', " & maxSeq & ", " & _
                                             gU.convdbNVCData(gU.dbEncode(rows.Item("RT_CODE").ToString.Trim)) & ",NULL, " & _
                                             "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                            End Select
                            REM **********************

                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                        Next
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode


                    sql_string = "update wms_sr_co set " & _
                                 "SRC_DATE = " & gU.convdbDate(gU.dbEncode(SRC_DATE.Text.Trim)) & ", " & _
                                 "SRC_WH = " & gU.convdbNVCData(gU.dbEncode(SRC_WH.SelectedValue)) & ", " & _
                                 "SRC_REM =" & gU.convdbNVCData(gU.dbEncode(SRC_REM.Text.Trim)) & ", " & _
                                 "SRC_WIT_NO =" & gU.convdbNVCData(gU.dbEncode(SRC_WIT_NO.Text.Trim)) & ", " & _
                                 "SRC_WIT_TOTAL_QTY =" & gU.convdbNumData(gU.dbEncode(SRC_WIT_TOTAL_QTY.Text.Trim)) & ", " & _
                                 "SRC_WIT_TOTAL_KG =" & gU.convdbNumData(gU.dbEncode(SRC_WIT_TOTAL_KG.Text.Trim)) & ", " & _
                                 "sys_lub = '" & Session("usr_id") & "', " & _
                                 "sys_lud = Getdate() " & _
                                "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and SRC_CODE = '" & gU.dbEncode(SRC_CODE.Text) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "SEQ_NO")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag").ToString.Trim
                                Case "N"
                                    maxSeq = DB.getValueFromSQL("Select ISNULL(Max(convert(int,SEQ_NO)) + 1,1) as value from WMS_SR_CO_D " & _
                                             " WHERE IMP_CODE='" & Session("IMP_CODE") & "' AND STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                             " AND SRC_CODE='" & gU.dbEncode(SRC_CODE.Text) & "'", gConn, transaction)

                                    itemSQL = "insert into wms_sr_co_D " & _
                                             "(IMP_CODE, STORER_CODE, SRC_CODE, SEQ_NO, RT_CODE, RT_STATUS_BEF, SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) " & _
                                             "values " & _
                                             "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(SRC_CODE.Text) & "', " & maxSeq & ", " & _
                                              gU.convdbNVCData(gU.dbEncode(rows.Item("RT_CODE").ToString.Trim)) & ",NULL, " & _
                                              "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                Case "D"
                                    itemSQL = "delete from wms_sr_co_D " & _
                                            "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and SRC_CODE = '" & gU.dbEncode(SRC_CODE.Text.Trim) & "' " & _
                                            "and SEQ_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("seq_no").ToString.Trim, "")) & "' "
                                Case Else
                                    itemSQL = ""
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
                    SRC_CODE.Text = nextNo
                    ViewState("SRC_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    SRC_CODE.ForeColor = Drawing.Color.Black
                    SRC_CODE.Font.Size = 10
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

        'Dim pk_code As String = ""
        Dim srcCode, storerCode As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("SRC_CODE") <> "" Then
            'pk_code = ViewState("GR_CODE")
            srcCode = ViewState("SRC_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            'pk_code = Request("GR_CODE")
            srcCode = Server.UrlDecode(Request("SRC_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))

            ViewState("SRC_CODE") = srcCode
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
            SRC_STATUS.Value = "NEW"
            DSP_SRC_STATUS.Text = "NEW"
            SRC_DATE.Text = Now.Date.ToString("dd/MM/yyyy")
            btnClose.Visible = False
            btnAttach.Visible = False
            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = "SELECT IMP_CODE, STORER_CODE, SRC_CODE, SRC_STATUS,convert(varchar, SRC_DATE," & DDFORMAT & ") as SRC_DATE, SRC_WIT_NO, SRC_WH, convert(varchar,SRC_CLOSE_DATE," & DDFORMAT & ") as SRC_CLOSE_DATE, SRC_CLOSE_BY, convert(varchar,SRC_CANCEL_DATE," & DDFORMAT & ") as SRC_CANCEL_DATE, SRC_CANCEL_BY, SRC_REM, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, " & _
                        "SRC_WIT_TOTAL_QTY, SRC_WIT_TOTAL_KG " & _
                        "FROM WMS_SR_CO " & _
                        "WHERE WMS_SR_CO.SRC_CODE = '" & gU.dbEncode(srcCode) & "' " & _
                        "AND WMS_SR_CO.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                        "AND WMS_SR_CO.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)
            If dt.Rows.Count > 0 Then
                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                SRC_CODE.Text = dt.Rows(0).Item("SRC_CODE").ToString
                SRC_STATUS.Value = dt.Rows(0).Item("SRC_STATUS").ToString

                DSP_SRC_STATUS.Text = gDB.getColValue(dt.Rows(0).Item("SRC_STATUS").ToString, "wms_sr_co.SRC_STATUS")
                SRC_CODE_HF.Value = dt.Rows(0).Item("SRC_CODE").ToString
                STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                STORER_CODE_HF.Value = dt.Rows(0).Item("STORER_CODE").ToString
                SRC_DATE.Text = dt.Rows(0).Item("SRC_DATE").ToString

                SRC_WIT_NO.Text = dt.Rows(0).Item("SRC_WIT_NO").ToString
                SRC_WH.SelectedValue = dt.Rows(0).Item("SRC_WH").ToString
                SRC_CLOSE_DATE.Text = dt.Rows(0).Item("SRC_CLOSE_DATE").ToString.Trim
                SRC_CLOSE_BY.Text = dt.Rows(0).Item("SRC_CLOSE_BY").ToString.Trim
                SRC_CANCEL_DATE.Text = dt.Rows(0).Item("SRC_CANCEL_DATE").ToString.Trim
                SRC_CANCEL_BY.Text = dt.Rows(0).Item("SRC_CANCEL_BY").ToString.Trim
                SRC_REM.Text = dt.Rows(0).Item("SRC_REM").ToString.Trim

                SRC_WIT_TOTAL_QTY.Text = dt.Rows(0).Item("SRC_WIT_TOTAL_QTY").ToString.Trim
                SRC_WIT_TOTAL_KG.Text = dt.Rows(0).Item("SRC_WIT_TOTAL_KG").ToString.Trim

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "SELECT 'U' as mFlag,WMS_SR_CO_D.IMP_CODE,WMS_SR_CO_D.STORER_CODE, WMS_SR_CO_D.SEQ_NO, WMS_SR_CO_D.RT_CODE, WMS_STOCK_RETURN.RT_STATUS, WMS_STOCK_RETURN.SYS_CD, WMS_STOCK_RETURN.RT_REF_NO2, WMS_STOCK_RETURN.RT_WH, " & _
                    "Convert(varchar,WMS_STOCK_RETURN.RT_DATE," & DDFORMAT & ") as RT_DATE, " & _
                    "SUM(WMS_STOCK_RETURN_D.RTD_RCV_QTY) as RTD_RCV_QTY, SUM(WMS_STOCK_RETURN_D.RTD_KG) as RTD_KG, " & _
                    "STUFF((SELECT ', ' + i.RTD_LOC  FROM WMS_STOCK_RETURN_D i where i.rt_code = WMS_SR_CO_D.rt_code group by i.RTD_LOC order by RTD_LOC FOR XML PATH('')),1,1,'') as LOCATION_STR, " & _
                    "STUFF((SELECT ', ' +  x.itm_sku_no from WMS_ITEM x,WMS_STOCK_RETURN_D where WMS_STOCK_RETURN_D.IMP_CODE = x.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = x.STORER_CODE AND " & _
                    "WMS_STOCK_RETURN_D.RTD_ITM_CODE = x.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = x.PACK_KEY and WMS_SR_CO_D.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_SR_CO_D.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND " & _
                    "WMS_SR_CO_D.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE GROUP BY x.itm_sku_no ORDER by x.itm_sku_no FOR XML PATH('')),1,1,'') as ITM_SKU_NO_STR " & _
                    "FROM WMS_SR_CO_D INNER JOIN " & _
                    "WMS_STOCK_RETURN ON WMS_SR_CO_D.IMP_CODE = WMS_STOCK_RETURN.IMP_CODE AND WMS_SR_CO_D.STORER_CODE = WMS_STOCK_RETURN.STORER_CODE AND " & _
                    "WMS_SR_CO_D.RT_CODE = WMS_STOCK_RETURN.RT_CODE " & _
                    "INNER JOIN WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE " & _
                    "INNER JOIN WMS_ITEM ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                    "WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                    "where WMS_SR_CO_D.SRC_CODE = '" & gU.dbEncode(srcCode) & "' " & _
                    "and WMS_SR_CO_D.STORER_CODE = '" & gU.dbEncode(storerCode) & "' " & _
                    "and WMS_SR_CO_D.IMP_CODE = '" & Session("IMP_CODE") & "'" & _
                    "GROUP BY WMS_SR_CO_D.IMP_CODE, WMS_SR_CO_D.STORER_CODE, WMS_SR_CO_D.SEQ_NO, WMS_SR_CO_D.RT_CODE, WMS_STOCK_RETURN.RT_STATUS, WMS_STOCK_RETURN.SYS_CD, WMS_STOCK_RETURN.RT_REF_NO2, WMS_STOCK_RETURN.RT_WH,WMS_STOCK_RETURN.RT_DATE "

        SQLString = SQLString & " order by convert(int, WMS_SR_CO_D.SEQ_NO)"
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
        Dim cancelSql As String = "update wms_sr_co " & _
                     "set SRC_status = 'CANCELLED', " & _
                     "SRC_CANCEL_DATE=getdate(),SRC_CANCEL_BY='" & Session("usr_id") & "'," & _
                     "sys_lub = '" & Session("usr_id") & "', " & _
                     "sys_lud = Getdate() " & _
                     "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                     "and SRC_code = '" & gU.dbEncode(SRC_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()

            SRC_STATUS.Value = "CANCELLED"
            DSP_SRC_STATUS.Text = gDB.getColValue("CANCELLED", "WMS_SR_CO.SRC_STATUS")

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

    Protected Sub addItemtoSR()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(convert(int, SEQ_NO)) + 1 from wms_sr_co_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and SRC_CODE = '" & gU.dbEncode(SRC_CODE.Text.Trim) & "' "
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
            Dim SR_dt As DataTable


            Dim itemListarray As String()
            Dim rtCodeList As String = ""

            Dim pref_loc As String = ""


            itemListarray = Split(itemList.Value, ", ")

            If itemListarray.Count = 0 Then
                If itemList.Value <> "" Then
                    rtCodeList = Server.HtmlDecode(itemList.Value)
                End If
            Else
                For i = 0 To itemListarray.Count - 1
                    rtCodeList = gU.appendToList(rtCodeList, Server.HtmlDecode(itemListarray(i)))
                Next
            End If

            SQLString = " SELECT WMS_STOCK_RETURN.IMP_CODE, WMS_STOCK_RETURN.STORER_CODE, WMS_STOCK_RETURN.RT_CODE, WMS_STOCK_RETURN.RT_STATUS, WMS_STOCK_RETURN.RT_WH, WMS_STOCK_RETURN.RT_REF_NO2, CONVERT(varchar,WMS_STOCK_RETURN.rt_date,103) as RT_DATE, " & _
                        " SUM(WMS_STOCK_RETURN_D.RTD_KG) as RTD_KG, SUM(WMS_STOCK_RETURN_D.RTD_RCV_QTY) as RTD_RCV_QTY, " & _
                        " STUFF((SELECT ', ' + i.RTD_LOC  FROM WMS_STOCK_RETURN_D i where i.rt_code = WMS_STOCK_RETURN.rt_code GROUP BY i.RTD_LOC ORDER BY i.RTD_LOC FOR XML PATH('')),1,1,'') as LOCATION_STR, " & _
                        " STUFF((SELECT ', ' +  x.itm_sku_no from WMS_ITEM x,WMS_STOCK_RETURN_D where WMS_STOCK_RETURN_D.IMP_CODE = x.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = x.STORER_CODE AND " & _
                        " WMS_STOCK_RETURN_D.RTD_ITM_CODE = x.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = x.PACK_KEY and WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND " & _
                        " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE GROUP BY x.itm_sku_no ORDER by x.itm_sku_no FOR XML PATH('')),1,1,'') as ITM_SKU_NO_STR " & _
                        " FROM WMS_STOCK_RETURN INNER JOIN " & _
                        " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND " & _
                        " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE INNER JOIN " & _
                        " WMS_ITEM ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                        " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                        " LEFT OUTER JOIN WMS_WH_BIN ON WMS_WH_BIN.LOC_KEY = WMS_STOCK_RETURN_D.RTD_LOC " & _
                        " WHERE WMS_STOCK_RETURN.RT_CODE IN ('" & Replace(rtCodeList, ", ", "', '") & "')  and WMS_STOCK_RETURN.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                        " and WMS_STOCK_RETURN.IMP_CODE = '" & Session("IMP_CODE") & "'" & _
                        " group by WMS_STOCK_RETURN.IMP_CODE, WMS_STOCK_RETURN.STORER_CODE, WMS_STOCK_RETURN.RT_CODE, WMS_STOCK_RETURN.RT_STATUS, WMS_STOCK_RETURN.RT_WH, WMS_STOCK_RETURN.RT_REF_NO2, WMS_STOCK_RETURN.rt_date " & _
                        " Order by WMS_STOCK_RETURN.RT_CODE "

            REM **********************
            SR_DT = gDB.getDataTable(SQLString)
            For i As Integer = 0 To SR_dt.Rows.Count - 1
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
                dt.Rows(rows_count - 1).Item("SEQ_NO") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("RT_CODE") = SR_dt.Rows(i).Item("RT_CODE").ToString
                dt.Rows(rows_count - 1).Item("RT_REF_NO2") = SR_dt.Rows(i).Item("RT_REF_NO2").ToString
                dt.Rows(rows_count - 1).Item("RT_STATUS") = SR_dt.Rows(i).Item("RT_STATUS").ToString
                dt.Rows(rows_count - 1).Item("RT_WH") = SR_dt.Rows(i).Item("RT_WH").ToString
                dt.Rows(rows_count - 1).Item("RT_DATE") = SR_dt.Rows(i).Item("RT_DATE").ToString
                dt.Rows(rows_count - 1).Item("RTD_KG") = SR_dt.Rows(i).Item("RTD_KG").ToString
                dt.Rows(rows_count - 1).Item("RTD_RCV_QTY") = SR_dt.Rows(i).Item("RTD_RCV_QTY").ToString
                dt.Rows(rows_count - 1).Item("LOCATION_STR") = SR_dt.Rows(i).Item("LOCATION_STR").ToString
                dt.Rows(rows_count - 1).Item("ITM_SKU_NO_STR") = SR_dt.Rows(i).Item("ITM_SKU_NO_STR").ToString

                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

            Next
            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub

    Protected Sub btnClose_Click(sender As Object, e As System.EventArgs) Handles btnClose.Click
        If 1 = 2 Then ' CInt(gU.decodeEmptyCInt(NO_OF_ISR.Text, 0)) > 0 Then

            uiFun.displayMsgNew(bcloseUDP, "closeBtn", "Pending for approval/ Un-Posted SR has selected. Close-out is not allowed!", "")

        Else
            cautionList.Items.Clear()
            Dim CautionCount As Integer = 0

            Dim nList As ListItem
            If NO_OF_ITM_SKU.Text > 1 Then
                nList = New ListItem
                nList.Text = "Multiple stock number has been selected."
                cautionList.Items.Add(nList)
                CautionCount += 1
            End If

            Dim TempDT As DataTable

            TempDT = ViewState("dt")
            If TempDT IsNot Nothing AndAlso TempDT.Rows.Count > 0 Then
                Dim sqlString As String = ""
                Dim lrt_type As String = ""
                For i = 0 To TempDT.Rows.Count - 1
                    If TempDT.Rows(i).Item("mFlag").ToString.Trim <> "D" Then
                        sqlString = "Select rt_type from wms_stock_return " & _
                                    " WHERE WMS_STOCK_RETURN.RT_CODE ='" & TempDT.Rows(i).Item("RT_CODE").ToString.Trim & "'  and WMS_STOCK_RETURN.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    " and WMS_STOCK_RETURN.IMP_CODE = '" & Session("IMP_CODE") & "'"

                        lrt_type = DB.getValueFromSQL(sqlString)

                        If lrt_type = "LOAN" Then
                            nList = New ListItem
                            nList.Text = "Unusable LOAN item selected."
                            cautionList.Items.Add(nList)
                            CautionCount += 1
                            Exit For
                        End If
                    End If
                Next
            End If

            btnClose_ModalPopupExtender.Show()
        End If

    End Sub
    Private Sub CloseSR()
        Dim successFlag As Boolean = False
        If validateAll() Then
            Call save("Y")

            Dim updateSQL As String = ""
            Dim selectSQL As String = ""
            Dim gConn As SqlConnection

            gConn = gDB.getConnection()
            Dim transaction As SqlTransaction

            transaction = gConn.BeginTransaction()

            Try
                Dim tempDT As DataTable

                selectSQL = "Select rt_code from wms_sr_co_d where src_code='" & gU.dbEncode(ViewState("SRC_CODE")) & "' and imp_code='" & Session("IMP_CODE") & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "'"
                tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                If tempDT.Rows.Count > 0 Then

                    For i = 0 To tempDT.Rows.Count - 1
                        updateSQL = "Update wms_stock_return set RT_STATUS='CLOSED', " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and RT_CODE = '" & gU.dbEncode(tempDT.Rows(i).Item("RT_CODE").ToString.Trim) & "' " & _
                                    "and RT_STATUS='PREWEIGHT' "
                        gDB.amendData(updateSQL, gConn, transaction)
                    Next

                    updateSQL = "update wms_sr_co " & _
                                "set SRC_status = 'CLOSED', " & _
                                "SRC_CLOSE_BY='" & Session("usr_id") & "'," & _
                                "SRC_CLOSE_DATE = Getdate(), " & _
                                "sys_lub = '" & Session("usr_id") & "', " & _
                                "sys_lud = Getdate() " & _
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and SRC_code = '" & gU.dbEncode(SRC_CODE.Text) & "' "
                    gDB.amendData(updateSQL, gConn, transaction)

                    successFlag = True

                    transaction.Commit()

                    SRC_STATUS.Value = "CLOSED"
                    DSP_SRC_STATUS.Text = gDB.getColValue("CLOSED", "wms_sr_co.SRC_STATUS")

                    'setPageCtrlAccess()
                    'ar.sec_viewMode = "Y"
                    'exceptionEditList.Add("btnUnClose")
                    'ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)

                    'uiFun.displayMsg(Me, "", "The Selected Stock Return have been closed!", Session("gLang"))

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

            If successFlag Then
                Dim rmtPost As New RemotePost
                rmtPost.Url = "SRCloseMain.aspx"
                rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                rmtPost.Add("SRC_CODE", ViewState("SRC_CODE"))
                rmtPost.alertMsg = "The Selected Stock Return have been closed!"
                rmtPost.Post()
            End If

        End If
    End Sub

    Private Sub PreweightSR()
        Dim successFlag As Boolean = False
        If validateAll() Then
            Call save("Y")

            Dim updateSQL As String = ""
            Dim selectSQL As String = ""
            Dim gConn As SqlConnection

            gConn = gDB.getConnection()
            Dim transaction As SqlTransaction

            transaction = gConn.BeginTransaction()

            Try
                Dim tempDT As DataTable

		selectSQL = "Select wms_sr_co_d.rt_code from wms_sr_co_d, wms_stock_return " & _
                    "where wms_sr_co_d.src_code='" & gU.dbEncode(ViewState("SRC_CODE")) & "' and wms_sr_co_d.imp_code='" & Session("IMP_CODE") & "' and wms_sr_co_d.storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                    "AND wms_stock_return.RT_CODE = wms_sr_co_d.rt_code " & _
                    "AND wms_stock_return.imp_code = wms_sr_co_d.imp_code " & _
                    "AND wms_stock_return.storer_code = wms_sr_co_d.storer_code " & _
                    "and wms_stock_return.RT_STATUS <> 'POSTED' "

                tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                If tempDT.Rows.Count > 0 Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "There are outstanding Stock return has not yet been posted, Pre-weight is not allowed.", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "There are outstanding Stock return has not yet been posted, Pre-weight is not allowed.", Session("gLang"))
                    End If
                Else

                selectSQL = "Select rt_code from wms_sr_co_d where src_code='" & gU.dbEncode(ViewState("SRC_CODE")) & "' and imp_code='" & Session("IMP_CODE") & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "'"
                tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                If tempDT.Rows.Count > 0 Then

                    For i = 0 To tempDT.Rows.Count - 1
                        updateSQL = "Update wms_stock_return set RT_STATUS='PREWEIGHT', " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and RT_CODE = '" & gU.dbEncode(tempDT.Rows(i).Item("RT_CODE").ToString.Trim) & "' " & _
                                    "and RT_STATUS='POSTED' "
                        gDB.amendData(updateSQL, gConn, transaction)
                    Next

                    updateSQL = "update wms_sr_co " & _
                                "set SRC_status = 'PREWEIGHT', " & _
                                "SRC_CLOSE_BY='" & Session("usr_id") & "'," & _
                                "SRC_CLOSE_DATE = Getdate(), " & _
                                "sys_lub = '" & Session("usr_id") & "', " & _
                                "sys_lud = Getdate() " & _
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and SRC_code = '" & gU.dbEncode(SRC_CODE.Text) & "' "
                    gDB.amendData(updateSQL, gConn, transaction)

                    successFlag = True

                    transaction.Commit()

                    SRC_STATUS.Value = "PREWEIGHT"
                    DSP_SRC_STATUS.Text = gDB.getColValue("PREWEIGHT", "wms_sr_co.SRC_STATUS")

                    uiFun.displayMsg(Me, "", "Preweight has been done!", Session("gLang"))

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

            If successFlag Then
                Dim rmtPost As New RemotePost
                rmtPost.Url = "SRCloseMain.aspx"
                rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
                rmtPost.Add("SRC_CODE", ViewState("SRC_CODE"))
                rmtPost.alertMsg = "Preweight is completed!"
                rmtPost.Post()
            End If

        End If
    End Sub

    Protected Sub reloadDSPfield()
        'If Not GridView1 Is Nothing AndAlso GridView1.Rows.Count > 0 Then
        '    For i = 0 To GridView1.Rows.Count - 1
        '        CType(GridView1.Rows(i).FindControl("dsp_rtd_loc"), Label).Text = CType(GridView1.Rows(i).FindControl("rtd_loc"), HiddenField).Value
        '    Next
        'End If
    End Sub

    Protected Sub selectItemBtn_Click(sender As Object, e As System.EventArgs) Handles selectItemBtn.Click
        Session("SR_MWH") = SRC_WH.SelectedValue
        Session("LOOKUP_MAIN_WH") = SRC_WH.SelectedValue

        ScriptManager.RegisterStartupScript(LOOKUPUDP, LOOKUPUDP.GetType, "SRLookUp", "SRLookUp('" & STORER_CODE.SelectedValue & "','" & SRC_WH.SelectedValue & "');", True)
    End Sub

    Private Sub updateCount()
        Dim tempDT As DataTable

        tempDT = ViewState("dt")

        If tempDT IsNot Nothing AndAlso tempDT.Rows.Count > 0 Then
            Dim SRStr As String = ""
            Dim SRCount As Integer = 0
            Dim I_SRCount As Integer = 0
            Dim TotalQTY1 As Integer = 0
            Dim TotalKG As Double = 0

            For i = 0 To tempDT.Rows.Count - 1
                If tempDT.Rows(i).Item("mFlag").ToString.Trim <> "D" Then
                    If Not gU.inList(SRStr, tempDT.Rows(i).Item("RT_CODE").ToString.Trim) Then
                        SRCount += 1
                        SRStr = gU.appendToList(SRStr, tempDT.Rows(i).Item("RT_CODE").ToString.Trim)

                        If tempDT.Rows(i).Item("RT_STATUS").ToString.Trim <> "POSTED" AndAlso tempDT.Rows(i).Item("RT_STATUS").ToString.Trim <> "CLOSED" AndAlso tempDT.Rows(i).Item("RT_STATUS").ToString.Trim <> "PREWEIGHT" Then
                            I_SRCount += 1
                        End If
                    End If

                    TotalQTY1 += gU.decodeEmptyCInt(tempDT.Rows(i).Item("RTD_RCV_QTY").ToString.Trim, 0)
                    TotalKG += gU.decodeEmptyCdbl(tempDT.Rows(i).Item("RTD_KG").ToString.Trim, 0)

                End If
            Next

            If SRStr <> "" Then
                Dim sqlString As String = "SELECT count(distinct wms_item.itm_sku_no) as count " & _
                                          "FROM WMS_STOCK_RETURN_D INNER JOIN " & _
                                          "WMS_ITEM ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                                          "WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                          "WHERE WMS_STOCK_RETURN_D.IMP_CODE='" & Session("IMP_CODE") & "' and WMS_STOCK_RETURN_D.storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                          "AND WMS_STOCK_RETURN_D.RT_CODE in ('" & SRStr.Replace(", ", "', '") & "')"

                NO_OF_ITM_SKU.Text = gU.decodeEmptyCInt(DB.getValueFromSQL(sqlString), 0)
                NO_OF_ISR.Text = I_SRCount
                NO_OF_SR.Text = SRCount

                TOTAL_QTY1.Text = TotalQTY1
                TOTAL_KG.Text = TotalKG

                If I_SRCount > 0 Then
                    ISR_COUNT_TD.Style.Add("background-color", "red")
                Else
                    ISR_COUNT_TD.Style.Add("background-color", "transparent")
                End If
            End If
        Else
            NO_OF_ITM_SKU.Text = 0
            NO_OF_ISR.Text = 0
            NO_OF_SR.Text = 0
            TOTAL_QTY1.Text = 0
            TOTAL_KG.Text = 0

        End If
    End Sub

    Protected Sub btnPOK_Click(sender As Object, e As System.EventArgs) Handles btnPOK.Click
        CloseSR()
    End Sub

    Protected Sub btnPreweight_Click(sender As Object, e As System.EventArgs) Handles btnPreweight.Click
        PreweightSR()
    End Sub

    Protected Sub btnUnClose_Click(sender As Object, e As System.EventArgs) Handles btnUnClose.Click
        UnCloseSR()
    End Sub

    Protected Sub btnUnPreweight_Click(sender As Object, e As System.EventArgs) Handles btnUnPreweight.Click
        UnPreweightSR()
    End Sub

    Private Sub UnCloseSR()
        Dim successFlag As Boolean = False

        Dim updateSQL As String = ""
        Dim selectSQL As String = ""
        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            Dim tempDT As DataTable

            selectSQL = "Select rt_code from wms_sr_co_d where src_code='" & gU.dbEncode(ViewState("SRC_CODE")) & "' and imp_code='" & Session("IMP_CODE") & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "'"
            tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

            If tempDT.Rows.Count > 0 Then

                For i = 0 To tempDT.Rows.Count - 1
                    updateSQL = "Update wms_stock_return set RT_STATUS='PREWEIGHT', " & _
                                "sys_lub = '" & Session("usr_id") & "', " & _
                                "sys_lud = Getdate() " & _
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and RT_CODE = '" & gU.dbEncode(tempDT.Rows(i).Item("RT_CODE").ToString.Trim) & "' " & _
                                "and RT_STATUS='CLOSED' "
                    gDB.amendData(updateSQL, gConn, transaction)
                Next

                updateSQL = "update wms_sr_co " & _
                            "set SRC_status = 'PREWEIGHT', " & _
                            "SRC_CLOSE_BY=NULL," & _
                            "SRC_CLOSE_DATE = NULL, " & _
                            "SRC_UNCLOSE_DATE=getdate()," & _
                            "SRC_UNCLOSE_BY='" & Session("usr_id") & "', " & _
                            "sys_lub = '" & Session("usr_id") & "', " & _
                            "sys_lud = Getdate() " & _
                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            "and SRC_code = '" & gU.dbEncode(ViewState("SRC_CODE")) & "' "
                gDB.amendData(updateSQL, gConn, transaction)

                successFlag = True
                transaction.Commit()

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

        If successFlag Then
            Dim rmtPost As New RemotePost
            rmtPost.Url = "SRCloseMain.aspx"
            rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
            rmtPost.Add("SRC_CODE", ViewState("SRC_CODE"))
            rmtPost.alertMsg = "UnClose is completed!"
            rmtPost.Post()
        End If

    End Sub

    Private Sub UnPreweightSR()
        Dim successFlag As Boolean = False

        Dim updateSQL As String = ""
        Dim selectSQL As String = ""
        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            Dim tempDT As DataTable

            selectSQL = "Select rt_code from wms_sr_co_d where src_code='" & gU.dbEncode(ViewState("SRC_CODE")) & "' and imp_code='" & Session("IMP_CODE") & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "'"
            tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

            If tempDT.Rows.Count > 0 Then

                For i = 0 To tempDT.Rows.Count - 1
                    updateSQL = "Update wms_stock_return set RT_STATUS='POSTED', " & _
                                "sys_lub = '" & Session("usr_id") & "', " & _
                                "sys_lud = Getdate() " & _
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and RT_CODE = '" & gU.dbEncode(tempDT.Rows(i).Item("RT_CODE").ToString.Trim) & "' " & _
                                "and RT_STATUS='PREWEIGHT' "
                    gDB.amendData(updateSQL, gConn, transaction)
                Next

                updateSQL = "update wms_sr_co " & _
                            "set SRC_status = 'NEW', " & _
                            "SRC_CLOSE_BY=NULL," & _
                            "SRC_CLOSE_DATE = NULL, " & _
                            "SRC_UNCLOSE_DATE=getdate()," & _
                            "SRC_UNCLOSE_BY='" & Session("usr_id") & "', " & _
                            "sys_lub = '" & Session("usr_id") & "', " & _
                            "sys_lud = Getdate() " & _
                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            "and SRC_code = '" & gU.dbEncode(ViewState("SRC_CODE")) & "' "
                gDB.amendData(updateSQL, gConn, transaction)

                successFlag = True
                transaction.Commit()

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

        If successFlag Then
            Dim rmtPost As New RemotePost
            rmtPost.Url = "SRCloseMain.aspx"
            rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
            rmtPost.Add("SRC_CODE", ViewState("SRC_CODE"))
            rmtPost.alertMsg = "UnPreweight is completed!"
            rmtPost.Post()
        End If

    End Sub

    Protected Sub addWIT()
        Dim witListArr() As String
        Dim resultList As String = ""
        If witmList.Value <> "" Then
            witListArr = Split(witmList.Value, ", ")
            resultList = gU.formatList(SRC_WIT_NO.Text.Trim)

            For i = 0 To witListArr.Length - 1
                If Not gU.inList(resultList, witListArr(i)) Then
                    resultList = gU.appendToList(resultList, witListArr(i))
                End If
            Next

            SRC_WIT_NO.Text = resultList

            If Not String.IsNullOrWhiteSpace(resultList) Then

                witListArr = Split(resultList, ", ")
                Dim tempSQL As String = ""

                For i = 0 To witListArr.Length - 1
                    tempSQL = gU.appendToList(tempSQL, "'" & witListArr(i) & "'")
                Next
                If tempSQL <> "" Then
                    Dim tempDT As DataTable
                    Dim selectSQL As String = "select sum(ISNULL(DOD_QTY,0)) as TOTAL_QTY1, sum(ISNULL(DOD_TOT_WGT,0)) as TOTAL_KG from wms_delv_order_d " & _
                                              "WHERE IMP_CODE='" & Session("IMP_CODE") & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND DO_CODE IN (" & tempSQL & ") "
                    '"group by imp_code, storer_code, DO_CODE"

                    tempDT = gDB.getDataTable(selectSQL)

                    If tempDT.Rows.Count > 0 Then
                        SRC_WIT_TOTAL_KG.Text = tempDT.Rows(0).Item("TOTAL_KG").ToString.Trim
                        SRC_WIT_TOTAL_QTY.Text = tempDT.Rows(0).Item("TOTAL_QTY1").ToString.Trim
                    End If
                End If
            End If

        End If
        

    End Sub

    Protected Sub btnWITLookup_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnWITLookup.Click
        'Session("SR_MWH") = SRC_WH.SelectedValue
        'Session("LOOKUP_MAIN_WH") = SRC_WH.SelectedValue

        Session("SR_MWH") = ""
        Session("LOOKUP_MAIN_WH") = ""


        ScriptManager.RegisterStartupScript(LOOKUPUDP, LOOKUPUDP.GetType, "WITLookUp", "WITLookUp('" & STORER_CODE.SelectedValue & "');", True)
    End Sub
End Class
