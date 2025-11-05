Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class UserGPMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Private dt As New DataTable
    Private dt_l As New DataTable

    Dim checkBoxValue() As String = {"Y", "N"}
    Private prgmSelectedText, prgmSelectedValue As ArrayList
    Private idSelectedText, idSelectedValue As ArrayList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            Session("dt") = Nothing
            Session("dt_l") = Nothing

            Session("n_cur_seq") = ""
            Session("grp_code") = ""

            Call BindGV()
        Else
            dt = Session("dt")
            dt_l = Session("dt_l")
        End If

        prgmSelectedText = New ArrayList
        prgmSelectedValue = New ArrayList
        idSelectedText = New ArrayList
        idSelectedValue = New ArrayList

        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("ADM_002", Session("usr_id"), Me)
        'ar = New AccessRightUtils("ADM_002", "FU", Me)

        ar.hideForm(Me)
        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "User Group Maintenance"
            userRoleHeader.Text = "User Role"
            lbl_grp_code.Text = "Group Code"
            lblPrgm.Text = "Program Access"
            lbl_grp_name.Text = "Name"
            lbl_grp_type.Text = "Type"
            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            newrow.Text = "Add"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            addPrgm.Text = "Add Program"
            GridView1.EmptyDataText = "No Record"
            GridView2.EmptyDataText = "No Record"
        ElseIf Session("gLang") = "C" Then
            lheader.Text = "维护"
            userRoleHeader.Text = "User Role"
            lbl_grp_code.Text = "Group Code"
            lblPrgm.Text = "Program Access"
            lbl_grp_name.Text = "Name"
            lbl_grp_type.Text = "类型"
            lbl_sys_cb.Text = "创建者"
            lbl_sys_lub.Text = "上次更新者"
            lbl_sys_cd.Text = "创建日期"
            lbl_sys_lud.Text = "上次更新日期"
            saveBtn1.Text = "保存"
            saveBtn2.Text = "保存"
            saveBtn1.OnClientClick = "return confirm(""确定保存资料?"");"
            saveBtn2.OnClientClick = "return confirm(""确定保存资料?"");"
            newrow.Text = "新增"
            addPrgm.Text = "新增"
            GridView1.EmptyDataText = "没有资料记录"
            GridView2.EmptyDataText = "没有资料记录"
        End If
        REM **********************

        REM **********************
        REM Additional CSS
        grp_code.CssClass = "REQUIRED"
        grp_name.CssClass = "REQUIRED"
        REM **********************

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here
                Call cU.changeGVLabel(oGridViewRow, e, "User ID", "User ID")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        End Select
    End Sub

    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here
                Call cU.changeGVLabel(oGridViewRow, e, "Program ID", "Program ID")
                'Call cU.changeGVLabel(oGridViewRow, e, "Read", "Read")
                Call cU.changeGVLabel(oGridViewRow, e, "Write", "Write")
                Call cU.changeGVLabel(oGridViewRow, e, "Report", "Report")
                Call cU.changeGVLabel(oGridViewRow, e, "Print", "Print")
                Call cU.changeGVLabel(oGridViewRow, e, "Access", "Access")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim
                Dim object_name As String = ""

                object_name = "usr_id"

                Dim uDDL As DropDownList = CType(e.Row.FindControl(object_name), DropDownList)
                Dim list_item As ListItem
                Dim isExit As Boolean = False
                Dim removeIndex As New ArrayList

                If idSelectedText Is Nothing Then idSelectedText = New ArrayList
                If idSelectedValue Is Nothing Then idSelectedValue = New ArrayList

                If Not DataBinder.Eval(e.Row.DataItem, "usr_id").ToString.Trim = "" Then

                    If xFlag <> "N" Then
                        list_item = New ListItem

                        list_item.Text = DB.getValueFromSQL("select usr_fname + ' ' + usr_sname as usr_name from wms_user where usr_id = '" & _
                                                            DataBinder.Eval(e.Row.DataItem, "usr_id").ToString.Trim & "'")
                        list_item.Value = DataBinder.Eval(e.Row.DataItem, "usr_id").ToString.Trim
                        uDDL.Items.Add(list_item)

                        idSelectedText.Add(list_item.Text)
                        idSelectedValue.Add(list_item.Value)
                    Else
                        uiFun.load_dropdown(uDDL, "select usr_id, usr_fname + ' ' + usr_sname as usr_name from wms_user", "usr_id", "usr_name")
                        For i As Integer = 0 To idSelectedText.Count - 1
                            uDDL.Items.Remove(New ListItem(idSelectedText(i), idSelectedValue(i)))
                        Next

                        uDDL.SelectedValue = DataBinder.Eval(e.Row.DataItem, "usr_id").ToString.Trim
                    End If
                Else
                    uiFun.load_dropdown(uDDL, "select usr_id, usr_fname + ' ' + usr_sname as usr_name from wms_user", "usr_id", "usr_name")

                    For i As Integer = 0 To idSelectedText.Count - 1
                        uDDL.Items.Remove(New ListItem(idSelectedText(i), idSelectedValue(i)))
                    Next

                End If

                REM **********************

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Remove"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否确定要删除这个资料?')")
                    nButton.Text = "移除"
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
            dt.Rows.Add()
            rows_count = dt.Rows.Count
            dt.Rows(rows_count - 1).Item("mFlag") = "N"
            dt.AcceptChanges()

            Session("dt") = dt
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

        If grp_code.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_grp_code.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_grp_code.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If grp_name.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_grp_name.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_grp_name.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim prgmSQL As String = ""
        Dim gConn As SqlConnection
        Dim nextNo As String = ""

        If validateAll() Then
            gConn = gDB.getConnection()
            Dim transaction As SqlTransaction = Nothing
            Dim user_coll As New ArrayList
            Dim prgm_coll As New ArrayList

            Try
                gConn = gDB.getConnection()
                transaction = gConn.BeginTransaction("localTransaction")
                ' Start a local transaction

                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "insert into wms_user_group " & _
                    "(grp_code, grp_name, grp_type, sys_cb, sys_cd, sys_lub, sys_lud)" & _
                    "values ( " & _
                    "'" & gU.dbEncode(grp_code.Text) & "'," & _
                    "'" & gU.dbEncode(grp_name.Text) & "'," & _
                    "'" & gU.dbEncode(grp_type.SelectedValue) & "'," & _
                    "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
                        For Each rows As DataRow In dt.Rows
                            If gU.decodeNull(rows.Item("usr_id").ToString.Trim, "") <> "" Then
                                itemSQL = ""
                                REM **********************
                                REM Modify Here
                                Select Case rows.Item("mFlag")
                                    Case "N"
                                        itemSQL = "insert into wms_user_group_alloc " & _
                                                "(grp_code, usr_id, sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                "values " & _
                                                "(N'" & gU.dbEncode(grp_code.Text) & "'," & _
                                                "'" & gU.dbEncode(gU.decodeNull(rows.Item("usr_id").ToString.Trim, "")) & "', " & _
                                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                End Select
                                REM **********************

                                If Not user_coll.Count < 2 Then
                                    For i As Integer = 0 To user_coll.Count - 1
                                        If user_coll(i) = gU.decodeNull(rows.Item("usr_id").ToString.Trim, "") Then
                                            If Not transaction Is Nothing Then
                                                transaction.Rollback()
                                                transaction = Nothing
                                            End If

                                            uiFun.displayMsg(Me, "", "Duplicate User ID is selected, please select an unique user.", Session("gLang"))

                                            Exit Sub
                                        End If
                                    Next
                                End If

                                user_coll.Add(gU.decodeNull(rows.Item("usr_id").ToString.Trim, ""))

                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                            End If
                        Next
                    End If

                    If cU.gfBuildDataTableforGridView(dt_l, GridView2, True, checkBoxValue) Then
                        For Each p_rows As DataRow In dt_l.Rows
                            If gU.dbEncode(gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, "")) <> "" Then
                                prgmSQL = ""
                                REM **********************
                                REM Modify Here
                                Select Case p_rows.Item("mFlag")
                                    Case "N"
                                        prgmSQL = "insert into wms_sec_access " & _
                                                "(grp_code, fun_code, " & _
                                                "sec_read, sec_write, sec_report, sec_print, sec_access, " & _
                                                "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                "values ( " & _
                                                "'" & gU.dbEncode(grp_code.Text) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, "")) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_read").ToString.Trim, "")) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_write").ToString.Trim, "")) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_report").ToString.Trim, "")) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_print").ToString.Trim, "")) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_access").ToString.Trim, "")) & "', " & _
                                                "'" & Session("usr_id") & "',Getdate(), " & _
                                                "'" & Session("usr_id") & "',Getdate() " & _
                                                ") "
                                    Case "D"
                                        prgmSQL = "delete from wms_sec_access where grp_code = '" & gU.dbEncode(grp_code.Text.Trim) & "' " & _
                                                    "and fun_code = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, "")) & "' "
                                    Case Else
                                        prgmSQL = "update wms_sec_access set " & _
                                                    "sec_read = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_read").ToString.Trim, "")) & "', " & _
                                                    "sec_write = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_write").ToString.Trim, "")) & "', " & _
                                                    "sec_report = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_report").ToString.Trim, "")) & "', " & _
                                                    "sec_print = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_print").ToString.Trim, "")) & "', " & _
                                                    "sec_access = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_access").ToString.Trim, "")) & "', " & _
                                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                                    "sys_lud = Getdate() where grp_code = '" & gU.dbEncode(grp_code.Text.Trim) & "' " & _
                                                    "and fun_code = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, "")) & "' "

                                End Select
                                REM **********************

                                If Not prgm_coll.Count < 2 Then
                                    For i As Integer = 0 To prgm_coll.Count - 1
                                        If prgm_coll(i) = gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, "") Then
                                            If Not transaction Is Nothing Then
                                                transaction.Rollback()
                                                transaction = Nothing
                                            End If

                                            uiFun.displayMsg(Me, "", "Duplicate Program Code is selected, please select an unique code.", Session("gLang"))

                                            Exit Sub
                                        End If
                                    Next
                                End If

                                prgm_coll.Add(gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, ""))

                                If prgmSQL <> "" Then gDB.amendData(prgmSQL, gConn, transaction)
                            End If
                        Next
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_user_group set " & _
                    "grp_name = '" & gU.dbEncode(grp_name.Text) & "', " & _
                    "grp_type = '" & gU.dbEncode(grp_type.SelectedValue) & "', " & _
                    "sys_lub = '" & Session("usr_id") & "', " & _
                    "sys_lud = Getdate() where grp_code = '" & gU.dbEncode(grp_code.Text.Trim) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
                        For Each rows As DataRow In dt.Rows
                            If gU.decodeNull(rows.Item("usr_id").ToString.Trim, "") <> "" Then
                                itemSQL = ""
                                REM **********************
                                REM Modify Here
                                Select Case rows.Item("mFlag")
                                    Case "N"
                                        itemSQL = "insert into wms_user_group_alloc " & _
                                                "(grp_code, usr_id, sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                "values " & _
                                                "('" & gU.dbEncode(grp_code.Text) & "'," & _
                                                "'" & gU.dbEncode(gU.decodeNull(rows.Item("usr_id").ToString.Trim, "")) & "', " & _
                                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                    Case "D"
                                        itemSQL = "delete from wms_user_group_alloc where grp_code = '" & gU.dbEncode(grp_code.Text.Trim) & "' and usr_id = '" & gU.dbEncode(gU.decodeNull(rows.Item("usr_id").ToString.Trim, "")) & "' "
                                    Case Else
                                        itemSQL = ""
                                End Select
                                REM **********************

                                'Response.Write(itemSQL)

                                If Not user_coll.Count < 2 Then
                                    For i As Integer = 0 To user_coll.Count - 1
                                        If user_coll(i) = gU.decodeNull(rows.Item("usr_id").ToString.Trim, "") Then
                                            If Not transaction Is Nothing Then
                                                transaction.Rollback()
                                                transaction = Nothing
                                            End If

                                            uiFun.displayMsg(Me, "", "Duplicate User ID is selected, please select an unique user.", Session("gLang"))

                                            Exit Sub
                                        End If
                                    Next
                                End If

                                user_coll.Add(gU.decodeNull(rows.Item("usr_id").ToString.Trim, ""))

                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                            End If
                        Next
                    End If

                    If cU.gfBuildDataTableforGridView(dt_l, GridView2, True, checkBoxValue) Then
                        For Each p_rows As DataRow In dt_l.Rows
                            If gU.dbEncode(gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, "")) <> "" Then
                                prgmSQL = ""
                                REM **********************
                                REM Modify Here
                                Select Case p_rows.Item("mFlag")
                                    Case "N"
                                        prgmSQL = "insert into wms_sec_access " & _
                                                "(grp_code, fun_code, " & _
                                                "sec_read, sec_write, sec_report, sec_print, sec_access, " & _
                                                "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                "values ( " & _
                                                "'" & gU.dbEncode(grp_code.Text) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, "")) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_read").ToString.Trim, "")) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_write").ToString.Trim, "")) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_report").ToString.Trim, "")) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_print").ToString.Trim, "")) & "', " & _
                                                "'" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_access").ToString.Trim, "")) & "', " & _
                                                "'" & Session("usr_id") & "',Getdate(), " & _
                                                "'" & Session("usr_id") & "',Getdate() " & _
                                                ") "
                                    Case "D"
                                        prgmSQL = "delete from wms_sec_access where grp_code = '" & gU.dbEncode(grp_code.Text.Trim) & "' " & _
                                                    "and fun_code = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, "")) & "' "
                                    Case Else
                                        prgmSQL = "update wms_sec_access set " & _
                                                    "sec_read = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_read").ToString.Trim, "")) & "', " & _
                                                    "sec_write = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_write").ToString.Trim, "")) & "', " & _
                                                    "sec_report = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_report").ToString.Trim, "")) & "', " & _
                                                    "sec_print = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_print").ToString.Trim, "")) & "', " & _
                                                    "sec_access = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("sec_access").ToString.Trim, "")) & "', " & _
                                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                                    "sys_lud = Getdate() where grp_code = '" & gU.dbEncode(grp_code.Text.Trim) & "' " & _
                                                    "and fun_code = '" & gU.dbEncode(gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, "")) & "' "
                                End Select
                                REM **********************

                                If Not prgm_coll.Count < 2 Then
                                    For i As Integer = 0 To prgm_coll.Count - 1
                                        If prgm_coll(i) = gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, "") Then
                                            If Not transaction Is Nothing Then
                                                transaction.Rollback()
                                                transaction = Nothing
                                            End If

                                            uiFun.displayMsg(Me, "", "Duplicate Program Code is selected, please select an unique code.", Session("gLang"))

                                            Exit Sub
                                        End If
                                    Next
                                End If

                                prgm_coll.Add(gU.decodeNull(p_rows.Item("fun_code").ToString.Trim, ""))

                                If prgmSQL <> "" Then gDB.amendData(prgmSQL, gConn, transaction)
                            End If
                        Next
                    End If
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                If Not transaction Is Nothing Then
                    transaction.Commit()
                    transaction = Nothing
                End If

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                End If

                If flag <> "Y" Then uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                Call BindGV()
            Catch ex As Exception
                If Not transaction Is Nothing Then
                    transaction.Rollback()
                    transaction = Nothing
                End If

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
        Dim dt_l As New DataTable


        REM **********************
        REM Modify Here
        REM Primary Key Session
        pk_code = Request("grp_code")
        REM **********************

        REM**********************
        REM Generate Dropdown List from WMS_COL_CODE Table
        'uiFun.load_dropdownBy_ColCode(quo_currency, "WMS_QUOTATION_HD.QUO_CURRENCY", Session("gLang"))
        REM **********************

        If Session("pagemode") <> "N" Then
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = "SELECT * " & _
                   " from wms_user_group WHERE grp_code = '" & gU.dbEncode(pk_code) & "' "
            dt = gDB.getDataTable(SQLString)
            If dt.Rows.Count > 0 Then
                grp_code.Text = dt.Rows(0).Item("grp_code").ToString
                grp_name.Text = dt.Rows(0).Item("grp_name").ToString
                grp_type.SelectedValue = dt.Rows(0).Item("grp_type").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                'grp_code.Enabled = False
                grp_code.ReadOnly = True
                grp_code.BackColor = Drawing.Color.Transparent
                grp_code.BorderWidth = 0
                grp_code.CssClass = ""
                REM **********************
            End If

            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "SELECT usr_id, 'U' as mFlag " & _
                    " from wms_user_group_alloc where grp_code = '" & gU.dbEncode(pk_code) & "' "
        SQLString = SQLString & " order by usr_id"
        REM **********************

        dt = gDB.getDataTable(SQLString)
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If

        REM **********************
        REM Modify Here
        REM Generate Data Table from access list
        SQLString = "SELECT s.fun_code,f.fun_type, s.sec_read, s.sec_write, s.sec_report, s.sec_print, s.sec_access, 'U' as mFlag, max(f.fun_display_seq) " & _
               " from wms_sec_access s, wms_function f " & _
               "WHERE s.grp_code = '" & gU.dbEncode(pk_code) & "' " & _
               "and s.fun_code = f.fun_code " & _
               "and f.fun_type in ('SCREEN', 'BUTTON') " & _
               "group by s.fun_code,f.fun_type, s.sec_read, s.sec_write, s.sec_report, s.sec_print, s.sec_access " & _
               "order by max(f.fun_display_seq) "
        REM **********************

        dt_l = gDB.getDataTable(SQLString)
        If dt_l.Rows.Count > 0 Then
            GridView2.DataSource = dt_l
        Else
            GridView2.DataSource = Nothing
        End If

        Session("dt") = dt
        Session("dt_l") = dt_l
        GridView1.DataBind()
        GridView2.DataBind()

        REM **********************
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        Dim ddsql As String = ""

        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim
                Dim object_name As String = ""

                object_name = "fun_code"

                Dim fDDL As DropDownList = CType(e.Row.FindControl(object_name), DropDownList)
                Dim list_item As ListItem
                Dim isExit As Boolean = False
                Dim removeIndex As New ArrayList


                If prgmSelectedText Is Nothing Then prgmSelectedText = New ArrayList
                If prgmSelectedValue Is Nothing Then prgmSelectedValue = New ArrayList

                If Not DataBinder.Eval(e.Row.DataItem, "fun_code").ToString.Trim = "" Then

                    If xFlag <> "N" Then
                        list_item = New ListItem

                        'list_item.Text = DB.getValueFromSQL("select fun_code + ' - ' + fun_eng_name as fun_name from wms_function where fun_type ='SCREEN' and fun_code = '" & _
                        '                                    DataBinder.Eval(e.Row.DataItem, "fun_code").ToString.Trim & "'")

                        ddSql = "select case " & _
                                        "when fun_type = 'SCREEN' then fun_code + ' - ' " & _
                                        "when fun_type = 'BUTTON' then fun_parent_code + ' ---------- ' end " & _
                                        " + fun_eng_name as fun_name " & _
                                "from wms_function " & _
                                "where fun_code = '" & DataBinder.Eval(e.Row.DataItem, "fun_code").ToString.Trim & "'"

                        list_item.Text = DB.getValueFromSQL(ddSql)

                        list_item.Value = DataBinder.Eval(e.Row.DataItem, "fun_code").ToString.Trim
                        fDDL.Items.Add(list_item)

                        prgmSelectedText.Add(list_item.Text)
                        prgmSelectedValue.Add(list_item.Value)
                    Else
                        'uiFun.load_dropdown(fDDL, "select fun_code, fun_code + ' - ' + fun_eng_name as fun_name from wms_function where fun_type ='SCREEN'", "fun_code", "fun_name")

                        ddSql = "select fun_code, " & _
                                   "case " & _
                                       "when fun_type = 'SCREEN' then fun_code + ' - ' " & _
                                       "when fun_type = 'BUTTON' then fun_parent_code + ' ---------- ' end " & _
                                       " + fun_eng_name as fun_name " & _
                               "from wms_function " & _
                               "where fun_type in ('SCREEN', 'BUTTON') " & _
                               "order by fun_display_seq"

                        uiFun.load_dropdown(fDDL, ddSql, "fun_code", "fun_name")

                        For i As Integer = 0 To prgmSelectedText.Count - 1
                            fDDL.Items.Remove(New ListItem(prgmSelectedText(i), prgmSelectedValue(i)))
                        Next

                        fDDL.SelectedValue = DataBinder.Eval(e.Row.DataItem, "fun_code").ToString.Trim
                    End If
                Else
                    'uiFun.load_dropdown(fDDL, "select fun_code, fun_code + ' - ' + fun_eng_name as fun_name from wms_function where fun_type ='SCREEN'", "fun_code", "fun_name")

                    ddSql = "select fun_code, " & _
                                    "case " & _
                                        "when fun_type = 'SCREEN' then fun_code + ' - ' " & _
                                        "when fun_type = 'BUTTON' then fun_parent_code + ' ---------- ' end " & _
                                        " + fun_eng_name as fun_name " & _
                                "from wms_function " & _
                                "where fun_type in ('SCREEN', 'BUTTON') " & _
                                "order by fun_display_seq"

                    uiFun.load_dropdown(fDDL, ddSql, "fun_code", "fun_name")

                    'fDDL.SelectedValue = ""

                    For i As Integer = 0 To prgmSelectedText.Count - 1
                        fDDL.Items.Remove(New ListItem(prgmSelectedText(i), prgmSelectedValue(i)))
                    Next

                End If

                If DataBinder.Eval(e.Row.DataItem, "fun_type").ToString.Trim = "BUTTON" Then
                    CType(e.Row.FindControl("sec_read"), CheckBox).Visible = False
                    CType(e.Row.FindControl("sec_write"), CheckBox).Visible = False
                    CType(e.Row.FindControl("sec_report"), CheckBox).Visible = False
                    CType(e.Row.FindControl("sec_print"), CheckBox).Visible = False
                    CType(e.Row.FindControl("sec_access"), CheckBox).Visible = False
                Else
                    object_name = "sec_read"
                    If DataBinder.Eval(e.Row.DataItem, "sec_read").ToString.Trim.ToUpper = "Y" Then
                        CType(e.Row.FindControl(object_name), CheckBox).Checked = True
                    Else
                        CType(e.Row.FindControl(object_name), CheckBox).Checked = False
                    End If

                    object_name = "sec_write"
                    If DataBinder.Eval(e.Row.DataItem, "sec_write").ToString.Trim.ToUpper = "Y" Then
                        CType(e.Row.FindControl(object_name), CheckBox).Checked = True
                    Else
                        CType(e.Row.FindControl(object_name), CheckBox).Checked = False
                    End If

                    object_name = "sec_report"
                    If DataBinder.Eval(e.Row.DataItem, "sec_report").ToString.Trim.ToUpper = "Y" Then
                        CType(e.Row.FindControl(object_name), CheckBox).Checked = True
                    Else
                        CType(e.Row.FindControl(object_name), CheckBox).Checked = False
                    End If

                    object_name = "sec_print"
                    If DataBinder.Eval(e.Row.DataItem, "sec_print").ToString.Trim.ToUpper = "Y" Then
                        CType(e.Row.FindControl(object_name), CheckBox).Checked = True
                    Else
                        CType(e.Row.FindControl(object_name), CheckBox).Checked = False
                    End If

                    object_name = "sec_access"
                    If DataBinder.Eval(e.Row.DataItem, "sec_access").ToString.Trim.ToUpper = "Y" Then
                        CType(e.Row.FindControl(object_name), CheckBox).Checked = True
                    Else
                        CType(e.Row.FindControl(object_name), CheckBox).Checked = False
                    End If
                End If
                REM **********************
                object_name = "btnDelete_P"
                Dim nButton As Button = CType(e.Row.FindControl(object_name), Button)

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否确定要删除这个资料?')")
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

    Protected Sub addPrgm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles addPrgm.Click
        If cU.gfBuildDataTableforGridView(dt_l, GridView2, True, checkBoxValue) Then
            Dim rows_count As Integer = 0
            dt_l.Rows.Add()
            rows_count = dt_l.Rows.Count
            dt_l.Rows(rows_count - 1).Item("mFlag") = "N"
            dt_l.AcceptChanges()

            Session("dt_l") = dt_l
            GridView2.DataSource = dt_l
            GridView2.DataBind()
        End If
    End Sub

    Protected Sub GridView2_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView2.RowDeleting
        Call ar.hideGVRow(GridView2, GridView2.Rows(e.RowIndex))
        dt_l.Rows(e.RowIndex).Item("mFlag") = "D"
        dt_l.AcceptChanges()
    End Sub
End Class
