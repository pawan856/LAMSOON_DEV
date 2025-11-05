Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class PROJ_ProjMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As AccessRightUtils

    Private imp_code As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim p_status As String = ""

        ar = New AccessRightUtils("PROJ_001", Session("usr_id"), Me)

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            Session("pagemode") = Request("mode")

            REM**********************
            REM Generate Dropdown List from CMS_COL_CODE Table
            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME  from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
            REM **********************

            REM **********************
            REM Modify Here
            If Session("pagemode") = "N" Then
                REM **********************
                REM Modify Here
                prj_code.ForeColor = Drawing.Color.Red
                prj_status.Text = "NEW"
                p_status = "NEW"
                CancelBtn.Visible = False
                REM **********************
            Else

                Dim SQLString As String = ""
                Dim dt As New DataTable
                Dim WhereStr As String = ""

                SQLString = "SELECT * from wms_project"

                If Request("prj_code") <> "" Then
                    WhereStr = " WHERE IMP_CODE = '" & gU.dbEncode(imp_code) & "' and STORER_CODE = '" & gU.dbEncode(Request("storer_code")) & "' and prj_code = '" & gU.dbEncode(Request("prj_code")) & "' "
                End If

                SQLString = SQLString & " " & WhereStr

                dt = gDB.getDataTable(SQLString)

                If dt.Rows.Count > 0 Then
                    prj_code.Text = dt.Rows(0).Item("prj_code").ToString
                    prj_name.Text = dt.Rows(0).Item("prj_name").ToString
                    STORER_CODE.SelectedValue = dt.Rows(0).Item("storer_code").ToString
                    prj_date.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("prj_date").ToString)
                    prj_end_date.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("prj_end_date").ToString)
                    prj_status.Text = dt.Rows(0).Item("prj_status").ToString
                    prj_desc.Text = dt.Rows(0).Item("prj_desc").ToString
                    sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                    sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                    sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                    sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                    p_status = dt.Rows(0).Item("prj_status").ToString

                    ViewState("p_status") = p_status

                End If
            End If
            REM **********************
        Else
            p_status = ViewState("p_status")
        End If

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Project Master"
            lprj_code.Text = "Project Code"
            lprj_name.Text = "Project Name"
            lprj_date.Text = "Project Start Date"
            lprj_end_date.Text = "Project End Date"
            lprj_status.Text = "Status"
            lbl_STORER_CODE.Text = "Storer"
            lprj_desc.Text = "Description"
            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            If Session("pagemode") = "N" Then
                prj_code.Text = "[Code will be auto generated]"
            End If
        ElseIf Session("gLang") = "C" Then
            lheader.Text = "项目维护"
            lprj_code.Text = "项目代码"
            lprj_name.Text = "项目名称"
            lprj_date.Text = "项目开始日期"
            lprj_end_date.Text = "项目结束日期"
            lprj_status.Text = "状态"
            lbl_STORER_CODE.Text = "貨主代碼"
            lprj_desc.Text = "描述 "
            lbl_sys_cb.Text = "创建者"
            lbl_sys_lub.Text = "上次更新者"
            lbl_sys_cd.Text = "创建日期"
            lbl_sys_lud.Text = "上次更新日期"
            saveBtn1.Text = "保存"
            saveBtn2.Text = "保存"
            CancelBtn.Text = "取消"
            CancelBtn.OnClientClick = "return confirm(""确定取消资料?"");"
            saveBtn1.OnClientClick = "return confirm(""确定保存资料?"");"
            saveBtn2.OnClientClick = "return confirm(""确定保存资料?"");"
            If Session("pagemode") = "N" Then
                prj_code.Text = "[代码会自动产生]"
            End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS
        prj_date.CssClass = "REQUIRED"
        STORER_CODE.CssClass = "REQUIRED"
        prj_name.CssClass = "REQUIRED"
        REM **********************

        REM ****************************
        REM Modify Access Right Here

        If p_status = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
        End If

        ar.hideForm(Me)
        REM ****************************

    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""

        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If prj_name.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lprj_name.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lprj_name.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If prj_date.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lprj_date.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lprj_date.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        Return True

    End Function

    Protected Sub save()
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim gConn As SqlConnection
        Dim nextNo As String = ""
        gConn = gDB.getConnection()

        If validateAll() Then
            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                REM **********************
                REM Modify Here
                If Session("pagemode") = "N" Then

                    nextNo = DB.getDocNo("PROJ", gConn, transaction)

                    sql_string = "insert into wms_project (imp_code, storer_code, prj_code, prj_name, prj_date, prj_end_date, " & _
                                                        "prj_status, prj_desc, sys_cb, sys_cd) values ( " & _
                                                        "'" & gU.dbEncode(imp_code) & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', " & _
                                                        "N'" & gU.dbEncode(nextNo) & "', '" & gU.dbEncode(prj_name.Text.Trim) & "', " & _
                                                        gU.convdbDate(gU.dbEncode(prj_date.Text)) & ", " & gU.convdbDate(gU.dbEncode(prj_end_date.Text)) & ", " & _
                                                        "N'" & gU.dbEncode(prj_status.Text) & "', " & _
                                                        "N'" & gU.dbEncode(prj_desc.Text) & "',N'" & Session("usr_id") & "',Getdate())"

                Else
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_project set storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', " & _
                                    "prj_code = N'" & gU.dbEncode(prj_code.Text) & "', " & _
                                    "prj_name = N'" & gU.dbEncode(prj_name.Text) & "', " & _
                                    "prj_date = " & gU.convdbDate(gU.dbEncode(prj_date.Text)) & ", " & _
                                    "prj_end_date = " & gU.convdbDate(gU.dbEncode(prj_end_date.Text)) & ", " & _
                                    "prj_status = N'" & gU.dbEncode(prj_status.Text) & "', " & _
                                    "prj_desc = N'" & gU.dbEncode(prj_desc.Text) & "', " & _
                                    "sys_lub = N'" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where prj_code = '" & gU.dbEncode(prj_code.Text.Trim) & "' "
                End If
                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    prj_code.Text = nextNo
                    prj_code.ForeColor = Drawing.Color.Black
                    prj_code.Font.Size = 10
                End If
                uiFun.displayMsg(Me, "1001", "", Session("gLang"))
                REM **********************
            Catch ex As Exception
                transaction.Rollback()
                uiFun.displayMsg(Me, "1002", "", Session("gLang"))
            End Try
        End If
    End Sub

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        prj_status.Text = "CANCELLED"
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

End Class
