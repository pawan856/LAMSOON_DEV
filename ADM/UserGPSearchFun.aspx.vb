Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data

Partial Class UserGPSearchFun
    Inherits System.Web.UI.Page
    Private uiFun As New UIfunc
    Private gDB As New GlobalDBFunc
    Private cU As New CommonUtils

    Protected Sub Submit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Submit.Click
        Call BindGV()
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lblTitle.Text = "User Group Maintenance"
            lblCode.Text = "Group Code"
            lblName.Text = "Group Name"
            Submit.Text = "Search"
            NewBtn.text = "New"
            gvrsList.EmptyDataText = "No Record Found."
        Else
            lblTitle.Text = "用户集团"
            lblCode.Text = "集团代码"
            lblName.Text = "集团名称"
            Submit.Text = "搜寻"
            NewBtn.Text = "新增"
            gvrsList.EmptyDataText = "找不到相关资料"
        End If
        REM **********************
    End Sub

    Protected Sub gvrsList_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvrsList.PageIndexChanging
        gvrsList.PageIndex = e.NewPageIndex
        Call BindGV()
    End Sub

    Protected Sub gvrsList_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvrsList.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here

                Call cU.changeGVLabel(oGridViewRow, e, lblCode.Text, lblCode.Text)
                Call cU.changeGVLabel(oGridViewRow, e, lblName.Text, lblName.Text)
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        End Select
    End Sub

    Protected Sub gvrsList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvrsList.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

            If Session("gLang") = "E" Then
                nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                nButton.Text = "Delete"
            ElseIf Session("gLang") = "C" Then
                '
                nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                nButton.Text = "删除"
            End If
        End If
    End Sub

    Protected Sub gvrsList_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gvrsList.RowDeleting
        REM **********************
        REM Modify Here
        Dim grp_code As String = gvrsList.DataKeys(e.RowIndex).Value
        Dim delete_sql As String = "delete from wms_user_group where grp_code = '" & grp_code & "' "
        REM **********************

        Dim alertstr As String = ""

        gDB.amendData(delete_sql)
        uiFun.displayMsg(Me, "1005", "", Session("gLang"))

        Call BindGV()
    End Sub


    Protected Sub NewBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NewBtn.Click
        REM **********************
        REM Modify Here
        Response.Redirect("UserGPMain.aspx?mode=N")
        REM **********************
        Response.End()
    End Sub

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"

        REM **********************
        REM Modify Here
        SQLString = "SELECT grp_code, grp_name from wms_user_group"

        If grp_code.Text <> "" Then
            If SCString <> "WHERE" Then
                SCString = SCString & " AND "
            End If
            SCString = SCString & " grp_code LIKE N'%" & grp_code.Text & "%' "
        End If

        If grp_name.Text <> "" Then
            If SCString <> "WHERE" Then
                SCString = SCString & " AND "
            End If
            SCString = SCString & " UPPER(grp_name) LIKE N'%" & UCase(grp_name.Text) & "%'"
        End If
        REM **********************

        If SCString <> "WHERE" Then
            SQLString = SQLString & " " & SCString
        End If
        dt = gDB.getDataTable(SQLString)

        gvrsList.DataSource = dt
        gvrsList.DataBind()
    End Sub
End Class
