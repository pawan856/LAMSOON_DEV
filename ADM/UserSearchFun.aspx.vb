Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data

Partial Class UserSearchFun
    Inherits System.Web.UI.Page
    Private uiFun As New UIfunc
    Private gDB As New GlobalDBFunc
    Private cU As New CommonUtils
    Private gU As New GeneralUtils
    Private usr_type As String = ""

    Protected Sub Submit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Submit.Click
        Call BindGV()
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        usr_type = Request("user_type")

        user_type.Value = usr_type

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            If usr_type = "S" Then
                lblTitle.Text = "Staff User Maintenance"
            ElseIf usr_type = "T" Then
                lblTitle.Text = "Storer User Maintenance"
            ElseIf usr_type = "C" Then
                lblTitle.Text = "Customer User Maintenance"
            ElseIf usr_type = "V" Then
                lblTitle.Text = "Vendor User Maintenance"
            Else
                lblTitle.Text = "User Maintenance"
            End If
            lblCode.Text = "User ID"
            lblName.Text = "User Name"
            Submit.Text = "Search"
            NewBtn.Text = "New"
            gvrsList.EmptyDataText = "No Record Found."
        Else
            If usr_type = "S" Then
                lblTitle.Text = "系統使用者維護"
            ElseIf usr_type = "T" Then
                lblTitle.Text = "貨主使用者維護"
            ElseIf usr_type = "C" Then
                lblTitle.Text = "客戶使用者維護"
            ElseIf usr_type = "V" Then
                lblTitle.Text = "供應商使用者維護"
            Else
                lblTitle.Text = "系統使用者資料"
            End If
            lblCode.Text = "使用者代碼"
            lblName.Text = "使用者姓名"
            Submit.Text = "插尋"
            NewBtn.Text = "新增"
            gvrsList.EmptyDataText = "找不到相關資料"
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
                Call cU.changeGVLabel(oGridViewRow, e, "Status", "狀態")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        End Select
    End Sub

    Protected Sub gvrsList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvrsList.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

            CType(e.Row.Cells(0).Controls(0), HyperLink).NavigateUrl = "UserMaster.aspx?usr_id=" & CType(e.Row.Cells(0).Controls(0), HyperLink).Text & _
                                                                        "&user_type=" & user_type.Value

            If Session("gLang") = "E" Then
                nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                nButton.Text = "Delete"
            ElseIf Session("gLang") = "C" Then
                '
                nButton.Attributes.Add("onclick", "javascript:return confirm('你是否确定要刪除這個資料?')")
                nButton.Text = "刪除"
            End If
        End If
    End Sub

    Protected Sub gvrsList_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gvrsList.RowDeleting
        REM **********************
        REM Modify Here
        Dim userID As String = gvrsList.DataKeys(e.RowIndex).Value
        Dim delete_sql As String = "delete from wms_user where usr_id = '" & userID & "' "
        REM **********************

        Dim alertstr As String = ""

        gDB.amendData(delete_sql)
        uiFun.displayMsg(Me, "1005", "", Session("gLang"))

        Call BindGV()
    End Sub


    Protected Sub NewBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NewBtn.Click
        REM **********************
        REM Modify Here
        Response.Redirect("UserMaster.aspx?user_type=" & usr_type & "&mode=N")
        REM **********************
        Response.End()
    End Sub

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = ""

        REM **********************
        REM Modify Here
        SQLString = "SELECT usr_id , usr_fname + ' ' + usr_sname as usr_name, wms_col_code.colc_eng_value as usr_status " & _
                    "from wms_user, wms_col_code " & _
                    "where 'WMS_USER.USR_STATUS' = wms_col_code.colc_tabcol " & _
                    "and wms_user.usr_status *= wms_col_code.colc_code " & _
                    "and wms_user.usr_type = '" & gU.dbEncode(usr_type) & "' "

        If usr_id.Text <> "" Then
            SCString = SCString & " and wms_user.usr_id LIKE N'%" & usr_id.Text & "%' "
        End If

        If usr_name.Text <> "" Then
            SCString = SCString & " and (UPPER(wms_user.usr_fname) LIKE N'%" & UCase(Usr_name.Text) & "%' OR "
            SCString = SCString & " UPPER(wms_user.usr_sname) LIKE N'%" & UCase(Usr_name.Text) & "%' OR "
            SCString = SCString & " UPPER(wms_user.usr_nickname) LIKE N'%" & UCase(Usr_name.Text) & "%')"
        End If
        REM **********************

        If SCString <> "" Then
            SQLString = SQLString & " " & SCString
        End If
        dt = gDB.getDataTable(SQLString)

        gvrsList.DataSource = dt
        gvrsList.DataBind()
    End Sub
End Class
