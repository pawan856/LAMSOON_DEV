
Partial Class CustomControls_Lookup_ListNo
    Inherits System.Web.UI.UserControl

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Session("SEARCH_SESSION_PAGE_CKD_CHECKER") <> "" Then
            txt_checker.Text = Session("SEARCH_SESSION_PAGE_CKD_CHECKER")
            'Else
            '    txt_checker.Text = "1"
        End If
    End Sub
End Class
