
Partial Class CustomControls_Lookup_ListNo
    Inherits System.Web.UI.UserControl

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Session("SEARCH_SESSION_PAGE_CKD_CC_LIST_NO") <> "" Then
            txt_list_no.Text = Session("SEARCH_SESSION_PAGE_CKD_CC_LIST_NO")
        Else
            txt_list_no.Text = "1"
        End If
    End Sub
End Class
