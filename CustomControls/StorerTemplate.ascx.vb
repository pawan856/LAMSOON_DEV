
Partial Class CustomControls_StorerTemplate
    Inherits System.Web.UI.UserControl

    Private uiFun As New UIfunc    

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Dim setReadonly As Boolean = False

            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))

            If Request("sc") IsNot Nothing Then
                STORER_CODE.SelectedValue = Request("sc")
            End If

            If Request("screadonly") IsNot Nothing Then
                setReadonly = Request("screadonly")
            End If

            If Session("usr_pref_storer") IsNot Nothing AndAlso Session("usr_pref_storer") <> "" Then
                If Request("ctemp") Is Nothing Then
                    STORER_CODE.SelectedValue = Session("usr_pref_storer")

                    Session("SEARCH_SESSION_PAGE_STORER_CODE") = STORER_CODE.SelectedValue

                    Page.Response.Redirect("~/cms_search.aspx?menu_code=" & Session("LR_SESSION_MENU_CODE") & "&sc=" & STORER_CODE.SelectedValue & "&ctemp=true")
                End If

                setReadonly = True
            End If

            If setReadonly Then
                If STORER_CODE.SelectedValue IsNot Nothing Then
                    lbl_storer_code.Text = STORER_CODE.SelectedItem.Text
                    hd_storer_code.Value = STORER_CODE.SelectedValue

                    lbl_storer_code.Visible = True
                    STORER_CODE.Visible = False
                End If
            End If
        End If
    End Sub

    Protected Sub STORER_CODE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles STORER_CODE.SelectedIndexChanged
        Session("SEARCH_SESSION_PAGE_STORER_CODE") = STORER_CODE.SelectedValue

        Page.Response.Redirect("~/cms_search.aspx?menu_code=" & Session("LR_SESSION_MENU_CODE") & "&sc=" & STORER_CODE.SelectedValue & "&ctemp=true")
    End Sub
End Class
