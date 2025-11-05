Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data

Partial Class user_group_fn
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private ar As New AccessRightUtils
    Private appCon As New AppConfig

    Private menu_code As String
    Private moduelCollection As New ArrayList
    Private prgmIdCollection As New ArrayList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        menu_code = Request("menu_code")

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        Dim SQLString As String = ""
        Dim dt As New DataTable

        appCon.RemoveTempSession("SEARCH_SESSION_PAGE_")

        SQLString = "SELECT fun_code, fun_eng_name, fun_chi_name, fun_display_seq, fun_web_page from wms_function where fun_type = 'SCREEN' and fun_parent_code = '" & menu_code & "' "

        SQLString = SQLString & " and exists (select 1 from wms_user_group_alloc uga, wms_sec_access sa " & _
        "where sa.sec_access = 'Y' and sa.grp_code = uga.grp_code and sa.fun_code = wms_function.fun_code and uga.usr_id = '" & Session("usr_id") & "')"
        SQLString = SQLString & " order by fun_display_seq"

        dt = gDB.getDataTable(SQLString)
        For i As Integer = 0 To dt.Rows.Count - 1
            If Session("gLang") = "E" Then
                moduelCollection.Add(dt.Rows(i).Item("fun_eng_name").ToString)
            Else
                moduelCollection.Add(dt.Rows(i).Item("fun_chi_name").ToString)
            End If

            If dt.Rows(i).Item("fun_web_page").ToString <> "" Then
                If dt.Rows(i).Item("fun_web_page").ToString.Last = "=" Then
                    prgmIdCollection.Add(dt.Rows(i).Item("fun_web_page").ToString & dt.Rows(i).Item("fun_code").ToString)
                Else
                    prgmIdCollection.Add(dt.Rows(i).Item("fun_web_page").ToString)
                End If
            Else
                prgmIdCollection.Add("#" & dt.Rows(i).Item("fun_code").ToString)
            End If
        Next

        Dim iL As Integer = 1

        Dim htmlString As String = "<table border=""0"" cellspacing=""0"" cellpadding=""0""  width=""100%"" >" & _
                                    "<tr><td colspan=""2"">&nbsp;</td></tr><tr><td width=""10%"" class=""main_page_text"">&nbsp;" & _
                                    "</td><td width=""74%"" class=""main_page_text""><table width=""100%"" border=""0"" cellspacing=""0"" " & _
                                    "cellpadding=""5""></table></td></tr>"

        Dim htmlEndTag As String = ""
        Dim prgmLink As String = ""

        If moduelCollection.Count > 0 Then
            htmlString = htmlString & "<tr><td width=""10%"" class=""main_page_text"">&nbsp;</td><td width=""74%"" class=""main_page_text"">" & _
                        "<table cellspacing=1 cellpadding=5 border=0 bgcolor=""#CCCCCC"" align=""center"" width=100%>"

            For x As Integer = 0 To moduelCollection.Count - 1
                prgmLink = ""
                htmlEndTag = ""

                If iL = 1 Then
                    htmlString = htmlString & "<tr bgcolor=""#F0F0F0"">"
                End If

                If iL <= 3 Then
                    htmlString = htmlString & "<td onMouseOver=""this.bgColor='#B8B8B8'"" onMouseOut=""this.bgColor='#F0F0F0'"" class=""main_menu"" style=""width: 33%"">"
                    htmlString = htmlString & "<a href=""" & prgmLink & prgmIdCollection.Item(x).ToString & """ style="""" class=""main_menu"">"
                    htmlString = htmlString & moduelCollection.Item(x).ToString & "</a>"
                    htmlString = htmlString & "</td>"

                    iL += 1
                Else
                    htmlEndTag = "</tr>"
                    iL = 1

					If x <> 0 Then	x -= 1 Else x = 0
                End If

                htmlString = htmlString & htmlEndTag
            Next

            htmlString = htmlString & "</tr></table></td></tr>"

        End If

        htmlString = htmlString & "</table>"

        menuform.InnerHtml = htmlString

    End Sub
End Class
