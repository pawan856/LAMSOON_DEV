Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data

Partial Class menu
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils

    Private moduelCollection As New ArrayList
    Private prgmIdCollection As New ArrayList

    Private ar As New AccessRightUtils
    Private appCon As New AppConfig

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
                ar.Force_PageEndCtrlClear(Me, False)
                Me.Visible = False
                Exit Sub
            End If

            Dim SQLString As String = ""
            Dim dt As New DataTable

            appCon.RemoveTempSession("SEARCH_SESSION_PAGE_")

            If Session("gLang") = "E" Then
                moduelCollection.Add("Sitemap")
            Else
                moduelCollection.Add("網頁指南")
            End If

            prgmIdCollection.Add("sitemap.aspx")

            SQLString = "SELECT fun_code, fun_eng_name, fun_chi_name, fun_display_seq, fun_web_page, fun_parent_code from wms_function "
            If Session("usr_type") = "S" Then
                SQLString = SQLString & "where fun_type = 'MENU' "
            ElseIf Session("usr_type") = "C" Then
                SQLString = SQLString & "where fun_type = 'MENU_CUST' "
            ElseIf Session("usr_type") = "T" Then
                SQLString = SQLString & "where fun_type = 'MENU_STORER' "
            ElseIf Session("usr_type") = "V" Then
                SQLString = SQLString & "where fun_type = 'MENU_VENDOR' "
            Else
                SQLString = SQLString & "where 1 = 0 "
            End If

            SQLString = SQLString & " order by fun_display_seq"

            dt = gDB.getDataTable(SQLString)

            For i As Integer = 0 To dt.Rows.Count - 1
                If Session("gLang") = "E" Then
                    moduelCollection.Add(dt.Rows(i).Item("fun_eng_name").ToString)
                Else
                    moduelCollection.Add(dt.Rows(i).Item("fun_chi_name").ToString)
                End If
                If dt.Rows(i).Item("fun_web_page").ToString <> "" Then
                    prgmIdCollection.Add(dt.Rows(i).Item("fun_web_page").ToString)
                Else
                    prgmIdCollection.Add("user_group_fn.aspx?menu_code=" & dt.Rows(i).Item("fun_code").ToString)
                End If

            Next

            Dim iL As Integer = 1

            Dim htmlString As String = "<table cellspacing=1 cellpadding=1 width=100% align=center bgcolor=#cccccc border=0><tbody>"
            Dim htmlEndTag As String = ""
            Dim prgmLink As String = ""

            'If moduelCollection.Count > 0 Then
            '    htmlString = htmlString & "<tr bgcolor=""#FFFFFF"">"
            '    For x As Integer = 0 To moduelCollection.Count - 1
            '        prgmLink = ""
            '        htmlEndTag = ""
            '        htmlString = htmlString & "<td onMouseOver=""this.bgColor='#eeeeee'"" onMouseOut=""this.bgColor=''"" width=""9%"" class=""mainMenuBarTD"" nowrap>"
            '        htmlString = htmlString & "<div align=""center""><a href=""" & prgmLink & prgmIdCollection.Item(x).ToString & """ class=""link"" target=""main"">"
            '        htmlString = htmlString & moduelCollection.Item(x).ToString & "</a></div>"
            '        htmlString = htmlString & "</td>"
            '        htmlString = htmlString & htmlEndTag
            '    Next
            '    htmlString = htmlString & "<td onMouseOver=""this.bgColor='#eeeeee'"" onMouseOut=""this.bgColor=''"" width=""10%"" class=""mainMenuBarTD"" nowrap>"
            '    htmlString = htmlString & "<div align=""center""><a href=""eoffice_login.aspx"" class=""link"" target=""_parent"">"
            '    If Session("gLang") = "E" Then
            '        htmlString = htmlString & "Logout</a></div>"
            '    Else
            '        htmlString = htmlString & "登出</a></div>"
            '    End If


            '    htmlString = htmlString & "</td>"
            '    htmlString = htmlString & "</tr>"
            'End If

            'htmlString = htmlString & "</tbody></table>"

            'menuform.InnerHtml = htmlString


            Dim menuDT As DataTable = gDB.getDataTable(SQLString)

            rept.DataSource = menuDT
            rept.DataBind()


            If Session("usr_id") IsNot Nothing Then
                If Session("l_conn") IsNot Nothing AndAlso Session("l_conn").ToString = "Production" Then
                    imgDB.ImageUrl = "~/images/version/Production.png"
                ElseIf Session("l_conn") IsNot Nothing AndAlso Session("l_conn").ToString = "Archived" Then
                    imgDB.ImageUrl = "~/images/version/Archived.png"
                End If
                logined_user.Text = Session("usr_id") & ", welcome!"

                If System.Configuration.ConfigurationManager.AppSettings.Item("OTS_SYSTEM_TYPE") <> "" Then
                    logined_user.Text &= " (" & System.Configuration.ConfigurationManager.AppSettings.Item("OTS_SYSTEM_TYPE") & ")"
                End If
            End If
        End If
    End Sub

    Private Sub SignOut()        
        FormsAuthentication.SignOut()


        'Dim updateSql As String
        'updateSql = "UPDATE LM_SECUR_LOG SET secl_logout=Getdate() WHERE secl_sessionid = '" & Session("secl_sessionid") & "'"
        'gDB.amendData(updateSql)

        Session.Abandon()

        'FormsAuthentication.RedirectToLoginPage()
        'Response.Redirect("~/eoffice_login.aspx", False)

        HttpContext.Current.Response.Write("<scr" & _
                                                "ipt>top.location.href='./';</scr" & "ipt>")
    End Sub

    Protected Sub rept_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rept.ItemDataBound
        Dim hl As HyperLink = TryCast(e.Item.FindControl("menuLink"), HyperLink)

        If hl IsNot Nothing Then
            If DataBinder.Eval(e.Item.DataItem, "fun_web_page").ToString.Trim = "" AndAlso _
                DataBinder.Eval(e.Item.DataItem, "fun_parent_code").ToString.Trim = "" Then

                hl.NavigateUrl = "~/user_group_fn.aspx?User_ID=" & Session("Userid") & _
                                                        "&menu_code=" & DataBinder.Eval(e.Item.DataItem, "fun_code").ToString

                'If DataBinder.Eval(e.Item.DataItem, "mu_rpt_filter_yn").ToString.ToUpper.Trim = "Y" Then
                '    hl.NavigateUrl += "&rf=Y"
                'End If


            ElseIf DataBinder.Eval(e.Item.DataItem, "fun_web_page").ToString.Trim <> "" Then
                Dim tmpLink As String = ""

                tmpLink = DataBinder.Eval(e.Item.DataItem, "fun_web_page").ToString.Trim

                If tmpLink = "" Then
                    tmpLink = "#"
                End If

                If tmpLink.Contains("?") Then
                    tmpLink += "&menu_code=" & gU.jsURLEncode(DataBinder.Eval(e.Item.DataItem, "fun_code").ToString.Trim)
                Else
                    tmpLink += "?menu_code=" & gU.jsURLEncode(DataBinder.Eval(e.Item.DataItem, "fun_code").ToString.Trim)
                End If

                hl.NavigateUrl = tmpLink

            ElseIf DataBinder.Eval(e.Item.DataItem, "fun_web_page").ToString.Trim = "" AndAlso _
                DataBinder.Eval(e.Item.DataItem, "fun_parent_code").ToString.Trim <> "" Then

                Dim accessRightSQL As String = ""

                accessRightSQL = "SELECT fun_code, fun_eng_name, fun_chi_name, fun_display_seq, fun_web_page "
                accessRightSQL += "from wms_function where fun_parent_code = '" & DataBinder.Eval(e.Item.DataItem, "fun_parent_code").ToString.Trim & "' "

                accessRightSQL += " and exists (select 1 from wms_user_group_alloc uga, wms_sec_access sa "
                accessRightSQL += "where sa.sec_access = 'Y' and sa.grp_code = uga.grp_code and sa.fun_code = wms_function.fun_code and uga.usr_id = '" & Session("usr_id") & "')"
                accessRightSQL += " order by fun_display_seq"

                Dim acceTable As New DataTable
                acceTable = gDB.getDataTable(accessRightSQL)

                If acceTable.Rows.Count > 0 Then
                    Dim tmpLink As String = ""

                    tmpLink = acceTable.Rows(0).Item("fun_web_page").ToString

                    If tmpLink = "" Then
                        tmpLink = "cms_search.aspx"
                    End If

                    If tmpLink.Contains("?") Then
                        tmpLink += "&menu_code=" & gU.jsURLEncode(DataBinder.Eval(e.Item.DataItem, "mu_id").ToString.Trim)
                    Else
                        tmpLink += "?menu_code=" & gU.jsURLEncode(DataBinder.Eval(e.Item.DataItem, "mu_id").ToString.Trim)
                    End If


                    Dim phyPath As String = ""
                    Dim attCode As String() = tmpLink.Split("&")

                    If Not tmpLink.Substring(0, tmpLink.IndexOf("?")).Contains("http") Then
                        If tmpLink.Substring(0, 1) = "/" Then
                            phyPath = Server.MapPath("~" & tmpLink.Substring(0, tmpLink.IndexOf("?")))
                        Else
                            phyPath = Server.MapPath("~/" & tmpLink.Substring(0, tmpLink.IndexOf("?")))
                        End If
                    End If

                    If phyPath <> "" Then
                        If Not IO.File.Exists(phyPath) Then
                            If Session("userid") = "OTSADMIN" Then
                                hl.NavigateUrl = "javascript:alert('The following page cannot be found:\n\r" & tmpLink & "');"
                            Else
                                hl.NavigateUrl = "javascript:alert('The page cannot be found.');"
                            End If
                        Else
                            hl.NavigateUrl = tmpLink
                        End If
                    Else
                        hl.NavigateUrl = "javascript:alert('The page cannot be found.');"
                    End If
                Else
                    hl.NavigateUrl = "#"
                End If
            End If

            hl.Target = "main"

            If Session("gLang") = "E" Then                
                hl.Text = DataBinder.Eval(e.Item.DataItem, "fun_eng_name").ToString
            Else
                hl.Text = DataBinder.Eval(e.Item.DataItem, "fun_chi_name").ToString
            End If
        End If
    End Sub

    Protected Sub btnLogout_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnLogout.Click
        SignOut()
    End Sub
End Class
