Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports System.Web.UI.WebControls

Partial Class main
    Inherits System.Web.UI.Page

    Private dt As New GlobalDBFunc
    Private func As New DBfunc
    Private ui As New UIfunc

    Private ar As New AccessRightUtils

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack Then
                If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
                    ar.Force_PageEndCtrlClear(Me, False)
                    Me.Visible = False
                    Exit Sub
                End If

                Dim nt As New DataTable

                nt = dt.getDataTable("SELECT 1 AS VALUE FROM wms_user_group_alloc WHERE usr_id = '' AND grp_code IN ('', '')")

               If nt.Rows.Count > 0 Then
                    main.Attributes.Add("src", "main_news.aspx?ranDate=" & Server.UrlEncode(Now))
                Else
                    'main.Attributes.Add("src", "lin_ots/myAc   tionList/myActionList.asp?login=Y")
                    main.Attributes.Add("src", "main_news.aspx?ranDate=" & Server.UrlEncode(Now))
                End If

            End If
        Catch ex As Exception
            WriteLine("Error: " & ex.Message)
        End Try
    End Sub
End Class
