Imports Microsoft.VisualBasic
Imports System.Web

Public Class RemotePost
    Private gU As New GeneralUtils

    Private Inputs As System.Collections.Specialized.NameValueCollection = New System.Collections.Specialized.NameValueCollection

    Public Url As String = ""
    Public Method As String = "post"
    Public FormName As String = "form1"
    Public Target As String = "_self"
    Public alertMsg As String = ""

    Public Sub Add(ByVal name As String, ByVal value As String)
        Inputs.Add(name, value)
    End Sub

    Public Sub Post()
        System.Web.HttpContext.Current.Response.Clear()
        System.Web.HttpContext.Current.Response.Write("<html><head>")

        If alertMsg <> "" Then
            System.Web.HttpContext.Current.Response.Write(String.Format("</head><body onload=""alert('" & gU.jsString(alertMsg) & "');document.{0}.submit();"">", FormName))
        Else
            System.Web.HttpContext.Current.Response.Write(String.Format("</head><body onload=""document.{0}.submit();"">", FormName))
        End If

        System.Web.HttpContext.Current.Response.Write(String.Format("<form name=""{0}"" method=""{1}"" action=""{2}"" target=""{3}"" >", FormName, Method, Url, Target))
        Dim i As Integer = 0
        Do While i < Inputs.Keys.Count
            System.Web.HttpContext.Current.Response.Write(String.Format("<input name=""{0}"" type=""hidden"" value=""{1}"">", System.Web.HttpContext.Current.Server.HtmlEncode(Inputs.Keys(i)), System.Web.HttpContext.Current.Server.HtmlEncode(Inputs(Inputs.Keys(i)))))
            i += 1
        Loop
        System.Web.HttpContext.Current.Response.Write("</form>")
        System.Web.HttpContext.Current.Response.Write("</body></html>")
        System.Web.HttpContext.Current.Response.End()
    End Sub
End Class
