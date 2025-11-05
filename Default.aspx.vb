Imports System.Web
Imports System.Net

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        REM*******************************************************************
        REM Applying web.config<appSettings> value into Cache,
        REM The cache entries for each browser would be shared by all pages.

        Dim app_config As New AppConfig()
        app_config.SetCache()
        app_config.Dispose()
        REM*******************************************************************

        If Membership.GetUser() IsNot Nothing AndAlso _
              (Session("userid") IsNot Nothing OrElse Not Session.IsNewSession) Then

            'Server.Transfer("otsLogin.aspx")
            Response.Redirect("eoffice_Login.aspx", False)
        End If

        If Not IsPostBack Then
            Dim checkwithAD As Boolean = False
            'Dim ipRange As String = System.Configuration.ConfigurationManager.AppSettings.Item("INTRANET_IP_RANGE")
            'Dim expList As New ArrayList

            'expList.AddRange(Split(System.Configuration.ConfigurationManager.AppSettings.Item("AD_IP_EXCEPTION"), ","))

            'Dim clientIP As String = GetClientIP()

            'If ipRange <> "" Then
            '    If Not expList.Contains(clientIP) Then
            '        Dim theIPRange As String = ""

            '        If ipRange.Contains("*") Then
            '            theIPRange = ipRange.Substring(0, ipRange.IndexOf("*"))

            '            If theIPRange.Substring(theIPRange.Length - 1, 1) = "." Then
            '                theIPRange = theIPRange.Substring(0, theIPRange.Length - 1)
            '            End If
            '        Else
            '            theIPRange = ipRange
            '        End If

            '        If clientIP = Dns.GetHostEntry("localhost").AddressList(0).ToString OrElse _
            '            clientIP = Dns.GetHostEntry("localhost").AddressList(1).ToString OrElse _
            '            clientIP = Dns.GetHostEntry(Dns.GetHostName).AddressList(1).ToString OrElse _
            '            clientIP.StartsWith(theIPRange) Then

            '            checkwithAD = True
            '        End If
            '    End If
            'End If

            If checkwithAD Then
                'Server.Transfer("otsLogin.aspx?ad=true")
                Response.Redirect("eoffice_Login.aspx?ad=true", False)
            Else
                'Server.Transfer("otsLogin.aspx")
                Response.Redirect("eoffice_Login.aspx", False)
            End If
        End If
    End Sub

    Private Function GetClientIP() As String
        Dim strIPAddr As String

        If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OrElse _
                Request.ServerVariables("HTTP_X_FORWARDED_FOR").ToString.IndexOf("unknown") > 0 Then

            strIPAddr = Request.ServerVariables("REMOTE_ADDR")

        ElseIf Request.ServerVariables("HTTP_X_FORWARDED_FOR").ToString.IndexOf(",") > 0 Then
            strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") - 1)
        ElseIf Request.ServerVariables("HTTP_X_FORWARDED_FOR").ToString.IndexOf(";") > 0 Then
            strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") - 1)
        Else
            strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        End If
        Return Mid(strIPAddr, 1, 30).Trim
    End Function
End Class
