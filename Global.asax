<%@ Application Language="VB" %>

<script runat="server">

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)        
        REM*******************************************************************
        REM Applying web.config<appSettings> value into Cache,
        REM The cache entries for each browser would be shared by all pages.

        Dim app_config As New AppConfig()
        app_config.SetCache()
        app_config.Dispose()
        REM*******************************************************************
    End Sub
    
    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application shutdown
    End Sub
        
    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when an unhandled error occurs
        
        Dim ex As Exception = Server.GetLastError
        
        If ex.GetBaseException.GetType.ToString = "System.Web.HttpRequestValidationException" Then
            '			Server.ClearError()
            '			Response.Redirect(Context.Request.Url.ToString)
		
            '            Response.Clear()
            '            Response.StatusCode = 200
            '            Response.Write("<html><head>" & _
            '                           "<title>HTTP Validation Fail</title>" & _
            '                           "<scr" & "ipt type=""text/javascript"" language=""javascript"">" & _
            '                           "function back() {history.go(-1); }" & _
            '                           "</scr" & "ipt>" & _
            '						   "</head>" & _
            '                           "<body>" & _
            '						   "<p>Invalid char detected in the input screen! Please remove any angle brackets like &lt; or &gt;.</p>" & _
            '                           "<p><a href=""javascript:back();"">Back</a></p>" & _
            '                           "</body>" & _
            '                           "</html>")
            '           Response.End
        End If
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        HttpContext.Current.Session.Timeout = 120
        
        REM*******************************************************************
        REM Applying web.config<appSettings> value into Cache,
        REM The cache entries for each browser would be shared by all pages.

        Dim app_config As New AppConfig()
        app_config.SetCache()
        app_config.Dispose()
        REM*******************************************************************
        
        'Dim request_cookies As String = Request.Headers("Cookie")
        
        'If (request_cookies IsNot Nothing) AndAlso (request_cookies.IndexOf("ASP.NET_SessionId") >= 0) Then            
        '    Session("SessionTimeOut") = True
        'End If
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a session ends. 
        ' Note: The Session_End event is raised only when the sessionstate mode
        ' is set to InProc in the Web.config file. If session mode is set to StateServer 
        ' or SQLServer, the event is not raised.
    End Sub
       
    Protected Sub Application_EndRequest(sender As Object, e As System.EventArgs)
        If HttpContext.Current.Response.StatusCode = 401 Then
            HttpContext.Current.Response.ClearContent()
            HttpContext.Current.Response.Write("<scr" & _
                                                "ipt>window.top.location='./eoffice_login.aspx';</scr" & "ipt>")
        End If
    End Sub

    Protected Sub Application_AcquireRequestState(sender As Object, e As System.EventArgs)
        If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session.IsNewSession Then
            Dim request_cookies As String = Request.Headers("Cookie")
            
            If request_cookies IsNot Nothing AndAlso request_cookies.IndexOf("ASP.NET_SessionId") >= 0 Then
                If User.Identity.IsAuthenticated Then
                    FormsAuthentication.SignOut()
                    HttpContext.Current.Response.ClearContent()
                    HttpContext.Current.Response.Write("<scr" & _
                                    "ipt>alert('Your session has expired. Maximum idle time is " & HttpContext.Current.Session.Timeout & " minutes. Please login again.');" & _
                                    "top.location.href='./eoffice_login.aspx';</scr" & "ipt>")
                    HttpContext.Current.Response.End()
                End If
            End If
        End If        
    End Sub
    
    'Protected Sub Application_PostRequestHandlerExecute(sender As Object, e As System.EventArgs)
    '    If TypeOf Context.Handler Is IRequiresSessionState OrElse TypeOf Context.Handler Is IReadOnlySessionState Then
    '        Dim authenticationCookie As HttpCookie = Request.Cookies(FormsAuthentication.FormsCookieName)
            
    '        If authenticationCookie IsNot Nothing Then
    '            Dim authenticationTicket As FormsAuthenticationTicket = FormsAuthentication.Decrypt(authenticationCookie.Value)
                
    '            If Not authenticationTicket.Expired AndAlso (Session("usr_id") Is Nothing OrElse Session("usr_id") = "") Then
    '                FormsAuthentication.SignOut()
    '                HttpContext.Current.Response.ClearContent()
    '                HttpContext.Current.Response.Write("<scr" & _
    '                                "ipt>alert('Your session has expired. Maximum idle time is " & HttpContext.Current.Session.Timeout & " minutes. Please login again.');" & _
    '                                "top.location.href='./otsLogin.aspx';</scr" & "ipt>")
    '            End If
    '        End If
    '    End If
    'End Sub
</script>