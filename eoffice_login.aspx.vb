Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class eoffice_login
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gUI As New UIfunc
    Private rmtPost As New RemotePost
    Private appCon As New AppConfig

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        REM*******************************************************************
        REM Applying web.config<appSettings> value into Cache,
        REM The cache entries for each browser would be shared by all pages.

        Dim app_config As New AppConfig()
        Dim cmdPa As GlobalDBFunc.DBCmdPara

        app_config.SetCache()
        app_config.Dispose()

        REM*******************************************************************

        If Not IsPostBack Then
            If Membership.GetUser() IsNot Nothing AndAlso Session("usr_id") IsNot Nothing Then
                ClientScript.RegisterStartupScript(Me.GetType, "accLock", "alert('You are already logged in.');", True)

                Login1.Enabled = False
                Exit Sub

            Else
                Session.Remove("gLang")
                Session.Remove("usr_id")

                appCon.RemoveTempSession("SEARCH_SESSION_PAGE_")
            End If

            If Request("ad") IsNot Nothing AndAlso Request("ad") = "true" Then
                'Server.Transfer("ad_check.aspx?nowarn=true")
                Response.Redirect("ad_check.aspx?nowarn=true", False)
            End If
        End If
    End Sub

    Protected Sub Login1_Authenticate(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.AuthenticateEventArgs) Handles Login1.Authenticate
        Dim cmdPa As GlobalDBFunc.DBCmdPara
        Dim SQLString As String = ""
        Dim dt As New DataTable

        Dim isValid As Boolean = True

        If Membership.GetUser(Login1.UserName) Is Nothing Then
            cmdPa = New GlobalDBFunc.DBCmdPara

            SQLString = "SELECT usr_id, usr_pwd, usr_type, usr_pref_storer " &
                        "from " & Cache("DB_USER") & " " &
                        "where usr_id = " & cmdPa.AP(Login1.UserName) & " " &
                        "and usr_login_yn = 'Y' "

            dt = gDB.getDataTable(SQLString, , , , cmdPa)

            If dt.Rows.Count > 0 Then

                If Login1.Password = decryptPwd(Login1.UserName, dt.Rows(0).Item("usr_pwd").ToString) Then
                    Membership.CreateUser(Login1.UserName, Login1.Password)
                Else
                    isValid = False
                End If
            Else
                isValid = False
            End If
        End If

        If isValid AndAlso Membership.ValidateUser(Login1.UserName, Login1.Password) Then


            cmdPa = New GlobalDBFunc.DBCmdPara

            SQLString = "SELECT usr_id, usr_pwd,usr_type, usr_pref_storer, USR_FRC_CHG_PWD_YN from " & Cache("DB_USER") &
                              " where usr_id = " & cmdPa.AP(Login1.UserName) '& _
            '" and usr_login_yn = 'Y' "
            Session("l_conn") = SelConn.SelectedItem
            dt = gDB.getDataTable(SQLString, , , , cmdPa)

            If dt.Rows.Count > 0 Then
                'If dt.Rows(0).Item("USR_FRC_CHG_PWD_YN").ToString = "Y" Then
                '    rmtPost.Url = "ForcechangePW.aspx"
                '    rmtPost.Add("cpmode", "Y")
                '    rmtPost.Add("cpid", Login1.UserName)
                '    rmtPost.Post()

                'Else
                FormsAuthentication.SetAuthCookie(Login1.UserName, False)
                AccessRightUtils.storeLoginInfo(Login1.UserName, dt.Rows.Item(0)("USR_TYPE").ToString, dt.Rows.Item(0)("usr_pref_storer").ToString, SelLang.SelectedValue)
                'FormsAuthentication.RedirectFromLoginPage(Login1.UserName, False)
                Response.Redirect("~/main.aspx", False)
                'End If

            End If
        End If

    End Sub

    Private Function decryptPwd(ByVal userID As String, ByVal dbPwd As String) As String
        Dim Key As String = "OTS_" & UCase(userID)
        Dim p As Encryption.Symmetric.Provider = Encryption.Symmetric.Provider.Rijndael
        Dim sym As New Encryption.Symmetric(p)
        Dim resultData As Encryption.Data

        sym.Key.Text = Key

        Dim decryptedData As New Encryption.Data

        decryptedData.Base64 = dbPwd

        resultData = sym.Decrypt(decryptedData)

        Return resultData.ToString
    End Function
End Class
