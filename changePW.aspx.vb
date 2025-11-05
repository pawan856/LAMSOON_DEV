Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data

Partial Class changePW
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As New AccessRightUtils

    Private dt As New DataTable

    Private menu_code As String

    Protected Sub Submit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Submit.Click
        Dim updateSql As String
        Dim cmdPa As GlobalDBFunc.DBCmdPara

        ar.checkLoginStatus()

        If Newpw.Text <> Confirmpw.Text Then
            uiFun.displayMsg(Me, "", "Confirm Password not match.", Session("gLang"))
            'ElseIf Newpw.Text.Trim.Length < 8 Then
            '    uiFun.displayMsg(Me, "", "Invalid New Password (at least 8 characters).", Session("gLang"))
            'ElseIf Not gU.isValidPWD(Newpw.Text.Trim) Then
            '    uiFun.displayMsg(Me, "", "Invalid New Password (3 different combinations of Upper Case Char, Lower Case Char, Digit and Symbols).", Session("gLang"))
        Else
            Dim seq_string As String = "SELECT usr_pwd FROM wms_user WHERE usr_id = '" & _
                           Session("usr_id") & "'"

            Dim nS_dt As New DataTable
            nS_dt = gDB.getDataTable(seq_string)

            If nS_dt.Rows.Count > 0 Then
                If Membership.ValidateUser(Session("usr_id"), Oldpw.Text) Then
                    If Newpw.Text.Trim <> Oldpw.Text.Trim Then
                        Dim user As MembershipUser = Membership.GetUser(Session("usr_id").ToString.Trim.ToUpper)
                        user.ChangePassword(user.ResetPassword(), Newpw.Text)
                        Membership.UpdateUser(user)

                        Dim Key As String = "OTS_" & UCase(Session("usr_id"))
                        Dim p As Encryption.Symmetric.Provider = Encryption.Symmetric.Provider.Rijndael
                        Dim sym As New Encryption.Symmetric(p)
                        sym.Key.Text = Key

                        Dim encryptedData As Encryption.Data
                        encryptedData = sym.Encrypt(New Encryption.Data(Newpw.Text.Trim))

                        cmdPa = New GlobalDBFunc.DBCmdPara

                        updateSql = "update wms_user set " & _
                                        "usr_pwd = " & cmdPa.AP(gU.decodeNull(encryptedData.Base64, "")) & ", " & _
                                        "usr_pwd_lud = Getdate() " & _
                                    "where usr_id = " & cmdPa.AP(Session("usr_id")) & " "

                        gDB.amendData(updateSql, , , cmdPa)

                        uiFun.displayMsg(Me, "", "Password has been changed.", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "New Password is same as Old Password.", Session("gLang"))
                    End If
                Else
                    uiFun.displayMsg(Me, "", "Incorrect Password.", Session("gLang"))
                End If
            Else
                uiFun.displayMsg(Me, "", "Invalid user.", Session("gLang"))
            End If
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        menu_code = Request("menu_code")

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lblTitle.Text = "Change Password"
            lblOldpw.Text = "Old Password"
            lblNewpw.Text = "New Password"
            lblConfirmpw.Text = "Confirm Password"
            Submit.Text = "Save"
        Else
            lblTitle.Text = "Change Password"
            lblOldpw.Text = "Old Password"
            lblNewpw.Text = "New Password"
            lblConfirmpw.Text = "Confirm Password"
            Submit.Text = "Save"
        End If
        REM **********************
    End Sub
End Class
