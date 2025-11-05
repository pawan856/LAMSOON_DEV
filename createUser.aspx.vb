Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class createUser
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As AccessRightUtils
    Private usr_type As String = ""
    Protected fun_code As String = ""

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

    End Sub


    Protected Sub btnADD_Click(sender As Object, e As System.EventArgs) Handles btnADD.Click
        Dim gConn As SqlConnection

        gConn = gDB.getConnection()

        Dim transaction As SqlTransaction
        transaction = gConn.BeginTransaction()
        Try

            Dim SQLString As String = "insert into wms_user (usr_id, usr_status,usr_type) " & _
                                "values ('" & gU.dbEncode(UCase(user_id.Text.Trim)) & "','A','S')"

            gDB.amendData(SQLString, gConn, transaction)

            Dim newuser As MembershipUser = Membership.GetUser(UCase(user_id.Text.Trim))

            If newuser Is Nothing Then
                Membership.CreateUser(UCase(user_id.Text.Trim), pwd.Text.Trim, UCase(user_id.Text.Trim) & "@otsbilling.com")
            End If


            SQLString = "INSERT INTO WMS_USER_GROUP_ALLOC " & _
                        " (GRP_CODE, USR_ID) " & _
                        " VALUES('ADMIN_GP','" & gU.dbEncode(UCase(user_id.Text.Trim)) & "') "

            gDB.amendData(SQLString, gConn, transaction)

            transaction.Commit()
            uiFun.displayMsg(Me, "", "User Added", Session("gLang"))
        Catch ex As Exception
            transaction.Rollback()
            uiFun.displayMsg(Me, "", "Add user Failed", Session("gLang"))
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Sub
End Class
