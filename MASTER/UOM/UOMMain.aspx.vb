Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class MASTER_UOM_UOMMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As AccessRightUtils

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)
        ar.hideForm(Me)
        REM ****************************

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If

        If Not IsPostBack Then

            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            REM **********************
            REM Modify Here
            If Session("gLang") = "E" Then
                lheader.Text = "UOM  Maintenance"
                lbl_UOM_CODE.Text = "UOM Code:"
                lbl_UOM_DESC.Text = "UOM Description:"
                lbl_UOM_PCS.Text = "UOM PCS:"

                lbl_sys_cb.Text = "CB"
                lbl_sys_lub.Text = "LUB"
                lbl_sys_cd.Text = "CD"
                lbl_sys_lud.Text = "LUD"
                saveBtn1.Text = "Save"
                saveBtn2.Text = "Save"
                saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
                saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"

                If Session("pagemode") = "N" Then
                    'WH_CODE.Text = "[Code will be auto generated]"
                End If

            ElseIf Session("gLang") = "C" Then

                lheader.Text = "貨料單位主資料庫"
                lbl_UOM_CODE.Text = "貨料單位代碼:"
                lbl_UOM_DESC.Text = "貨料單位描述:"
                lbl_UOM_PCS.Text = "貨料單位件數:"

                lbl_sys_cb.Text = "創建者"
                lbl_sys_lub.Text = "最後更新者"
                lbl_sys_cd.Text = "創建日期"
                lbl_sys_lud.Text = "最後更新日期"
                saveBtn1.Text = "儲存"
                saveBtn2.Text = "儲存"
                saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
                saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"

                If Session("pagemode") = "N" Then
                    'UOM_CODE.Text = "[代碼會自動產生]"
                End If
            End If

            REM **********************
            UOM_CODE.CssClass = "REQUIRED"

            If Session("pagemode") = "N" Then
                'UOM_CODE.ForeColor = Drawing.Color.Red
            Else
                Call BindGV()
            End If

            If Session("pagemode") <> "N" And UOM_CODE.Text <> "" Then
                UOM_CODE.BorderWidth = 0
                UOM_CODE.BackColor = Drawing.Color.Transparent
                UOM_CODE.ReadOnly = True
            End If

            REM **********************
        End If
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""

        If UOM_CODE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_UOM_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_UOM_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim gConn As SqlConnection
        Dim nextNo As String = ""
        Dim dupSQL As String
        Dim dupTbl As New DataTable

        If validateAll() Then
            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction

            Try
                REM **********************
                REM Modify Here
                If Session("pagemode") = "N" Then

                    dupSQL = "select 1 from wms_uom " & _
                                "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and uom_code = '" & gU.dbEncode(UOM_CODE.Text) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then
                        'nextNo = DB.getDocNo("UOM", gConn, transaction)
                        nextNo = UOM_CODE.Text

                        sql_string = "insert into wms_uom (" & _
                        "imp_code, uom_code, uom_desc, uom_pcs," & _
                        "sys_cb, sys_cd, sys_lub, sys_lud) values ( " & _
                        gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & _
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(UOM_DESC.Text)) & "," & _
                         gU.dbEncode(gU.decodeNullOrEmpty(UOM_PCS.Text, "0")) & "," & _
                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in UOM!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "貨料單位資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_uom set " & _
                                    "uom_desc = " & gU.convdbNVCData(gU.dbEncode(UOM_DESC.Text)) & ", " & _
                                    "uom_pcs = " & gU.dbEncode(gU.decodeNullOrEmpty(UOM_PCS.Text, "0")) & ", " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                    "and uom_code = '" & gU.dbEncode(UOM_CODE.Text) & "' "
                End If
                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    UOM_CODE.Text = nextNo
                    UOM_CODE.ForeColor = Drawing.Color.Black
                    UOM_CODE.Font.Size = 10
                    'uom_code.ReadOnly = True
                    'uom_code.BackColor = Drawing.Color.Transparent
                    'uom_code.BorderWidth = 0
                End If

                uiFun.displayMsg(Me, "1001", "", Session("gLang"))

                Call BindGV()
                REM **********************
            Catch ex As Exception
                transaction.Rollback()
                uiFun.displayMsg(Me, "1002", "", Session("gLang"))
            Finally
                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try
        End If
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"
        Dim WhereStr As String = ""
        'Dim pk_code As String = ""
        Dim pk_code, storerCode As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("uom_code") <> "" Then
            pk_code = ViewState("uom_code")
        Else
            pk_code = Server.UrlDecode(Request("uom_code"))

            ViewState("uom_code") = pk_code
        End If
        REM **********************


        If Session("pagemode") <> "N" Then
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header

            SQLString = "SELECT * FROM WMS_UOM WHERE UOM_CODE = '" & gU.dbEncode(pk_code) & "' AND IMP_CODE = '" & Session("IMP_CODE") & "'"

            SQLString = SQLString & " " & WhereStr

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                UOM_CODE.Text = dt.Rows(0).Item("UOM_CODE").ToString
                UOM_DESC.Text = dt.Rows(0).Item("UOM_DESC").ToString
                UOM_PCS.Text = cU.FormatIntegerString(dt.Rows(0).Item("UOM_PCS").ToString)
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)
            End If
        End If
    End Sub
End Class
