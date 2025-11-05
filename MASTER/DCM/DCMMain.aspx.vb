Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class MASTER_DCM_DCMMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As AccessRightUtils
    Private DDFORMAT As String = ""
    Private imp_code As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DDFORMAT = gU.getConfig("DDFORMATNO")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)
        ar.hideForm(Me)
        REM ****************************

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then

            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            REM **********************
            REM Modify Here
            If Session("gLang") = "E" Then
                lheader.Text = "Batch No. Maintenance"
                lbl_DC_DATE_CODE.Text = "Batch No.:"
                lbl_DC_CONV_DATE.Text = "Date of Batch:"
                lbl_STORER_CODE.Text = "Storer:"

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

                lheader.Text = "批次號碼主資料庫"
                lbl_DC_DATE_CODE.Text = "批次號碼:"
                lbl_DC_CONV_DATE.Text = "日期:"
                lbl_STORER_CODE.Text = "貨主代碼:"

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
            DC_DATE_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME  from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))

            If Session("pagemode") = "N" Then
                'UOM_CODE.ForeColor = Drawing.Color.Red
            Else
                Call BindGV()
            End If

            If Session("pagemode") <> "N" Then
                If DC_DATE_CODE.Text <> "" Then
                    DC_DATE_CODE.BorderWidth = 0
                    DC_DATE_CODE.BackColor = Drawing.Color.Transparent
                    DC_DATE_CODE.ReadOnly = True
                End If

                STORER_CODE.CssClass = ""
            End If

            REM **********************
        End If
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""

        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If DC_DATE_CODE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_DC_DATE_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_DC_DATE_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If DC_CONV_DATE.Text.Trim <> "" And Not gU.isValidDate(DC_CONV_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_DC_CONV_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_DC_CONV_DATE.Text & "!", Session("gLang"))
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

                    dupSQL = "select 1 from WMS_DATE_CODE " & _
                                "where DC_DATE_CODE = '" & gU.dbEncode(DC_DATE_CODE.Text) & "' " & _
                                "and imp_code='" & gU.dbEncode(imp_code.Trim) & "' " & _
                                "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue.Trim) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then
                        'nextNo = DB.getDocNo("UOM", gConn, transaction)
                        nextNo = DC_DATE_CODE.Text

                        sql_string = "insert into WMS_DATE_CODE (" & _
                        "IMP_CODE, STORER_CODE, DC_DATE_CODE, DC_CONV_DATE, " & _
                        "sys_cb, sys_cd, sys_lub, sys_lud) values ( " & _
                         gU.convdbNVCData(gU.dbEncode(imp_code)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue.Trim)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & _
                        gU.convdbDate(gU.dbEncode(DC_CONV_DATE.Text)) & "," & _
                        "N'" & Session("usr_id") & "',Getdate(),N'" & Session("usr_id") & "',Getdate())"
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Batch No. Maintenance!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "貨料單位資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM char "N" is use for update Unicode
                    sql_string = "update WMS_DATE_CODE set " & _
                                    "DC_CONV_DATE = " & gU.convdbDate(gU.dbEncode(DC_CONV_DATE.Text)) & ", " & _
                                    "sys_lub = N'" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where DC_DATE_CODE = '" & gU.dbEncode(DC_DATE_CODE.Text) & "' " & _
                                    "and imp_code='" & gU.dbEncode(imp_code.Trim) & "' " & _
                                    "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue.Trim) & "' "
                End If
                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    DC_DATE_CODE.Text = nextNo
                    DC_DATE_CODE.ForeColor = Drawing.Color.Black
                    DC_DATE_CODE.Font.Size = 10
                    'uom_code.ReadOnly = True
                    'uom_code.BackColor = Drawing.Color.Transparent
                    'uom_code.BorderWidth = 0

                    STORER_CODE.CssClass = ""
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
        If ViewState("DC_DATE_CODE") <> "" Then
            pk_code = ViewState("DC_DATE_CODE")            
        Else
            pk_code = Server.UrlDecode(Request("DC_DATE_CODE"))

            ViewState("DC_DATE_CODE") = pk_code
        End If

        If ViewState("STORER_CODE") <> "" Then
            storerCode = ViewState("STORER_CODE")
        Else
            storerCode = Server.UrlDecode(Request("STORER_CODE"))

            ViewState("STORER_CODE") = storerCode
        End If
        REM **********************


        If Session("pagemode") <> "N" Then
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header

            SQLString = "SELECT wms_date_code.IMP_CODE, wms_date_code.STORER_CODE, wms_date_code.DC_DATE_CODE," & _
                        " CONVERT(nvarchar(30), wms_date_code.DC_CONV_DATE, " & DDFORMAT & ") as DC_CONV_DATE, wms_date_code.SYS_CB, wms_date_code.SYS_CD, wms_date_code.SYS_LUB, wms_date_code.SYS_LUD " & _
                        " FROM WMS_DATE_CODE WHERE DC_DATE_CODE = '" & gU.dbEncode(pk_code) & "' AND IMP_CODE = '" & gU.dbEncode(imp_code) & "' AND STORER_CODE = '" & gU.dbEncode(storerCode) & "' "

            SQLString = SQLString & " " & WhereStr

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                If dt.Rows(0).Item("storer_code").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME AS STO_NAME from WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
                End If

                STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                DC_DATE_CODE.Text = dt.Rows(0).Item("DC_DATE_CODE").ToString
                DC_CONV_DATE.Text = dt.Rows(0).Item("DC_CONV_DATE").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)
            End If
        End If
    End Sub
End Class
