Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class INV_invoiceSetting
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

            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(INVR_CURR, "select CCY_CODE, CCY_NAME from WMS_CURRENCY ORDER BY 2", "CCY_CODE", "CCY_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(INVR_TYPE, "select colc_code, colc_eng_value from wms_col_code where colc_tabcol = 'WMS_INV_RATES.INVR_TYPE' ORDER BY colc_display_seq", "colc_code", "colc_eng_value", , Session("gSelectLabel"))
            uiFun.load_dropdownBy_ColCode(INVR_STATUS, "WMS_INV_RATES.INVR_STATUS", Session("sLang"))

            Image_INVR_LOC.Attributes.Add("onclick", "LocLookUp('" & INVR_LOC.ClientID & "')")

            REM **********************
            REM Modify Here
            If Session("gLang") = "E" Then
                lheader.Text = "Invoice Setting"
                lbl_STORER_CODE.Text = "Storer"
                lbl_INV_DISPLAY_SEQ.Text = "Display Seq."
                lbl_INVR_TYPE.Text = "Type"
                lbl_INVR_LOC.Text = "Location"
                lbl_INVR_DESC.Text = "Description"
                lbl_INVR_EFF_DATE.Text = "Eff Date"
                lbl_INVR_EXP_DATE.Text = "Exp date"
                lbl_INVR_UNIT.Text = "Unit"
                lbl_INVR_RATE.Text = "Rate"
                lbl_INVR_CURR.Text = "Currency"
                lbl_INVR_STATUS.Text = "Status"

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

                lheader.Text = "Invoice Setting"
                lbl_STORER_CODE.Text = "Storer"
                lbl_INV_DISPLAY_SEQ.Text = "Display Seq."
                lbl_INVR_TYPE.Text = "Type"
                lbl_INVR_LOC.Text = "Location"
                lbl_INVR_DESC.Text = "Description"
                lbl_INVR_EFF_DATE.Text = "Eff Date"
                lbl_INVR_EXP_DATE.Text = "Exp date"
                lbl_INVR_UNIT.Text = "Unit"
                lbl_INVR_RATE.Text = "Rate"
                lbl_INVR_CURR.Text = "Currency"
                lbl_INVR_STATUS.Text = "Status"

                lbl_sys_cb.Text = "創建者"
                lbl_sys_lub.Text = "最後更新者"
                lbl_sys_cd.Text = "創建日期"
                lbl_sys_lud.Text = "最後更新日期"
                saveBtn1.Text = "儲存"
                saveBtn2.Text = "儲存"
                saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
                saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"

                If Session("pagemode") = "N" Then

                End If
            End If

            REM **********************
            STORER_CODE.CssClass = "REQUIRED"
            INVR_TYPE.CssClass = "REQUIRED"
            INVR_LOC.CssClass = "REQUIRED"

            If Session("pagemode") = "N" Then
                If STORER_CODE.SelectedValue = "" Then
                    STORER_CODE.SelectedValue = Session("usr_pref_storer")
                End If

                If INVR_STATUS.SelectedValue = "" Then
                    INVR_STATUS.SelectedValue = "ACTIVE"
                End If

                INVR_LOC.Text = "ALL"
            Else

                Call BindGV()
            End If

            REM **********************
        End If
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""

        If STORER_CODE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If INVR_TYPE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_INVR_TYPE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_INVR_TYPE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If INVR_LOC.Text.Trim = "" Then
            INVR_LOC.Text = "ALL"
            'If Session("gLang") = "E" Then
            '    uiFun.displayMsg(Me, "", lbl_INVR_LOC.Text & " cannot be empty!", Session("gLang"))
            ' Else
            '      uiFun.displayMsg(Me, "", lbl_INVR_LOC.Text & "不能空白!", Session("gLang"))
            '   End If
            '    Return False
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

                    dupSQL = "select 1 from WMS_INV_RATES " & _
                                "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and INV_SEQ = '" & INV_SEQ.Value & "' " & _
                                "and INVR_TYPE = '" & gU.dbEncode(INVR_TYPE.SelectedValue) & "' " & _
                                "and INVR_LOC = '" & gU.dbEncode(INVR_LOC.Text) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then
                        nextNo = DB.getDocNo("INVRATE", gConn, transaction)

                        sql_string = "insert into WMS_INV_RATES (" & _
                        "imp_code, STORER_CODE, INV_SEQ, INV_DISPLAY_SEQ," & _
                        "INVR_TYPE, INVR_LOC, INVR_DESC, INVR_EFF_DATE, INVR_EXP_DATE , " & _
                        "INVR_UNIT, INVR_RATE, INVR_CURR, INVR_STATUS, " & _
                        "sys_cb, sys_cd, sys_lub, sys_lud) values ( " & _
                        gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & _
                        gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & _
                        gU.dbEncode(gU.decodeNullOrEmpty(INV_DISPLAY_SEQ.Text, "0")) & "," & _
                        gU.convdbNVCData(gU.dbEncode(INVR_TYPE.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(INVR_LOC.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(INVR_DESC.Text)) & "," & _
                        gU.convdbDate(gU.dbEncode(INVR_EFF_DATE.Text)) & "," & _
                        gU.convdbDate(gU.dbEncode(INVR_EXP_DATE.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(INVR_UNIT.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(INVR_RATE.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(INVR_CURR.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(INVR_STATUS.SelectedValue)) & "," & _
                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Invoice Setting table!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "貨料單位資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM char "N" is use for update Unicode
                    sql_string = "update WMS_INV_RATES set " & _
                                    "STORER_CODE = " & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & _
                                    "INV_DISPLAY_SEQ = " & gU.dbEncode(gU.decodeNullOrEmpty(INV_DISPLAY_SEQ.Text, "0")) & "," & _
                                    "INVR_TYPE = " & gU.convdbNVCData(gU.dbEncode(INVR_TYPE.SelectedValue)) & "," & _
                                    "INVR_LOC = " & gU.convdbNVCData(gU.dbEncode(INVR_LOC.Text)) & "," & _
                                    "INVR_DESC = " & gU.convdbNVCData(gU.dbEncode(INVR_DESC.Text)) & "," & _
                                    "INVR_EFF_DATE = " & gU.convdbDate(gU.dbEncode(INVR_EFF_DATE.Text)) & "," & _
                                    "INVR_EXP_DATE = " & gU.convdbDate(gU.dbEncode(INVR_EXP_DATE.Text)) & "," & _
                                    "INVR_UNIT = " & gU.convdbNVCData(gU.dbEncode(INVR_UNIT.Text)) & "," & _
                                    "INVR_RATE = " & gU.convdbNVCData(gU.dbEncode(INVR_RATE.Text)) & "," & _
                                    "INVR_CURR = " & gU.convdbNVCData(gU.dbEncode(INVR_CURR.SelectedValue)) & "," & _
                                    "INVR_STATUS = " & gU.convdbNVCData(gU.dbEncode(INVR_STATUS.SelectedValue)) & "," & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                    "and INV_SEQ = " & gU.convdbNVCData(gU.dbEncode(INV_SEQ.Value)) & " " & _
                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and INVR_TYPE = '" & gU.dbEncode(INVR_TYPE.SelectedValue) & "' " & _
                                    "and INVR_LOC = '" & gU.dbEncode(INVR_LOC.Text) & "' "

                End If
                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    INV_SEQ.Value = nextNo
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
        Dim lSTORER_CODE, lINV_SEQ, lINVR_TYPE, lINVR_LOC As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("uom_code") <> "" Then
            lSTORER_CODE = ViewState("STORER_CODE")
            lINV_SEQ = ViewState("INV_SEQ")
            lINVR_TYPE = ViewState("INVR_TYPE")
            lINVR_LOC = ViewState("INVR_LOC")
        Else
            lSTORER_CODE = Request("STORER_CODE")
            lINV_SEQ = Request("INV_SEQ")
            lINVR_TYPE = Request("INVR_TYPE")
            lINVR_LOC = Request("INVR_LOC")

            ViewState("STORER_CODE") = lSTORER_CODE
            ViewState("INV_SEQ") = lINV_SEQ
            ViewState("INVR_TYPE") = lINVR_TYPE
            ViewState("INVR_LOC") = lINVR_LOC
        End If
        REM **********************


        If Session("pagemode") <> "N" Then
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header

            SQLString = "SELECT wms_inv_rates.IMP_CODE,  wms_inv_rates.STORER_CODE,  wms_inv_rates.INV_SEQ,  wms_inv_rates.INV_DISPLAY_SEQ, " & _
                        " wms_inv_rates.INVR_TYPE,  wms_inv_rates.INVR_LOC,  wms_inv_rates.INVR_STATUS,  wms_inv_rates.INVR_DESC, " & _
                        " Convert(varchar,wms_inv_rates.INVR_EFF_DATE,103) as INVR_EFF_DATE,  Convert(varchar, wms_inv_rates.INVR_EXP_DATE,103) as INVR_EXP_DATE,  wms_inv_rates.INVR_UNIT,  wms_inv_rates.INVR_RATE, " & _
                        " wms_inv_rates.INVR_CURR,  wms_inv_rates.SYS_LUB,  wms_inv_rates.SYS_LUD,  wms_inv_rates.SYS_CD,  wms_inv_rates.SYS_CB FROM WMS_INV_RATES WHERE IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                        "AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
                        "AND INV_SEQ = '" & gU.dbEncode(lINV_SEQ) & "' " & _
                        "AND INVR_TYPE = '" & gU.dbEncode(lINVR_TYPE) & "' " & _
                        "AND INVR_LOC = '" & gU.dbEncode(lINVR_LOC) & "' "

            SQLString = SQLString & " " & WhereStr

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
                End If
                If dt.Rows(0).Item("INVR_TYPE").ToString = "" Then
                    uiFun.load_dropdown(INVR_TYPE, "select colc_code, colc_eng_value from wms_col_code where colc_tabcol = 'WMS_INV_RATES.INVR_TYPE' ORDER BY colc_display_seq", "colc_code", "colc_eng_value", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(INVR_TYPE, "select colc_code, colc_eng_value from wms_col_code where colc_tabcol = 'WMS_INV_RATES.INVR_TYPE' and colc_code = '" & gU.dbEncode(dt.Rows(0).Item("INVR_TYPE").ToString) & "' ORDER BY colc_display_seq", "colc_code", "colc_eng_value", , Session("gSelectLabel"), , True)
                End If
                If dt.Rows(0).Item("INVR_STATUS").ToString = "" Then
                    uiFun.load_dropdownBy_ColCode(INVR_STATUS, "WMS_INV_RATES.INVR_STATUS", Session("gLang"))
                Else
                    uiFun.load_dropdownBy_ColCode(INVR_STATUS, "WMS_INV_RATES.INVR_STATUS", Session("gLang"), , , dt.Rows(0).Item("INVR_STATUS").ToString)
                End If

                INV_SEQ.Value = dt.Rows(0).Item("INV_SEQ").ToString
                If dt.Rows(0).Item("INVR_LOC").ToString <> "" Then
                    INVR_LOC.ReadOnly = True
                    Image_INVR_LOC.Visible = False
                End If
                INVR_LOC.Text = dt.Rows(0).Item("INVR_LOC").ToString
                INV_DISPLAY_SEQ.Text = dt.Rows(0).Item("INV_DISPLAY_SEQ").ToString
                INVR_DESC.Text = dt.Rows(0).Item("INVR_DESC").ToString
                INVR_EFF_DATE.Text = dt.Rows(0).Item("INVR_EFF_DATE").ToString
                INVR_EXP_DATE.Text = dt.Rows(0).Item("INVR_EXP_DATE").ToString
                INVR_UNIT.Text = dt.Rows(0).Item("INVR_UNIT").ToString
                INVR_RATE.Text = dt.Rows(0).Item("INVR_RATE").ToString
                INVR_CURR.SelectedValue = dt.Rows(0).Item("INVR_CURR").ToString

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)
            End If
        End If
    End Sub

    Protected Sub INVR_TYPE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles INVR_TYPE.SelectedIndexChanged
        INVR_DESC.Text = INVR_TYPE.SelectedItem.Text

    End Sub
End Class
