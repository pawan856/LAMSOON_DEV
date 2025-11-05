Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class MASTER_SM_storerMaster
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

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If

        REM ****************************

        If Not IsPostBack Then

            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            REM **********************
            REM Modify Here
            If Session("gLang") = "E" Then
                lheader.Text = "Storer Master Maintenance"
                lbl_STORER_CODE.Text = "Storer Code:"
                lbl_STO_STATUS.Text = "Status:"
                lbl_STO_SHORTNAME.Text = "Storer Short Name:"
                lbl_STO_NAME.Text = "Storer Name:"
                lbl_STO_NAME_CH.Text = "Storer Name in Chinese:"
                lbl_STO_ADDR.Text = "Address:"
                lbl_STO_REGION.Text = "Region:"
                lbl_STO_AREA.Text = "Area:"
                lbl_STO_COUNTRY.Text = "Country:"

                lbl_STO_CONT_PER_GEN.Text = "Display Trail Only:"
                lbl_STO_CONT_TEL_GEN.Text = "Tel.:"
                lbl_STO_CONT_DEPT_GEN.Text = "Department:"
                lbl_STO_CONT_EMAIL_GEN.Text = "Email:"

                lbl_STO_CONT_PER_MGR.Text = "Manager:"
                lbl_STO_CONT_TEL_MGR.Text = "Tel.:"
                lbl_STO_CONT_DEPT_MGR.Text = "Department:"
                lbl_STO_CONT_EMAIL_MGR.Text = "Email:"

                lbl_STO_CONT_PER_ACC.Text = "Accountant:"
                lbl_STO_CONT_TEL_ACC.Text = "Tel.:"
                lbl_STO_CONT_DEPT_ACC.Text = "Department:"
                lbl_STO_CONT_EMAIL_ACC.Text = "Email:"

                lbl_STO_CONT_PER_ORD.Text = "Contact Person(Order):"
                lbl_STO_CONT_TEL_ORD.Text = "Tel.:"
                lbl_STO_CONT_DEPT_ORD.Text = "Department:"
                lbl_STO_CONT_EMAIL_ORD.Text = "Email:"

                lbl_STO_CONT_PER_LOG.Text = "Contact Person(Logistic):"
                lbl_STO_CONT_TEL_LOG.Text = "Tel.:"
                lbl_STO_CONT_DEPT_LOG.Text = "Department:"
                lbl_STO_CONT_EMAIL_LOG.Text = "Email:"

                lbl_STO_FAX.Text = "Fax:"
                lbl_STO_MAIN_TEL.Text = "Main Tel."
                lbl_STO_WEBSITE.Text = "Web Site:"
                lbl_STO_CURR.Text = "Currency:"
                lbl_STO_REM.Text = "Lot Allocation Logic (NEW for new logic or blank):"
                lbl_sys_cb.Text = "CB"
                lbl_sys_lub.Text = "LUB"
                lbl_sys_cd.Text = "CD"
                lbl_sys_lud.Text = "LUD"
                saveBtn1.Text = "Save"
                saveBtn2.Text = "Save"
                saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
                saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"

                If Session("pagemode") = "N" Then
                    STORER_CODE.Text = "[Code will be auto generated]"
                End If

            ElseIf Session("gLang") = "C" Then
                lheader.Text = "貨主主資料庫維護"
                lbl_STORER_CODE.Text = "貨主代碼:"
                lbl_STO_STATUS.Text = "狀態:"
                lbl_STO_SHORTNAME.Text = "貨主簡稱:"
                lbl_STO_NAME.Text = "貨主名稱:"
                lbl_STO_NAME_CH.Text = "貨主中文名稱:"
                lbl_STO_ADDR.Text = "地址:"
                lbl_STO_REGION.Text = "地區:"
                lbl_STO_AREA.Text = "區域:"
                lbl_STO_COUNTRY.Text = "國家:"

                lbl_STO_CONT_PER_GEN.Text = "Display Trail Only:"
                lbl_STO_CONT_TEL_GEN.Text = "電話:"
                lbl_STO_CONT_DEPT_GEN.Text = "部門:"
                lbl_STO_CONT_EMAIL_GEN.Text = "電子郵件:"

                lbl_STO_CONT_PER_MGR.Text = "經理:"
                lbl_STO_CONT_TEL_MGR.Text = "電話:"
                lbl_STO_CONT_DEPT_MGR.Text = "部門:"
                lbl_STO_CONT_EMAIL_MGR.Text = "電子郵件:"

                lbl_STO_CONT_PER_ACC.Text = "會計:"
                lbl_STO_CONT_TEL_ACC.Text = "電話:"
                lbl_STO_CONT_DEPT_ACC.Text = "部門:"
                lbl_STO_CONT_EMAIL_ACC.Text = "電子郵件:"

                lbl_STO_CONT_PER_ORD.Text = "聯絡人(Order):"
                lbl_STO_CONT_TEL_ORD.Text = "電話:"
                lbl_STO_CONT_DEPT_ORD.Text = "部門:"
                lbl_STO_CONT_EMAIL_ORD.Text = "電子郵件:"

                lbl_STO_CONT_PER_LOG.Text = "聯絡人(物流):"
                lbl_STO_CONT_TEL_LOG.Text = "電話:"
                lbl_STO_CONT_DEPT_LOG.Text = "部門:"
                lbl_STO_CONT_EMAIL_LOG.Text = "電子郵件:"

                lbl_STO_FAX.Text = "傳真:"
                lbl_STO_MAIN_TEL.Text = "主要電話:"
                lbl_STO_WEBSITE.Text = "網頁:"

                lbl_STO_CURR.Text = "貨幣:"
                lbl_STO_REM.Text = "Lot Allocation Logic (NEW for new logic or blank):"
                lbl_sys_cb.Text = "創建者"
                lbl_sys_lub.Text = "最後更新者"
                lbl_sys_cd.Text = "創建日期"
                lbl_sys_lud.Text = "最後更新日期"
                saveBtn1.Text = "儲存"
                saveBtn2.Text = "儲存"
                saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
                saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"
                If Session("pagemode") = "N" Then
                    STORER_CODE.Text = "[代碼會自動產生]"
                End If
            End If

            REM **********************
            STO_NAME.CssClass = "REQUIRED"
            STO_ADDR1.CssClass = "REQUIRED"
            STO_ADDR2.CssClass = "REQUIRED"
            STO_ADDR3.CssClass = "REQUIRED"
            STO_COUNTRY.CssClass = "REQUIRED"
            STO_SHORTNAME.CssClass = "REQUIRED"
            'STORER_CODE.CssClass = "REQUIRED"

            REM**********************
            uiFun.load_dropdown(STO_CURR, "select CCY_CODE, CCY_NAME from WMS_CURRENCY ORDER BY 2", "CCY_CODE", "CCY_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdownBy_ColCode(STO_STATUS, "WMS_STORER.STO_STATUS", Session("gLang"), , Session("gSelectLabel"))
            uiFun.load_radioList(STO_BATCH_FIELD_REF, "select '' as colc_code, 'By Batch No.' as colc_eng_value, 0 as colc_display_seq  union Select colc_code, colc_eng_value, wms_col_code.COLC_DISPLAY_SEQ from wms_col_code where colc_tabcol='WMS_STORER.STO_BATCH_FIELD_REF' order by colc_display_seq", "colc_code", "colc_eng_value")
            REM **********************

            If Session("pagemode") = "N" Then
                STORER_CODE.ForeColor = Drawing.Color.Red
            Else
                Call BindGV()
            End If

            If Session("pagemode") <> "N" And STORER_CODE.Text <> "" Then
                'STORER_CODE.BorderWidth = 0
                'STORER_CODE.BackColor = Drawing.Color.Transparent
                'STORER_CODE.ReadOnly = True
            End If

            If Session("pagemode") = "N" Then
                STO_STATUS.SelectedValue = "ACTIVE"
                btnFieldMast.Visible = False
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

        If STO_SHORTNAME.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STO_SHORTNAME.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STO_SHORTNAME.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If STO_NAME.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STO_NAME.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STO_NAME.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If STO_ADDR1.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STO_ADDR.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STO_ADDR.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If STO_COUNTRY.Text = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STO_COUNTRY.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STO_COUNTRY.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If STO_CONT_EMAIL_GEN.Text.Trim <> "" And Not gU.isValidEmail(STO_CONT_EMAIL_GEN.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_STO_CONT_EMAIL_GEN.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_STO_CONT_EMAIL_GEN.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If STO_CONT_EMAIL_MGR.Text.Trim <> "" And Not gU.isValidEmail(STO_CONT_EMAIL_MGR.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_STO_CONT_EMAIL_MGR.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_STO_CONT_EMAIL_MGR.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If STO_CONT_EMAIL_ACC.Text.Trim <> "" And Not gU.isValidEmail(STO_CONT_EMAIL_ACC.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_STO_CONT_EMAIL_ACC.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_STO_CONT_EMAIL_ACC.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If STO_CONT_EMAIL_ORD.Text.Trim <> "" And Not gU.isValidEmail(STO_CONT_EMAIL_ORD.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_STO_CONT_EMAIL_ORD.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_STO_CONT_EMAIL_ORD.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If STO_CONT_EMAIL_LOG.Text.Trim <> "" And Not gU.isValidEmail(STO_CONT_EMAIL_LOG.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_STO_CONT_EMAIL_LOG.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_STO_CONT_EMAIL_LOG.Text & "!", Session("gLang"))
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
        Dim paP As GlobalDBFunc.DBCmdPara

        If validateAll() Then
            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction

            Try
                REM **********************
                REM Modify Here


                If Session("pagemode") = "N" Then


                    dupSQL = "select 1 from wms_storer " & _
                              "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.Text) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then

                        nextNo = DB.getDocNo("SM", gConn, transaction)
                        'nextNo = STORER_CODE.Text

                        sql_string = "insert into wms_storer (" &
                        "storer_code, imp_code, STO_status, " &
                        "sto_shortname, sto_name, sto_name_ch, " &
                        "sto_addr1, sto_addr2, sto_addr3, " &
                        "sto_area, sto_region, sto_country, " &
                        "sto_cont_per_gen, sto_cont_tel_gen, sto_cont_dept_gen, sto_cont_email_gen, " &
                        "sto_cont_per_mgr, sto_cont_tel_mgr, sto_cont_dept_mgr, sto_cont_email_mgr, " &
                        "sto_cont_per_acc, sto_cont_tel_acc, sto_cont_dept_acc, sto_cont_email_acc, " &
                        "sto_cont_per_ord, sto_cont_tel_ord, sto_cont_dept_ord, sto_cont_email_ord, " &
                        "sto_cont_per_log, sto_cont_tel_log, sto_cont_dept_log, sto_cont_email_log, " &
                        "sto_fax, sto_main_tel, sto_website, " &
                        "sto_curr, sto_rem, STO_BATCH_FIELD_REF, " &
                        "sys_cb, sys_cd, sys_lub, sys_lud) values ( " &
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STO_STATUS.SelectedValue)) & "," &
                        gU.convdbNVCData(gU.dbEncode(STO_SHORTNAME.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_NAME.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_NAME_CH.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(STO_ADDR1.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_ADDR2.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_ADDR3.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(STO_AREA.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_REGION.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_COUNTRY.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(STO_CONT_PER_GEN.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_TEL_GEN.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_DEPT_GEN.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_EMAIL_GEN.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(STO_CONT_PER_MGR.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_TEL_MGR.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_DEPT_MGR.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_EMAIL_MGR.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(STO_CONT_PER_ACC.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_TEL_ACC.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_DEPT_ACC.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_EMAIL_ACC.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(STO_CONT_PER_ORD.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_TEL_ORD.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_DEPT_ORD.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_EMAIL_ORD.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(STO_CONT_PER_LOG.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_TEL_LOG.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_DEPT_LOG.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_CONT_EMAIL_LOG.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(STO_FAX.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_MAIN_TEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(STO_WEBSITE.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(STO_CURR.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(STO_REM.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(STO_BATCH_FIELD_REF.SelectedValue)) & "," &
                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"

                        gDB.amendData(sql_string, gConn, transaction)


                        'sql_string = " insert into wms_field_option (SELECT '" & gU.dbEncode(Session("IMP_CODE")) & "',  '" & gU.dbEncode(nextNo) & "',  wms_field_option.FUN_CODE,  wms_field_option.FLDO_FIELD_NAME,  wms_field_option.FLDO_FIELD_DESC,  wms_field_option.FLDO_TABLE_NAME, " & _
                        '             " wms_field_option.FLDO_FIELD_REF,  wms_field_option.FLDO_FIELD_OPTION,  wms_field_option.FLDO_DISPLAY_SEQ, wms_field_option.FLDO_STATUS,  wms_field_option.FLDO_MASTER_TABLE_NAME,  wms_field_option.FLDO_FILED_CTRL, " & _
                        '             " wms_field_option.FLDO_FIELD_TYPE FROM wms_field_option WHERE IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "' AND STORER_CODE='TEMPLATE') "

                        'gDB.amendData(sql_string, gConn, transaction)

                        paP = New GlobalDBFunc.DBCmdPara
                        sql_string = " insert into wms_vendor (IMP_CODE,STORER_CODE,VND_CODE,VND_STATUS,VND_NAME,sys_cd,sys_cb,sys_lud,sys_lub) values ( " & _
                                     paP.AP(Session("IMP_CODE")) & ", " & paP.AP(nextNo) & ", 'DEF_VEND', 'ACTIVE', 'Default Vendor'," & _
                                     " Getdate()," & paP.AP(Session("usr_id")) & ",Getdate()," & paP.AP(Session("usr_id")) & _
                                     ") "
                        gDB.amendData(sql_string, gConn, transaction, paP)
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Storer Master!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "貨主資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_storer set " &
                                    "sto_status = " & gU.convdbNVCData(gU.dbEncode(STO_STATUS.SelectedValue)) & ", " &
                                    "sto_shortname = " & gU.convdbNVCData(gU.dbEncode(STO_SHORTNAME.Text)) & ", " &
                                    "sto_name = " & gU.convdbNVCData(gU.dbEncode(STO_NAME.Text)) & ", " &
                                    "sto_name_ch = " & gU.convdbNVCData(gU.dbEncode(STO_NAME_CH.Text)) & ", " &
                                    "sto_addr1 = " & gU.convdbNVCData(gU.dbEncode(STO_ADDR1.Text)) & ", " &
                                    "sto_addr2 = " & gU.convdbNVCData(gU.dbEncode(STO_ADDR2.Text)) & ", " &
                                    "sto_addr3 = " & gU.convdbNVCData(gU.dbEncode(STO_ADDR3.Text)) & ", " &
                                    "sto_area = " & gU.convdbNVCData(gU.dbEncode(STO_AREA.Text)) & ", " &
                                    "sto_region = " & gU.convdbNVCData(gU.dbEncode(STO_REGION.Text)) & ", " &
                                    "sto_country = " & gU.convdbNVCData(gU.dbEncode(STO_COUNTRY.Text)) & ", " &
                                    "sto_cont_per_gen = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_PER_GEN.Text)) & ", " &
                                    "sto_cont_tel_gen = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_TEL_GEN.Text)) & ", " &
                                    "sto_cont_dept_gen = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_DEPT_GEN.Text)) & ", " &
                                    "sto_cont_email_gen = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_EMAIL_GEN.Text)) & ", " &
                                    "sto_cont_per_mgr = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_PER_MGR.Text)) & ", " &
                                    "sto_cont_tel_mgr = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_TEL_MGR.Text)) & ", " &
                                    "sto_cont_dept_mgr = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_DEPT_MGR.Text)) & ", " &
                                    "sto_cont_email_mgr = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_EMAIL_MGR.Text)) & ", " &
                                    "sto_cont_per_acc = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_PER_ACC.Text)) & ", " &
                                    "sto_cont_tel_acc = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_TEL_ACC.Text)) & ", " &
                                    "sto_cont_dept_acc = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_DEPT_ACC.Text)) & ", " &
                                    "sto_cont_email_acc = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_EMAIL_ACC.Text)) & ", " &
                                    "sto_cont_per_ord = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_PER_ORD.Text)) & ", " &
                                    "sto_cont_tel_ord = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_TEL_ORD.Text)) & ", " &
                                    "sto_cont_dept_ord = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_DEPT_ORD.Text)) & ", " &
                                    "sto_cont_email_ord = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_EMAIL_ORD.Text)) & ", " &
                                    "sto_cont_per_log = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_PER_LOG.Text)) & ", " &
                                    "sto_cont_tel_log = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_TEL_LOG.Text)) & ", " &
                                    "sto_cont_dept_log = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_DEPT_LOG.Text)) & ", " &
                                    "sto_cont_email_log = " & gU.convdbNVCData(gU.dbEncode(STO_CONT_EMAIL_LOG.Text)) & ", " &
                                    "sto_fax = " & gU.convdbNVCData(gU.dbEncode(STO_FAX.Text)) & ", " &
                                    "sto_main_tel = " & gU.convdbNVCData(gU.dbEncode(STO_MAIN_TEL.Text)) & ", " &
                                    "sto_website = " & gU.convdbNVCData(gU.dbEncode(STO_WEBSITE.Text)) & ", " &
                                    "sto_curr = " & gU.convdbNVCData(gU.dbEncode(STO_CURR.SelectedValue)) & ", " &
                                    "sto_rem = " & gU.convdbNVCData(gU.dbEncode(STO_REM.SelectedValue)) & ", " &
                                    "STO_BATCH_FIELD_REF = " & gU.convdbNVCData(gU.dbEncode(STO_BATCH_FIELD_REF.SelectedValue)) & ", " &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                    "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.Text) & "' "
                    gDB.amendData(sql_string, gConn, transaction)
                End If


                'Response.Write(sql_string)


                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    STORER_CODE.Text = nextNo
                    STORER_CODE.ForeColor = Drawing.Color.Black
                    STORER_CODE.Font.Size = 10

                    'STORER_CODE.BackColor = Drawing.Color.Transparent
                    'STORER_CODE.BorderWidth = 0
                    'STORER_CODE.ReadOnly = True
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
        Dim pk_code As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("storer_code") <> "" Then
            pk_code = ViewState("storer_code")

        Else
            pk_code = Server.UrlDecode(Request("storer_code"))

            ViewState("storer_code") = pk_code
        End If
        REM **********************

        'STO_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            'CO_CODE.ForeColor = Drawing.Color.Red
            STO_STATUS.SelectedValue = "ACTIVE"
            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header

            SQLString = "SELECT * FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(pk_code) & "' AND IMP_CODE = '" & Session("IMP_CODE") & "'"

            SQLString = SQLString & " " & WhereStr

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                STORER_CODE.Text = dt.Rows(0).Item("storer_code").ToString

                STO_STATUS.SelectedValue = gU.decodeNull(dt.Rows(0).Item("sto_status").ToString.ToUpper, "ACTIVE")
                STO_SHORTNAME.Text = dt.Rows(0).Item("sto_shortname").ToString
                STO_NAME.Text = dt.Rows(0).Item("sto_name").ToString
                STO_NAME_CH.Text = dt.Rows(0).Item("sto_name_ch").ToString
                STO_ADDR1.Text = dt.Rows(0).Item("sto_addr1").ToString
                STO_ADDR2.Text = dt.Rows(0).Item("sto_addr2").ToString
                STO_ADDR3.Text = dt.Rows(0).Item("sto_addr3").ToString
                STO_REGION.Text = dt.Rows(0).Item("sto_region").ToString
                STO_AREA.Text = dt.Rows(0).Item("sto_area").ToString
                STO_COUNTRY.Text = dt.Rows(0).Item("sto_country").ToString
                STO_CONT_PER_GEN.Text = dt.Rows(0).Item("sto_cont_per_gen").ToString
                STO_CONT_TEL_GEN.Text = dt.Rows(0).Item("sto_cont_tel_gen").ToString
                STO_CONT_DEPT_GEN.Text = dt.Rows(0).Item("sto_cont_dept_gen").ToString
                STO_CONT_EMAIL_GEN.Text = dt.Rows(0).Item("sto_cont_email_gen").ToString
                STO_CONT_PER_MGR.Text = dt.Rows(0).Item("sto_cont_per_mgr").ToString
                STO_CONT_TEL_MGR.Text = dt.Rows(0).Item("sto_cont_tel_mgr").ToString
                STO_CONT_DEPT_MGR.Text = dt.Rows(0).Item("sto_cont_dept_mgr").ToString
                STO_CONT_EMAIL_MGR.Text = dt.Rows(0).Item("sto_cont_email_mgr").ToString
                STO_CONT_PER_ACC.Text = dt.Rows(0).Item("sto_cont_per_acc").ToString
                STO_CONT_TEL_ACC.Text = dt.Rows(0).Item("sto_cont_tel_acc").ToString
                STO_CONT_DEPT_ACC.Text = dt.Rows(0).Item("sto_cont_dept_acc").ToString
                STO_CONT_EMAIL_ACC.Text = dt.Rows(0).Item("sto_cont_email_acc").ToString
                STO_CONT_PER_ORD.Text = dt.Rows(0).Item("sto_cont_per_ord").ToString
                STO_CONT_TEL_ORD.Text = dt.Rows(0).Item("sto_cont_tel_ord").ToString
                STO_CONT_DEPT_ORD.Text = dt.Rows(0).Item("sto_cont_dept_ord").ToString
                STO_CONT_DEPT_ORD.Text = dt.Rows(0).Item("sto_cont_dept_ord").ToString
                STO_CONT_PER_LOG.Text = dt.Rows(0).Item("sto_cont_per_log").ToString
                STO_CONT_TEL_LOG.Text = dt.Rows(0).Item("sto_cont_tel_log").ToString
                STO_CONT_DEPT_LOG.Text = dt.Rows(0).Item("sto_cont_dept_log").ToString
                STO_CONT_EMAIL_LOG.Text = dt.Rows(0).Item("sto_cont_email_log").ToString
                STO_FAX.Text = dt.Rows(0).Item("sto_fax").ToString
                STO_MAIN_TEL.Text = dt.Rows(0).Item("sto_main_tel").ToString
                STO_WEBSITE.Text = dt.Rows(0).Item("sto_website").ToString
                STO_CURR.SelectedValue = dt.Rows(0).Item("sto_curr").ToString
                STO_REM.SelectedValue = dt.Rows(0).Item("sto_rem").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                STO_BATCH_FIELD_REF.SelectedValue = dt.Rows(0).Item("STO_BATCH_FIELD_REF").ToString.Trim

            End If

            'btnFieldMast.Visible = True
        End If
    End Sub
End Class
