Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class vendorMaster
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
                lheader.Text = "Vendor Master Maintenance"
                lSTORER_CODE.Text = "Storer:"
                lVND_CODE.Text = "Vendor Code:"
                lVND_STATUS.Text = "Status:"
                lVND_SHORTNAME.Text = "Vendor Short Name:"
                lVND_NAME.Text = "Vendor Name:"
                lVND_NAME_CH.Text = "Vendor Name in Chinese:"
                lVND_ADDR.Text = "Address:"
                lVND_REGION.Text = "Region:"
                lVND_AREA.Text = "Area:"
                lVND_COUNTRY.Text = "Country:"
                lVND_CONT_PER1.Text = "Contact Person 1:"
                lVND_CONT_PER2.Text = "Contact Person 2:"
                lVND_CONT_TEL1.Text = "Tel.:"
                lVND_CONT_TEL2.Text = "Tel.:"
                lVND_CONT_DEPT1.Text = "Department:"
                lVND_CONT_DEPT2.Text = "Department:"
                lVND_CONT_EMAIL1.Text = "Email:"
                lVND_CONT_EMAIL2.Text = "Email:"
                lVND_FAX.Text = "Fax:"
                lVND_MAIN_TEL.Text = "Main Tel.:"
                lVND_WEBSITE.Text = "Web Site:"
                lVND_PROD_CAT.Text = "Product Cat.:"
                lVND_BLACKLIST.Text = "BlackList:"
                lVND_PERF_QC.Text = "Prefered QC:"
                lVND_PERF_DELIV.Text = "Prefered Delivery:"
                lVND_CURR.Text = "Currency:"
                lVND_PAYTERM.Text = "Payment Term:"
                lVND_REM.Text = "Remarks:"
                lbl_sys_cb.Text = "CB"
                lbl_sys_lub.Text = "LUB"
                lbl_sys_cd.Text = "CD"
                lbl_sys_lud.Text = "LUD"
                saveBtn1.Text = "Save"
                saveBtn2.Text = "Save"
                saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
                saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"

                If Session("pagemode") = "N" Then
                    'VND_CODE.Text = "[Code will be auto generated]"
                    VND_CODE_HELP.Visible = True
                Else
                    VND_CODE_HELP.Visible = False
                End If

            ElseIf Session("gLang") = "C" Then
                lheader.Text = "供應商主資料庫維護"
                lSTORER_CODE.Text = "貨主:"
                lVND_CODE.Text = "供應商代碼:"
                lVND_SHORTNAME.Text = "供應商簡稱:"
                lVND_NAME.Text = "供應商名稱:"
                lVND_NAME_CH.Text = "供應商中文名稱:"
                lVND_ADDR.Text = "地址:"
                lVND_REGION.Text = "地區:"
                lVND_AREA.Text = "區域:"
                lVND_COUNTRY.Text = "國家:"
                lVND_CONT_PER1.Text = "聯絡人 1:"
                lVND_CONT_PER2.Text = "聯絡人 2:"
                lVND_CONT_TEL1.Text = "電話:"
                lVND_CONT_TEL2.Text = "電話:"
                lVND_CONT_DEPT1.Text = "部門:"
                lVND_CONT_DEPT2.Text = "部門:"
                lVND_CONT_EMAIL1.Text = "電子郵件:"
                lVND_CONT_EMAIL2.Text = "電子郵件:"
                lVND_FAX.Text = "傳真:"
                lVND_MAIN_TEL.Text = "主要電話:"
                lVND_WEBSITE.Text = "網頁:"
                lVND_PROD_CAT.Text = "產品分类:"
                lVND_BLACKLIST.Text = "黑名單:"
                lVND_PERF_QC.Text = "優先品質管理:"
                lVND_PERF_DELIV.Text = "優先送貨:"
                lVND_CURR.Text = "貨幣:"
                lVND_PAYTERM.Text = "付款方式:"
                lVND_REM.Text = "備註:"
                lbl_sys_cb.Text = "創建者"
                lbl_sys_lub.Text = "最後更新者"
                lbl_sys_cd.Text = "創建日期"
                lbl_sys_lud.Text = "最後更新日期"
                saveBtn1.Text = "儲存"
                saveBtn2.Text = "儲存"
                saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
                saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"

                If Session("pagemode") = "N" Then
                    'VND_CODE.Text = "[代碼會自動產生]"
                    VND_CODE_HELP.Visible = True
                Else
                    VND_CODE_HELP.Visible = False
                End If
            End If

            REM **********************
            VND_NAME.CssClass = "REQUIRED"
            VND_ADDR1.CssClass = "REQUIRED"
            VND_ADDR2.CssClass = "REQUIRED"
            VND_ADDR3.CssClass = "REQUIRED"
            VND_COUNTRY.CssClass = "REQUIRED"
            'VND_CODE.CssClass = "REQUIRED"

            REM**********************
            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdownBy_ColCode(VND_STATUS, "WMS_VENDOR.VND_STATUS", Session("gLang"), , Session("gSelectLabel"))
            uiFun.load_dropdown(VND_CURR, "select CCY_CODE, CCY_NAME from WMS_CURRENCY ORDER BY 2", "CCY_CODE", "CCY_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(VND_PAYTERM, "select PAY_CODE, PAY_NAME from WMS_PAYTERMS ORDER BY 2", "PAY_CODE", "PAY_NAME", , Session("gSelectLabel"))
            'uiFun.load_dropdownBy_ColCode(VND_BLACKLIST, "WMS_VENDOR.VND_BLACKLIST", Session("gLang"))

            REM **********************

            If Session("pagemode") = "N" Then
                VND_CODE.CssClass = "REQUIRED"                
                STORER_CODE.CssClass = "REQUIRED"

                If STORER_CODE.SelectedValue = "" Then
                    STORER_CODE.SelectedValue = Session("usr_pref_storer")
                End If
            Else
                Call BindGV()
            End If

            If Session("pagemode") <> "N" And VND_CODE.Text <> "" Then
                VND_CODE.CssClass = ""
                VND_CODE.BorderWidth = 0
                VND_CODE.BackColor = Drawing.Color.Transparent
                VND_CODE.ReadOnly = True
            End If

            If Session("pagemode") = "N" Then
                VND_STATUS.SelectedValue = "ACTIVE"
            End If

            REM **********************
        End If
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""

        If VND_CODE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lVND_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lVND_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If


        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lSTORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lSTORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If VND_NAME.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lVND_NAME.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lVND_NAME.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If VND_ADDR1.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lVND_ADDR.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lVND_ADDR.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If VND_COUNTRY.Text = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lVND_COUNTRY.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lVND_COUNTRY.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If VND_CONT_EMAIL1.Text.Trim <> "" And Not gU.isValidEmail(VND_CONT_EMAIL1.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lVND_CONT_EMAIL1.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lVND_CONT_EMAIL1.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If VND_CONT_EMAIL2.Text.Trim <> "" And Not gU.isValidEmail(VND_CONT_EMAIL2.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lVND_CONT_EMAIL2.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lVND_CONT_EMAIL2.Text & "!", Session("gLang"))
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

            Dim BL_YN As String = ""

            If VND_BLACKLIST.Checked Then
                BL_YN = "Y"
            End If

            Try
                REM **********************
                REM Modify Here
                If Session("pagemode") = "N" Then

                    dupSQL = "select 1 from wms_vendor " & _
                                "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and upper(vnd_code) = upper('" & gU.dbEncode(VND_CODE.Text) & "') "

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then
                        'nextNo = DB.getDocNo("VM", gConn, transaction)
                        nextNo = VND_CODE.Text.ToUpper

                        sql_string = "insert into wms_vendor (" & _
                        "vnd_code, imp_code, storer_code, " & _
                        "vnd_status, vnd_shortname, vnd_name, " & _
                        "vnd_name_ch, vnd_addr1, vnd_addr2, " & _
                        "vnd_addr3, vnd_area, vnd_region, " & _
                        "vnd_country, vnd_cont_per1, vnd_cont_per2, " & _
                        "vnd_cont_tel1, vnd_cont_tel2, vnd_cont_dept1, " & _
                        "vnd_cont_dept2, vnd_cont_email1, vnd_cont_email2, " & _
                        "vnd_fax, vnd_main_tel, vnd_website, " & _
                        "vnd_prod_cat, vnd_blacklist, vnd_blacklist_reason, " & _
                        "vnd_perf_qc, vnd_perf_deliv, vnd_curr, " & _
                        "vnd_payterm, vnd_rem, " & _
                        "sys_cb, sys_cd, sys_lub, sys_lud) values ( " & _
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(VND_STATUS.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(VND_SHORTNAME.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_NAME.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(VND_NAME_CH.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_ADDR1.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_ADDR2.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(VND_ADDR3.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_AREA.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_REGION.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(VND_COUNTRY.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_CONT_PER1.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_CONT_PER2.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(VND_CONT_TEL1.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_CONT_TEL2.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_CONT_DEPT1.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(VND_CONT_DEPT2.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_CONT_EMAIL1.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_CONT_EMAIL2.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(VND_FAX.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_MAIN_TEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_WEBSITE.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(VND_PROD_CAT.Text)) & "," & gU.convdbNVCData(gU.dbEncode(BL_YN)) & "," & gU.convdbNVCData(gU.dbEncode(VND_BLACKLIST_REASON.Text)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(VND_PERF_QC.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_PERF_DELIV.Text)) & "," & gU.convdbNVCData(gU.dbEncode(VND_CURR.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(VND_PAYTERM.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(VND_REM.Text)) & "," & _
                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Vendor Master!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "銷售商資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_vendor set " & _
                                    "vnd_status = " & gU.convdbNVCData(gU.dbEncode(VND_STATUS.SelectedValue)) & ", " & _
                                    "vnd_shortname = " & gU.convdbNVCData(gU.dbEncode(VND_SHORTNAME.Text)) & ", " & _
                                    "vnd_name = " & gU.convdbNVCData(gU.dbEncode(VND_NAME.Text)) & ", " & _
                                    "vnd_name_ch = " & gU.convdbNVCData(gU.dbEncode(VND_NAME_CH.Text)) & ", " & _
                                    "vnd_addr1 = " & gU.convdbNVCData(gU.dbEncode(VND_ADDR1.Text)) & ", " & _
                                    "vnd_addr2 = " & gU.convdbNVCData(gU.dbEncode(VND_ADDR2.Text)) & ", " & _
                                    "vnd_addr3 = " & gU.convdbNVCData(gU.dbEncode(VND_ADDR3.Text)) & ", " & _
                                    "vnd_area = " & gU.convdbNVCData(gU.dbEncode(VND_AREA.Text)) & ", " & _
                                    "vnd_region = " & gU.convdbNVCData(gU.dbEncode(VND_REGION.Text)) & ", " & _
                                    "vnd_country = " & gU.convdbNVCData(gU.dbEncode(VND_COUNTRY.Text)) & ", " & _
                                    "vnd_cont_per1 = " & gU.convdbNVCData(gU.dbEncode(VND_CONT_PER1.Text)) & ", " & _
                                    "vnd_cont_per2 = " & gU.convdbNVCData(gU.dbEncode(VND_CONT_PER2.Text)) & ", " & _
                                    "vnd_cont_tel1 = " & gU.convdbNVCData(gU.dbEncode(VND_CONT_TEL1.Text)) & ", " & _
                                    "vnd_cont_tel2 = " & gU.convdbNVCData(gU.dbEncode(VND_CONT_TEL2.Text)) & ", " & _
                                    "vnd_cont_dept1 = " & gU.convdbNVCData(gU.dbEncode(VND_CONT_DEPT1.Text)) & ", " & _
                                    "vnd_cont_dept2 = " & gU.convdbNVCData(gU.dbEncode(VND_CONT_DEPT2.Text)) & ", " & _
                                    "vnd_cont_email1 = " & gU.convdbNVCData(gU.dbEncode(VND_CONT_EMAIL1.Text)) & ", " & _
                                    "vnd_cont_email2 = " & gU.convdbNVCData(gU.dbEncode(VND_CONT_EMAIL2.Text)) & ", " & _
                                    "vnd_fax = " & gU.convdbNVCData(gU.dbEncode(VND_FAX.Text)) & ", " & _
                                    "vnd_main_tel = " & gU.convdbNVCData(gU.dbEncode(VND_MAIN_TEL.Text)) & ", " & _
                                    "vnd_website = " & gU.convdbNVCData(gU.dbEncode(VND_WEBSITE.Text)) & ", " & _
                                    "vnd_prod_cat = " & gU.convdbNVCData(gU.dbEncode(VND_PROD_CAT.Text)) & ", " & _
                                    "vnd_blacklist = " & gU.convdbNVCData(gU.dbEncode(BL_YN)) & ", " & _
                                    "vnd_blacklist_reason = " & gU.convdbNVCData(gU.dbEncode(VND_BLACKLIST_REASON.Text)) & ", " & _
                                    "vnd_perf_qc = " & gU.convdbNVCData(gU.dbEncode(VND_PERF_QC.Text)) & ", " & _
                                    "vnd_perf_deliv = " & gU.convdbNVCData(gU.dbEncode(VND_PERF_DELIV.Text)) & ", " & _
                                    "vnd_curr = " & gU.convdbNVCData(gU.dbEncode(VND_CURR.Text)) & ", " & _
                                    "vnd_payterm = " & gU.convdbNVCData(gU.dbEncode(VND_PAYTERM.Text)) & ", " & _
                                    "vnd_rem = " & gU.convdbNVCData(gU.dbEncode(VND_REM.Text)) & ", " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                    "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    "and vnd_code = '" & gU.dbEncode(VND_CODE.Text) & "' "
                End If
                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    VND_CODE.Text = nextNo
                    VND_CODE.ForeColor = Drawing.Color.Black
                    VND_CODE.Font.Size = 10
                    VND_CODE.ReadOnly = True
                    VND_CODE.BackColor = Drawing.Color.Transparent
                    VND_CODE.BorderWidth = 0

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
        If ViewState("vnd_code") <> "" Then
            pk_code = ViewState("vnd_code")
            storerCode = ViewState("storer_code")
        Else
            pk_code = Server.UrlDecode(Request("vnd_code"))
            storerCode = Server.UrlDecode(Request("storer_code"))

            ViewState("vnd_code") = pk_code
            ViewState("storer_code") = storerCode
        End If
        REM **********************

        VND_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            'CO_CODE.ForeColor = Drawing.Color.Red
            VND_STATUS.SelectedValue = "ACTIVE"
            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header

            SQLString = "SELECT * FROM WMS_VENDOR WHERE VND_CODE = '" & _
                        gU.dbEncode(pk_code) & "' AND STORER_CODE = '" & gU.dbEncode(storerCode) & "' AND IMP_CODE = '" & Session("IMP_CODE") & "'"

            SQLString = SQLString & " " & WhereStr

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                VND_CODE.Text = dt.Rows(0).Item("VND_CODE").ToString

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
                End If

                VND_STATUS.SelectedValue = gU.decodeNull(dt.Rows(0).Item("vnd_status").ToString, "ACTIVE")
                VND_SHORTNAME.Text = dt.Rows(0).Item("vnd_shortname").ToString
                VND_NAME.Text = dt.Rows(0).Item("vnd_name").ToString
                VND_NAME_CH.Text = dt.Rows(0).Item("vnd_name_ch").ToString
                VND_ADDR1.Text = dt.Rows(0).Item("vnd_addr1").ToString
                VND_ADDR2.Text = dt.Rows(0).Item("vnd_addr2").ToString
                VND_ADDR3.Text = dt.Rows(0).Item("vnd_addr3").ToString
                VND_REGION.Text = dt.Rows(0).Item("vnd_region").ToString
                VND_AREA.Text = dt.Rows(0).Item("vnd_area").ToString
                VND_COUNTRY.Text = dt.Rows(0).Item("vnd_country").ToString
                VND_CONT_PER1.Text = dt.Rows(0).Item("vnd_cont_per1").ToString
                VND_CONT_PER2.Text = dt.Rows(0).Item("vnd_cont_per2").ToString
                VND_CONT_TEL1.Text = dt.Rows(0).Item("vnd_cont_tel1").ToString
                VND_CONT_TEL2.Text = dt.Rows(0).Item("vnd_cont_tel2").ToString
                VND_CONT_DEPT1.Text = dt.Rows(0).Item("vnd_cont_dept1").ToString
                VND_CONT_DEPT2.Text = dt.Rows(0).Item("vnd_cont_dept2").ToString
                VND_CONT_EMAIL1.Text = dt.Rows(0).Item("vnd_cont_email1").ToString
                VND_CONT_EMAIL2.Text = dt.Rows(0).Item("vnd_cont_email2").ToString
                VND_FAX.Text = dt.Rows(0).Item("vnd_fax").ToString
                VND_MAIN_TEL.Text = dt.Rows(0).Item("vnd_main_tel").ToString
                VND_WEBSITE.Text = dt.Rows(0).Item("vnd_website").ToString
                VND_PROD_CAT.Text = dt.Rows(0).Item("vnd_prod_cat").ToString

                If dt.Rows(0).Item("vnd_blacklist").ToString.Trim.ToUpper = "Y" Then
                    VND_BLACKLIST.Checked = True
                End If

                VND_BLACKLIST_REASON.Text = dt.Rows(0).Item("vnd_blacklist_reason").ToString
                VND_PERF_QC.Text = dt.Rows(0).Item("vnd_perf_qc").ToString
                VND_PERF_DELIV.Text = dt.Rows(0).Item("vnd_perf_deliv").ToString
                VND_CURR.SelectedValue = dt.Rows(0).Item("vnd_curr").ToString
                VND_PAYTERM.SelectedValue = dt.Rows(0).Item("vnd_payterm").ToString
                VND_REM.Text = dt.Rows(0).Item("vnd_rem").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)
            End If
        End If
    End Sub
End Class
