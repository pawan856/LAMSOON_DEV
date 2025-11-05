Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization

Partial Class MASTER_CM_CustomerMaster
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As AccessRightUtils
    Private thumb As ThumbGenerator
    Private DDFORMAT As String = gU.getConfig("DDFORMATNO")
    Private DDFormat2 As String = gU.getConfig("DDFORMAT2")

    Private PHY_PS_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("PHY_PS_DIR")
    Private SYSP_TEMP_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("SYSP_TEMP_DIR")

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
                lheader.Text = "Customer Master Maintenance"
                lbl_CUS_CODE.Text = "Customer Code:"
                lbl_STORER_CODE.Text = "Storer Code:"
                lbl_CUS_STATUS.Text = "Status:"
                lbl_CUS_SHORTNAME.Text = "Customer Short Name:"
                lbl_CUS_NAME.Text = "Customer Name:"
                lbl_CUS_NAME_CH.Text = "Customer Name in Chinese:"

                lbl_CUS_LOGO_L.Text = "Large Logo:"
                lbl_CUS_LOGO_S.Text = "Small Logo:"

                lbl_CUS_ADDR_DEL.Text = "Delivery Address:"
                lbl_CUS_REGION_DEL.Text = "Delivery Region:"
                lbl_CUS_AREA_DEL.Text = "Delivery Area:"
                lbl_CUS_COUNTRY_DEL.Text = "Delivery Country:"

                lbl_CUS_ADDR_BILL.Text = "Billing Address:"
                lbl_CUS_REGION_BILL.Text = "Billing Region:"
                lbl_CUS_AREA_BILL.Text = "Billing Area:"
                lbl_CUS_COUNTRY_BILL.Text = "Billing Country:"

                lbl_CUS_CONT_PER_GEN.Text = "General Manager:"
                lbl_CUS_CONT_TEL_GEN.Text = "Tel.:"
                lbl_CUS_CONT_DEPT_GEN.Text = "Department:"
                lbl_CUS_CONT_EMAIL_GEN.Text = "Email:"

                lbl_CUS_CONT_PER_ACC.Text = "Accountant:"
                lbl_CUS_CONT_TEL_ACC.Text = "Tel.:"
                lbl_CUS_CONT_DEPT_ACC.Text = "Department:"
                lbl_CUS_CONT_EMAIL_ACC.Text = "Email:"

                lbl_CUS_CONT_PER_ORD.Text = "Contact Person(Order):"
                lbl_CUS_CONT_TEL_ORD.Text = "Tel.:"
                lbl_CUS_CONT_DEPT_ORD.Text = "Department:"
                lbl_CUS_CONT_EMAIL_ORD.Text = "Email:"

                lbl_CUS_CONT_PER_LOG.Text = "Contact Person(Logistic):"
                lbl_CUS_CONT_TEL_LOG.Text = "Tel.:"
                lbl_CUS_CONT_DEPT_LOG.Text = "Department:"
                lbl_CUS_CONT_EMAIL_LOG.Text = "Email:"

                lbl_CUS_FAX.Text = "Fax:"
                lbl_CUS_MAIN_TEL.Text = "Main Tel.:"
                lbl_CUS_MAIN_EMAIL.Text = "Main Email:"
                lbl_CUS_WEBSITE.Text = "Web Site:"

                lbl_CUS_CURR.Text = "Currency:"
                lbl_CUS_PAYTERM.Text = "Payment Term:"
                lbl_CUS_REM.Text = "Remarks:"
                lbl_sys_cb.Text = "CB"
                lbl_sys_lub.Text = "LUB"
                lbl_sys_cd.Text = "CD"
                lbl_sys_lud.Text = "LUD"
                saveBtn1.Text = "Save"
                saveBtn2.Text = "Save"
                saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
                saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"

                If Session("pagemode") = "N" Then
                    'CUS_CODE.Text = "[Code will be auto generated]"
                End If

            ElseIf Session("gLang") = "C" Then
                lheader.Text = "客戶主資料庫維護"
                lbl_CUS_CODE.Text = "客戶代碼:"
                lbl_STORER_CODE.Text = "貨主代碼:"
                lbl_CUS_STATUS.Text = "狀態:"
                lbl_CUS_SHORTNAME.Text = "客戶簡稱:"
                lbl_CUS_NAME.Text = "客戶名稱:"
                lbl_CUS_NAME_CH.Text = "客戶中文名稱:"

                lbl_CUS_LOGO_L.Text = "大標誌:"
                lbl_CUS_LOGO_S.Text = "小標誌:"

                lbl_CUS_ADDR_DEL.Text = "送貨地址:"
                lbl_CUS_REGION_DEL.Text = "送貨地區:"
                lbl_CUS_AREA_DEL.Text = "送貨區域:"
                lbl_CUS_COUNTRY_DEL.Text = "送貨國家:"

                lbl_CUS_ADDR_BILL.Text = "帳單地址:"
                lbl_CUS_REGION_BILL.Text = "帳單地區:"
                lbl_CUS_AREA_BILL.Text = "帳單區域:"
                lbl_CUS_COUNTRY_BILL.Text = "帳單國家:"

                lbl_CUS_CONT_PER_GEN.Text = "總經理:"
                lbl_CUS_CONT_TEL_GEN.Text = "電話:"
                lbl_CUS_CONT_DEPT_GEN.Text = "部門:"
                lbl_CUS_CONT_EMAIL_GEN.Text = "電子郵件:"

                lbl_CUS_CONT_PER_ACC.Text = "會計:"
                lbl_CUS_CONT_TEL_ACC.Text = "電話:"
                lbl_CUS_CONT_DEPT_ACC.Text = "部門:"
                lbl_CUS_CONT_EMAIL_ACC.Text = "電子郵件:"

                lbl_CUS_CONT_PER_ORD.Text = "聯絡人(Order):"
                lbl_CUS_CONT_TEL_ORD.Text = "電話:"
                lbl_CUS_CONT_DEPT_ORD.Text = "部門:"
                lbl_CUS_CONT_EMAIL_ORD.Text = "電子郵件:"

                lbl_CUS_CONT_PER_LOG.Text = "聯絡人(物流):"
                lbl_CUS_CONT_TEL_LOG.Text = "電話:"
                lbl_CUS_CONT_DEPT_LOG.Text = "部門:"
                lbl_CUS_CONT_EMAIL_LOG.Text = "電子郵件:"

                lbl_CUS_FAX.Text = "傳真:"
                lbl_CUS_MAIN_TEL.Text = "主要電話:"
                lbl_CUS_MAIN_EMAIL.Text = "主要電子郵件:"
                lbl_CUS_WEBSITE.Text = "網頁:"

                lbl_CUS_CURR.Text = "貨幣:"
                lbl_CUS_PAYTERM.Text = "付款方式:"
                lbl_CUS_REM.Text = "備註:"
                lbl_sys_cb.Text = "創建者"
                lbl_sys_lub.Text = "最後更新者"
                lbl_sys_cd.Text = "創建日期"
                lbl_sys_lud.Text = "最後更新日期"
                saveBtn1.Text = "儲存"
                saveBtn2.Text = "儲存"
                saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
                saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"
                If Session("pagemode") = "N" Then
                    'CUS_CODE.Text = "[代碼會自動產生]"
                End If
            End If

            REM **********************
            CUS_NAME.CssClass = "REQUIRED"
            'CUS_ADDR1.CssClass = "REQUIRED"
            'CUS_ADDR2.CssClass = "REQUIRED"
            'CUS_ADDR3.CssClass = "REQUIRED"
            'CUS_COUNTRY.CssClass = "REQUIRED"

            REM**********************

            uiFun.load_dropdown(CUS_CURR, "select CCY_CODE, CCY_NAME from WMS_CURRENCY ORDER BY 2", "CCY_CODE", "CCY_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdownBy_ColCode(CUS_STATUS, "WMS_CUSTOMER.CUS_STATUS", Session("gLang"), , Session("gSelectLabel"))
            uiFun.load_dropdown(CUS_PAYTERM, "select PAY_CODE, PAY_NAME from WMS_PAYTERMS ORDER BY 2", "PAY_CODE", "PAY_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME  from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))

            REM **********************

            ViewState("image_name_L") = ""
            ViewState("image_name_S") = ""

            ViewState("remove_imageL") = False
            ViewState("remove_imageS") = False

            Session.Remove(CUS_LOGO_L_UPLOAD.ClientID)
            Session.Remove(CUS_LOGO_S_UPLOAD.ClientID)

            If Session("pagemode") = "N" Then
                'CUS_CODE.ForeColor = Drawing.Color.Red
                STORER_CODE.CssClass = "REQUIRED"
                CUS_CODE.CssClass = "REQUIRED"

                CUS_LOGO_L_UPLOAD.Visible = True
                CUS_LOGO_S_UPLOAD.Visible = True

                CUS_LOGO_L.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=&refresh=true&width=320&height=120&userid=" & Session("usr_id")
                CUS_LOGO_S.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=&refresh=true&width=150&height=70&userid=" & Session("usr_id")

                If STORER_CODE.SelectedValue = "" Then
                    STORER_CODE.SelectedValue = Session("usr_pref_storer")
                End If

            Else
                Call BindGV()
            End If

            If Session("pagemode") <> "N" And CUS_CODE.Text <> "" Then
                CUS_CODE.BorderWidth = 0
                CUS_CODE.BackColor = Drawing.Color.Transparent
                CUS_CODE.ReadOnly = True
            End If

            If Session("pagemode") = "N" Then
                CUS_STATUS.SelectedValue = "ACTIVE"
            End If

            REM **********************
        Else
            If CUS_LOGO_L_UPLOAD.HasFile Then
                Session(CUS_LOGO_L_UPLOAD.ClientID) = CUS_LOGO_L_UPLOAD.PostedFile
                CUS_LOGO_L_UPLOAD_lit.Text = CUS_LOGO_L_UPLOAD.PostedFile.FileName & "<br />"
                CUS_LOGO_L_UPLOAD.Visible = False
                CUS_LOGO_L_UPLOAD_lit.Visible = True
                CUS_LOGO_L_EDIT.Visible = True
            End If

            If CUS_LOGO_S_UPLOAD.HasFile Then
                Session(CUS_LOGO_S_UPLOAD.ClientID) = CUS_LOGO_S_UPLOAD.PostedFile
                CUS_LOGO_S_UPLOAD_lit.Text = CUS_LOGO_S_UPLOAD.PostedFile.FileName
                CUS_LOGO_S_UPLOAD.Visible = False
                CUS_LOGO_S_UPLOAD_lit.Visible = True
                CUS_LOGO_S_EDIT.Visible = True
            End If
        End If
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""

        If CUS_CODE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_CUS_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_CUS_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If CUS_NAME.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_CUS_NAME.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_CUS_NAME.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        'If CUS_ADDR1.Text.Trim = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_CUS_ADDR.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_CUS_ADDR.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        'If CUS_COUNTRY.Text = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_CUS_COUNTRY.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_CUS_COUNTRY.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        If CUS_CONT_EMAIL_GEN.Text.Trim <> "" And Not gU.isValidEmail(CUS_CONT_EMAIL_GEN.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_CUS_CONT_EMAIL_GEN.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_CUS_CONT_EMAIL_GEN.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If CUS_CONT_EMAIL_ACC.Text.Trim <> "" And Not gU.isValidEmail(CUS_CONT_EMAIL_ACC.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_CUS_CONT_EMAIL_ACC.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_CUS_CONT_EMAIL_ACC.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If CUS_CONT_EMAIL_ORD.Text.Trim <> "" And Not gU.isValidEmail(CUS_CONT_EMAIL_ORD.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_CUS_CONT_EMAIL_ORD.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_CUS_CONT_EMAIL_ORD.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If CUS_CONT_EMAIL_LOG.Text.Trim <> "" And Not gU.isValidEmail(CUS_CONT_EMAIL_LOG.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_CUS_CONT_EMAIL_LOG.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_CUS_CONT_EMAIL_LOG.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If CUS_MAIN_EMAIL.Text.Trim <> "" And Not gU.isValidEmail(CUS_MAIN_EMAIL.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_CUS_MAIN_EMAIL.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_CUS_MAIN_EMAIL.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        'Dim intValue As Integer
        'If Not Integer.TryParse(txtMaxBatch.Text, intValue) AndAlso Not intValue > 0 AndAlso intValue < 11 Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", "Please Enter a Number,Max Batch!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", "請輸入數字，最大批量!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        Return True
    End Function

    Private Sub AddNewRowToGrid()
        Dim dtCurrentTable As DataTable = TryCast(ViewState("CurrentTable"), DataTable)
        Dim drCurrentRow As DataRow = Nothing
        If Not ViewState("CurrentTable") Is Nothing Then

            If dtCurrentTable.Rows.Count > 0 Then

                Dim RowNo = 0
                For Each item As DataRow In dtCurrentTable.Rows
                    item("ITEM_CODE") = TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("gvDdlITEMCODE"), DropDownList).SelectedValue
                    item("MPL_FLAG") = If(TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("gvChkMPL"), CheckBox).Checked, 1, 0)
                    item("MSL_FLAG") = If(TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("gvChkMSL"), CheckBox).Checked, 1, 0)
                    item("MB_FLAG") = If(TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("gvChkMB"), CheckBox).Checked, 1, 0)
                    item("LOTS_CANNOT_BE_EARLIER") = If(TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("gvChkValidLot"), CheckBox).Checked, 1, 0)
                    item("MIN_PROD_LIFE") = TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("gvTxtMPLMonths"), TextBox).Text
                    item("MIN_SELF_LIFE") = TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("gvTxtMSLMonths"), TextBox).Text

                    'Dim MIN_PROD_LIFE = TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("MIN_PROD_LIFE"), TextBox).Text
                    'item("MIN_PROD_LIFE") = CDate(MIN_PROD_LIFE).ToString("dd-MMM-yyyy")
                    ''item("MIN_PROD_LIFE") = If(MIN_PROD_LIFE <> "", MIN_PROD_LIFE, "0")
                    'Dim MIN_SELF_LIFE = TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("MIN_SELF_LIFE"), TextBox).Text
                    'item("MIN_SELF_LIFE") = CDate(MIN_SELF_LIFE).ToString("dd-MMM-yyyy")
                    'item("MIN_SELF_LIFE") = Format(CDate(MIN_SELF_LIFE), "dd-MMM-yyyy").ToString()
                    'item("MIN_SELF_LIFE") = If(MIN_SELF_LIFE <> "", MIN_SELF_LIFE, "0")
                    Dim MAX_BATCHES = TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("gvTxtMBMonths"), TextBox).Text
                    item("MAX_BATCHES") = If(MAX_BATCHES <> "", MAX_BATCHES, "0")

                    Dim MAX_LOTS = TryCast(gvRules.Rows(RowNo).Cells(0).FindControl("gvTxtMaxLots"), TextBox).Text
                    item("MAX_LOTS") = If(MAX_LOTS <> "", MAX_LOTS, "0")

                    RowNo = RowNo + 1
                Next
                drCurrentRow = dtCurrentTable.NewRow()
                'drCurrentRow("SrNo") = dtCurrentTable.Rows.Count + 1
                dtCurrentTable.Rows.Add(drCurrentRow)
                ViewState("CurrentTable") = dtCurrentTable
                gvRules.DataSource = dtCurrentTable
                gvRules.DataBind()

            Else
                drCurrentRow = dtCurrentTable.NewRow()
                drCurrentRow("ITEM_CODE") = 0
                drCurrentRow("MAX_LOTS") = 0
                drCurrentRow("MPL_FLAG") = 0
                drCurrentRow("MSL_FLAG") = 0
                drCurrentRow("MB_FLAG") = 0
                drCurrentRow("LOTS_CANNOT_BE_EARLIER") = 0
                drCurrentRow("MIN_PROD_LIFE") = 0
                drCurrentRow("MIN_SELF_LIFE") = 0
                drCurrentRow("MAX_BATCHES") = 0

                dtCurrentTable.Rows.Add(drCurrentRow)
                ViewState("CurrentTable") = dtCurrentTable
                gvRules.DataSource = dtCurrentTable
                gvRules.DataBind()

            End If
        Else
            dtCurrentTable = New DataTable()
            dtCurrentTable.Columns.Add("ITEM_CODE")
            dtCurrentTable.Columns.Add("MPL_FLAG")
            dtCurrentTable.Columns.Add("MSL_FLAG")
            dtCurrentTable.Columns.Add("MB_FLAG")
            dtCurrentTable.Columns.Add("LOTS_CANNOT_BE_EARLIER")
            dtCurrentTable.Columns.Add("MIN_PROD_LIFE")
            dtCurrentTable.Columns.Add("MIN_SELF_LIFE")
            dtCurrentTable.Columns.Add("MAX_BATCHES")
            dtCurrentTable.Columns.Add("MAX_LOTS")
            drCurrentRow = dtCurrentTable.NewRow()
            'drCurrentRow("SrNo") = dtCurrentTable.Rows.Count + 1
            drCurrentRow("ITEM_CODE") = "SELECT"
            drCurrentRow("MPL_FLAG") = 0
            drCurrentRow("MSL_FLAG") = 0
            drCurrentRow("MB_FLAG") = 0
            drCurrentRow("LOTS_CANNOT_BE_EARLIER") = 0
            drCurrentRow("MIN_PROD_LIFE") = ""
            drCurrentRow("MIN_SELF_LIFE") = ""
            drCurrentRow("MAX_BATCHES") = ""
            drCurrentRow("MAX_LOTS") = ""
            dtCurrentTable.Rows.Add(drCurrentRow)
            ViewState("CurrentTable") = dtCurrentTable
            gvRules.DataSource = dtCurrentTable
            gvRules.DataBind()

        End If
        gvRules.EditIndex = 0
    End Sub

    Protected Sub saveRule(ByVal CUST_CODE As String, ByRef con As SqlConnection, ByRef pTransaction As SqlTransaction)

        'Dim con As SqlConnection = gDB.getConnection()
        Dim CmdText As String = ""
        CmdText = "Delete from WMS_CUSTOMER_RULE Where CUST_CODE='" & CUST_CODE & "'"
        Dim cmd As SqlCommand = New SqlCommand(CmdText, con)
        cmd.Transaction = pTransaction
        cmd.ExecuteNonQuery()

        Dim RowCount As Int16 = 1
        For Each item As GridViewRow In gvRules.Rows
            Dim DDL_ITEM_CODE = TryCast(item.Cells(0).FindControl("gvDdlITEMCODE"), DropDownList)
            Dim chkIsMPL = TryCast(item.Cells(0).FindControl("gvChkMPL"), CheckBox)
            Dim IsMPL = If(chkIsMPL.Checked, 1, 0)
            Dim chkIsMSL = TryCast(item.Cells(0).FindControl("gvChkMSL"), CheckBox)
            Dim IsMSL = If(chkIsMSL.Checked, 1, 0)
            Dim chkIsMB = TryCast(item.Cells(0).FindControl("gvChkMB"), CheckBox)
            Dim IsMB = If(chkIsMB.Checked, 1, 0)
            Dim chkIsValidLot = TryCast(item.Cells(0).FindControl("gvChkValidLot"), CheckBox)
            Dim IsValidLot = If(chkIsValidLot.Checked, 1, 0)
            Dim ITEM_CODE = TryCast(item.Cells(0).FindControl("ITEM_CODE"), DropDownList)
            Dim MPLMonths = TryCast(item.Cells(0).FindControl("gvtxtMPLMonths"), TextBox)
            Dim MSLMonths = TryCast(item.Cells(0).FindControl("gvtxtMSLMonths"), TextBox)
            'Dim MPLMonths = TryCast(item.Cells(0).FindControl("MIN_PROD_LIFE"), TextBox)
            'Dim MSLMonths = TryCast(item.Cells(0).FindControl("MIN_SELF_LIFE"), TextBox)
            Dim MBMonths = TryCast(item.Cells(0).FindControl("gvtxtMBMonths"), TextBox)
            Dim MAXLOTS = TryCast(item.Cells(0).FindControl("gvtxtMaxLots"), TextBox)
            CmdText = "Insert into WMS_CUSTOMER_RULE(ITEM_CODE,CUST_CODE,MPL_FLAG,MIN_PROD_LIFE,MSL_FLAG,MIN_SELF_LIFE,MB_FLAG,MAX_BATCHES,LOTS_CANNOT_BE_EARLIER,MAX_LOTS,Series)" &
                                            "Values(@ITEM_CODE,@CUST_CODE,@MPL_FLAG,@MIN_PROD_LIFE,@MSL_FLAG,@MIN_SELF_LIFE,@MB_FLAG,@MAX_BATCHES,@LOTS_CANNOT_BE_EARLIER,@MAX_LOTS,@Series)"


            cmd = New SqlCommand(CmdText, con)
            cmd.Transaction = pTransaction
            cmd.Parameters.AddWithValue("@ITEM_CODE", DDL_ITEM_CODE.SelectedValue.ToString())
            cmd.Parameters.AddWithValue("@CUST_CODE", CUST_CODE)
            cmd.Parameters.AddWithValue("@MPL_FLAG", If(chkIsMPL.Checked, 1, 0))
            cmd.Parameters.AddWithValue("@MSL_FLAG", If(chkIsMSL.Checked, 1, 0))
            cmd.Parameters.AddWithValue("@MB_FLAG", If(chkIsMB.Checked, 1, 0))
            cmd.Parameters.AddWithValue("@LOTS_CANNOT_BE_EARLIER", If(chkIsValidLot.Checked, 1, 0))
            'cmd.Parameters.AddWithValue("@MIN_PROD_LIFE", CDate(MPLMonths.Text).ToString("dd-MMM-yyyy"))
            'cmd.Parameters.AddWithValue("@MIN_SELF_LIFE", CDate(MSLMonths.Text).ToString("dd-MMM-yyyy"))
            cmd.Parameters.AddWithValue("@MIN_PROD_LIFE", If(MPLMonths.Text.Trim() <> "", MPLMonths.Text.Trim(), "0"))
            cmd.Parameters.AddWithValue("@MIN_SELF_LIFE", If(MSLMonths.Text.Trim() <> "", MSLMonths.Text.Trim(), "0"))
            cmd.Parameters.AddWithValue("@MAX_BATCHES", If(MBMonths.Text.Trim() <> "", MBMonths.Text.Trim(), "0"))
            cmd.Parameters.AddWithValue("@MAX_LOTS", If(MAXLOTS.Text.Trim() <> "", MAXLOTS.Text.Trim(), "0"))
            cmd.Parameters.AddWithValue("@Series", RowCount)
            cmd.ExecuteNonQuery()
            RowCount = RowCount + 1
        Next
    End Sub

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim gConn As SqlConnection
        Dim nextNo As String = ""
        Dim dupSQL As String
        Dim dupTbl As New DataTable

        Dim logoName_L As String = ""
        Dim logoName_S As String = ""
        Dim tmpImpCode As String = ""
        Dim tempRemovelogoName_L As Boolean = False
        Dim tempRemovelogoName_S As Boolean = False

        Dim nUPL As HttpPostedFile = Nothing
        Dim nUPS As HttpPostedFile = Nothing

        If Not Session(CUS_LOGO_L_UPLOAD.ClientID) Is Nothing Then nUPL = TryCast(Session(CUS_LOGO_L_UPLOAD.ClientID), HttpPostedFile)
        'If Not Session(itm_picture2_upload.ClientID) Is Nothing Then nUP2 = TryCast(Session(itm_picture2_upload.ClientID), HttpPostedFile)

        If Session("IMP_CODE").ToString = "" Then tmpImpCode = "W" Else tmpImpCode = Session("IMP_CODE").ToString

        Dim old_L As String = ViewState("image_name_L")
        Dim old_S As String = ViewState("image_name_S")

        Dim tmpFInfo As IO.FileInfo

        If validateAll() Then
            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction

            Try
                REM **********************
                REM Modify Here

                If Session("pagemode") = "N" Then


                    dupSQL = "Select 1 from wms_customer " &
                              "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                    "and cus_code = '" & gU.dbEncode(CUS_CODE.Text) & "' " &
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then

                        'nextNo = DB.getDocNo("CM", gConn, transaction)
                        nextNo = CUS_CODE.Text

                        If nUPL Is Nothing Then
                            If CUS_LOGO_L_UPLOAD.HasFile Then
                                tmpFInfo = New FileInfo(CUS_LOGO_L_UPLOAD.PostedFile.FileName)
                                logoName_L = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_CM_1_" & gU.decodeNullOrEmpty(nextNo, "0") & gU.decodeNullOrEmpty(STORER_CODE.SelectedValue, "0") & tmpFInfo.Extension
                            End If
                        Else
                            tmpFInfo = New FileInfo(nUPL.FileName)
                            logoName_L = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_CM_1_" & gU.decodeNullOrEmpty(nextNo, "0") & gU.decodeNullOrEmpty(STORER_CODE.SelectedValue, "0") & tmpFInfo.Extension
                        End If

                        If nUPS Is Nothing Then
                            If CUS_LOGO_S_UPLOAD.HasFile Then
                                tmpFInfo = New FileInfo(CUS_LOGO_S_UPLOAD.PostedFile.FileName)
                                logoName_S = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_CM_2_" & gU.decodeNullOrEmpty(nextNo, "0") & gU.decodeNullOrEmpty(STORER_CODE.SelectedValue, "0") & tmpFInfo.Extension
                            End If
                        Else
                            tmpFInfo = New FileInfo(nUPS.FileName)
                            logoName_S = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_CM_2_" & gU.decodeNullOrEmpty(nextNo, "0") & gU.decodeNullOrEmpty(STORER_CODE.SelectedValue, "0") & tmpFInfo.Extension
                        End If

                        sql_string = "insert into wms_customer (" &
                        "cus_code, storer_code, imp_code, CUS_status, " &
                        "cus_shortname, cus_name, cus_name_ch, " &
                        "cus_addr1_del, cus_addr2_del, cus_addr3_del, " &
                        "cus_area_del, cus_region_del, cus_country_del, " &
                        "cus_addr1_bill, cus_addr2_bill, cus_addr3_bill, " &
                        "cus_area_bill, cus_region_bill, cus_country_bill, " &
                        "cus_cont_per_gen, cus_cont_tel_gen, cus_cont_dept_gen, cus_cont_email_gen, " &
                        "cus_cont_per_acc, cus_cont_tel_acc, cus_cont_dept_acc, cus_cont_email_acc, " &
                        "cus_cont_per_ord, cus_cont_tel_ord, cus_cont_dept_ord, cus_cont_email_ord, " &
                        "cus_cont_per_log, cus_cont_tel_log, cus_cont_dept_log, cus_cont_email_log, " &
                        "cus_fax, cus_main_tel, cus_main_email, " &
                        "cus_website, cus_curr, cus_payterm, cus_rem," &
                        "cus_logo_l, cus_logo_s," &
                        "sys_cb, sys_cd, sys_lub, sys_lud,MPL_FLAG,MIN_PROD_LIFE,MSL_FLAG,MIN_SELF_LIFE,MB_FLAG,MAX_BATCHES,MAX_LOTS,LOTS_CANNOT_BE_EARLIER) values ( " &
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(CUS_STATUS.SelectedValue)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_SHORTNAME.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_NAME.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_NAME_CH.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_ADDR1_DEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_ADDR2_DEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_ADDR3_DEL.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_AREA_DEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_REGION_DEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_COUNTRY_DEL.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_ADDR1_BILL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_ADDR2_BILL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_ADDR3_BILL.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_AREA_BILL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_REGION_BILL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_COUNTRY_BILL.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_CONT_PER_GEN.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_TEL_GEN.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_DEPT_GEN.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_EMAIL_GEN.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_CONT_PER_ACC.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_TEL_ACC.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_DEPT_ACC.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_EMAIL_ACC.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_CONT_PER_ORD.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_TEL_ORD.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_DEPT_ORD.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_EMAIL_ORD.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_CONT_PER_LOG.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_TEL_LOG.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_DEPT_LOG.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CONT_EMAIL_LOG.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_FAX.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_MAIN_TEL.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_MAIN_EMAIL.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_WEBSITE.Text)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CURR.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_PAYTERM.SelectedValue)) & "," &
                        gU.convdbNVCData(gU.dbEncode(CUS_REM.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(logoName_L)) & "," & gU.convdbNVCData(gU.dbEncode(logoName_S)) & "," &
                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()," &
                        gU.convdbNVCData(gU.dbEncode(If(chkMPL.Checked, "1", "0"))) & "," &
                        gU.convdbNVCData(gU.dbEncode(txtMPLMonths.Text.Trim())) & "," &
                        gU.convdbNVCData(gU.dbEncode(If(chkMSL.Checked, "1", "0"))) & "," &
                        gU.convdbNVCData(gU.dbEncode(txtMSLMonths.Text.Trim())) & "," &
                        gU.convdbNVCData(gU.dbEncode(If(chkMaxBatches.Checked, "1", "0"))) & "," &
                        gU.convdbNVCData(gU.dbEncode(txtMaxBatch.Text.Trim())) & "," &
                        gU.convdbNVCData(gU.dbEncode(txtMaxLots.Text.Trim())) & "," &
                        gU.convdbNVCData(gU.dbEncode(If(chkValidLot.Checked, "1", "0"))) & ")"

                        ' gU.convdbDate(gU.dbEncode(MIN_PROD_LIFE.Text)) & ", " &
                        ' gU.convdbDate(gU.dbEncode(MIN_SELF_LIFE.Text)) & ", " &
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Customer Master!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "客戶資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    If nUPL Is Nothing Then
                        If CUS_LOGO_L_UPLOAD.HasFile Then
                            tmpFInfo = New FileInfo(CUS_LOGO_L_UPLOAD.PostedFile.FileName)
                            logoName_L = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_CM_1_" & CUS_CODE.Text.Trim & gU.decodeNullOrEmpty(STORER_CODE.SelectedValue, "0") & tmpFInfo.Extension
                        End If
                    Else
                        tmpFInfo = New FileInfo(nUPL.FileName)
                        logoName_L = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_CM_1_" & CUS_CODE.Text.Trim & gU.decodeNullOrEmpty(STORER_CODE.SelectedValue, "0") & tmpFInfo.Extension
                    End If

                    If nUPS Is Nothing Then
                        If CUS_LOGO_S_UPLOAD.HasFile Then
                            tmpFInfo = New FileInfo(CUS_LOGO_S_UPLOAD.PostedFile.FileName)
                            logoName_S = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_CM_2_" & CUS_CODE.Text.Trim & gU.decodeNullOrEmpty(STORER_CODE.SelectedValue, "0") & tmpFInfo.Extension
                        End If
                    Else
                        tmpFInfo = New FileInfo(nUPS.FileName)
                        logoName_S = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_CM_2_" & CUS_CODE.Text.Trim & gU.decodeNullOrEmpty(STORER_CODE.SelectedValue, "0") & tmpFInfo.Extension
                    End If

                    REM char "N" is use for update Unicode
                    sql_string = "update wms_customer set " &
                                    "cus_status = " & gU.convdbNVCData(gU.dbEncode(CUS_STATUS.SelectedValue)) & ", " &
                                    "cus_shortname = " & gU.convdbNVCData(gU.dbEncode(CUS_SHORTNAME.Text)) & ", " &
                                    "cus_name = " & gU.convdbNVCData(gU.dbEncode(CUS_NAME.Text)) & ", " &
                                    "cus_name_ch = " & gU.convdbNVCData(gU.dbEncode(CUS_NAME_CH.Text)) & ", " &
                                    "cus_addr1_del = " & gU.convdbNVCData(gU.dbEncode(CUS_ADDR1_DEL.Text)) & ", " &
                                    "cus_addr2_del = " & gU.convdbNVCData(gU.dbEncode(CUS_ADDR2_DEL.Text)) & ", " &
                                    "cus_addr3_del = " & gU.convdbNVCData(gU.dbEncode(CUS_ADDR3_DEL.Text)) & ", " &
                                    "cus_area_del = " & gU.convdbNVCData(gU.dbEncode(CUS_AREA_DEL.Text)) & ", " &
                                    "cus_region_del = " & gU.convdbNVCData(gU.dbEncode(CUS_REGION_DEL.Text)) & ", " &
                                    "cus_country_del = " & gU.convdbNVCData(gU.dbEncode(CUS_COUNTRY_DEL.Text)) & ", " &
                                    "cus_addr1_bill = " & gU.convdbNVCData(gU.dbEncode(CUS_ADDR1_BILL.Text)) & ", " &
                                    "cus_addr2_bill = " & gU.convdbNVCData(gU.dbEncode(CUS_ADDR2_BILL.Text)) & ", " &
                                    "cus_addr3_bill = " & gU.convdbNVCData(gU.dbEncode(CUS_ADDR3_BILL.Text)) & ", " &
                                    "cus_area_bill = " & gU.convdbNVCData(gU.dbEncode(CUS_AREA_BILL.Text)) & ", " &
                                    "cus_region_bill = " & gU.convdbNVCData(gU.dbEncode(CUS_REGION_BILL.Text)) & ", " &
                                    "cus_country_bill = " & gU.convdbNVCData(gU.dbEncode(CUS_COUNTRY_BILL.Text)) & ", " &
                                    "cus_cont_per_gen = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_PER_GEN.Text)) & ", " &
                                    "cus_cont_tel_gen = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_TEL_GEN.Text)) & ", " &
                                    "cus_cont_dept_gen = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_DEPT_GEN.Text)) & ", " &
                                    "cus_cont_email_gen = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_EMAIL_GEN.Text)) & ", " &
                                    "cus_cont_per_acc = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_PER_ACC.Text)) & ", " &
                                    "cus_cont_tel_acc = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_TEL_ACC.Text)) & ", " &
                                    "cus_cont_dept_acc = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_DEPT_ACC.Text)) & ", " &
                                    "cus_cont_email_acc = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_EMAIL_ACC.Text)) & ", " &
                                    "cus_cont_per_ord = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_PER_ORD.Text)) & ", " &
                                    "cus_cont_tel_ord = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_TEL_ORD.Text)) & ", " &
                                    "cus_cont_dept_ord = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_DEPT_ORD.Text)) & ", " &
                                    "cus_cont_email_ord = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_EMAIL_ORD.Text)) & ", " &
                                    "cus_cont_per_log = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_PER_LOG.Text)) & ", " &
                                    "cus_cont_tel_log = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_TEL_LOG.Text)) & ", " &
                                    "cus_cont_dept_log = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_DEPT_LOG.Text)) & ", " &
                                    "cus_cont_email_log = " & gU.convdbNVCData(gU.dbEncode(CUS_CONT_EMAIL_LOG.Text)) & ", " &
                                    "cus_fax = " & gU.convdbNVCData(gU.dbEncode(CUS_FAX.Text)) & ", " &
                                    "cus_main_tel = " & gU.convdbNVCData(gU.dbEncode(CUS_MAIN_TEL.Text)) & ", " &
                                    "cus_main_email = " & gU.convdbNVCData(gU.dbEncode(CUS_MAIN_EMAIL.Text)) & ", " &
                                    "cus_website = " & gU.convdbNVCData(gU.dbEncode(CUS_WEBSITE.Text)) & ", " &
                                    "cus_curr = " & gU.convdbNVCData(gU.dbEncode(CUS_CURR.SelectedValue)) & ", " &
                                    "cus_payterm = " & gU.convdbNVCData(gU.dbEncode(CUS_PAYTERM.SelectedValue)) & ", " &
                                    "MPL_FLAG = " & gU.convdbNVCData(gU.dbEncode(If(chkMPL.Checked, "1", "0"))) & ", " &
                                    "MSL_FLAG = " & gU.convdbNVCData(gU.dbEncode(If(chkMSL.Checked, "1", "0"))) & ", " &
                                      "MIN_PROD_LIFE = " & gU.convdbNVCData(gU.dbEncode(txtMPLMonths.Text.Trim())) & ", " &
                                    "MIN_SELF_LIFE = " & gU.convdbNVCData(gU.dbEncode(txtMSLMonths.Text.Trim())) & ", " &
                                    "MB_FLAG = " & gU.convdbNVCData(gU.dbEncode(If(chkMaxBatches.Checked, "1", "0"))) & ", " &
                                    "LOTS_CANNOT_BE_EARLIER = " & gU.convdbNVCData(gU.dbEncode(If(chkValidLot.Checked, "1", "0"))) & ", " &
                                    "MAX_BATCHES = " & gU.convdbNVCData(gU.dbEncode(txtMaxBatch.Text.Trim())) & ", " &
                                    "MAX_LOTS = " & gU.convdbNVCData(gU.dbEncode(txtMaxLots.Text.Trim())) & ", " &
                                    "cus_rem = " & gU.convdbNVCData(gU.dbEncode(CUS_REM.Text)) & ","

                    '"MIN_SELF_LIFE = " & gU.convdbDate(gU.dbEncode(MIN_SELF_LIFE.Text)) & ", " &
                    '"MIN_PROD_LIFE = " & gU.convdbDate(gU.dbEncode(MIN_PROD_LIFE.Text)) & ", " &

                    If Not ViewState("remove_imageL") Is Nothing Then
                        tempRemovelogoName_L = ViewState("remove_imageL")
                    Else
                        tempRemovelogoName_L = False
                    End If

                    If Not ViewState("remove_imageS") Is Nothing Then
                        tempRemovelogoName_S = ViewState("remove_imageS")
                    Else
                        tempRemovelogoName_S = False
                    End If

                    If Not tempRemovelogoName_L Then
                        If logoName_L <> "" Then sql_string = sql_string & "cus_logo_l = " & gU.convdbNVCData(gU.dbEncode(logoName_L)) & ", "
                    Else
                        sql_string = sql_string & "cus_logo_l = null, "
                    End If

                    If Not tempRemovelogoName_S Then
                        If logoName_S <> "" Then sql_string = sql_string & "cus_logo_s = " & gU.convdbNVCData(gU.dbEncode(logoName_S)) & ", "
                    Else
                        sql_string = sql_string & "cus_logo_s = null, "
                    End If

                    sql_string = sql_string & "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                    "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                    "and cus_code = '" & gU.dbEncode(CUS_CODE.Text) & "' " &
                                    "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                Call saveRule(CUS_CODE.Text, gConn, transaction)
                If Not transaction Is Nothing Then
                    transaction.Commit()
                    transaction = Nothing

                    Dim isSavePic1 As Boolean = False
                    Dim isSavePic2 As Boolean = False

                    If Not IO.Directory.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE")) Then IO.Directory.CreateDirectory(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE"))
                    If Not IO.Directory.Exists(SYSP_TEMP_DIR) Then IO.Directory.CreateDirectory(SYSP_TEMP_DIR)

                    If Not tempRemovelogoName_L Then
                        If nUPL Is Nothing Then
                            If CUS_LOGO_L_UPLOAD.HasFile Then
                                CUS_LOGO_L_UPLOAD.SaveAs(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & logoName_L)
                                isSavePic1 = True
                            End If
                        Else
                            nUPL.SaveAs(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & logoName_L)
                            isSavePic1 = True
                        End If

                        If isSavePic1 Then
                            If IO.File.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_L) Then

                                thumb = New ThumbGenerator
                                thumb.SetParams(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_L, 150, 180)
                                Dim unID As String = thumb.GetUniqueThumbName
                                Cache.Remove(unID)

                                IO.File.Delete(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_L)
                            End If

                            IO.File.Move(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & logoName_L, PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & logoName_L)
                        End If

                    Else
                        removeImage(old_L)
                        ViewState("remove_imageL") = False
                    End If

                    If Not tempRemovelogoName_S Then
                        If nUPS Is Nothing Then
                            If CUS_LOGO_S_UPLOAD.HasFile Then
                                CUS_LOGO_S_UPLOAD.SaveAs(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & logoName_S)
                                isSavePic2 = True
                            End If
                        Else
                            nUPS.SaveAs(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & logoName_S)
                            isSavePic2 = True
                        End If

                        If isSavePic2 Then
                            If IO.File.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_S) Then

                                thumb = New ThumbGenerator
                                thumb.SetParams(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_S, 150, 180)
                                Dim unID As String = thumb.GetUniqueThumbName
                                Cache.Remove(unID)

                                IO.File.Delete(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_S)
                            End If

                            IO.File.Move(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & logoName_S, PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & logoName_S)
                        End If
                    Else
                        removeImage(old_S)
                        ViewState("remove_imageS") = False
                    End If

                    Session.Remove(CUS_LOGO_L_UPLOAD.ClientID)
                    Session.Remove(CUS_LOGO_S_UPLOAD.ClientID)

                    CUS_LOGO_L_UPLOAD.Visible = True
                    CUS_LOGO_S_UPLOAD.Visible = True

                    CUS_LOGO_L_UPLOAD_lit.Visible = False
                    CUS_LOGO_S_UPLOAD_lit.Visible = False

                End If

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    CUS_CODE.Text = nextNo
                    CUS_CODE.ForeColor = Drawing.Color.Black
                    CUS_CODE.Font.Size = 10

                    CUS_CODE.BackColor = Drawing.Color.Transparent
                    CUS_CODE.BorderWidth = 0
                    CUS_CODE.ReadOnly = True

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
        Dim pk_code, storercode As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("cus_code") <> "" Then
            pk_code = ViewState("cus_code")
            storercode = ViewState("storer_code")
        Else
            pk_code = Server.UrlDecode(Request("cus_code"))
            storercode = Server.UrlDecode(Request("storer_code"))

            ViewState("cus_code") = pk_code
            ViewState("storer_code") = storercode
        End If
        REM **********************

        'CUS_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here

            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header

            SQLString = "SELECT * FROM wms_customer WHERE CUS_CODE = '" & gU.dbEncode(pk_code) & "' AND STORER_CODE = '" & gU.dbEncode(storercode) & "' AND IMP_CODE = '" & Session("IMP_CODE") & "'"

            SQLString = SQLString & " " & WhereStr

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                CUS_CODE.Text = dt.Rows(0).Item("cus_code").ToString

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , , True)
                End If

                CUS_STATUS.SelectedValue = gU.decodeNull(dt.Rows(0).Item("cus_status").ToString.ToUpper, "active")
                CUS_SHORTNAME.Text = dt.Rows(0).Item("cus_shortname").ToString
                CUS_NAME.Text = dt.Rows(0).Item("cus_name").ToString
                CUS_NAME_CH.Text = dt.Rows(0).Item("cus_name_ch").ToString
                CUS_ADDR1_DEL.Text = dt.Rows(0).Item("cus_addr1_del").ToString
                CUS_ADDR2_DEL.Text = dt.Rows(0).Item("cus_addr2_del").ToString
                CUS_ADDR3_DEL.Text = dt.Rows(0).Item("cus_addr3_del").ToString
                CUS_REGION_DEL.Text = dt.Rows(0).Item("cus_region_del").ToString
                CUS_AREA_DEL.Text = dt.Rows(0).Item("cus_area_del").ToString
                CUS_COUNTRY_DEL.Text = dt.Rows(0).Item("cus_country_del").ToString
                CUS_ADDR1_BILL.Text = dt.Rows(0).Item("cus_addr1_bill").ToString
                CUS_ADDR2_BILL.Text = dt.Rows(0).Item("cus_addr2_bill").ToString
                CUS_ADDR3_BILL.Text = dt.Rows(0).Item("cus_addr3_bill").ToString
                CUS_REGION_BILL.Text = dt.Rows(0).Item("cus_region_bill").ToString
                CUS_AREA_BILL.Text = dt.Rows(0).Item("cus_area_bill").ToString
                CUS_COUNTRY_BILL.Text = dt.Rows(0).Item("cus_country_bill").ToString
                CUS_CONT_PER_GEN.Text = dt.Rows(0).Item("cus_cont_per_gen").ToString
                CUS_CONT_TEL_GEN.Text = dt.Rows(0).Item("cus_cont_tel_gen").ToString
                CUS_CONT_DEPT_GEN.Text = dt.Rows(0).Item("cus_cont_dept_gen").ToString
                CUS_CONT_EMAIL_GEN.Text = dt.Rows(0).Item("cus_cont_email_gen").ToString
                CUS_CONT_PER_ACC.Text = dt.Rows(0).Item("cus_cont_per_acc").ToString
                CUS_CONT_TEL_ACC.Text = dt.Rows(0).Item("cus_cont_tel_acc").ToString
                CUS_CONT_DEPT_ACC.Text = dt.Rows(0).Item("cus_cont_dept_acc").ToString
                CUS_CONT_EMAIL_ACC.Text = dt.Rows(0).Item("cus_cont_email_acc").ToString
                CUS_CONT_PER_ORD.Text = dt.Rows(0).Item("cus_cont_per_ord").ToString
                CUS_CONT_TEL_ORD.Text = dt.Rows(0).Item("cus_cont_tel_ord").ToString
                CUS_CONT_DEPT_ORD.Text = dt.Rows(0).Item("cus_cont_dept_ord").ToString
                CUS_CONT_DEPT_ORD.Text = dt.Rows(0).Item("cus_cont_dept_ord").ToString
                CUS_CONT_PER_LOG.Text = dt.Rows(0).Item("cus_cont_per_log").ToString
                CUS_CONT_TEL_LOG.Text = dt.Rows(0).Item("cus_cont_tel_log").ToString
                CUS_CONT_DEPT_LOG.Text = dt.Rows(0).Item("cus_cont_dept_log").ToString
                CUS_CONT_EMAIL_LOG.Text = dt.Rows(0).Item("cus_cont_email_log").ToString
                CUS_FAX.Text = dt.Rows(0).Item("cus_fax").ToString
                CUS_MAIN_TEL.Text = dt.Rows(0).Item("cus_main_tel").ToString
                CUS_MAIN_EMAIL.Text = dt.Rows(0).Item("cus_main_email").ToString
                CUS_WEBSITE.Text = dt.Rows(0).Item("cus_website").ToString
                CUS_CURR.SelectedValue = dt.Rows(0).Item("cus_curr").ToString
                CUS_PAYTERM.SelectedValue = dt.Rows(0).Item("cus_payterm").ToString
                CUS_REM.Text = dt.Rows(0).Item("cus_rem").ToString
                'gvddlITEMCODE.SelectedValue = dt.Rows(0).Item("ITEM_CODE").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                txtMPLMonths.Text = cU.chgToFullDF(dt.Rows(0).Item("MIN_PROD_LIFE").ToString)
                txtMSLMonths.Text = cU.chgToFullDF(dt.Rows(0).Item("MIN_SELF_LIFE").ToString)


                'Dim strDate As String = If(dt.Rows(0).Item("MIN_PROD_LIFE") IsNot DBNull.Value, dt.Rows(0).Item("MIN_PROD_LIFE").ToString, "")
                'If Not String.IsNullOrEmpty(strDate) Then
                '    Dim d As DateTime = Convert.ToDateTime(strDate)
                '    Dim reformatted As String = d.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture)
                '    MIN_PROD_LIFE.Text = reformatted
                'End If

                'strDate = If(dt.Rows(0).Item("MIN_PROD_LIFE") IsNot DBNull.Value, dt.Rows(0).Item("MIN_SELF_LIFE").ToString, "")
                'If Not String.IsNullOrEmpty(strDate) Then
                '    Dim d As DateTime = Convert.ToDateTime(strDate)
                '    Dim reformatted As String = d.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture)
                '    MIN_SELF_LIFE.Text = reformatted
                'End If

                txtMaxBatch.Text = cU.chgToFullDF(dt.Rows(0).Item("MAX_BATCHES").ToString)
                txtMaxLots.Text = cU.chgToFullDF(dt.Rows(0).Item("MAX_LOTS").ToString)
                chkMaxBatches.Checked = If(dt.Rows(0).Item("MB_FLAG").ToString = "1", True, False)
                chkMPL.Checked = If(dt.Rows(0).Item("MPL_FLAG").ToString = "1", True, False)
                chkMSL.Checked = If(dt.Rows(0).Item("MSL_FLAG").ToString = "1", True, False)
                chkValidLot.Checked = If(dt.Rows(0).Item("LOTS_CANNOT_BE_EARLIER").ToString = "1", True, False)

                If chkMPL.Checked Then
                    txtMPLMonths.Enabled = True
                Else
                    txtMPLMonths.Enabled = False
                End If

                If chkMSL.Checked Then
                    txtMSLMonths.Enabled = True
                Else
                    txtMSLMonths.Enabled = False
                End If

                If chkMaxBatches.Checked Then
                    txtMaxBatch.Enabled = True
                Else
                    txtMaxBatch.Enabled = False
                End If

                CUS_LOGO_L.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & dt.Rows(0).Item("CUS_LOGO_L").ToString & "&refresh=true&width=320&height=120&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                CUS_LOGO_S.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & dt.Rows(0).Item("CUS_LOGO_S").ToString & "&refresh=true&width=150&height=70&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")

                If dt.Rows(0).Item("CUS_LOGO_L").ToString <> "" Then
                    CUS_LOGO_L.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & dt.Rows(0).Item("CUS_LOGO_L").ToString & "&refresh=true&ds=true&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                    CUS_LOGO_L.Target = "_new1"

                    'preview_pic1.Visible = True

                    CUS_LOGO_L_UPLOAD.Visible = False
                    CUS_LOGO_L_EDIT.Visible = True
                    CUS_LOGO_L_REMOVE.Visible = True
                    CUS_LOGO_L.Enabled = True
                Else
                    'preview_pic1.Visible = False
                    CUS_LOGO_L_UPLOAD.Visible = True
                    CUS_LOGO_L_EDIT.Visible = False
                    CUS_LOGO_L_REMOVE.Visible = False
                End If

                If dt.Rows(0).Item("CUS_LOGO_S").ToString <> "" Then
                    CUS_LOGO_S.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & dt.Rows(0).Item("CUS_LOGO_S").ToString & "&refresh=true&ds=true&Height=180&width=180&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                    CUS_LOGO_S.Target = "_new2"

                    'preview_pic2.Visible = True
                    CUS_LOGO_S_UPLOAD.Visible = False
                    CUS_LOGO_S_EDIT.Visible = True
                    CUS_LOGO_S_REMOVE.Visible = True
                    CUS_LOGO_S.Enabled = True
                Else
                    'preview_pic2.Visible = False
                    CUS_LOGO_S_UPLOAD.Visible = True
                    CUS_LOGO_S_EDIT.Visible = False
                    CUS_LOGO_S_REMOVE.Visible = False
                End If

                ViewState("image_name_L") = dt.Rows(0).Item("CUS_LOGO_L").ToString
                ViewState("image_name_S") = dt.Rows(0).Item("CUS_LOGO_S").ToString

                Dim dupSQL = "Select * from wms_customer_rule " &
                          "where cust_code = '" & gU.dbEncode(CUS_CODE.Text) & "'"

                Dim dupTbl = gDB.getDataTable(dupSQL)

                If dupTbl.Rows.Count > 0 Then
                    'gvddlITEMCODE.SelectedValue = dupTbl.Rows(0)("ITEM_CODE")
                    ViewState("CurrentTable") = dupTbl
                End If

                gvRules.DataSource = dupTbl
                gvRules.DataBind()
            End If
        End If
    End Sub

    Private Sub removeImage(ByVal FileName As String)
        If FileName <> "" Then
            Try
                If IO.File.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & FileName) Then

                    thumb = New ThumbGenerator
                    thumb.SetParams(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & FileName, 150, 180)
                    Dim unID As String = thumb.GetUniqueThumbName
                    Cache.Remove(unID)

                    IO.File.Delete(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & FileName)
                End If

                CUS_LOGO_L_REMOVE.Enabled = True
                CUS_LOGO_L_REMOVE.Visible = False

                CUS_LOGO_L_EDIT.Enabled = True
                CUS_LOGO_L_EDIT.Visible = False

            Catch ex As Exception
                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            End Try
        End If
    End Sub

    Protected Sub CUS_LOGO_L_REMOVE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CUS_LOGO_L_REMOVE.Click
        ViewState("remove_imageL") = True
        Session.Remove(CUS_LOGO_L_UPLOAD.ClientID)

        CUS_LOGO_L_REMOVE.Enabled = False
        CUS_LOGO_L_EDIT.Enabled = False
        CUS_LOGO_L.Enabled = False
        'preview_pic1.Enabled = False
        CUS_LOGO_L_UPLOAD.Visible = False
        CUS_LOGO_L_UPLOAD_lit.Text = ""
        CUS_LOGO_L_UPLOAD_lit.Visible = False
    End Sub

    Protected Sub CUS_LOGO_L_EDIT_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CUS_LOGO_L_EDIT.Click
        Session.Remove(CUS_LOGO_L_UPLOAD.ClientID)
        CUS_LOGO_L_UPLOAD.Visible = True
        CUS_LOGO_L_UPLOAD_lit.Text = ""
        CUS_LOGO_L_UPLOAD_lit.Visible = False
        CUS_LOGO_L_EDIT.Visible = False
        CUS_LOGO_L_REMOVE.Visible = False
    End Sub

    Protected Sub CUS_LOGO_S_REMOVE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CUS_LOGO_S_REMOVE.Click
        ViewState("remove_imageS") = True
        Session.Remove(CUS_LOGO_S_UPLOAD.ClientID)

        CUS_LOGO_S_REMOVE.Enabled = False
        CUS_LOGO_S_EDIT.Enabled = False
        CUS_LOGO_S.Enabled = False
        'preview_pic2.Enabled = False
        CUS_LOGO_S_UPLOAD.Visible = False
        CUS_LOGO_S_UPLOAD_lit.Text = ""
        CUS_LOGO_S_UPLOAD_lit.Visible = False
    End Sub

    Protected Sub CUS_LOGO_S_EDIT_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CUS_LOGO_S_EDIT.Click
        Session.Remove(CUS_LOGO_S_UPLOAD.ClientID)
        CUS_LOGO_S_UPLOAD.Visible = True
        CUS_LOGO_S_UPLOAD_lit.Text = ""
        CUS_LOGO_S_UPLOAD_lit.Visible = False
        CUS_LOGO_S_EDIT.Visible = False
        CUS_LOGO_S_REMOVE.Visible = False
    End Sub

    Protected Sub btnAddRow_Click(sender As Object, e As EventArgs) Handles btnAddRow.Click
        Call AddNewRowToGrid()
    End Sub

    Protected Sub gvRules_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles gvRules.RowCommand
        If (e.CommandName = "Remove") Then

            Dim RowIndex As Int16 = Int16.Parse(e.CommandArgument.ToString)
            If ViewState("CurrentTable") IsNot Nothing Then
                Dim dtCurrentTable As DataTable = TryCast(ViewState("CurrentTable"), DataTable)
                dtCurrentTable.Rows.RemoveAt(RowIndex)
                gvRules.DataSource = dtCurrentTable
                gvRules.DataBind()
                ViewState("CurrentTable") = dtCurrentTable
            End If

        End If
    End Sub

    Protected Sub chkMPL_CheckedChanged(sender As Object, e As EventArgs) Handles chkMPL.CheckedChanged
        If chkMPL.Checked Then
            txtMPLMonths.Enabled = True
            txtMPLMonths.Text = ""
        Else
            txtMPLMonths.Enabled = False
            txtMPLMonths.Text = ""
        End If
    End Sub

    Protected Sub chkMSL_CheckedChanged(sender As Object, e As EventArgs) Handles chkMSL.CheckedChanged
        If chkMSL.Checked Then
            txtMSLMonths.Enabled = True
            txtMSLMonths.Text = ""
        Else
            txtMSLMonths.Enabled = False
            txtMSLMonths.Text = ""
        End If
    End Sub

    Protected Sub chkMaxBatches_CheckedChanged(sender As Object, e As EventArgs) Handles chkMaxBatches.CheckedChanged
        If chkMaxBatches.Checked Then
            txtMaxBatch.Enabled = True
            txtMaxBatch.Text = ""
        Else
            txtMaxBatch.Enabled = False
            txtMaxBatch.Text = ""
        End If
    End Sub

    'Protected Sub DDL_ITEM_CODE_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    'chkMPL.Checked = False
    '    'MIN_PROD_LIFE.Text = ""
    '    'chkMSL.Checked = False
    '    'MIN_SELF_LIFE.Text = ""
    '    'chkMaxBatches.Checked = False
    '    'txtMaxBatch.Text = ""
    '    'chkValidLot.Checked = False

    '    Dim dupSQL = "Select * from wms_customer_rule " &
    '                          "where cust_code = '" & gU.dbEncode(CUS_CODE.Text) & "'"
    '    Dim dupTbl = gDB.getDataTable(dupSQL)

    '    'If dupTbl.Rows.Count > 0 Then
    '    '    DDL_ITEM_CODE.SelectedValue = dupTbl.Rows(0)("ITEM_CODE")
    '    ViewState("CurrentTable") = dupTbl
    '    'End If

    '    gvRules.DataSource = dupTbl
    '    gvRules.DataBind()
    'End Sub

    Protected Sub gvRules_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvRules.RowDataBound
        If e.Row.RowIndex >= 0 And e.Row.RowType = DataControlRowType.DataRow Then
            Dim drv As DataRowView = CType(e.Row.DataItem, DataRowView)
            Dim ddlITEMCODE As DropDownList = CType(e.Row.FindControl("gvDdlITEMCODE"), DropDownList)
            'Dim ddlITEMCODE As DropDownList = TryCast(gvRules.Rows(e.Row.RowIndex).FindControl("gvDdlITEMCODE"), DropDownList)
            uiFun.load_dropdown(ddlITEMCODE, "Select ITM_CODE,ITM_NAME from WMS_ITEM where IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "ITM_CODE", "ITM_NAME", , Session("gSelectLabel"))
            If drv("ITEM_CODE") IsNot DBNull.Value Then
                ddlITEMCODE.SelectedValue = drv("ITEM_CODE")
            End If
        End If
    End Sub
End Class
