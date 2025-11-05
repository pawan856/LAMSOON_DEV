Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class UserMaster
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As AccessRightUtils
    Private usr_type As String = ""
    Protected fun_code As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        usr_type = Request("user_type")
        user_type.Value = usr_type

        REM ****************************
        REM Modify Access Right Here
        If usr_type = "S" Then
            fun_code = "ADM_UP"
        ElseIf usr_type = "T" Then
            fun_code = "ADM_SUP"
        ElseIf usr_type = "C" Then
            fun_code = "ADM_CUP"
        ElseIf usr_type = "V" Then
            fun_code = "ADM_VUP"
        Else
            Response.Write("<script language='javascript'>alert('Unknown user type!');</script>")
            Response.End()
        End If

        ar = New AccessRightUtils(fun_code, Session("usr_id"), Me)

        ar.hideForm(Me)
        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            If usr_type = "S" Then
                lheader.Text = "Staff User Maintenance"
            ElseIf usr_type = "T" Then
                lheader.Text = "Storer User Maintenance"
            ElseIf usr_type = "C" Then
                lheader.Text = "Customer User Maintenance"
            ElseIf usr_type = "V" Then
                lheader.Text = "Vendor User Maintenance"
            End If
            lusr_id.Text = "User ID"
            lusr_status.Text = "Status"
            lusr_pwd.Text = "Password"
            lusr_cus_code.Text = "Customer Code"
            lusr_manager.Text = "Manager"
            lusr_storer_code.Text = "Storer"
            lusr_pref_storer.Text = "Prefer Storer"
            lusr_vnd_no.Text = "Vendor"
            lusr_group.Text = "Group"
            lusr_dob.Text = "D.O.B."
            lusr_start_date.Text = "Start Date"
            lusr_last_salary.Text = "Last Salary"
            lusr_current_salary.Text = "Current Salary"
            lusr_fname.Text = "First Name"
            lusr_sname.Text = "Surname"
            lusr_nickname.Text = "Nick Name"
            lusr_addr.Text = "Address"
            lusr_region.Text = "Region"
            lusr_area.Text = "Area"
            lusr_country.Text = "Country"
            lusr_home_tel.Text = "Home Tel."
            lusr_mobile_tel.Text = "Mobile No."
            lusr_email.Text = "Email"
            lusr_title.Text = "Title"
            lusr_work_title.Text = "Position"
            lusr_department.Text = "Department"
            lusr_responsibility.Text = "Responsibility"
            lusr_remarks.Text = "Remarks"
            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"

        ElseIf Session("gLang") = "C" Then
            If usr_type = "S" Then
                lheader.Text = "系統使用者維護"
            ElseIf usr_type = "T" Then
                lheader.Text = "貨主使用者維護"
            ElseIf usr_type = "C" Then
                lheader.Text = "客戶使用者維護"
            ElseIf usr_type = "V" Then
                lheader.Text = "供應商使用者維護"
            End If
            lusr_id.Text = "使用者代號"
            lusr_status.Text = "狀態"
            lusr_pwd.Text = "密碼"
            lusr_cus_code.Text = "客戶代碼"
            lusr_manager.Text = "主管"
            lusr_storer_code.Text = "貨主"
            lusr_pref_storer.Text = "貨主"
            lusr_vnd_no.Text = "供應商"
            lusr_group.Text = "團隊"
            lusr_dob.Text = "出生日期"
            lusr_start_date.Text = "開始日期"
            lusr_last_salary.Text = "最終薪金"
            lusr_current_salary.Text = "目前薪金"
            lusr_fname.Text = "名"
            lusr_sname.Text = "姓"
            lusr_nickname.Text = "別名"
            lusr_addr.Text = "地址"
            lusr_region.Text = "地域"
            lusr_area.Text = "區域"
            lusr_country.Text = "國家"
            lusr_home_tel.Text = "住宅電話"
            lusr_mobile_tel.Text = "手機號碼"
            lusr_email.Text = "電子郵箱"
            lusr_title.Text = "頭銜"
            lusr_work_title.Text = "職位"
            lusr_department.Text = "部門"
            lusr_responsibility.Text = "职責"
            lusr_remarks.Text = "備注"
            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "上次更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "上次更新日期"
            saveBtn1.Text = "保存"
            saveBtn2.Text = "保存"
            saveBtn1.OnClientClick = "return confirm(""确定保存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""确定保存資料?"");"

        End If

        REM **********************

        REM **********************
        REM Additional CSS

        If usr_type = "S" Then
            'usr_manager.CssClass = "REQUIRED"
            tr_storer_code.Visible = False
            tr_vnd_no.Visible = False
        ElseIf usr_type = "T" Then
            usr_storer_code.CssClass = "REQUIRED"
            tr_storer_code.Visible = True
            tr_pref_storer.Visible = False
            tr_cus_code.Visible = False
            tr_vnd_no.Visible = False
        ElseIf usr_type = "C" Then
            usr_cus_code.CssClass = "REQUIRED"
            tr_pref_storer.Visible = False
            tr_storer_code.Visible = False
            tr_vnd_no.Visible = False
        ElseIf usr_type = "V" Then
            usr_vnd_no.CssClass = "REQUIRED"
            tr_pref_storer.Visible = False
            tr_cus_code.Visible = False
            tr_storer_code.Visible = False
        End If

        'usr_addr1.CssClass = "REQUIRED"
        'usr_addr2.CssClass = "REQUIRED"
        'usr_addr3.CssClass = "REQUIRED"
        usr_email.CssClass = "REQUIRED"
        'usr_home_tel.CssClass = "REQUIRED"
        'usr_mobile_tel.CssClass = "REQUIRED"
        usr_fname.CssClass = "REQUIRED"
        usr_sname.CssClass = "REQUIRED"
        usr_status.CssClass = "REQUIRED"
        REM **********************

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            REM **********************
            REM Modify Here

            REM**********************
            uiFun.load_dropdownBy_ColCode(usr_status, "WMS_USER.USR_STATUS", Session("gLang"))
            uiFun.load_dropdown(usr_cus_code, "select cus_code, cus_name from wms_customer where cus_status = 'ACTIVE'", "cus_code", "cus_code", , Session("gSelectLabel"))
            uiFun.load_dropdown(usr_storer_code, "select storer_code, sto_shortname as sto_name from wms_storer where sto_status = 'ACTIVE' order by 2", "storer_code", "sto_name", , Session("gSelectLabel"), , False)
            uiFun.load_dropdown(usr_pref_storer, "select storer_code, sto_shortname as sto_name from wms_storer where sto_status = 'ACTIVE' order by 2", "storer_code", "sto_name", , Session("gSelectLabel"), , False)
            uiFun.load_dropdown(usr_vnd_no, "select vnd_code, vnd_name from wms_vendor order by 2", "vnd_code", "vnd_name", , Session("gSelectLabel"))
            uiFun.load_dropdown(usr_manager, "select usr_id, usr_fname + ' ' + usr_sname as usr_name from wms_user", "usr_id", "usr_name", , Session("gSelectLabel"))

            REM **********************

            If Session("pagemode") = "N" Then
                usr_id.ReadOnly = False
                usr_pref_storer.SelectedIndex = 0
            Else
                Dim SQLString As String = ""
                Dim dt As New DataTable
                Dim WhereStr As String = ""

                SQLString = "SELECT usr_id, usr_status, usr_pwd, usr_cus_code, " & _
                                "Convert(varchar,usr_dob, " & gU.getConfig("DDFORMATNO") & ") as usr_dob, " & _
                                "Convert(varchar,usr_start_date, " & gU.getConfig("DDFORMATNO") & ") as usr_start_date, " & _
                                "usr_last_salary, usr_current_salary, " & _
                                "usr_fname, usr_sname, usr_nickname, usr_addr1, usr_addr2, usr_addr3, usr_region, usr_area, usr_country, " & _
                                "usr_home_tel, usr_mobile_tel, usr_email, usr_title, usr_work_title, usr_department, " & _
                                "usr_responsibility, usr_remarks, usr_manager, usr_group, " & _
                                "sys_cb, sys_lub, sys_cd, sys_lud, " & _
                                "usr_type, usr_pref_storer, usr_storer_code, usr_vnd_no, usr_pwd_lud, usr_login_yn, usr_contr_type, " & _
                                "usr_chi_name, usr_sex, usr_category, usr_salary_rank, usr_ot_cost, usr_staff_no, usr_team, " & _
                                "usr_role, usr_bonus, usr_allowance, usr_resign_date, usr_frc_chg_pwd_yn " & _
                            "from wms_user"

                If Request("usr_id") <> "" Then
                    WhereStr = " WHERE usr_id = '" & gU.dbEncode(Request("usr_id")) & "' "
                End If

                SQLString = SQLString & " " & WhereStr

                dt = gDB.getDataTable(SQLString)

                If dt.Rows.Count > 0 Then
                    usr_id.Text = dt.Rows(0).Item("usr_id").ToString
                    usr_status.SelectedValue = dt.Rows(0).Item("usr_status").ToString
                    usr_pwd.Text = dt.Rows(0).Item("usr_pwd").ToString
                    usr_cus_code.SelectedValue = dt.Rows(0).Item("usr_cus_code").ToString
                    usr_manager.SelectedValue = dt.Rows(0).Item("usr_manager").ToString
                    usr_group.Text = dt.Rows(0).Item("usr_group").ToString
                    'usr_dob.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("usr_dob").ToString)
                    'usr_start_date.Text = cU.chgToYYYYMMDD(dt.Rows(0).Item("usr_start_date").ToString)

                    usr_dob.Text = dt.Rows(0).Item("usr_dob").ToString
                    usr_start_date.Text = dt.Rows(0).Item("usr_start_date").ToString

                    usr_last_salary.Text = cU.FormatDecimalwString(dt.Rows(0).Item("usr_last_salary").ToString)
                    usr_current_salary.Text = cU.FormatDecimalwString(dt.Rows(0).Item("usr_current_salary").ToString)
                    usr_fname.Text = dt.Rows(0).Item("usr_fname").ToString
                    usr_sname.Text = dt.Rows(0).Item("usr_sname").ToString
                    usr_nickname.Text = dt.Rows(0).Item("usr_nickname").ToString
                    usr_addr1.Text = dt.Rows(0).Item("usr_addr1").ToString
                    usr_addr2.Text = dt.Rows(0).Item("usr_addr2").ToString
                    usr_addr3.Text = dt.Rows(0).Item("usr_addr3").ToString
                    usr_home_tel.Text = dt.Rows(0).Item("usr_home_tel").ToString
                    usr_mobile_tel.Text = dt.Rows(0).Item("usr_mobile_tel").ToString
                    usr_email.Text = dt.Rows(0).Item("usr_email").ToString
                    usr_title.Text = dt.Rows(0).Item("usr_title").ToString
                    usr_work_title.Text = dt.Rows(0).Item("usr_work_title").ToString
                    usr_department.Text = dt.Rows(0).Item("usr_department").ToString
                    usr_responsibility.Text = dt.Rows(0).Item("usr_responsibility").ToString
                    usr_remarks.Text = dt.Rows(0).Item("usr_remarks").ToString
                    sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                    sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                    sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                    sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)
                    usr_region.Text = dt.Rows(0).Item("usr_region").ToString
                    usr_area.Text = dt.Rows(0).Item("usr_area").ToString
                    usr_country.Text = dt.Rows(0).Item("usr_country").ToString
                    usr_storer_code.SelectedValue = dt.Rows(0).Item("usr_storer_code").ToString
                    usr_pref_storer.SelectedValue = dt.Rows(0).Item("usr_pref_storer").ToString
                    usr_vnd_no.SelectedValue = dt.Rows(0).Item("usr_vnd_no").ToString

                    usr_id.ReadOnly = True
                    usr_id.CssClass = ""
                    usr_id.BackColor = Drawing.Color.Transparent
                    usr_id.BorderWidth = 0
                End If
            End If
        End If

        If Session("pagemode") = "N" Then
            usr_id.CssClass = "REQUIRED"
            usr_pwd.CssClass = "REQUIRED"
        End If
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""

        If usr_id.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lusr_id.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lusr_id.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If usr_status.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lusr_status.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lusr_status.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If usr_fname.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lusr_fname.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lusr_fname.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If usr_sname.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lusr_sname.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lusr_sname.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        'If usr_addr1.Text.Trim = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lusr_addr.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lusr_addr.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        'If usr_mobile_tel.Text.Trim = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lusr_mobile_tel.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lusr_mobile_tel.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'Else
        '    If Not gU.checkPhoneNumber(usr_mobile_tel.Text) Then
        '        If Session("gLang") = "E" Then
        '            uiFun.displayMsg(Me, "", lusr_mobile_tel.Text & " invalid phone number!", Session("gLang"))
        '        Else
        '            uiFun.displayMsg(Me, "", lusr_mobile_tel.Text & "無效!", Session("gLang"))
        '        End If

        '        Return False
        '    End If
        'End If

        If usr_email.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lusr_email.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lusr_email.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        Else
            If Not gU.isValidEmail(usr_email.Text) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", lusr_email.Text & " is invalid!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", lusr_email.Text & "無效!", Session("gLang"))
                End If
                Return False
            End If

        End If

        If usr_home_tel.Text <> "" Then
            If Not gU.checkPhoneNumber(usr_home_tel.Text) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", lusr_home_tel.Text & " invalid phone number!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", lusr_home_tel.Text & "無效!", Session("gLang"))
                End If

                Return False
            End If
        End If

        If Session("pagemode") = "N" Then
            If usr_pwd.Text.Trim = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", lusr_pwd.Text & " cannot be empty!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", lusr_pwd.Text & "不能空白!", Session("gLang"))
                End If
                Return False

            Else
                'If Not gU.isValidPWD(usr_pwd.Text.Trim) OrElse (Len(usr_pwd.Text.Trim) < 8 OrElse Len(usr_pwd.Text.Trim) > 12) Then

                '    If Session("gLang") = "E" Then
                '        uiFun.displayMsg(Me, "", "Password requirements:\n1. use at least 8 digits (max. 12 digits)\n2. the use of both upper- and lower-case letters (case sensitivity)\n3. inclusion of one or more numerical digits", Session("gLang"))
                '    Else
                '        uiFun.displayMsg(Me, "", "Password requirements:\n1. use at least 8 digits (max. 12 digits)\n2. the use of both upper- and lower-case letters (case sensitivity)\n3. inclusion of one or more numerical digits", Session("gLang"))
                '    End If
                '    Return False

                'End If


            End If
        End If

        If usr_type = "S" Then
            'If usr_manager.Text.Trim = "" Then
            '    If Session("gLang") = "E" Then
            '        uiFun.displayMsg(Me, "", lusr_manager.Text & " cannot be empty!", Session("gLang"))
            '    Else
            '        uiFun.displayMsg(Me, "", lusr_manager.Text & "不能空白!", Session("gLang"))
            '    End If
            '    Return False
            'End If
        ElseIf usr_type = "T" Then
            If usr_storer_code.SelectedValue.Trim = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", lusr_storer_code.Text & " cannot be empty!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", lusr_storer_code.Text & "不能空白!", Session("gLang"))
                End If
                Return False
            End If
        ElseIf usr_type = "C" Then
            If usr_cus_code.SelectedValue.Trim = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", lusr_cus_code.Text & " cannot be empty!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", lusr_cus_code.Text & "不能空白!", Session("gLang"))
                End If
                Return False
            End If
        ElseIf usr_type = "V" Then
            If usr_vnd_no.SelectedValue.Trim = "" Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", lusr_vnd_no.Text & " cannot be empty!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", lusr_vnd_no.Text & "不能空白!", Session("gLang"))
                End If
                Return False
            End If
        End If

        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim gConn As SqlConnection
        'Dim nextNo As String

        If validateAll() Then
            Dim Key As String = "OTS_" & UCase(usr_id.Text.Trim)
            Dim p As Encryption.Symmetric.Provider = Encryption.Symmetric.Provider.Rijndael
            Dim sym As New Encryption.Symmetric(p)
            sym.Key.Text = Key

            Dim storer, pref_storer As String

            'Dim encryptedData As Encryption.Data
            'encryptedData = sym.Encrypt(New Encryption.Data(usr_pwd.Text))

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                REM **********************
                REM Modify Here

                If usr_type = "T" Then

                    storer = usr_storer_code.SelectedValue
                    pref_storer = usr_storer_code.SelectedValue

                ElseIf usr_type = "C" Then

                    storer = gU.decodeNullOrEmpty(DB.getValueFromSQL("select storer_code from wms_customer where IMP_CODE='" & Session("IMP_CODE") & "' AND cus_code='" & gU.dbEncode(usr_cus_code.SelectedValue) & "'", gConn, transaction), "")
                    pref_storer = storer

                Else
                    storer = usr_storer_code.SelectedValue
                    pref_storer = usr_pref_storer.SelectedValue

                End If

                If Session("pagemode") = "N" Then

                    sql_string = "insert into wms_user (" & _
                                    "usr_id, usr_status, usr_pwd, usr_cus_code, usr_manager, usr_group, usr_dob, usr_start_date, usr_last_salary, " & _
                                    "usr_current_salary, usr_fname, usr_sname, usr_nickname, usr_addr1, usr_addr2, usr_addr3, usr_region, usr_area, " & _
                                    "usr_country, usr_home_tel, usr_mobile_tel, usr_email, usr_title, usr_work_title, usr_department, " & _
                                    "usr_responsibility, usr_remarks, usr_storer_code, usr_vnd_no, usr_type, usr_pref_storer, " & _
                                    "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                "values ( " & _
                                "'" & gU.dbEncode(UCase(usr_id.Text.Trim)) & "'," & _
                                "'" & gU.dbEncode(usr_status.SelectedValue) & "'," & _
                                "''," & _
                                "'" & gU.dbEncode(usr_cus_code.SelectedValue) & "'," & _
                                "'" & gU.dbEncode(usr_manager.SelectedValue) & "'," & _
                                "'" & gU.dbEncode(usr_group.Text) & "'," & _
                                gU.convdbDate(gU.dbEncode(usr_dob.Text)) & "," & _
                                gU.convdbDate(gU.dbEncode(usr_start_date.Text)) & "," & _
                                gU.dbEncode(gU.decodeNullOrEmpty(usr_last_salary.Text, "0")) & "," & _
                                gU.dbEncode(gU.decodeNullOrEmpty(usr_current_salary.Text, "0")) & "," & _
                                "N'" & gU.dbEncode(usr_fname.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_sname.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_nickname.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_addr1.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_addr2.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_addr3.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_region.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_area.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_country.Text) & "'," & _
                                "'" & gU.dbEncode(usr_home_tel.Text) & "'," & _
                                "'" & gU.dbEncode(usr_mobile_tel.Text) & "'," & _
                                "'" & gU.dbEncode(usr_email.Text) & "'," & _
                                "'" & gU.dbEncode(usr_title.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_work_title.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_department.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_responsibility.Text) & "'," & _
                                "N'" & gU.dbEncode(usr_remarks.Text) & "'," & _
                                "'" & gU.dbEncode(storer) & "'," & _
                                "'" & gU.dbEncode(usr_vnd_no.SelectedValue) & "'," & _
                                "'" & gU.dbEncode(usr_type) & "'," & _
                                "'" & gU.dbEncode(pref_storer) & "'," & _
                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"

                    Dim newuser As MembershipUser = Membership.GetUser(usr_id.Text.Trim)

                    If newuser Is Nothing Then
                        Membership.CreateUser(usr_id.Text.Trim, usr_pwd.Text, usr_email.Text)
                        If usr_status.SelectedValue = "I" Then
                            newuser = Membership.GetUser(usr_id.Text.Trim)
                            If Not newuser Is Nothing Then
                                newuser.IsApproved = False
                            End If
                        End If
                    End If

                Else
                    REM char "N" is use for update Unicode
                    sql_string = "update wms_user set " & _
                                "usr_status= '" & gU.dbEncode(usr_status.SelectedValue) & "',"

                    sql_string = sql_string & "usr_cus_code= '" & gU.dbEncode(usr_cus_code.SelectedValue) & "'," & _
                                "usr_manager= '" & gU.dbEncode(usr_manager.SelectedValue) & "'," & _
                                "usr_group= '" & gU.dbEncode(usr_group.Text) & "'," & _
                                "usr_dob= " & gU.convdbDate(gU.dbEncode(usr_dob.Text)) & "," & _
                                "usr_start_date= " & gU.convdbDate(gU.dbEncode(usr_start_date.Text)) & "," & _
                                "usr_last_salary= " & gU.dbEncode(gU.decodeNullOrEmpty(usr_last_salary.Text, "0")) & "," & _
                                "usr_current_salary= " & gU.dbEncode(gU.decodeNullOrEmpty(usr_current_salary.Text, "0")) & "," & _
                                "usr_fname= N'" & gU.dbEncode(usr_fname.Text) & "'," & _
                                "usr_sname= N'" & gU.dbEncode(usr_sname.Text) & "'," & _
                                "usr_nickname= N'" & gU.dbEncode(usr_nickname.Text) & "'," & _
                                "usr_addr1= N'" & gU.dbEncode(usr_addr1.Text) & "'," & _
                                "usr_addr2= N'" & gU.dbEncode(usr_addr2.Text) & "'," & _
                                "usr_addr3= N'" & gU.dbEncode(usr_addr3.Text) & "'," & _
                                "usr_region= N'" & gU.dbEncode(usr_region.Text) & "'," & _
                                "usr_area= N'" & gU.dbEncode(usr_area.Text) & "'," & _
                                "usr_country= N'" & gU.dbEncode(usr_country.Text) & "'," & _
                                "usr_home_tel= '" & gU.dbEncode(usr_home_tel.Text) & "'," & _
                                "usr_mobile_tel= '" & gU.dbEncode(usr_mobile_tel.Text) & "'," & _
                                "usr_email= '" & gU.dbEncode(usr_email.Text) & "'," & _
                                "usr_title= N'" & gU.dbEncode(usr_title.Text) & "'," & _
                                "usr_work_title= N'" & gU.dbEncode(usr_work_title.Text) & "'," & _
                                "usr_department= N'" & gU.dbEncode(usr_department.Text) & "'," & _
                                "usr_responsibility= N'" & gU.dbEncode(usr_responsibility.Text) & "'," & _
                                "usr_remarks= N'" & gU.dbEncode(usr_remarks.Text) & "'," & _
                                "usr_storer_code= '" & gU.dbEncode(storer) & "'," & _
                                "usr_vnd_no= '" & gU.dbEncode(usr_vnd_no.SelectedValue) & "'," & _
                                "usr_pref_storer= '" & gU.dbEncode(pref_storer) & "'," & _
                                "sys_lub = '" & Session("usr_id") & "', " & _
                                "sys_lud = Getdate() " & _
                                "where usr_id = '" & gU.dbEncode(usr_id.Text.Trim) & "' "

                    If usr_pwd.Text <> "" Then
                        ' sql_string = sql_string & "usr_pwd= '" & gU.dbEncode(gU.decodeNull(encryptedData.Base64, "")) & "',"
                        Dim Upuser As MembershipUser = Membership.GetUser(usr_id.Text.Trim)

                        If Not Upuser Is Nothing Then

                            Upuser.ChangePassword(Upuser.ResetPassword, usr_pwd.Text.ToString.Trim)


                            If usr_status.SelectedValue = "I" Then
                                Upuser.IsApproved = False
                            Else
                                Upuser.UnlockUser()
                                Upuser.IsApproved = True
                            End If

                            Membership.UpdateUser(Upuser)
                        Else
                            'Membership.CreateUser(uid, user_pwd.Text.ToString.Trim, user_email.Text.ToString.Trim)
                        End If

                    End If

                End If
                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    usr_id.ReadOnly = True
                    usr_id.CssClass.Remove(0)
                    usr_id.BackColor = Drawing.Color.Transparent
                    usr_id.BorderWidth = 0
                End If
                uiFun.displayMsg(Me, "1001", "", Session("gLang"))
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
End Class
