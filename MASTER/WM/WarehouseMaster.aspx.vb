Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class MASTER_WM_WarehouseMaster
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As AccessRightUtils
    Private thumb As ThumbGenerator

    Private PHY_PS_DIR As String = System.Configuration.ConfigurationManager.AppSettings.Item("PHY_PS_DIR")

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
                lheader.Text = "Subinventory Master Maintenance"
                lbl_WH_CODE.Text = "Subinventory:"
                lbl_WH_MAIN_WH.Text = "Subinventory:"
                lbl_WH_OWNER.Text = "Warehouse Owner:"
                lbl_WH_NAME.Text = "Subinventory Name:"
                lbl_WH_NAME_CH.Text = "Subinventory Name in Chinese:"
                lbl_WH_ADDR.Text = "Address:"
                lbl_WH_REGION.Text = "Region:"
                lbl_WH_TYPE.Text = "Type:"
                lbl_WH_COUNTRY.Text = "Country:"

                lbl_WH_CONT_PER1.Text = "Display in Item Balance Report:"
                lbl_WH_CONT_TEL1.Text = "Tel.:"
                lbl_WH_CONT_EMAIL1.Text = "Email:"

                lbl_WH_CONT_PER2.Text = "Contact Person 2:"
                lbl_WH_CONT_TEL2.Text = "Tel.:"
                lbl_WH_CONT_EMAIL2.Text = "Email:"

                lbl_WH_GROSS_AREA.Text = "Gross Area:"
                lbl_WH_GROSS_CBM.Text = "Gross CBM:"

                lbl_img.Text = "Image:"
                lbl_WH_REM.Text = "Remarks:"

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

                lheader.Text = "子庫存 Master Maintenance"
                lbl_WH_CODE.Text = "子庫存:"
                lbl_WH_MAIN_WH.Text = "子庫存:"
                lbl_WH_OWNER.Text = "業主:"
                lbl_WH_NAME.Text = "子庫存名稱:"
                lbl_WH_NAME_CH.Text = "子庫存中文名稱:"
                lbl_WH_ADDR.Text = "地址:"
                lbl_WH_REGION.Text = "地區:"
                lbl_WH_TYPE.Text = "類型:"
                lbl_WH_COUNTRY.Text = "國家:"

                lbl_WH_CONT_PER1.Text = "聯絡人1:"
                lbl_WH_CONT_TEL1.Text = "電話:"
                lbl_WH_CONT_EMAIL1.Text = "電子郵件:"

                lbl_WH_CONT_PER2.Text = "聯絡人2:"
                lbl_WH_CONT_TEL2.Text = "電話:"
                lbl_WH_CONT_EMAIL2.Text = "電子郵件:"

                lbl_WH_GROSS_AREA.Text = "總面積:"
                lbl_WH_GROSS_CBM.Text = "總CBM:"

                lbl_img.Text = "圖片:"
                lbl_WH_REM.Text = "備註:"
                lbl_sys_cb.Text = "創建者"
                lbl_sys_lub.Text = "最後更新者"
                lbl_sys_cd.Text = "創建日期"
                lbl_sys_lud.Text = "最後更新日期"
                saveBtn1.Text = "儲存"
                saveBtn2.Text = "儲存"
                saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
                saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"

                If Session("pagemode") = "N" Then
                    'WH_CODE.Text = "[代碼會自動產生]"
                End If
            End If

            REM **********************
            WH_CODE.CssClass = "REQUIRED"
            WH_NAME.CssClass = "REQUIRED"
            WH_ADDR1.CssClass = "REQUIRED"
            WH_ADDR2.CssClass = "REQUIRED"
            WH_ADDR3.CssClass = "REQUIRED"
            'WH_COUNTRY.CssClass = "REQUIRED"
            'STORER_CODE.CssClass = "REQUIRED"

            REM**********************
            'uiFun.load_dropdown(WH_CURR, "select CCY_CODE, CCY_NAME from WMS_CURRENCY ORDER BY 2", "CCY_CODE", "CCY_NAME", , Session("gSelectLabel"))
            'uiFun.load_dropdownBy_ColCode(WH_STATUS, "WMS_STORER.WH_STATUS", Session("gLang"), , Session("gSelectLabel"))
            REM **********************
            uiFun.load_dropdown(WH_TYPE, "select WH_TYPE_CODE, WH_TYPE_DESC from WMS_WH_TYPE ORDER BY 2", "WH_TYPE_CODE", "WH_TYPE_DESC", , Session("gSelectLabel"))
            ViewState("image_name") = ""

            ViewState("remove_image") = False

            Session.Remove(wh_picture_upload.ClientID)

            If Session("pagemode") = "N" Then
                'WH_CODE.ForeColor = Drawing.Color.Red
            Else
                Call BindGV()
            End If

            If Session("pagemode") <> "N" And WH_CODE.Text <> "" Then
                WH_CODE.BorderWidth = 0
                WH_CODE.BackColor = Drawing.Color.Transparent
                WH_CODE.ReadOnly = True
            End If

            REM **********************
        Else
            If wh_picture_upload.HasFile Then
                Session(wh_picture_upload.ClientID) = wh_picture_upload.PostedFile
                wh_picture_upload_lit.Text = wh_picture_upload.PostedFile.FileName
                wh_picture_upload.Visible = False
                wh_picture_upload_lit.Visible = True
                wh_picture_edit.Visible = True
            End If
        End If
    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""

        If WH_CODE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_WH_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_WH_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        Else
            If Not gU.isAlphaNumueric(WH_CODE.Text.Trim) Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", lbl_WH_CODE.Text & " Must be alphanumeric!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", lbl_WH_CODE.Text & "必須為數字及英文字母!", Session("gLang"))
                End If
                Return False
            End If

        End If

        If WH_NAME.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_WH_NAME.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_WH_NAME.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If WH_ADDR1.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_WH_ADDR.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_WH_ADDR.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        'If WH_COUNTRY.Text = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_WH_COUNTRY.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_WH_COUNTRY.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        If WH_CONT_EMAIL1.Text.Trim <> "" And Not gU.isValidEmail(WH_CONT_EMAIL1.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_WH_CONT_EMAIL1.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_WH_CONT_EMAIL1.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If WH_CONT_EMAIL2.Text.Trim <> "" And Not gU.isValidEmail(WH_CONT_EMAIL2.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid email, " & lbl_WH_CONT_EMAIL2.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的電子郵箱, " & lbl_WH_CONT_EMAIL2.Text & "!", Session("gLang"))
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

        Dim nFileName As String = ""

        Dim nUP As HttpPostedFile = Nothing

        If Not Session(wh_picture_upload.ClientID) Is Nothing Then nUP = TryCast(Session(wh_picture_upload.ClientID), HttpPostedFile)

        Dim old_filename As String = ViewState("image_name")
        Dim tmpFInfo As IO.FileInfo
        Dim tempRemoveFile As Boolean = False
        Dim tmpImpCode As String = ""

        If validateAll() Then
            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction

            Try
                REM **********************
                REM Modify Here


                If Session("pagemode") = "N" Then
                    dupSQL = "select 1 from wms_warehouse " & _
                              "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                    "and wh_code = '" & gU.dbEncode(WH_CODE.Text) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    If dupTbl.Rows.Count = 0 Then

                        'nextNo = DB.getDocNo("WH", gConn, transaction)
                        nextNo = WH_CODE.Text

                        If nUP Is Nothing Then
                            If wh_picture_upload.HasFile Then
                                tmpFInfo = New FileInfo(wh_picture_upload.PostedFile.FileName)
                                nFileName = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_WM_" & gU.decodeNullOrEmpty(WH_CODE.Text, "0") & tmpFInfo.Extension
                            End If
                        Else
                            tmpFInfo = New FileInfo(nUP.FileName)
                            nFileName = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_WM_" & gU.decodeNullOrEmpty(WH_CODE.Text, "0") & tmpFInfo.Extension
                        End If

                        sql_string = "insert into wms_warehouse (" &
                        "wh_code, WH_MAIN_WH, imp_code, wh_owner, " &
                        "wh_name, wh_name_ch, wh_addr1, " &
                        "wh_addr2, wh_addr3, wh_type, " &
                        "wh_region, wh_country, wh_cont_per1, " &
                        "wh_cont_tel1, wh_cont_email1, wh_cont_per2, " &
                        "wh_cont_tel2, wh_cont_email2, wh_gross_area, " &
                        "wh_gross_cbm, wh_picture, wh_rem, " &
                        "sys_cb, sys_cd, sys_lub, sys_lud) values ( " &
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(WH_MAIN_WH.Text)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(WH_OWNER.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(WH_NAME.Text)) & "," & gU.convdbNVCData(gU.dbEncode(WH_NAME_CH.Text)) & "," & gU.convdbNVCData(gU.dbEncode(WH_ADDR1.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(WH_ADDR2.Text)) & "," & gU.convdbNVCData(gU.dbEncode(WH_ADDR3.Text)) & "," & gU.convdbNVCData(gU.dbEncode(WH_TYPE.SelectedValue)) & "," &
                        gU.convdbNVCData(gU.dbEncode(WH_REGION.Text)) & "," & gU.convdbNVCData(gU.dbEncode(WH_COUNTRY.Text)) & "," & gU.convdbNVCData(gU.dbEncode(WH_CONT_PER1.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(WH_CONT_TEL1.Text)) & "," & gU.convdbNVCData(gU.dbEncode(WH_CONT_EMAIL1.Text)) & "," & gU.convdbNVCData(gU.dbEncode(WH_CONT_PER2.Text)) & "," &
                        gU.convdbNVCData(gU.dbEncode(WH_CONT_TEL2.Text)) & "," & gU.convdbNVCData(gU.dbEncode(WH_CONT_EMAIL2.Text)) & "," & gU.dbEncode(gU.decodeNullOrEmpty(WH_GROSS_AREA.Text, "0")) & "," &
                        gU.dbEncode(gU.decodeNullOrEmpty(WH_GROSS_CBM.Text, "0")) & "," & gU.convdbNVCData(gU.dbEncode(wh_picture.Text)) & "," & gU.convdbNVCData(gU.dbEncode(WH_REM.Text)) & "," &
                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate())"
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Warehouse Master!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "倉庫資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM char "N" is use for update Unicode
                    If nUP Is Nothing Then
                        If wh_picture_upload.HasFile Then
                            tmpFInfo = New FileInfo(wh_picture_upload.PostedFile.FileName)
                            nFileName = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_WM_" & WH_CODE.Text.Trim & tmpFInfo.Extension
                        End If
                    Else
                        tmpFInfo = New FileInfo(nUP.FileName)
                        nFileName = tmpImpCode & "_" & Session("usr_id") & Format(Now, "yyyyMMddhhmmss") & "_WM_" & WH_CODE.Text.Trim & tmpFInfo.Extension
                    End If

                    sql_string = "update wms_warehouse set " &
                                    "wh_owner = " & gU.convdbNVCData(gU.dbEncode(WH_OWNER.Text)) & ", " &
                                    "WH_MAIN_WH = " & gU.convdbNVCData(gU.dbEncode(WH_MAIN_WH.Text)) & ", " &
                                    "wh_name = " & gU.convdbNVCData(gU.dbEncode(WH_NAME.Text)) & ", " &
                                    "wh_name_ch = " & gU.convdbNVCData(gU.dbEncode(WH_NAME_CH.Text)) & ", " &
                                    "wh_addr1 = " & gU.convdbNVCData(gU.dbEncode(WH_ADDR1.Text)) & ", " &
                                    "wh_addr2 = " & gU.convdbNVCData(gU.dbEncode(WH_ADDR2.Text)) & ", " &
                                    "wh_addr3 = " & gU.convdbNVCData(gU.dbEncode(WH_ADDR3.Text)) & ", " &
                                    "wh_type = " & gU.convdbNVCData(gU.dbEncode(WH_TYPE.SelectedValue)) & ", " &
                                    "wh_region = " & gU.convdbNVCData(gU.dbEncode(WH_REGION.Text)) & ", " &
                                    "wh_country = " & gU.convdbNVCData(gU.dbEncode(WH_COUNTRY.Text)) & ", " &
                                    "wh_cont_per1 = " & gU.convdbNVCData(gU.dbEncode(WH_CONT_PER1.Text)) & ", " &
                                    "wh_cont_tel1 = " & gU.convdbNVCData(gU.dbEncode(WH_CONT_TEL1.Text)) & ", " &
                                    "wh_cont_email1 = " & gU.convdbNVCData(gU.dbEncode(WH_CONT_EMAIL1.Text)) & ", " &
                                    "wh_cont_per2 = " & gU.convdbNVCData(gU.dbEncode(WH_CONT_PER2.Text)) & ", " &
                                    "wh_cont_tel2 = " & gU.convdbNVCData(gU.dbEncode(WH_CONT_TEL2.Text)) & ", " &
                                    "wh_cont_email2 = " & gU.convdbNVCData(gU.dbEncode(WH_CONT_EMAIL2.Text)) & ", " &
                                    "wh_gross_area = " & gU.dbEncode(gU.decodeNullOrEmpty(WH_GROSS_AREA.Text, "0")) & ", " &
                                    "wh_gross_cbm = " & gU.dbEncode(gU.decodeNullOrEmpty(WH_GROSS_CBM.Text, "0")) & ", " &
                                    "wh_rem = " & gU.convdbNVCData(gU.dbEncode(WH_REM.Text)) & ", "

                    If Not ViewState("remove_image") Is Nothing Then
                        tempRemoveFile = ViewState("remove_image")
                    Else
                        tempRemoveFile = False
                    End If

                    If Not tempRemoveFile Then
                        If nFileName <> "" Then sql_string = sql_string & "wh_picture = " & gU.convdbNVCData(gU.dbEncode(nFileName)) & ", "
                    Else
                        sql_string = sql_string & "wh_picture = null, "
                    End If

                    sql_string = sql_string & "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                    "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                    "and wh_code = '" & gU.dbEncode(WH_CODE.Text) & "' "
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                If Not transaction Is Nothing Then
                    transaction.Commit()
                    transaction = Nothing

                    Dim isSavePic As Boolean = False

                    If Not IO.Directory.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE")) Then IO.Directory.CreateDirectory(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE"))

                    If Not tempRemoveFile Then
                        If nUP Is Nothing Then
                            If wh_picture_upload.HasFile Then
                                wh_picture_upload.SaveAs(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & nFileName)
                                isSavePic = True
                            End If
                        Else
                            nUP.SaveAs(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & nFileName)
                            isSavePic = True
                        End If

                        If isSavePic Then
                            If IO.File.Exists(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_filename) Then

                                thumb = New ThumbGenerator
                                thumb.SetParams(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_filename, 150, 180)
                                Dim unID As String = thumb.GetUniqueThumbName
                                Cache.Remove(unID)

                                IO.File.Delete(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & old_filename)
                            End If

                            IO.File.Move(PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & "new_" & nFileName, PHY_PS_DIR & "\" & Session("PAGE_SESSION_MENU_CODE") & "\" & nFileName)
                        End If
                    Else
                        removeImage(old_filename)
                        ViewState("remove_image") = False
                    End If
                End If


                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'WH_CODE.Text = nextNo
                    'WH_CODE.ForeColor = Drawing.Color.Black
                    'WH_CODE.Font.Size = 10

                    WH_CODE.BackColor = Drawing.Color.Transparent
                    WH_CODE.BorderWidth = 0
                    WH_CODE.ReadOnly = True
                End If

                Session.Remove(wh_picture_upload.ClientID)
                wh_picture_upload.Visible = True
                wh_picture_upload_lit.Visible = False
                wh_picture_edit.Visible = False

                uiFun.displayMsg(Me, "1001", "", Session("gLang"))

                Call BindGV()
                REM **********************
            Catch ex As Exception
                transaction.Rollback()

                Session.Remove(wh_picture_upload.ClientID)
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
        Dim pic_preview_yn As Boolean = False

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("wh_code") <> "" Then
            pk_code = ViewState("wh_code")

        Else
            pk_code = Server.UrlDecode(Request("wh_code"))

            ViewState("wh_code") = pk_code
        End If
        REM **********************

        'WH_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            ''CO_CODE.ForeColor = Drawing.Color.Red
            'wh_picture.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=&refresh=true&Height=180&userid=" & Session("usr_id")
            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header

            SQLString = "SELECT * FROM WMS_WAREHOUSE WHERE WH_CODE = '" & gU.dbEncode(pk_code) & "' AND IMP_CODE = '" & Session("IMP_CODE") & "'"

            SQLString = SQLString & " " & WhereStr

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                WH_CODE.Text = dt.Rows(0).Item("WH_CODE").ToString

                WH_OWNER.Text = dt.Rows(0).Item("WH_OWNER").ToString
                WH_NAME.Text = dt.Rows(0).Item("WH_NAME").ToString
                WH_NAME_CH.Text = dt.Rows(0).Item("WH_NAME_CH").ToString
                WH_ADDR1.Text = dt.Rows(0).Item("WH_ADDR1").ToString
                WH_ADDR2.Text = dt.Rows(0).Item("WH_ADDR2").ToString
                WH_ADDR3.Text = dt.Rows(0).Item("WH_ADDR3").ToString
                WH_TYPE.SelectedValue = dt.Rows(0).Item("WH_TYPE").ToString
                WH_REGION.Text = dt.Rows(0).Item("WH_REGION").ToString
                WH_COUNTRY.Text = dt.Rows(0).Item("WH_COUNTRY").ToString
                WH_CONT_PER1.Text = dt.Rows(0).Item("WH_CONT_PER1").ToString
                WH_CONT_TEL1.Text = dt.Rows(0).Item("WH_CONT_TEL1").ToString
                WH_CONT_EMAIL1.Text = dt.Rows(0).Item("WH_CONT_EMAIL1").ToString
                WH_CONT_PER2.Text = dt.Rows(0).Item("WH_CONT_PER2").ToString
                WH_CONT_TEL2.Text = dt.Rows(0).Item("WH_CONT_TEL2").ToString
                WH_CONT_EMAIL2.Text = dt.Rows(0).Item("WH_CONT_EMAIL2").ToString
                WH_GROSS_AREA.Text = cU.FormatDecimalwString(dt.Rows(0).Item("WH_GROSS_AREA").ToString)
                WH_GROSS_CBM.Text = cU.FormatDecimalwString(dt.Rows(0).Item("WH_GROSS_CBM").ToString)
                WH_MAIN_WH.Text = dt.Rows(0).Item("WH_MAIN_WH").ToString
                WH_REM.Text = dt.Rows(0).Item("WH_REM").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                wh_picture.ImageUrl = "~/ThumbnailHandler.ashx?VFilePath=" & dt.Rows(0).Item("wh_picture").ToString & "&refresh=true&width=250&height=250&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")

                If dt.Rows(0).Item("wh_picture").ToString <> "" Then
                    wh_picture.NavigateUrl = "~/ThumbnailHandler.ashx?VFilePath=" & dt.Rows(0).Item("wh_picture").ToString & "&refresh=true&ds=true&code=" & Session("PAGE_SESSION_MENU_CODE") & "&userid=" & Session("usr_id")
                    wh_picture.Target = "_new1"

                    preview_pic.Visible = True
                    wh_picture_remove.Visible = True

                    pic_preview_yn = True
                Else
                    preview_pic.Visible = False
                    wh_picture_remove.Visible = False
                End If

                If pic_preview_yn Then
                    pic_tr.Visible = True
                    remove_tr.Visible = True
                Else
                    pic_tr.Visible = False
                    remove_tr.Visible = False
                End If

                ViewState("image_name") = dt.Rows(0).Item("wh_picture").ToString

                If ViewState("remove_image") Then
                    wh_picture_upload.Enabled = False
                    wh_picture_remove.Enabled = False
                    wh_picture.Enabled = False
                    preview_pic.Enabled = False
                Else
                    wh_picture_upload.Enabled = True
                    wh_picture_remove.Enabled = True
                    wh_picture.Enabled = True
                    preview_pic.Enabled = True
                End If
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

            Catch ex As Exception
                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            End Try
        End If
    End Sub

    Protected Sub wh_picture_remove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles wh_picture_remove.Click
        ViewState("remove_image") = True
        Session.Remove(wh_picture_upload.ClientID)
        wh_picture_upload.Enabled = False
        wh_picture_remove.Enabled = False
        wh_picture.Enabled = False
        preview_pic.Enabled = False
        wh_picture_upload.Visible = True
        wh_picture_edit.Visible = False
        wh_picture_upload_lit.Text = ""
        wh_picture_upload_lit.Visible = False
    End Sub

    Protected Sub wh_picture_edit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles wh_picture_edit.Click
        Session.Remove(wh_picture_upload.ClientID)
        wh_picture_upload.Visible = True
        wh_picture_upload_lit.Text = ""
        wh_picture_upload_lit.Visible = False
        wh_picture_edit.Visible = False
    End Sub
End Class
