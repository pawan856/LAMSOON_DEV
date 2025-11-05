Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Globalization
Imports System.Data.SqlClient

Partial Class TMS_ATTACH

    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As New AccessRightUtils

    Private dt As New DataTable
    Private SystemPath As String = ""
    Private MaxFileSize As Integer = 20000000 'in bytes


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SystemPath = Cache("SYSP_ATTACH_FILE_DIR")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)
        ar.hideForm(Me)

        'moduleAction = Request("moduleAction")

        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If


        If Not IsPostBack Then
            ViewState("dt") = Nothing
            ViewState("UPFL_DOC_NO") = ""
            ViewState("DOC_TYPE") = ""
            ViewState("Parent_Editmode") = ""
            ViewState("UPFL_DOC_NO") = Server.UrlDecode(Request("UPFL_DOC_NO"))
            ViewState("DOC_TYPE") = Server.UrlDecode(Request("doc_type"))
            ViewState("Parent_Editmode") = Server.UrlDecode(Request("PARENT_EDIT_MODE"))
            Call BindGV()
        Else
            dt = ViewState("dt")
        End If

        If Session("gLang") = "E" Then

            lheader.Text = "Upload Attachment"
            AddBtn1.Text = "New File"
            lbl_sys_cb.Text = "Upload By"
            lbl_sys_lub.Text = "Last Updated By"
            lbl_sys_cd.Text = "Upload Date"
            lbl_sys_lud.Text = "Last Updated Date"
            CloseBtn1.Text = "Close"
            lbl_Pnl_DOC_TYPE.Text = "Document Type:"
            lbl_Pnl_FILE_UPLOAD.Text = "File Upload:"
            lbl_Pnl_FileSize.Text = "File Size:"
            lbl_PNL_UPFL_DESC.Text = "Description:"

        ElseIf Session("gLang") = "C" Then

            lheader.Text = "上載附件"
            AddBtn1.Text = "新增"
            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            CloseBtn1.Text = "關閉"
            lbl_Pnl_DOC_TYPE.Text = "檔案類型:"
            lbl_Pnl_FILE_UPLOAD.Text = "檔案上傳:"
            lbl_Pnl_FileSize.Text = "檔案大小:"
            lbl_PNL_UPFL_DESC.Text = "檔案說明:"
        End If

        If ViewState("Parent_Editmode") = "V" Then
            AddBtn1.Enabled = False
            Pnl_UPFL_TITLE.Enabled = False
            PNL_UPFL_DESC.Enabled = False
            pnlDelBtn.Enabled = False
            pnlSaveBtn.Enabled = False
        End If

    End Sub


    Protected Sub BindGV()
        Dim sql_string As String = ""
        Dim doc_code As String = ""
        Dim doc_type As String = ""
        Dim pa1 As GlobalDBFunc.DBCmdPara
        Dim dt As New DataTable

        pa1 = New GlobalDBFunc.DBCmdPara

        If ViewState("UPFL_DOC_NO") <> "" Then
            doc_code = ViewState("UPFL_DOC_NO")
            doc_type = ViewState("DOC_TYPE")
        Else
            doc_code = Server.UrlDecode(Request("UPFL_DOC_NO"))
            doc_type = Server.UrlDecode(Request("DOC_TYPE"))
        End If

        pa1 = New GlobalDBFunc.DBCmdPara
        sql_string = " SELECT wms_upload_file.UPFL_SYS_FILENAME,  wms_upload_file.DOC_TYPE, wms_upload_file.UPFL_DOC_NO, wms_upload_file.UPFL_USER_FILENAME, " & _
                     " wms_upload_file.UPFL_TITLE,  wms_upload_file.UPFL_DESC,  wms_upload_file.UPFL_FILESIZE, CONVERT(varchar,wms_upload_file.SYS_CD,131) as SYS_CD " & _
                     " FROM wms_upload_file " & _
                     " WHERE wms_upload_file.DOC_TYPE =" & pa1.AP(doc_type) & " AND wms_upload_file.UPFL_DOC_NO=" & pa1.AP(doc_code) & _
                     " Order by SYS_CD desc"

        dt = gDB.getDataTable(sql_string, , , , pa1)
        'Response.Write(gDB.getCmdSql(sql_string, pa1))
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If

        ViewState("dt") = dt
        GridView1.DataBind()

    End Sub

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        If GridView1.DataSource IsNot Nothing Then
            Dim ddlPager As DropDownList = TryCast(GridView1.BottomPagerRow.FindControl("pager_select"), DropDownList)
            Dim tmpPagerBtn As ImageButton

            If ddlPager.Items.Count = 0 Then
                If ddlPager IsNot Nothing Then
                    For i As Integer = 1 To GridView1.PageCount
                        ddlPager.Items.Add(New ListItem(i.ToString(), i.ToString()))
                    Next
                End If
            End If

            ddlPager.SelectedValue = GridView1.PageIndex + 1

            If GridView1.PageIndex = 0 Then
                tmpPagerBtn = TryCast(GridView1.BottomPagerRow.FindControl("pager_first"), ImageButton)
                tmpPagerBtn.Enabled = False

                tmpPagerBtn = TryCast(GridView1.BottomPagerRow.FindControl("pager_previous"), ImageButton)
                tmpPagerBtn.Enabled = False
            End If

            If GridView1.PageIndex = GridView1.PageCount - 1 Then
                tmpPagerBtn = TryCast(GridView1.BottomPagerRow.FindControl("pager_next"), ImageButton)
                tmpPagerBtn.Enabled = False

                tmpPagerBtn = TryCast(GridView1.BottomPagerRow.FindControl("pager_last"), ImageButton)
                tmpPagerBtn.Enabled = False
            End If
        End If
    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        Dim ddlPager As DropDownList = TryCast(GridView1.BottomPagerRow.FindControl("pager_select"), DropDownList)

        If ddlPager.SelectedValue <> GridView1.PageIndex + 1 Then
            GridView1.PageIndex = ddlPager.SelectedValue - 1
            GridView1.DataSource = ViewState("dt")
            GridView1.DataBind()
        End If
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
        GridView1.DataSource = ViewState("dt")
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Dim v_sql As String = ""
        Dim doc_type As String = ""
        Dim doc_code As String = ""
        Dim srv_file_name As String = ""
        Dim srv_dt As New DataTable
        Dim paP As GlobalDBFunc.DBCmdPara


        If ViewState("UPFL_DOC_NO") <> "" Then
            doc_type = ViewState("DOC_TYPE")
            doc_code = ViewState("UPFL_DOC_NO")
        Else
            doc_type = Server.UrlDecode(Request("DOC_TYPE"))
            doc_code = Server.UrlDecode(Request("UPFL_DOC_NO"))
        End If

        Select Case e.CommandName
            Case "EDITInfo"
                srv_file_name = e.CommandArgument
                pnlDelBtn.Visible = True

                paP = New GlobalDBFunc.DBCmdPara
                v_sql = "select UPFL_USER_FILENAME,UPFL_SYS_FILENAME,UPFL_FileSize,UPFL_TITLE,UPFL_DESC,sys_cd,sys_cb,sys_lud,sys_lub from wms_upload_file where " & _
                        " DOC_TYPE=" & paP.AP(doc_type) & " AND UPFL_DOC_NO=" & paP.AP(doc_code) & " AND UPFL_SYS_FILENAME=" & paP.AP(srv_file_name)

                srv_dt = gDB.getDataTable(v_sql, , , , paP)

                If srv_dt.Rows.Count > 0 Then
                    Pnl_UPFL_DOC_NO.Value = doc_code
                    DSP_Pnl_DOC_TYPE.Text = doc_type
                    Pnl_DOC_TYPE.Value = doc_type
                    EDITMODE.Value = "U"
                    Pnl_FILEUPLOAD.Visible = False
                    Pnl_UPFL_USER_FILENAME.Visible = True
                    Pnl_UPFL_USER_FILENAME.Text = srv_dt.Rows(0).Item("UPFL_USER_FILENAME").ToString
                    Pnl_UPFL_SYS_FILENAME.Value = srv_dt.Rows(0).Item("UPFL_SYS_FILENAME").ToString
                    Pnl_UPFL_TITLE.Text = srv_dt.Rows(0).Item("UPFL_TITLE").ToString
                    Pnl_UPFL_FILESIZE.Text = srv_dt.Rows(0).Item("UPFL_FILESIZE").ToString
                    PNL_UPFL_DESC.Text = srv_dt.Rows(0).Item("UPFL_DESC").ToString

                    sys_cb.Text = srv_dt.Rows(0).Item("sys_cb").ToString
                    sys_lub.Text = srv_dt.Rows(0).Item("sys_lub").ToString
                    sys_cd.Text = cU.chgToFullDF(srv_dt.Rows(0).Item("sys_cd").ToString)
                    sys_lud.Text = cU.chgToFullDF(srv_dt.Rows(0).Item("sys_lud").ToString)
                End If

                btnAmend_ModalPopupExtender.Show()

            Case "DLFile"
                Dim index = e.CommandArgument
                Call GetDownloadFile(CType(GridView1.Rows(index).FindControl("UPFL_SYS_FILENAME"), HiddenField).Value.ToString.Trim, CType(GridView1.Rows(index).FindControl("UPFL_USER_FILENAME"), LinkButton).Text.ToString.Trim)

            Case "DelFile"
                Call Deleted_File(e.CommandArgument)

        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("SYS_CD"), Label).Text = DataBinder.Eval(e.Row.DataItem, "SYS_CD").ToString.Trim
                CType(e.Row.FindControl("UPFL_TITLE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "UPFL_TITLE").ToString.Trim
                CType(e.Row.FindControl("UPFL_USER_FILENAME"), LinkButton).Text = DataBinder.Eval(e.Row.DataItem, "UPFL_USER_FILENAME").ToString.Trim
                CType(e.Row.FindControl("UPFL_USER_FILENAME"), LinkButton).CommandArgument = DataBinder.Eval(e.Row.DataItem, "UPFL_SYS_FILENAME").ToString.Trim
                CType(e.Row.FindControl("UPFL_SYS_FILENAME"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "UPFL_SYS_FILENAME").ToString.Trim
                CType(e.Row.FindControl("UPFL_FILESIZE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "UPFL_FILESIZE").ToString.Trim & "KB"
                CType(e.Row.FindControl("UPFL_DESC"), Label).Text = DataBinder.Eval(e.Row.DataItem, "UPFL_DESC").ToString.Trim
                CType(e.Row.FindControl("btnDL"), Button).CommandArgument = e.Row.RowIndex
                CType(e.Row.FindControl("btnDel"), Button).CommandArgument = DataBinder.Eval(e.Row.DataItem, "UPFL_SYS_FILENAME").ToString.Trim
                If ViewState("Parent_Editmode") = "V" Then
                    CType(e.Row.FindControl("btnDEL"), Button).Enabled = False
                End If

        End Select
    End Sub

    Protected Sub AddBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AddBtn1.Click
        Dim doc_code As String = ""
        Dim doc_type As String = ""

        If ViewState("UPFL_DOC_NO") <> "" Then
            doc_type = ViewState("DOC_TYPE")
            doc_code = ViewState("UPFL_DOC_NO")
        Else
            doc_type = Server.UrlDecode(Request("DOC_TYPE"))
            doc_code = Server.UrlDecode(Request("UPFL_DOC_NO"))
        End If

        pnlDelBtn.Visible = False
        Pnl_UPFL_DOC_NO.Value = doc_code
        DSP_Pnl_DOC_TYPE.Text = doc_type
        Pnl_DOC_TYPE.Value = doc_type
        Pnl_FILEUPLOAD.Visible = True
        Pnl_UPFL_USER_FILENAME.Visible = False
        Pnl_UPFL_SYS_FILENAME.Value = ""
        Pnl_UPFL_TITLE.Text = ""
        Pnl_UPFL_FILESIZE.Text = ""
        PNL_UPFL_DESC.Text = ""
        EDITMODE.Value = "N"
        sys_cb.Text = ""
        sys_lub.Text = ""
        sys_cd.Text = ""
        sys_lud.Text = ""
        btnAmend_ModalPopupExtender.Show()

    End Sub

    Protected Function ValidatePanel() As Boolean
        If EDITMODE.Value = "N" Then
            If Not Pnl_FILEUPLOAD.HasFile Then
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "Please Select a file for uploading!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "請選擇上傳檔案!", Session("gLang"))
                End If
                Return False

            Else
                If Pnl_FILEUPLOAD.PostedFile.ContentLength > MaxFileSize Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "File Size Cannot Exceed " & MaxFileSize / 1000 & "KB!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "請選擇上傳檔案!", Session("gLang"))
                    End If
                    Return False
                End If


                If Pnl_FILEUPLOAD.PostedFile.FileName.Length > 128 Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "File Name too long!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "檔案名稱過長!", Session("gLang"))
                    End If
                    Return False
                End If
            End If
        End If

        If PNL_UPFL_DESC.Text.Length > 500 Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "File Desc. too long!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "檔案描述過長!", Session("gLang"))
            End If
            Return False
        End If

        Return True
    End Function

    Protected Sub pnlSaveBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlSaveBtn.Click

        Call SavePanel()

    End Sub


    Protected Sub SavePanel()

        Dim v_sql As String = ""
        Dim sno_sql As String = ""

        Dim new_file_name As String = ""

        Dim doc_type As String = ""
        Dim doc_code As String = ""
        Dim cnn As SqlConnection
        Dim transaction As SqlTransaction
        Dim paP As GlobalDBFunc.DBCmdPara
        paP = New GlobalDBFunc.DBCmdPara

        Dim mode As String = EDITMODE.Value

        If ViewState("UPFL_DOC_NO") <> "" Then
            doc_type = ViewState("DOC_TYPE")
            doc_code = ViewState("UPFL_DOC_NO")
        Else
            doc_type = Server.UrlDecode(Request("DOC_TYPE"))
            doc_code = Server.UrlDecode(Request("UPFL_DOC_NO"))
        End If


        If ValidatePanel() Then

            cnn = gDB.getConnection()
            transaction = cnn.BeginTransaction()
            Try
                Select Case mode

                    Case "N"

                        Dim upfile As String = Pnl_FILEUPLOAD.FileName
                        Dim uextention As String = System.IO.Path.GetExtension(upfile)

                        sno_sql = "select convert(varchar, getdate(),112) + replace(convert(varchar, getdate(),108),':','')  + replicate('0',3 - len(convert(varchar,next value for intf_log_seq))) + convert(varchar,next value for intf_log_seq) as value "
                        new_file_name = DB.getValueFromSQL(sno_sql) & uextention


                        Dim savepath As String = SystemPath & "/" & new_file_name
                        If File.Exists(savepath) Then
                            While File.Exists(savepath)
                                sno_sql = "select convert(varchar, getdate(),112) + replace(convert(varchar, getdate(),108),':','')  + replicate('0',3 - len(convert(varchar,next value for intf_log_seq))) + convert(varchar,next value for intf_log_seq) as value; "

                                new_file_name = DB.getValueFromSQL(sno_sql) & uextention
                                savepath = SystemPath & "/" & new_file_name
                            End While
                        End If

                        Pnl_FILEUPLOAD.SaveAs(savepath)

                        paP = New GlobalDBFunc.DBCmdPara
                        v_sql = "Insert into wms_upload_file (UPFL_SYS_FILENAME ,UPFL_TYPE, DOC_TYPE,FUN_CODE, UPFL_DOC_NO,UPFL_USER_FILENAME, UPFL_TITLE,UPFL_VERSION, " & _
                                " UPFL_DESC,UPFL_REMARKS,UPFL_FILESIZE,UPFL_STATUS,SYS_LUB,SYS_LUD,SYS_CB,SYS_CD ) values (" & _
                                paP.AP(new_file_name) & ",'A', " & paP.AP(doc_type) & ", " & paP.AP(Session("PAGE_SESSION_MENU_CODE")) & ", " & paP.AP(doc_code) & ", " & paP.AP(Pnl_FILEUPLOAD.FileName, SqlDbType.NVarChar) & ", " & _
                                paP.AP(Pnl_UPFL_TITLE.Text.ToString.Trim, SqlDbType.NVarChar) & ", '', " & paP.AP(PNL_UPFL_DESC.Text.ToString.Trim, SqlDbType.NVarChar) & ",'', " & paP.AP(Pnl_FILEUPLOAD.PostedFile.ContentLength / 1000) & ", 'A', " & _
                                paP.AP(Session("usr_id")) & ", Getdate(), " & paP.AP(Session("usr_id")) & ", Getdate()" & _
                                ") "

                    Case "U"
                        paP = New GlobalDBFunc.DBCmdPara

                        v_sql = "Update wms_upload_file set " & _
                                " UPFL_TITLE=" & paP.AP(Pnl_UPFL_TITLE.Text, SqlDbType.NVarChar) & ", " & _
                                " UPFL_DESC=" & paP.AP(PNL_UPFL_DESC.Text, SqlDbType.NVarChar) & ", " & _
                                " sys_lud=Getdate(), SYS_LUB=" & paP.AP(Session("usr_id")) & _
                                " Where UPFL_DOC_NO=" & paP.AP(Pnl_UPFL_DOC_NO.Value) & _
                                " AND DOC_TYPE=" & paP.AP(Pnl_DOC_TYPE.Value) & _
                                " AND UPFL_SYS_FILENAME=" & paP.AP(Pnl_UPFL_SYS_FILENAME.Value)

                End Select

                If v_sql <> "" Then

                    gDB.amendData(v_sql, cnn, transaction, paP)
                    uiFun.displayMsg(Me, "", "File info Saved Successfully!", Session("gLang"))
                    transaction.Commit()
                End If


                Call BindGV()
            Catch ex As Exception
                If Not transaction Is Nothing Then
                    transaction.Rollback()
                    transaction = Nothing
                End If
                'Throw ex
                Response.Write(ex.Message)
                Response.Write("<BR>" & gDB.getCmdSql(v_sql, paP))
            Finally
                If cnn IsNot Nothing Then
                    If cnn.State = ConnectionState.Open Then
                        cnn.Close()
                        cnn.Dispose()
                    End If
                End If
            End Try

        Else

            btnAmend_ModalPopupExtender.Show()
        End If

    End Sub

    Protected Sub pnlDelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlDelBtn.Click
        Call Deleted_File(Pnl_UPFL_SYS_FILENAME.Value)
    End Sub


    Public Sub Deleted_File(ByVal FileName As String)
        Dim v_sql As String = ""
        Dim doc_code, doc_type, f_name As String
        Dim cnn As SqlConnection
        Dim transaction As SqlTransaction

        Dim paP As GlobalDBFunc.DBCmdPara
        paP = New GlobalDBFunc.DBCmdPara

        If ViewState("UPFL_DOC_NO") <> "" Then
            doc_type = ViewState("DOC_TYPE")
            doc_code = ViewState("UPFL_DOC_NO")
        Else
            doc_type = Server.UrlDecode(Request("DOC_TYPE"))
            doc_code = Server.UrlDecode(Request("UPFL_DOC_NO"))
        End If

        f_name = FileName

        cnn = gDB.getConnection()
        transaction = cnn.BeginTransaction()

        Try


            paP = New GlobalDBFunc.DBCmdPara
            v_sql = "Delete from wms_upload_file where DOC_TYPE=" & paP.AP(doc_type) & " AND UPFL_DOC_NO=" & paP.AP(doc_code) & " AND UPFL_SYS_FILENAME=" & paP.AP(f_name)
            gDB.amendData(v_sql, cnn, transaction, paP)


            If File.Exists(SystemPath & "\" & f_name) Then
                File.Delete(SystemPath & "\" & f_name)
            End If


            uiFun.displayMsg(Me, "", "File Deleted Successfully!", Session("gLang"))
            transaction.Commit()

            Call BindGV()

        Catch ex As Exception
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If
            'Throw ex
            Response.Write(ex.Message)
            Response.Write("<BR>" & gDB.getCmdSql(v_sql, paP))
        Finally
            If cnn IsNot Nothing Then
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                    cnn.Dispose()
                End If
            End If
        End Try

    End Sub


    Private Sub GetDownloadFile(ByVal SysFilename As String, ByVal UserFileName As String)
        Dim fullFilePath As String = SystemPath & "\" & SysFilename
        Dim uFileName As String = UserFileName

        If fullFilePath <> "" Then
            Dim nFile As System.IO.FileInfo = New System.IO.FileInfo(fullFilePath)

            If nFile.Exists Then
                Response.Clear()
                Response.AddHeader("Content-Disposition", "attachment; filename=" & UserFileName)
                Response.AddHeader("Content-Length", nFile.Length.ToString())
                Response.ContentType = "application/octet-stream"
                Response.WriteFile(nFile.FullName)
                Response.End()
            Else
                ClientScript.RegisterStartupScript(Me.GetType(), "JSFUN", "alert('File Not Exists!');", True)
                'Response.Write("This file does not exist.")
            End If
        End If
    End Sub
End Class
