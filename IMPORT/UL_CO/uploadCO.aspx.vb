
Imports Microsoft.VisualBasic
Imports System.IO
Imports NPOI.HSSF.UserModel
Imports NPOI.HPSF
Imports NPOI.POIFS.FileSystem
Imports System.Data.SqlClient
Imports System.Data
Imports System.Collections.Generic

Partial Class EXCEL_uploadCO
    Inherits System.Web.UI.Page

    Private SecUtil As New securityUtil
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Protected isSubmit As Boolean = False
    Dim DDFormat As String = gU.getConfig("DDFORMATNO")
    Private batchID As String = ""
    Private SuccessCount As Integer = 0
    Private tempCount As Integer = 0
    Private CO_CODE As String = ""
    Private WICO_BATCH_ID As Integer = 0
    Private LastCol As Integer = 0
    Private nFun As New NPOIFunc
    'Private msgLog As PrgmLog = New PrgmLog(System.Configuration.ConfigurationManager.AppSettings.Item("LOG_PATH").ToString & "\AP_EXCEL", "uploadCFMaster.txt")

    Private itemDict As New Dictionary(Of String, String)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        ar = New AccessRightUtils("UL_CO", Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        user_id.Value = Session("usr_id")

        If Not IsPostBack Then
            If Session("gLang") = "E" Then
                lblHD.Text = "Upload CO"
                btnReset.Text = "Reset"
                btnDL.Text = "Download Import Excel Template"
                lbl_OUTPUT.Text = "Output"
                lbl_status.Text = "Status"
                statusSpan.Text = "Please upload your CO EXCEL Data file."
                lbl_storer_code.Text = "Storer"
                lbl_datafile.Text = "Data File"
                lbl_result.Text = "Upload Result"
                btnSubmit.Text = "Submit"

            ElseIf Session("gLang") = "C" Then
                lblHD.Text = "上載客戶訂單"
                btnReset.Text = "重置"
                btnDL.Text = "下載Excel表格範本"
                lbl_OUTPUT.Text = "輸出顯示"
                lbl_status.Text = "狀態"
                statusSpan.Text = "請上載你的客戶訂單資料檔案."
                lbl_storer_code.Text = "貨主"
                lbl_datafile.Text = "資料檔案"
                lbl_result.Text = "上載結果"
                btnSubmit.Text = "開始上載"

            End If




            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            If Session("usr_type") <> "S" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
                STORER_CODE.Enabled = False
            End If
        End If

    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim fileExt As String

        isSubmit = False

        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Please select storer!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "請選擇貨主!", Session("gLang"))
            End If

            Exit Sub
        End If

        If (FileUpload1.HasFile) Then
            fileExt = Path.GetExtension(FileUpload1.FileName)
            If fileExt = ".xls" Then 'And FileUpload1.PostedFile.ContentType = "application/vnd.ms-excel" 
                If FileUpload1.PostedFile.ContentLength <= 5242880 Then
                    isSubmit = True
                    'FileUpload1.Visible = False
                    'btnSubmit.Visible = False
                    '                    cust_no.Visible = False

                Else
                    If Session("gLang") = "E" Then
                        outputSpan.Text = "File size too large! Maximum upload file size is 5MB."
                        outputSpan.ForeColor = Drawing.Color.Red
                        Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "fail", "<script language=""JavaScript"">alert(""File size too large! Maximum upload file size is 5MB."");</script>")
                    Else
                        outputSpan.Text = "檔案容量過大!最大容量為5MB."
                        outputSpan.ForeColor = Drawing.Color.Red
                        Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "fail", "<script language=""JavaScript"">alert(""檔案容量過大!最大容量為5MB."");</script>")
                    End If

                End If
            Else
                If Session("gLang") = "E" Then
                    outputSpan.Text = "Incorrect file format! Excel format Only."
                    outputSpan.ForeColor = Drawing.Color.Red
                    Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "fail", "<script language=""JavaScript"">alert(""Incorrect file format! Excel format Only."");</script>")
                Else
                    outputSpan.Text = "不正確的檔案類型!只能使用Excel 檔案."
                    outputSpan.ForeColor = Drawing.Color.Red
                    Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "fail", "<script language=""JavaScript"">alert(""不正確的檔案類型!只能使用Excel 檔案."");</script>")
                End If

            End If
        Else
            outputSpan.Text = "---"
            outputSpan.ForeColor = Drawing.Color.Black
            If Session("gLang") = "E" Then
                Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "fail", "<script language=""JavaScript"">alert(""Please select data file."");</script>")
            Else
                Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "fail", "<script language=""JavaScript"">alert(""請選擇檔案."");</script>")
            End If

        End If
    End Sub


    Private Function ReadFile() As Integer
        Dim ExcelWBObj As HSSFWorkbook
        Dim ExcelWSObj, ParamWsobj As HSSFSheet
        Dim ExcelRow As HSSFRow
        Dim ExcelCell As HSSFCell
        'Dim uStream As Stream
        Dim StartRow As Integer = 0
        Dim errMsg, statusMsg As String
        Dim isEmptyRow As Boolean
        Dim getValueSql, updateSql As String
        Dim sysFileName, sysFilePath, tmpFilePath, errFilePath As String
        Dim tFile As FileStream

        'Dim clobPkMatch

        Dim paP As GlobalDBFunc.DBCmdPara

        errMsg = ""
        statusMsg = ""
        Try
            'Save the uploaded file
            Path.GetExtension(FileUpload1.FileName)
            sysFileName = user_id.Value & Now.ToString("yyyyMMddHHmmssfff") & Path.GetExtension(FileUpload1.FileName)
            sysFilePath = gU.getConfig("SYSP_UPLD_DIR") & "\" & sysFileName
            tmpFilePath = gU.getConfig("SYSP_UPLD_DIR") & "\TEMP\" & sysFileName
            errFilePath = gU.getConfig("SYSP_UPLD_DIR") & "\ERROR\" & sysFileName
            FileUpload1.PostedFile.SaveAs(tmpFilePath)

            tFile = New FileStream(tmpFilePath, FileMode.Open, FileAccess.Read)

            ExcelWBObj = New HSSFWorkbook(tFile)
            ExcelWSObj = ExcelWBObj.GetSheetAt(0)
            ParamWsobj = ExcelWBObj.GetSheet("PARAM")

            ExcelRow = ExcelWSObj.GetRow(0)

            'Check format - header
            If ExcelRow Is Nothing Then
                errMsg = "Invalid Excel format! Header row is missing."
                statusMsg = "Upload fail! Please correct data file."
                Throw New Exception()

            End If

            ''Check format - header 2
            For i = 0 To 9
                ExcelCell = ExcelRow.GetCell(i)
                If ExcelCell Is Nothing OrElse ExcelCell.ToString = "" Then
                    errMsg = "Invalid Excel format! Column Header " & CStr(i + 1) & " is missing."
                    statusMsg = "Upload fail! Please correct data file."
                    Throw New Exception()
                End If
            Next

            ExcelRow = ExcelWSObj.GetRow(1)

            isEmptyRow = True

            'Check if file is empty
            If Not ExcelRow Is Nothing Then
                For i = 0 To 9
                    ExcelCell = ExcelRow.GetCell(i)
                    If Not ExcelCell Is Nothing AndAlso Trim(ExcelCell.ToString) <> "" Then
                        isEmptyRow = False
                    End If
                Next
            End If

            If isEmptyRow Then
                errMsg = "Empty Excel file! No data is found in first row."
                statusMsg = "Upload skipped! Please input data."
                Throw New Exception()
            End If

            If ParamWsobj Is Nothing OrElse nFun.getCell(ParamWsobj, 1, 5).ToString <> "CO" Then
                errMsg = "Template Type does not match. Please use correct template for upload CO."
                statusMsg = "Upload skipped! Please check your template file type."
                Throw New Exception()
            End If


            StartRow = 1

            Dim MaxLengthCol() As Integer = {100, 30, 100, 500, 100,
                                             100, 20, 30, 100, 20, 20}



            Dim cnn As SqlConnection
            cnn = gDB.getConnection()

            Dim oTrans As SqlTransaction
            oTrans = cnn.BeginTransaction()

            Try

                getValueSql = " select ISNULL(max(convert(int, WICO_BATCH_ID)), 0) + 1 as max_batch_id " & _
                              " from WMS_IMP_CO_LOG "

                WICO_BATCH_ID = gU.decodeEmptyCInt(DB.getValueFromSQL(getValueSql, cnn, oTrans), 0)

                paP = New GlobalDBFunc.DBCmdPara
                updateSql = "Insert into WMS_IMP_CO_LOG (WICO_BATCH_ID, WICO_START_DATE, WICO_STATUS, WICO_FILE_NAME) Values (" & _
                            paP.AP(WICO_BATCH_ID) & ",Getdate(),'F'," & paP.AP(sysFileName) & ")"

                gDB.amendData(updateSql, cnn, oTrans, paP)

                Dim tempSEQ As Integer = 0
                Dim str_code As String = STORER_CODE.SelectedValue
                Dim WICO_SHIP_MODE, WICO_DELI_DATE, WICO_SKU, WICO_ITM_NAME, WICO_BATCH_NO, WICO_PALLET_NO, WICO_QTY, WICO_VND_CODE,
                    WICO_UOM, WICO_REF_NO, WICO_PCS_PER_UOM As String

                Dim LastRownum As Integer = ExcelWSObj.LastRowNum

                Dim isEmptyCell As Boolean = True

                For i = StartRow To LastRownum
                    ExcelRow = ExcelWSObj.GetRow(i)
                    tempSEQ += 1

                    'Check if row is empty
                    isEmptyCell = True
                    For x = 0 To 9
                        If Not nFun.getCell(ExcelWSObj, i, x) Is Nothing AndAlso nFun.getCell(ExcelWSObj, i, x).ToString <> "" Then
                            isEmptyCell = False
                        End If
                    Next

                    If isEmptyCell Then
                        Exit For
                    End If

                    WICO_SHIP_MODE = Left(nFun.getCell(ExcelWSObj, i, 0).ToString, MaxLengthCol(0))
                    WICO_DELI_DATE = Left(nFun.getCell(ExcelWSObj, i, 1).ToString, MaxLengthCol(1))
                    WICO_SKU = Left(nFun.getCell(ExcelWSObj, i, 2).ToString, MaxLengthCol(2))
                    WICO_ITM_NAME = Left(nFun.getCell(ExcelWSObj, i, 3).ToString, MaxLengthCol(3))
                    WICO_BATCH_NO = Left(nFun.getCell(ExcelWSObj, i, 4).ToString, MaxLengthCol(4))
                    WICO_PALLET_NO = Left(nFun.getCell(ExcelWSObj, i, 5).ToString, MaxLengthCol(5))
                    WICO_QTY = Left(nFun.getCell(ExcelWSObj, i, 6).ToString, MaxLengthCol(6))
                    WICO_UOM = Left(nFun.getCell(ExcelWSObj, i, 7).ToString, MaxLengthCol(7))
                    WICO_PCS_PER_UOM = Left(nFun.getCell(ExcelWSObj, i, 8).ToString, MaxLengthCol(9))
                    WICO_REF_NO = Left(nFun.getCell(ExcelWSObj, i, 9).ToString, MaxLengthCol(8))
                    WICO_VND_CODE = Left(nFun.getCell(ExcelWSObj, i, 10).ToString, MaxLengthCol(10))

                    REM XXX
                    paP = New GlobalDBFunc.DBCmdPara
                    updateSql = "Insert into wms_IMP_CO_data (WICO_BATCH_ID,WICO_LINE_NO,WICO_IMP_ALLOW_YN,WICO_IMP_YN,WICO_IMP_MAP_CODE," & _
                                "WICO_IMP_REMARKS, IMP_CODE, STORER_CODE, WICO_SHIP_MODE, WICO_DELI_DATE," & _
                                "WICO_SKU,WICO_ITM_NAME,WICO_BATCH_NO,WICO_PALLET_NO,WICO_QTY,WICO_UOM," & _
                                "WICO_REF_NO,WICO_PCS_PER_UOM,WICO_VND_CODE) Values(" & _
                                paP.AP(WICO_BATCH_ID) & "," & paP.AP(tempSEQ) & ",'N','N','',''," & _
                                paP.AP(Session("imp_code")) & "," & paP.AP(STORER_CODE.SelectedValue) & "," & paP.AP(WICO_SHIP_MODE) & "," & paP.AP(WICO_DELI_DATE) & "," & _
                                paP.AP(WICO_SKU) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(WICO_ITM_NAME, ""))) & "," & paP.AP(WICO_BATCH_NO) & "," & paP.AP(WICO_PALLET_NO) & "," & paP.AP(WICO_QTY) & "," & paP.AP(WICO_UOM) & "," & _
                                paP.AP(WICO_REF_NO) & "," & paP.AP(WICO_PCS_PER_UOM) & "," & paP.AP(WICO_VND_CODE) & _
                                ")"
                    'Dim tempsql As String = gDB.getCmdSql(updateSql, paP)
                    gDB.amendData(updateSql, cnn, oTrans, paP)


                    If i Mod 5 = 0 Then
                        Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='black'>Reading... (" & CStr(i) & ")</font>"";</script>")
                        Response.Flush()
                    End If
                Next

                'Save the uploaded file
                tFile.Close()
                File.Move(tmpFilePath, sysFilePath)


                oTrans.Commit()

            Catch ex As Exception

                oTrans.Rollback()

                If File.Exists(tmpFilePath) Then
                    If tFile IsNot Nothing Then
                        tFile.Close()
                    End If
                    File.Move(tmpFilePath, errFilePath)
                End If

                Throw ex
            Finally
                ExcelWSObj = Nothing
                ExcelWBObj = Nothing
                If cnn IsNot Nothing Then
                    If cnn.State = ConnectionState.Open Then
                        cnn.Close()
                        cnn.Dispose()
                    End If
                End If
            End Try

            Return WICO_BATCH_ID

        Catch ex As Exception
            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>Upload Failed.</font>"";</script>")

            If errMsg <> "" Then
                Response.Write("<script language=""JavaScript"">alert(""" & SecUtil.jsString(SecUtil.htmlEncodeNull(errMsg)) & """);</script>")
                Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & SecUtil.jsString(SecUtil.htmlEncodeNull(errMsg)) & "';</script>")
            Else
                Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & SecUtil.jsString(SecUtil.htmlEncodeNull(ex.Message)) & "';</script>")
            End If

            If statusMsg <> "" Then
                'Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>" & SecUtil.jsString(SecUtil.htmlEncodeNull(statusMsg)) & "</font>"";</script>")
                Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & SecUtil.jsString(SecUtil.htmlEncodeNull(statusMsg)) & "';</script>")
            Else
                'Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload fail! Please check your excel file data.</font>"";</script>")
                Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='Upload fail! Please check your excel file data.';</script>")
            End If

            Response.Flush()

            Return 0
        End Try
    End Function

    Protected Sub uploadFile()
        If Session("gLang") = "E" Then
            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='black'>Preparing...</font>"";</script>")
            Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Processing</font>"";</script>")
            Response.Flush()
        Else
            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='black'>準備中...</font>"";</script>")
            Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>處理中</font>"";</script>")
            Response.Flush()
        End If
        

        Dim selectSQL As String = ""
        Dim tempDT As DataTable
        Dim Resultmsg As String = ""
        Dim updateSQL As String = ""

        Dim processFlag As Boolean = False

        Try
            Select Case STORER_CODE.SelectedValue
                Case "002"
                    WICO_BATCH_ID = ReadFile_ARIIX()
                Case Else
                    WICO_BATCH_ID = ReadFile_ARIIX()
            End Select

            If WICO_BATCH_ID = 0 Then
                If Session("gLang") = "E" Then
                    Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>Terminated.</font>"";</script>")
                    Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload Failed.</font>"";</script>")
                Else
                    Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>意外中止.</font>"";</script>")
                    Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>上載失敗.</font>"";</script>")
                End If
                Response.Flush()
                Exit Sub

            Else
                Select Case STORER_CODE.SelectedValue
                    Case "002"
                        processFlag = validateFile_ARIIX(WICO_BATCH_ID)
                    Case Else
                        processFlag = validateFile_ARIIX(WICO_BATCH_ID)
                End Select
                If processFlag Then
                    Select Case STORER_CODE.SelectedValue
                        Case "002"
                            processFlag = ImportData_ARIIX(WICO_BATCH_ID)
                        Case Else
                            processFlag = ImportData_ARIIX(WICO_BATCH_ID)
                    End Select

                    If processFlag Then
                        'Response.Write("<script language=""JavaScript"">alert('File is uploaded successfully!');</script>")

                        'disable email function 
                        'If Session("usr_type") = "C" Or Session("usr_type") = "T" Then
                        '    SendCustCOMail(CO_CODE, STORER_CODE.SelectedValue)
                        'End If
                        If Session("gLang") = "E" Then
                            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='black'>-----</font>"";</script>")
                            Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='blue'>DONE.</font>"";</script>")
                            Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='File is uploaded successfully!';</script>")
                        Else
                            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='black'>-----</font>"";</script>")
                            Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='blue'>完成.</font>"";</script>")
                            Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='資料成功上載!';</script>")
                        End If
                        TR_FILE.Visible = False
                        Select Case STORER_CODE.SelectedValue
                            Case "002"

                            Case Else
                                'If Session("usr_type") = "S" Then
                                '    Response.Write("<script language=""JavaScript"">showLink('../../OUTBOUND/CO/COMain.aspx?FrmUP=Y&storer_code=" & STORER_CODE.SelectedValue & "&CO_CODE=" & CO_CODE & "');</script>")
                                'Else
                                '    Response.Write("<script language=""JavaScript"">showLink('../../OUTBOUND/CO/COMain_C.aspx?FrmUP=Y&storer_code=" & STORER_CODE.SelectedValue & "&CO_CODE=" & CO_CODE & "');</script>")
                                'End If
                        End Select
                        'Response.Write("<script language=""JavaScript"">showLogList();</script>")

                    Else
                        If Session("gLang") = "E" Then
                            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>Terminated.</font>"";</script>")
                            Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload Failed.</font>"";</script>")
                            Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & Resultmsg & "';</script>")
                            Response.Write("<script language=""JavaScript"">alert(""Data Insert Failed! Please check your excel file data."");</script>")
                        Else
                            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>上載中止.</font>"";</script>")
                            Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>上載失敗.</font>"";</script>")
                            Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & Resultmsg & "';</script>")
                            Response.Write("<script language=""JavaScript"">alert(""資料輸入失敗! 請檢查你的檔案資料."");</script>")
                        End If

                        Response.Flush()
                    End If

                Else
                    Dim strBld As New StringBuilder

                    selectSQL = "Select WICO_IMP_REMARKS from WMS_IMP_CO_DATA where WICO_BATCH_ID='" & gU.dbEncode(WICO_BATCH_ID) & "' and WICO_IMP_ALLOW_YN='N'"
                    tempDT = gDB.getDataTable(selectSQL)

                    If tempDT.Rows.Count > 0 Then
                        For i = 0 To tempDT.Rows.Count - 1
                            strBld.Append(tempDT.Rows(i).Item("WICO_IMP_REMARKS").ToString.Trim & "\n")
                        Next
                    End If

                    Resultmsg = strBld.ToString

                    updateSQL = "Update WMS_IMP_CO_LOG set WICO_END_DATE=Getdate(), WICO_STATUS='F', WICO_REMARKS='Validation Failed'"
                    gDB.amendData(updateSQL)

                    Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>Terminated.</font>"";</script>")
                    Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload Failed.</font>"";</script>")
                    Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & Resultmsg & "';</script>")
                    Response.Write("<script language=""JavaScript"">alert(""Data Validation failed! Please check your excel file data."");</script>")
                    Response.Flush()
                End If

            End If

        Catch ex As Exception
            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='black'>Terminated.</font>"";</script>")
            Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & SecUtil.jsString(SecUtil.htmlEncodeNull(ex.Message)) & "';</script>")
            Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload fail! Please Check your excel file data.</font>"";</script>")
            Response.Write("<script language=""JavaScript"">alert(""Upload fail! Please Check you excel file data."");</script>")
            Response.Flush()
        Finally
            '    cnn.Close()
        End Try

    End Sub

    Protected Function validateFile(ByVal batch_id As String) As Boolean
        Dim selectSQL, getvalueSQL, updateSQL As String
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim selectDT, tempDT As DataTable

        Dim successflag As Boolean = False

        Dim tempSQL As String = ""

        Dim cnn As SqlConnection
        cnn = gDB.getConnection()

        Dim oTrans As SqlTransaction
        oTrans = cnn.BeginTransaction()
        Try
            paP = New GlobalDBFunc.DBCmdPara
            selectSQL = "Select * from wms_IMP_CO_data where WICO_BATCH_ID=" & paP.AP(batch_id) & " order by WICO_LINE_NO"

            selectDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

            If selectDT.Rows.Count > 0 Then
                Dim hasError As Boolean = False
                Dim errMsg As String = ""

                Dim str_code, WICO_SHIP_MODE, WICO_DELI_DATE, WICO_SKU, WICO_ITM_NAME, WICO_BATCH_NO, WICO_PALLET_NO, WICO_QTY, WICO_UOM, WICO_REF_NO, WICO_PCS_PER_UOM, WICO_VND_CODE As String
                Dim COD_ITM_CODE, COD_PACK_KEY As String
                COD_ITM_CODE = ""
                COD_PACK_KEY = ""

                Dim WICO_LINE_NO As Integer
                Dim tempUOM, tempPCSperUOM As String

                Dim tempcount As Integer = 0

                For i = 0 To selectDT.Rows.Count - 1
                    hasError = False
                    errMsg = ""

                    str_code = STORER_CODE.SelectedValue
                    WICO_SHIP_MODE = selectDT.Rows(i).Item("WICO_SHIP_MODE").ToString.Trim
                    WICO_DELI_DATE = selectDT.Rows(i).Item("WICO_DELI_DATE").ToString.Trim
                    WICO_SKU = selectDT.Rows(i).Item("WICO_SKU").ToString.Trim.ToUpper
                    WICO_ITM_NAME = selectDT.Rows(i).Item("WICO_ITM_NAME").ToString.Trim
                    WICO_BATCH_NO = selectDT.Rows(i).Item("WICO_BATCH_NO").ToString.Trim
                    WICO_PALLET_NO = selectDT.Rows(i).Item("WICO_PALLET_NO").ToString.Trim
                    WICO_QTY = selectDT.Rows(i).Item("WICO_QTY").ToString.Trim
                    WICO_UOM = selectDT.Rows(i).Item("WICO_UOM").ToString.Trim
                    tempUOM = selectDT.Rows(i).Item("WICO_UOM").ToString.Trim
                    WICO_REF_NO = selectDT.Rows(i).Item("WICO_REF_NO").ToString.Trim
                    WICO_PCS_PER_UOM = selectDT.Rows(i).Item("WICO_PCS_PER_UOM").ToString.Trim
                    tempPCSperUOM = selectDT.Rows(i).Item("WICO_PCS_PER_UOM").ToString.Trim
                    WICO_VND_CODE = selectDT.Rows(i).Item("WICO_VND_CODE").ToString.Trim

                    WICO_LINE_NO = gU.decodeEmptyCInt(selectDT.Rows(i).Item("WICO_LINE_NO").ToString.Trim, 0)

                    If WICO_SHIP_MODE <> "" Then

                        getvalueSQL = "select count(*) from wms_col_code where COLC_TABCOL='WMS_DELV_ORDER.DO_SHIP_MODE' and upper(COLC_CODE)=upper('" & gU.dbEncode(WICO_SHIP_MODE) & "')"
                        tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                        If tempcount <= 0 Then
                            hasError = True
                            errMsg &= "Invalid Ship Mode || "
                        End If

                    End If

                    If WICO_DELI_DATE <> "" AndAlso Not gU.isValidDate(WICO_DELI_DATE, "DD/MM/YYYY") Then
                        hasError = True
                        errMsg &= "Invalid Delivery Date || "
                    End If

                    If WICO_UOM <> "" Then
                        getvalueSQL = "select count(*) from wms_UOM where UOM_CODE='" & gU.dbEncode(WICO_UOM) & "' AND IMP_CODE='" & gU.dbEncode(Session("imp_code")) & "'"
                        tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                        If tempcount <= 0 Then
                            hasError = True
                            errMsg &= "Invalid UOM || "
                        End If
                    End If

                    If WICO_QTY <> "" Then
                        If Not gU.IsWholeNumber(WICO_QTY) Then
                            hasError = True
                            errMsg &= "Invalid Quantity || "
                        End If
                    Else
                        hasError = True
                        errMsg &= "Empty Quantity || "
                    End If

                    If WICO_PCS_PER_UOM <> "" Then
                        If Not gU.IsWholeNumber(WICO_PCS_PER_UOM) Then
                            hasError = True
                            errMsg &= "Invalid No. of UOM || "
                        End If
                    End If

                    If WICO_BATCH_NO <> "" Then
                        paP = New GlobalDBFunc.DBCmdPara
                        getvalueSQL = "Select count(*) from wms_date_code where imp_code=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(str_code)

                        tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans, paP)
                        If tempcount <= 0 Then
                            hasError = True
                            errMsg &= "Invalid Batch No. || "
                        End If
                    End If


                    If WICO_SKU <> "" AndAlso WICO_ITM_NAME <> "" Then
                        paP = New GlobalDBFunc.DBCmdPara

                        tempSQL = ""
                        'If WICO_UOM <> "" Then
                        '    tempSQL = " AND ITM_UOM=" & paP.AP(WICO_UOM.ToUpper)
                        'End If

                        'If WICO_PCS_PER_UOM <> "" Then
                        '    If gU.IsWholeNumber(WICO_PCS_PER_UOM) Then
                        '        tempSQL = " AND ITM_PCS_PER_UOM=" & paP.AP(WICO_PCS_PER_UOM)
                        '    End If
                        'End If

                        selectSQL = "select itm_code,itm_name, pack_key, ITM_UOM, ITM_PCS_PER_UOM from wms_item where itm_status <> 'CANCELLED' AND imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(str_code) & " and upper(ITM_SKU_NO)=" & paP.AP(WICO_SKU.ToUpper) & tempSQL
                        Dim tempshow As String = gDB.getCmdSql(selectSQL, paP)
                        tempDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

                        If tempDT.Rows.Count > 0 Then
                            COD_ITM_CODE = tempDT.Rows(0).Item("itm_code").ToString.Trim
                            COD_PACK_KEY = tempDT.Rows(0).Item("pack_key").ToString.Trim
                            WICO_PCS_PER_UOM = tempDT.Rows(0).Item("ITM_PCS_PER_UOM").ToString.Trim
                            WICO_UOM = tempDT.Rows(0).Item("ITM_UOM").ToString.Trim
                            WICO_ITM_NAME = tempDT.Rows(0).Item("ITM_NAME").ToString.Trim

                            If tempUOM <> "" AndAlso tempUOM.ToUpper <> WICO_UOM.ToUpper Then
                                hasError = True
                                errMsg &= "UOM does not match the item record || "
                            End If

                            If tempPCSperUOM <> "" AndAlso tempPCSperUOM <> WICO_PCS_PER_UOM Then
                                hasError = True
                                errMsg &= "No. per UOM does not match the item record || "
                            End If

                            If WICO_VND_CODE <> "" Then
                                getvalueSQL = "Select count(*) from WMS_ALT_VEND_ITEM where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(str_code) & "' and itm_code='" & COD_ITM_CODE & "' and pack_key='" & COD_PACK_KEY & "' and upper(vnd_code)='" & gU.dbEncode(WICO_VND_CODE.ToUpper) & "'"
                                tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                                If tempcount <= 0 Then
                                    hasError = True
                                    errMsg &= "Vendor Code does not match the item || "
                                End If

                            Else

                                getvalueSQL = "Select count(*) from WMS_ALT_VEND_ITEM where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(str_code) & "' and itm_code='" & COD_ITM_CODE & "' and pack_key='" & COD_PACK_KEY & "' and vnd_code='DEF_VEND' "
                                tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                                If tempcount > 0 Then
                                    WICO_VND_CODE = "DEF_VEND"
                                End If
                            End If

                        Else
                            hasError = True
                            errMsg &= "Cannot Locate the item || "

                        End If
                    Else
                        hasError = True
                        errMsg &= "Empty Stock No./ Item Name || "
                    End If

                    If hasError Then
                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "update WMS_IMP_CO_DATA set WICO_IMP_ALLOW_YN='N'," & _
                                    " WICO_IMP_REMARKS=" & paP.AP("ROW No.(" & WICO_LINE_NO & ") " & Left(errMsg, Len(errMsg) - 3)) & _
                                    " where WICO_BATCH_ID=" & paP.AP(WICO_BATCH_ID) & " AND WICO_LINE_NO=" & paP.AP(WICO_LINE_NO)
                        gDB.amendData(updateSQL, cnn, oTrans, paP)
                    Else
                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "update WMS_IMP_CO_DATA set WICO_IMP_ALLOW_YN='Y'," & _
                                    "COD_ITM_CODE=" & paP.AP(COD_ITM_CODE) & "," & _
                                    "COD_PACK_KEY=" & paP.AP(COD_PACK_KEY) & "," & _
                                    "WICO_PCS_PER_UOM=" & paP.AP(WICO_PCS_PER_UOM) & "," & _
                                    "WICO_UOM=" & paP.AP(WICO_UOM) & "," & _
                                    "WICO_ITM_NAME=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(WICO_ITM_NAME, ""))) & "," & _
                                    "WICO_VND_CODE=" & paP.AP(WICO_VND_CODE) & _
                                    " where WICO_BATCH_ID=" & paP.AP(WICO_BATCH_ID) & " AND WICO_LINE_NO=" & paP.AP(WICO_LINE_NO)
                        gDB.amendData(updateSQL, cnn, oTrans, paP)

                        'COD_ITM_CODE
                    End If

                Next

                getvalueSQL = "Select Count(*) from WMS_IMP_CO_DATA where WICO_BATCH_ID='" & gU.dbEncode(WICO_BATCH_ID) & "' and WICO_IMP_ALLOW_YN='N'"
                tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)


                If tempcount > 0 Then
                    successflag = False
                Else
                    hasError = False
                    errMsg = ""

                    getvalueSQL = "Select Count(distinct upper(WICO_REF_NO)) from wms_IMP_CO_data where WICO_BATCH_ID='" & gU.dbEncode(WICO_BATCH_ID) & "' and WICO_IMP_ALLOW_YN='Y' "
                    tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    If tempcount > 1 Then
                        hasError = True
                        errMsg &= "More than 1 REF# Exists || "
                    End If

                    getvalueSQL = "Select Count(distinct upper(WICO_SHIP_MODE)) from wms_IMP_CO_data where WICO_BATCH_ID='" & gU.dbEncode(WICO_BATCH_ID) & "' and WICO_IMP_ALLOW_YN='Y' "
                    tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    If tempcount > 1 Then
                        hasError = True
                        errMsg &= "More than 1 Ship Mode Exists || "
                    End If

                    getvalueSQL = "Select Count(distinct upper(WICO_DELI_DATE)) from wms_IMP_CO_data where WICO_BATCH_ID='" & gU.dbEncode(WICO_BATCH_ID) & "' and WICO_IMP_ALLOW_YN='Y' "
                    tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    If tempcount > 1 Then
                        hasError = True
                        errMsg &= "More than 1 Delivery Date Exists || "
                    End If

                    If hasError Then
                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "update WMS_IMP_CO_DATA set WICO_IMP_ALLOW_YN='N'," & _
                                    " WICO_IMP_REMARKS=" & paP.AP("ROW No.(" & WICO_LINE_NO & ") " & Left(errMsg, Len(errMsg) - 3)) & _
                                    " where WICO_BATCH_ID=" & paP.AP(WICO_BATCH_ID) & " AND WICO_LINE_NO=1"
                        gDB.amendData(updateSQL, cnn, oTrans, paP)

                        successflag = False
                    Else
                        successflag = True
                    End If



                End If



            Else
                paP = New GlobalDBFunc.DBCmdPara
                updateSQL = "update WMS_IMP_CO_LOG set WICO_END_DATE=Getdate(), WICO_remarks='No Import detail Data is found' where WICO_BATCH_ID=" & paP.AP(batch_id)
                gDB.amendData(updateSQL, cnn, oTrans, paP)


                successflag = False
            End If


            oTrans.Commit()

        Catch ex As Exception
            oTrans.Rollback()
            cnn.Close()

            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>Terminated.</font>"";</script>")
            Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & SecUtil.jsString(SecUtil.htmlEncodeNull(ex.Message)) & "';</script>")
            Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload fail! Please Check your excel file data.</font>"";</script>")
            Response.Write("<script language=""JavaScript"">alert(""Upload fail! Please Check you excel file data."");</script>")

            Return False

        Finally
            If cnn IsNot Nothing Then
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                    cnn.Dispose()
                End If
            End If
        End Try

        Return successflag
    End Function

    Protected Function ImportData(ByVal batch_id As String) As Boolean
        Dim successFlag As Boolean = False

        Dim cnn As SqlConnection
        cnn = gDB.getConnection()

        Dim oTrans As SqlTransaction
        oTrans = cnn.BeginTransaction()

        Dim selectSQL, updateSQL, insertSQL As String
        Dim selectDT, tempDT As DataTable

        Dim paP As GlobalDBFunc.DBCmdPara
        Try
            paP = New GlobalDBFunc.DBCmdPara
            selectSQL = "Select * from WMS_IMP_CO_DATA where WICO_BATCH_ID=" & paP.AP(batch_id) & " AND WICO_IMP_ALLOW_YN='Y' "

            selectDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

            If selectDT.Rows.Count > 0 Then
                CO_CODE = ""
                Dim co_track_no As String = ""

                CO_CODE = DB.getDocNo("CO", cnn, oTrans)
                co_track_no = DB.getDocNo("TRACKNO", cnn, oTrans)

                Dim WICO_SHIP_MODE As String = ""
                Dim WICO_REF_NO As String = ""
                Dim WICO_DELI_DATE As String = ""

                selectSQL = "Select max(WICO_DELI_DATE) as WICO_DELI_DATE, max(WICO_SHIP_MODE) as WICO_SHIP_MODE, max(WICO_REF_NO) as WICO_REF_NO " & _
                            " from WMS_IMP_CO_DATA where WICO_BATCH_ID='" & gU.dbEncode(batch_id) & "' AND WICO_IMP_ALLOW_YN='Y' "
                tempDT = gDB.getDataTable(selectSQL, cnn, oTrans)

                If tempDT.Rows.Count > 0 Then
                    WICO_SHIP_MODE = tempDT.Rows(0).Item("WICO_SHIP_MODE").ToString.Trim
                    WICO_REF_NO = tempDT.Rows(0).Item("WICO_REF_NO").ToString.Trim
                    WICO_DELI_DATE = tempDT.Rows(0).Item("WICO_DELI_DATE").ToString.Trim
                End If

                paP = New GlobalDBFunc.DBCmdPara
                insertSQL = "insert into wms_cust_order (" & _
                            "co_code, imp_code, storer_code, co_status,  " & _
                            "co_date, CO_TARGET_DELDATE, CO_CUS_REF_NO,  co_track_no, CO_SHIP_MODE," & _
                            "sys_cb, sys_cd, sys_lub, sys_lud) Values (" & _
                            paP.AP(CO_CODE) & "," & paP.AP(Session("imp_code")) & "," & paP.AP(selectDT.Rows(0).Item("STORER_CODE").ToString.Trim) & ",'NEW'," & _
                            "Getdate(),Convert(datetime, " & paP.AP(WICO_DELI_DATE) & "," & DDFormat & ")," & paP.AP(WICO_REF_NO) & "," & paP.AP(co_track_no) & "," & paP.AP(WICO_SHIP_MODE) & "," & _
                             "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                gDB.amendData(insertSQL, cnn, oTrans, paP)


                If CUS_CODE.SelectedValue <> "" Then
                    selectSQL = "select cus_name, cus_addr1_del, cus_addr2_del, cus_addr3_del, " & _
                                "cus_area_del, cus_region_del, cus_country_del, cus_cont_per_ord, cus_cont_tel_ord " & _
                                "from wms_customer where cus_code = '" & CUS_CODE.SelectedValue & "' " & _
                                "and storer_code = '" & selectDT.Rows(0).Item("STORER_CODE").ToString.Trim & "' " & _
                                "and imp_code = '" & Session("IMP_CODE") & "'"
                    tempDT = gDB.getDataTable(selectSQL, cnn, oTrans)


                    If tempDT.Rows.Count > 0 Then
                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = " update wms_cust_order set " & _
                                    " CUS_CODE=" & paP.AP(CUS_CODE.SelectedValue) & ", " & _
                                    " CUS_NAME =" & paP.AP(tempDT.Rows(0).Item("cus_name").ToString.Trim) & ", " & _
                                    " CO_ADDR1 =" & paP.AP(tempDT.Rows(0).Item("cus_addr1_del").ToString.Trim) & ", " & _
                                    " CO_ADDR2 =" & paP.AP(tempDT.Rows(0).Item("cus_addr2_del").ToString.Trim) & ", " & _
                                    " CO_ADDR3 =" & paP.AP(tempDT.Rows(0).Item("cus_addr3_del").ToString.Trim) & ", " & _
                                    " CO_AREA_DEL =" & paP.AP(tempDT.Rows(0).Item("cus_area_del").ToString.Trim) & ", " & _
                                    " CO_REGION_DEL =" & paP.AP(tempDT.Rows(0).Item("cus_region_del").ToString.Trim) & ", " & _
                                    " CO_COUNTRY_DEL =" & paP.AP(tempDT.Rows(0).Item("cus_country_del").ToString.Trim) & ", " & _
                                    " CO_CUS_CONT =" & paP.AP(tempDT.Rows(0).Item("cus_cont_per_ord").ToString.Trim) & ", " & _
                                    " CO_CUS_CONT_TEL =" & paP.AP(tempDT.Rows(0).Item("cus_cont_tel_ord").ToString.Trim) & _
                                    " where imp_code=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(0).Item("STORER_CODE").ToString.Trim) & _
                                    " AND co_code=" & paP.AP(CO_CODE)
                        gDB.amendData(updateSQL, cnn, oTrans, paP)
                    End If

                End If


                Dim tot_pcs As Double = 0
                Dim tempCount As Integer = 0
                For i = 0 To selectDT.Rows.Count - 1

                    tot_pcs = gU.decodeEmptyCdbl(selectDT.Rows(i).Item("WICO_QTY").ToString.Trim, 0) * gU.decodeEmptyCdbl(selectDT.Rows(i).Item("WICO_PCS_PER_UOM").ToString.Trim, 1)

                    paP = New GlobalDBFunc.DBCmdPara
                    insertSQL = "insert into wms_cust_order_d ( " & _
                                "CO_CODE, imp_code, storer_code, COD_seq, COD_disp_seq, " & _
                                "COD_pallet_no, COD_batch_no, COD_itm_code, COD_pack_key, COD_ITM_DESC, " & _
                                "COD_qty, COD_uom, COD_PCS_UOM, cod_totpcs, " & _
                                " sys_cb, sys_cd, sys_lub, sys_lud) Values (" & _
                                paP.AP(CO_CODE) & "," & paP.AP(Session("imp_code")) & "," & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WICO_LINE_NO").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WICO_LINE_NO").ToString.Trim) & "," & _
                                paP.AP(gU.decodeNullOrEmpty(selectDT.Rows(i).Item("WICO_PALLET_NO").ToString.Trim, "000")) & "," & paP.AP(selectDT.Rows(i).Item("WICO_BATCH_NO").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("COD_ITM_CODE").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("COD_pack_key").ToString.Trim) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(selectDT.Rows(i).Item("WICO_ITM_NAME").ToString.Trim, ""))) & "," & _
                                paP.AP(selectDT.Rows(i).Item("WICO_QTY").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WICO_UOM").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WICO_PCS_PER_UOM").ToString.Trim) & "," & paP.AP(tot_pcs) & "," & _
                                 "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "

                    'Dim showSQL As String = gDB.getCmdSql(insertSQL, paP)
                    gDB.amendData(insertSQL, cnn, oTrans, paP)

                    Dim COD_PCS_CARTON As String = ""
                    Dim COD_TOT_WGT As String = ""
                    Dim COD_TOT_CBM As String = ""

                    'If selectDT.Rows(i).Item("WICO_VND_CODE").ToString.Trim <> "" Then
                    paP = New GlobalDBFunc.DBCmdPara
                    selectSQL = "Select AITM_QTY_PER_CTN, AITM_VOL, CARTON_CBM from V_ALT_VEND_ITEM " & _
                                "where imp_code=" & paP.AP(Session("imp_code")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                " AND ITM_CODE=" & paP.AP(selectDT.Rows(i).Item("COD_ITM_CODE").ToString.Trim) & " AND pack_key=" & paP.AP(selectDT.Rows(i).Item("COD_PACK_KEY").ToString.Trim)
                    tempDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

                    If tempDT.Rows.Count > 0 Then
                        COD_PCS_CARTON = tempDT.Rows(0).Item("AITM_QTY_PER_CTN").ToString.Trim
                        COD_TOT_WGT = tempDT.Rows(0).Item("AITM_VOL").ToString.Trim
                        COD_TOT_CBM = tempDT.Rows(0).Item("CARTON_CBM").ToString.Trim

                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "Update wms_cust_order_d set " & _
                                    " COD_PCS_CARTON =" & paP.AP(COD_PCS_CARTON) & "," & _
                                    " COD_TOT_WGT =" & paP.AP(COD_TOT_WGT) & "," & _
                                    " COD_TOT_CBM =" & paP.AP(COD_TOT_CBM) & _
                                    " Where imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                    " AND co_code=" & paP.AP(CO_CODE) & " AND cOD_SEQ=" & paP.AP(selectDT.Rows(i).Item("WICO_LINE_NO").ToString.Trim)


                        Dim temp1 As String = gDB.getCmdSql(updateSQL, paP)
                        gDB.amendData(updateSQL, cnn, oTrans, paP)
                    End If
                    ' End If



                    updateSQL = "Update WMS_IMP_CO_DATA set WICO_IMP_YN='Y'" & _
                                " where WICO_BATCH_ID=" & gU.dbEncode(batch_id) & " AND WICO_LINE_NO='" & selectDT.Rows(i).Item("WICO_LINE_NO").ToString.Trim & "' "
                    gDB.amendData(updateSQL, cnn, oTrans)
                Next

                paP = New GlobalDBFunc.DBCmdPara
                updateSQL = "Update WMS_IMP_CO_LOG set WICO_END_DATE=Getdate(), WICO_STATUS='D', WICO_REMARKS='Success' where WICO_BATCH_ID=" & paP.AP(batch_id)
                gDB.amendData(updateSQL, cnn, oTrans, paP)

                successFlag = True
            Else
                paP = New GlobalDBFunc.DBCmdPara
                updateSQL = "Update WMS_IMP_CO_LOG set WICO_END_DATE=Getdate(), WICO_STATUS='F', WICO_REMARKS='No import rows' where WICO_BATCH_ID=" & paP.AP(batch_id)
                gDB.amendData(updateSQL, cnn, oTrans, paP)
                successFlag = False
            End If

            oTrans.Commit()

        Catch ex As Exception

            oTrans.Rollback()
            Return False
        Finally
            If cnn IsNot Nothing Then
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                    cnn.Dispose()
                End If
            End If
        End Try

        Return successFlag
    End Function


    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Dim rmtpost As New RemotePost

        rmtpost.Url = "uploadCO.aspx"
        rmtpost.Target = "_self"
        rmtpost.Post()

    End Sub

    Protected Sub btnDL_Click(sender As Object, e As System.EventArgs) Handles btnDL.Click
        Dim sysPath As String = ""
        Select Case STORER_CODE.SelectedValue
            Case "002"
                sysPath = Server.MapPath("IMPORT_CO_TMPT_ARIIX.xls")

            Case Else
                sysPath = Server.MapPath("IMPORT_CO_TMPT_ARIIX.xls")
        End Select

        Dim nFile As System.IO.FileInfo = New System.IO.FileInfo(sysPath)

        If nFile.Exists Then
            Response.Clear()
            Response.AddHeader("Content-Disposition", "attachment; filename=IMPORT_CO_TMPT" & nFile.Extension)
            Response.AddHeader("Content-Length", nFile.Length.ToString())
            Response.ContentType = "application/octet-stream"
            Response.WriteFile(nFile.FullName)
            Response.End()
        Else
            Response.Write("This file does not exist.")
        End If
    End Sub

    Protected Sub STORER_CODE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles STORER_CODE.SelectedIndexChanged
        Dim selectSQL As String = ""
        selectSQL = "Select CUS_CODE, cus_name from wms_customer where imp_code='" & Session("imp_code") & "' AND storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "'"

        uiFun.load_dropdown(CUS_CODE, selectSQL, "CUS_CODE", "CUS_NAME", , Session("gSelectLabel"))

    End Sub

    Protected Sub SendCustCOMail(ByRef c_code As String, ByRef st_code As String)
        Dim app_email As String = ""
        Dim mail_title As String = "A New Customer Order has been created by customer."
        Dim mail_body As String = ""

        Dim selectSQL As String = "Select distinct lower(usr_email) as mailadd from wms_user, wms_user_group_alloc where " & _
                                  " wms_user.usr_id = wms_user_group_alloc.usr_id " & _
                                  " and wms_user_group_alloc.grp_code = 'CO_MAIL_GRP' "

        Dim mailDT As New DataTable


        Dim getValSql = "select sto_name as VALUE from wms_storer where storer_code = '" & gU.dbEncode(st_code) & "' "
        Dim lSTO_NAME = DB.getValueFromSQL(getValSql)


        mailDT = gDB.getDataTable(selectSQL)

        If mailDT.Rows.Count > 0 Then
            For i = 0 To mailDT.Rows.Count - 1
                app_email &= mailDT.Rows(i).Item("mailadd").ToString.Trim & ","
            Next
            app_email = Left(app_email, Len(app_email) - 1)
        Else
            app_email = gU.getConfig("AdminEmail")
        End If

        mail_body &= "The following Customer order has been created:" & vbNewLine
        mail_body &= "  CO Code: " & c_code & vbNewLine
        'mail_body &= "  Storer Name: " & st_code & vbNewLine
        mail_body &= "  Storer Name: " & lSTO_NAME & vbNewLine
        mail_body &= "  Created Date: " & Now.Date.ToString("dd/MM/yyyy HH:mm:ss")

        Dim mailog As PrgmLog

        If Not System.IO.Directory.Exists(HttpRuntime.Cache("SYSP_LOG_DIR")) Then
            System.IO.Directory.CreateDirectory(HttpRuntime.Cache("SYSP_LOG_DIR"))
        End If

        mailog = New PrgmLog(HttpRuntime.Cache("SYSP_LOG_DIR"), "emailLog" & Now.Date.ToString("ddMMyyyy") & ".txt")

        Call gU.sendEmail(app_email, mail_title, mail_body, mailog)


    End Sub

    Private Function ReadFile_ARIIX() As Integer
        Dim ExcelWBObj As HSSFWorkbook
        Dim ExcelWSObj As HSSFSheet
        Dim ExcelRow As HSSFRow
        Dim ExcelCell As HSSFCell
        'Dim uStream As Stream
        Dim StartRow As Integer = 0
        Dim errMsg, statusMsg As String
        Dim isEmptyRow As Boolean
        Dim getValueSql, updateSql As String
        Dim sysFileName, sysFilePath, tmpFilePath, errFilePath As String
        Dim tFile As FileStream

        'Dim clobPkMatch

        Dim paP As GlobalDBFunc.DBCmdPara

        errMsg = ""
        statusMsg = ""

        Try
            'Save the uploaded file
            Path.GetExtension(FileUpload1.FileName)
            sysFileName = user_id.Value & Now.ToString("yyyyMMddHHmmssfff") & Path.GetExtension(FileUpload1.FileName)
            sysFilePath = gU.getConfig("SYSP_UPLD_DIR") & "\" & sysFileName
            tmpFilePath = gU.getConfig("SYSP_UPLD_DIR") & "\TEMP\" & sysFileName
            errFilePath = gU.getConfig("SYSP_UPLD_DIR") & "\ERROR\" & sysFileName

            If Not System.IO.Directory.Exists(gU.getConfig("SYSP_UPLD_DIR")) Then
                System.IO.Directory.CreateDirectory(gU.getConfig("SYSP_UPLD_DIR"))
            End If

            If Not System.IO.Directory.Exists(gU.getConfig("SYSP_UPLD_DIR") & "\TEMP\") Then
                System.IO.Directory.CreateDirectory(gU.getConfig("SYSP_UPLD_DIR") & "\TEMP\")
            End If

            If Not System.IO.Directory.Exists(gU.getConfig("SYSP_UPLD_DIR") & "\ERROR\") Then
                System.IO.Directory.CreateDirectory(gU.getConfig("SYSP_UPLD_DIR") & "\ERROR\")
            End If

            FileUpload1.PostedFile.SaveAs(tmpFilePath)

            tFile = New FileStream(tmpFilePath, FileMode.Open, FileAccess.Read)

            ExcelWBObj = New HSSFWorkbook(tFile)
            ExcelWSObj = ExcelWBObj.GetSheetAt(1)
            'ParamWsobj = ExcelWBObj.GetSheet("PARAM")

            ExcelRow = ExcelWSObj.GetRow(0)

            'Check format - header
            If ExcelRow Is Nothing Then
                errMsg = "Invalid Excel format! Header row is missing."
                statusMsg = "Upload fail! Please correct data file."
                Throw New Exception()

            End If

            ''Check format - header 2

            ExcelRow = ExcelWSObj.GetRow(1)

            For i = 0 To 19
                ExcelCell = ExcelRow.GetCell(i)
                If ExcelCell Is Nothing OrElse ExcelCell.ToString = "" Then
                    errMsg = "Invalid Excel format! Column Header " & CStr(i + 1) & " is missing."
                    statusMsg = "Upload fail! Please correct data file."
                    Throw New Exception()
                End If
            Next

            ExcelRow = ExcelWSObj.GetRow(2)

            isEmptyRow = True

            'Check if file is empty
            If Not ExcelRow Is Nothing Then
                For i = 0 To 19
                    ExcelCell = ExcelRow.GetCell(i)
                    If Not ExcelCell Is Nothing AndAlso Trim(ExcelCell.ToString) <> "" Then
                        isEmptyRow = False
                    End If
                Next
            End If

            If isEmptyRow Then
                errMsg = "Empty Excel file! No data is found in first row."
                statusMsg = "Upload skipped! Please input data."
                Throw New Exception()
            End If

            ExcelRow = ExcelWSObj.GetRow(0)

            Dim isEnd As Boolean = False
            Dim colCount As Integer = 20
            Dim emptyCount As Integer = 0
            Dim itmStr As String = ""
            Dim formEval As HSSFFormulaEvaluator = ExcelWBObj.GetCreationHelper.CreateFormulaEvaluator


            Dim itemSKU As String = ""
            LastCol = ExcelRow.LastCellNum

            If LastCol = colCount Then
                errMsg = "Invalid Excel format! No Item Exists!"
                statusMsg = "Upload fail! Please correct data file."
                Throw New Exception()
            End If

            For i = colCount To LastCol

                ExcelCell = ExcelRow.GetCell(i)
                If ExcelCell IsNot Nothing Then
                    If ExcelCell.CellType = NPOI.SS.UserModel.CellType.Formula Then
                        formEval.EvaluateFormulaCell(nFun.getCell(ExcelWSObj, 0, i))

                        Select Case ExcelCell.CachedFormulaResultType
                            Case NPOI.SS.UserModel.CellType.Numeric
                                itemSKU = nFun.getCell(ExcelWSObj, 0, i).NumericCellValue
                            Case NPOI.SS.UserModel.CellType.String
                                itemSKU = nFun.getCell(ExcelWSObj, 0, i).StringCellValue
                            Case NPOI.SS.UserModel.CellType.Blank
                                itemSKU = ""
                            Case NPOI.SS.UserModel.CellType.Boolean
                                itemSKU = nFun.getCell(ExcelWSObj, 0, i).BooleanCellValue
                            Case Else
                                itemSKU = ""
                        End Select

                    Else
                        itemSKU = nFun.getCell(ExcelWSObj, 0, i).ToString
                    End If
                    

                    If itemSKU <> "" AndAlso itemSKU <> "N/A" Then
                        paP = New GlobalDBFunc.DBCmdPara
                        getValueSql = "Select max(itm_code + ', ' + pack_key + ', ' + itm_name) as itm_key from wms_item where itm_status <> 'CANCELLED' AND upper(itm_sku_no)=" & paP.AP(itemSKU.ToUpper.Trim.Replace("'", "")) & " and storer_code=" & paP.AP(STORER_CODE.SelectedValue)
                        itmStr = DB.getValueFromSQL(getValueSql, , , paP)
                    End If

                    If itemSKU <> "" AndAlso itmStr = "" AndAlso itemSKU <> "N/A" Then
                        errMsg = "Upload fail! Please correct data file."
                        statusMsg = "col. " & i + 1 & " Item NO. " & itemSKU & " not exists! Please check you datefile."
                        Throw New Exception()
                    End If
                End If
            Next


            StartRow = 2

            Dim MaxLengthCol() As Integer = {80, 0, 500, 0, 400, 80, 400, 400, 80, 20,
                                             200, 200, 100, 100, 100, 100, 500, 100, 200}

            Dim cnn As SqlConnection
            cnn = gDB.getConnection()

            Dim oTrans As SqlTransaction
            oTrans = cnn.BeginTransaction()

            Try

                getValueSql = " select ISNULL(max(convert(int, WICO_BATCH_ID)), 0) + 1 as max_batch_id " & _
                              " from WMS_IMP_CO_LOG "

                WICO_BATCH_ID = gU.decodeEmptyCInt(DB.getValueFromSQL(getValueSql, cnn, oTrans), 0)

                paP = New GlobalDBFunc.DBCmdPara
                updateSql = "Insert into WMS_IMP_CO_LOG (WICO_BATCH_ID, WICO_START_DATE, WICO_STATUS, WICO_FILE_NAME) Values (" & _
                            paP.AP(WICO_BATCH_ID) & ",Getdate(),'F'," & paP.AP(sysFileName) & ")"

                gDB.amendData(updateSql, cnn, oTrans, paP)

                Dim tempSEQ As Integer = 0
                Dim str_code As String = STORER_CODE.SelectedValue
                Dim WICO_SHIP_MODE, WICO_DELI_DATE, WICO_SKU, WICO_ITM_NAME, WICO_BATCH_NO, WICO_PALLET_NO, WICO_QTY, WICO_VND_CODE,
                    WICO_UOM, WICO_REF_NO, WICO_PCS_PER_UOM As String
                Dim WICO_CUS_CODE, WICO_CUS_NAME, WICO_CUS_CONT_TEL, WICO_CUS_CONT, WICO_REM, WICO_ADDR, WIDO_FTRACK_NO, WIDO_CONF_DELI_DATE, WIDO_DELI_STATUS, WIDO_DELI_RMKS As String

                Dim WICO_SENDER, WICO_SENDER_COUNTRY, WICO_SENDER_PROVINCE, WICO_SENDER_REGION, WICO_SENDER_ADDR, WICO_SENDER_TEL, WICO_PROVINCE, WICO_CITY, WICO_REGION, WICO_COUNTRY As String

                Dim LastRownum As Integer = ExcelWSObj.LastRowNum
                Dim tempCOGRP As Integer = 0
                Dim preCusCode As String = ""

                Dim isEmptyCell As Boolean = True

                For i = StartRow To LastRownum
                    ExcelRow = ExcelWSObj.GetRow(i)

                    If ExcelRow IsNot Nothing Then

                        'Check if row is empty
                        isEmptyCell = True

                        For x = 0 To 18
                            If Not nFun.getCell(ExcelWSObj, i, x) Is Nothing AndAlso nFun.getCell(ExcelWSObj, i, x).ToString <> "" Then
                                isEmptyCell = False
                            End If
                        Next

                        If isEmptyCell Then
                            Exit For
                        End If
                        WICO_REF_NO = Left(nFun.getCell(ExcelWSObj, i, 10).ToString, MaxLengthCol(4))

                        If WICO_REF_NO <> "" Then

                            tempCOGRP += 1

                            For x = colCount To LastCol
                                If nFun.getCell(ExcelWSObj, i, x) IsNot Nothing AndAlso nFun.getCell(ExcelWSObj, i, x).ToString <> "" AndAlso nFun.getCell(ExcelWSObj, i, x).ToString <> "0" Then
                                    tempSEQ += 1

                                    WIDO_FTRACK_NO = Left(nFun.getCell(ExcelWSObj, i, 0).ToString, MaxLengthCol(0))
                                    WIDO_DELI_STATUS = nFun.getCell(ExcelWSObj, i, 1).ToString
                                    WIDO_DELI_RMKS = Left(nFun.getCell(ExcelWSObj, i, 2).ToString, MaxLengthCol(2))
                                    WIDO_CONF_DELI_DATE = nFun.getCell(ExcelWSObj, i, 3).ToString
                                    WICO_REF_NO = Left(nFun.getCell(ExcelWSObj, i, 10).ToString, MaxLengthCol(4))
                                    WICO_CUS_CODE = Left(nFun.getCell(ExcelWSObj, i, 11).ToString, MaxLengthCol(5))

                                    WICO_CUS_NAME = Left(nFun.getCell(ExcelWSObj, i, 12).ToString, MaxLengthCol(6))

                                    WICO_CUS_CONT = Left(nFun.getCell(ExcelWSObj, i, 13).ToString, 200)
                                    WICO_ADDR = Left(nFun.getCell(ExcelWSObj, i, 18).ToString, MaxLengthCol(7))
                                    WICO_CUS_CONT_TEL = Left(nFun.getCell(ExcelWSObj, i, 19).ToString, MaxLengthCol(8))

                                    WICO_SENDER = Left(nFun.getCell(ExcelWSObj, i, 4).ToString, MaxLengthCol(12))
                                    WICO_SENDER_COUNTRY = Left(nFun.getCell(ExcelWSObj, i, 5).ToString, MaxLengthCol(13))
                                    WICO_SENDER_PROVINCE = Left(nFun.getCell(ExcelWSObj, i, 6).ToString, MaxLengthCol(14))
                                    WICO_SENDER_REGION = Left(nFun.getCell(ExcelWSObj, i, 7).ToString, MaxLengthCol(15))
                                    WICO_SENDER_ADDR = Left(nFun.getCell(ExcelWSObj, i, 8).ToString, MaxLengthCol(16))
                                    WICO_SENDER_TEL = Left(nFun.getCell(ExcelWSObj, i, 9).ToString, MaxLengthCol(17))

                                    WICO_COUNTRY = Left(nFun.getCell(ExcelWSObj, i, 14).ToString, 100)
                                    WICO_PROVINCE = Left(nFun.getCell(ExcelWSObj, i, 15).ToString, 100)
                                    WICO_REGION = Left(nFun.getCell(ExcelWSObj, i, 16).ToString, 200)
                                    WICO_CITY = Left(nFun.getCell(ExcelWSObj, i, 17).ToString, 200)

                                    WICO_QTY = Left(nFun.getCell(ExcelWSObj, i, x).ToString, MaxLengthCol(9))

                                    If nFun.getCell(ExcelWSObj, 0, x).CellType = NPOI.SS.UserModel.CellType.Formula Then
                                        formEval.EvaluateFormulaCell(nFun.getCell(ExcelWSObj, 0, x))
                                        Select Case nFun.getCell(ExcelWSObj, 0, x).CachedFormulaResultType
                                            Case NPOI.SS.UserModel.CellType.Numeric
                                                WICO_SKU = nFun.getCell(ExcelWSObj, 0, x).NumericCellValue
                                            Case NPOI.SS.UserModel.CellType.String
                                                WICO_SKU = nFun.getCell(ExcelWSObj, 0, x).StringCellValue.Replace("'", "")
                                            Case NPOI.SS.UserModel.CellType.Blank
                                                WICO_SKU = ""
                                            Case NPOI.SS.UserModel.CellType.Boolean
                                                WICO_SKU = nFun.getCell(ExcelWSObj, 0, x).BooleanCellValue
                                            Case Else
                                                WICO_SKU = ""
                                        End Select
                                    Else
                                        WICO_SKU = nFun.getCell(ExcelWSObj, 0, x).ToString
                                    End If

                                    If nFun.getCell(ExcelWSObj, 1, x).CellType = NPOI.SS.UserModel.CellType.Formula Then
                                        formEval.EvaluateFormulaCell(nFun.getCell(ExcelWSObj, 1, x))

                                        Select Case nFun.getCell(ExcelWSObj, 1, x).CachedFormulaResultType
                                            Case NPOI.SS.UserModel.CellType.Numeric
                                                WICO_ITM_NAME = nFun.getCell(ExcelWSObj, 1, x).NumericCellValue
                                            Case NPOI.SS.UserModel.CellType.String
                                                WICO_ITM_NAME = nFun.getCell(ExcelWSObj, 1, x).StringCellValue.Replace("'", "")
                                            Case NPOI.SS.UserModel.CellType.Blank
                                                WICO_ITM_NAME = ""
                                            Case NPOI.SS.UserModel.CellType.Boolean
                                                WICO_ITM_NAME = nFun.getCell(ExcelWSObj, 1, x).BooleanCellValue
                                            Case Else
                                                WICO_ITM_NAME = ""
                                        End Select
                                    Else
                                        WICO_ITM_NAME = nFun.getCell(ExcelWSObj, 1, x).ToString
                                    End If


                                    paP = New GlobalDBFunc.DBCmdPara
                                    updateSql = "Insert into wms_IMP_CO_data (WICO_BATCH_ID,WICO_LINE_NO,WICO_IMP_ALLOW_YN,WICO_IMP_YN,WICO_IMP_MAP_CODE," & _
                                                "WICO_IMP_REMARKS, IMP_CODE, STORER_CODE, WICO_SHIP_MODE, WICO_DELI_DATE," & _
                                                "WICO_SKU,WICO_ITM_NAME,WICO_BATCH_NO,WICO_PALLET_NO,WICO_QTY,WICO_UOM," & _
                                                "WICO_REF_NO,WICO_PCS_PER_UOM,WICO_VND_CODE, " & _
                                                "WICO_CUS_CODE, WICO_CUS_NAME, WICO_CUS_CONT_TEL, WICO_CUS_CONT, WICO_REM, WICO_ADDR, WIDO_FTRACK_NO, WIDO_CONF_DELI_DATE, WIDO_DELI_STATUS, WIDO_DELI_RMKS, CO_CODE, " & _
                                                "WICO_SENDER, WICO_SENDER_COUNTRY, WICO_SENDER_PROVINCE, WICO_SENDER_REGION, WICO_SENDER_ADDR, WICO_SENDER_TEL, WICO_PROVINCE, WICO_CITY, WICO_REGION, WICO_COUNTRY " & _
                                                ") Values(" & _
                                                paP.AP(WICO_BATCH_ID) & "," & paP.AP(tempSEQ) & ",'N','N','',''," & _
                                                paP.AP(Session("imp_code")) & "," & paP.AP(STORER_CODE.SelectedValue) & "," & paP.AP("") & "," & paP.AP("") & "," & _
                                                paP.AP(WICO_SKU) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(WICO_ITM_NAME, ""))) & "," & paP.AP("") & "," & paP.AP("000") & "," & paP.AP(WICO_QTY) & "," & paP.AP("") & "," & _
                                                paP.AP(WICO_REF_NO) & "," & paP.AP("") & "," & paP.AP("") & "," & _
                                                paP.AP(WICO_CUS_CODE) & "," & paP.AP(WICO_CUS_NAME, SqlDbType.NVarChar) & "," & paP.AP(WICO_CUS_CONT_TEL) & "," & paP.AP(WICO_CUS_CONT, SqlDbType.NVarChar) & "," & paP.AP("") & "," & paP.AP(WICO_ADDR, SqlDbType.NVarChar) & "," & _
                                                paP.AP(WIDO_FTRACK_NO) & "," & paP.AP(WIDO_CONF_DELI_DATE) & "," & paP.AP(WIDO_DELI_STATUS) & "," & paP.AP(WIDO_DELI_RMKS, SqlDbType.NVarChar) & "," & paP.AP(tempCOGRP) & "," & _
                                                paP.AP(WICO_SENDER, SqlDbType.NVarChar) & "," & paP.AP(WICO_SENDER_COUNTRY, SqlDbType.NVarChar) & "," & paP.AP(WICO_SENDER_PROVINCE, SqlDbType.NVarChar) & "," & paP.AP(WICO_SENDER_REGION, SqlDbType.NVarChar) & "," & paP.AP(WICO_SENDER_ADDR, SqlDbType.NVarChar) & "," & paP.AP(WICO_SENDER_TEL) & "," & paP.AP(WICO_PROVINCE, SqlDbType.NVarChar) & "," & paP.AP(WICO_CITY, SqlDbType.NVarChar) & "," & _
                                                paP.AP(WICO_REGION, SqlDbType.NVarChar) & ", " & paP.AP(WICO_COUNTRY, SqlDbType.NVarChar) & _
                                                ")"
                                    'Dim tempsql As String = gDB.getCmdSql(updateSql, paP)
                                    gDB.amendData(updateSql, cnn, oTrans, paP)


                                End If
                            Next
                        End If
                       

                        If i Mod 5 = 0 Then
                            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='black'>Reading... (" & CStr(i) & ")</font>"";</script>")
                            Response.Flush()
                        End If
                    End If
                Next

                'Save the uploaded file
                tFile.Close()
                File.Move(tmpFilePath, sysFilePath)


                oTrans.Commit()

            Catch ex As Exception

                oTrans.Rollback()

                If File.Exists(tmpFilePath) Then
                    If tFile IsNot Nothing Then
                        tFile.Close()
                    End If
                    File.Move(tmpFilePath, errFilePath)
                End If

                Throw ex
            Finally
                ExcelWSObj = Nothing
                ExcelWBObj = Nothing
                If cnn IsNot Nothing Then
                    If cnn.State = ConnectionState.Open Then
                        cnn.Close()
                        cnn.Dispose()
                    End If
                End If
            End Try

            Return WICO_BATCH_ID

        Catch ex As Exception
            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>Upload Failed.</font>"";</script>")

            If errMsg <> "" Then
                Response.Write("<script language=""JavaScript"">alert(""" & SecUtil.jsString(SecUtil.htmlEncodeNull(errMsg)) & """);</script>")
                Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & SecUtil.jsString(SecUtil.htmlEncodeNull(errMsg)) & "';</script>")
            Else
                Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & SecUtil.jsString(SecUtil.htmlEncodeNull(ex.Message)) & "';</script>")
            End If

            If statusMsg <> "" Then
                'Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>" & SecUtil.jsString(SecUtil.htmlEncodeNull(statusMsg)) & "</font>"";</script>")
                Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & SecUtil.jsString(SecUtil.htmlEncodeNull(statusMsg)) & "';</script>")
            Else
                'Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload fail! Please check your excel file data.</font>"";</script>")
                Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='Upload fail! Please check your excel file data.';</script>")
            End If

            Response.Flush()

            Return 0
        End Try
    End Function

    Protected Function validateFile_ARIIX(ByVal batch_id As String) As Boolean
        Dim selectSQL, getvalueSQL, updateSQL As String
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim selectDT, tempDT As DataTable

        Dim successflag As Boolean = False

        Dim tempSQL As String = ""

        Dim cnn As SqlConnection
        cnn = gDB.getConnection()

        Dim oTrans As SqlTransaction
        oTrans = cnn.BeginTransaction()
        Try
            paP = New GlobalDBFunc.DBCmdPara
            selectSQL = "Select * from wms_IMP_CO_data where WICO_BATCH_ID=" & paP.AP(batch_id) & " order by WICO_LINE_NO"

            selectDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

            If selectDT.Rows.Count > 0 Then
                Dim hasError As Boolean = False
                Dim errMsg As String = ""

                Dim str_code As String = ""
                Dim WICO_SHIP_MODE, WICO_DELI_DATE, WICO_SKU, WICO_ITM_NAME, WICO_BATCH_NO, WICO_PALLET_NO, WICO_QTY, WICO_VND_CODE,
                    WICO_UOM, WICO_REF_NO, WICO_PCS_PER_UOM As String
                Dim WICO_CUS_CODE, WICO_CUS_NAME, WICO_CUS_CONT_TEL, WICO_CUS_CONT, WICO_REM, WICO_ADDR, WIDO_FTRACK_NO, WIDO_CONF_DELI_DATE, WIDO_DELI_STATUS, WIDO_DELI_RMKS As String

                Dim COD_ITM_CODE, COD_PACK_KEY As String
                COD_ITM_CODE = ""
                COD_PACK_KEY = ""

                Dim WICO_LINE_NO As Integer
                Dim tempUOM, tempPCSperUOM As String

                Dim tempcount As Integer = 0

                For i = 0 To selectDT.Rows.Count - 1
                    hasError = False
                    errMsg = ""

                    str_code = STORER_CODE.SelectedValue
                    WICO_SHIP_MODE = selectDT.Rows(i).Item("WICO_SHIP_MODE").ToString.Trim
                    WICO_DELI_DATE = selectDT.Rows(i).Item("WICO_DELI_DATE").ToString.Trim
                    WICO_SKU = selectDT.Rows(i).Item("WICO_SKU").ToString.Trim.ToUpper
                    WICO_ITM_NAME = selectDT.Rows(i).Item("WICO_ITM_NAME").ToString.Trim
                    WICO_BATCH_NO = selectDT.Rows(i).Item("WICO_BATCH_NO").ToString.Trim
                    WICO_PALLET_NO = selectDT.Rows(i).Item("WICO_PALLET_NO").ToString.Trim
                    WICO_QTY = selectDT.Rows(i).Item("WICO_QTY").ToString.Trim
                    WICO_UOM = selectDT.Rows(i).Item("WICO_UOM").ToString.Trim
                    tempUOM = selectDT.Rows(i).Item("WICO_UOM").ToString.Trim
                    WICO_REF_NO = selectDT.Rows(i).Item("WICO_REF_NO").ToString.Trim
                    WICO_PCS_PER_UOM = selectDT.Rows(i).Item("WICO_PCS_PER_UOM").ToString.Trim
                    tempPCSperUOM = selectDT.Rows(i).Item("WICO_PCS_PER_UOM").ToString.Trim
                    WICO_VND_CODE = selectDT.Rows(i).Item("WICO_VND_CODE").ToString.Trim

                    WICO_CUS_CODE = selectDT.Rows(i).Item("WICO_CUS_CODE").ToString.Trim
                    WICO_CUS_NAME = selectDT.Rows(i).Item("WICO_CUS_NAME").ToString.Trim
                    WICO_CUS_CONT_TEL = selectDT.Rows(i).Item("WICO_CUS_CONT_TEL").ToString.Trim
                    WICO_CUS_CONT = selectDT.Rows(i).Item("WICO_CUS_CONT").ToString.Trim
                    WICO_REM = selectDT.Rows(i).Item("WICO_REM").ToString.Trim
                    WICO_ADDR = selectDT.Rows(i).Item("WICO_ADDR").ToString.Trim
                    WIDO_FTRACK_NO = selectDT.Rows(i).Item("WIDO_FTRACK_NO").ToString.Trim
                    WIDO_CONF_DELI_DATE = selectDT.Rows(i).Item("WIDO_CONF_DELI_DATE").ToString.Trim
                    WIDO_DELI_STATUS = selectDT.Rows(i).Item("WIDO_DELI_STATUS").ToString.Trim
                    WIDO_DELI_RMKS = selectDT.Rows(i).Item("WIDO_DELI_RMKS").ToString.Trim


                    WICO_LINE_NO = gU.decodeEmptyCInt(selectDT.Rows(i).Item("WICO_LINE_NO").ToString.Trim, 0)

                    If WIDO_CONF_DELI_DATE <> "" AndAlso Not gU.isValidDate(WIDO_CONF_DELI_DATE, "d/M/yyyy") Then
                        hasError = True
                        errMsg &= "Invalid Delivery Date || "
                    End If

                    If WICO_QTY <> "" Then
                        If Not gU.IsWholeNumber(WICO_QTY) Then
                            hasError = True
                            errMsg &= "Invalid Quantity || "
                        End If
                    End If

                    If WICO_SKU <> "" Then
                        paP = New GlobalDBFunc.DBCmdPara

                        tempSQL = ""
                        'If WICO_UOM <> "" Then
                        '    tempSQL = " AND ITM_UOM=" & paP.AP(WICO_UOM.ToUpper)
                        'End If

                        'If WICO_PCS_PER_UOM <> "" Then
                        '    If gU.IsWholeNumber(WICO_PCS_PER_UOM) Then
                        '        tempSQL = " AND ITM_PCS_PER_UOM=" & paP.AP(WICO_PCS_PER_UOM)
                        '    End If
                        'End If

                        selectSQL = "select itm_code,itm_name, pack_key, ITM_UOM, ITM_PCS_PER_UOM from wms_item where itm_status <> 'CANCELLED' AND imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(str_code) & " and upper(ITM_SKU_NO)=" & paP.AP(WICO_SKU.ToUpper) & tempSQL
                        Dim tempshow As String = gDB.getCmdSql(selectSQL, paP)
                        tempDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

                        If tempDT.Rows.Count > 0 Then
                            COD_ITM_CODE = tempDT.Rows(0).Item("itm_code").ToString.Trim
                            COD_PACK_KEY = tempDT.Rows(0).Item("pack_key").ToString.Trim
                            WICO_PCS_PER_UOM = tempDT.Rows(0).Item("ITM_PCS_PER_UOM").ToString.Trim
                            WICO_UOM = tempDT.Rows(0).Item("ITM_UOM").ToString.Trim
                            WICO_ITM_NAME = tempDT.Rows(0).Item("ITM_NAME").ToString.Trim

                        Else
                            hasError = True
                            errMsg &= "Cannot Locate the item || "

                        End If
                    Else
                        hasError = True
                        errMsg &= "Empty Stock No./ || "
                    End If

                    If hasError Then
                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "update WMS_IMP_CO_DATA set WICO_IMP_ALLOW_YN='N'," & _
                                    " WICO_IMP_REMARKS=" & paP.AP("ROW No.(" & WICO_LINE_NO & ") " & Left(errMsg, Len(errMsg) - 3)) & _
                                    " where WICO_BATCH_ID=" & paP.AP(WICO_BATCH_ID) & " AND WICO_LINE_NO=" & paP.AP(WICO_LINE_NO)
                        gDB.amendData(updateSQL, cnn, oTrans, paP)
                    Else
                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "update WMS_IMP_CO_DATA set WICO_IMP_ALLOW_YN='Y'," & _
                                    "COD_ITM_CODE=" & paP.AP(COD_ITM_CODE) & "," & _
                                    "COD_PACK_KEY=" & paP.AP(COD_PACK_KEY) & "," & _
                                    "WICO_PCS_PER_UOM=" & paP.AP(WICO_PCS_PER_UOM) & "," & _
                                    "WICO_UOM=" & paP.AP(WICO_UOM) & "," & _
                                    "WICO_ITM_NAME=" & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(WICO_ITM_NAME, ""))) & "," & _
                                    "WICO_VND_CODE=" & paP.AP(WICO_VND_CODE) & _
                                    " where WICO_BATCH_ID=" & paP.AP(WICO_BATCH_ID) & " AND WICO_LINE_NO=" & paP.AP(WICO_LINE_NO)
                        gDB.amendData(updateSQL, cnn, oTrans, paP)

                        'COD_ITM_CODE
                    End If

                Next

                getvalueSQL = "Select Count(*) from WMS_IMP_CO_DATA where WICO_BATCH_ID='" & gU.dbEncode(WICO_BATCH_ID) & "' and WICO_IMP_ALLOW_YN='N'"
                tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)


                If tempcount > 0 Then
                    successflag = False
                Else
                    hasError = False
                    errMsg = ""

                    'getvalueSQL = "Select Count(distinct upper(WICO_REF_NO)) from wms_IMP_CO_data where WICO_BATCH_ID='" & gU.dbEncode(WICO_BATCH_ID) & "' and WICO_IMP_ALLOW_YN='Y' "
                    'tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    'If tempcount > 1 Then
                    '    hasError = True
                    '    errMsg &= "More than 1 REF# Exists || "
                    'End If

                    'getvalueSQL = "Select Count(distinct upper(WICO_SHIP_MODE)) from wms_IMP_CO_data where WICO_BATCH_ID='" & gU.dbEncode(WICO_BATCH_ID) & "' and WICO_IMP_ALLOW_YN='Y' "
                    'tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    'If tempcount > 1 Then
                    '    hasError = True
                    '    errMsg &= "More than 1 Ship Mode Exists || "
                    'End If

                    'getvalueSQL = "Select Count(distinct upper(WICO_DELI_DATE)) from wms_IMP_CO_data where WICO_BATCH_ID='" & gU.dbEncode(WICO_BATCH_ID) & "' and WICO_IMP_ALLOW_YN='Y' "
                    'tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    'If tempcount > 1 Then
                    '    hasError = True
                    '    errMsg &= "More than 1 Delivery Date Exists || "
                    'End If

                    If hasError Then
                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "update WMS_IMP_CO_DATA set WICO_IMP_ALLOW_YN='N'," & _
                                    " WICO_IMP_REMARKS=" & paP.AP("ROW No.(" & WICO_LINE_NO & ") " & Left(errMsg, Len(errMsg) - 3)) & _
                                    " where WICO_BATCH_ID=" & paP.AP(WICO_BATCH_ID) & " AND WICO_LINE_NO=1"
                        gDB.amendData(updateSQL, cnn, oTrans, paP)

                        successflag = False
                    Else
                        successflag = True
                    End If

                End If
            Else
                paP = New GlobalDBFunc.DBCmdPara
                updateSQL = "update WMS_IMP_CO_LOG set WICO_END_DATE=Getdate(), WICO_remarks='No Import detail Data is found' where WICO_BATCH_ID=" & paP.AP(batch_id)
                gDB.amendData(updateSQL, cnn, oTrans, paP)


                successflag = False
            End If


            oTrans.Commit()

        Catch ex As Exception
            oTrans.Rollback()
            cnn.Close()

            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>Terminated.</font>"";</script>")
            Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & SecUtil.jsString(SecUtil.htmlEncodeNull(ex.Message)) & "';</script>")
            Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload fail! Please Check your excel file data.</font>"";</script>")
            Response.Write("<script language=""JavaScript"">alert(""Upload fail! Please Check you excel file data."");</script>")

            Return False

        Finally
            If cnn IsNot Nothing Then
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                    cnn.Dispose()
                End If
            End If
        End Try

        Return successflag
    End Function

    Protected Function ImportData_ARIIX(ByVal batch_id As String) As Boolean
        Dim successFlag As Boolean = False

        Dim cnn As SqlConnection
        cnn = gDB.getConnection()

        Dim oTrans As SqlTransaction
        oTrans = cnn.BeginTransaction()

        Dim selectSQL, getvalueSQL, updateSQL, insertSQL As String
        Dim selectDT, tempDT As DataTable

        Dim paP As GlobalDBFunc.DBCmdPara
        Try
            paP = New GlobalDBFunc.DBCmdPara
            selectSQL = "Select * from WMS_IMP_CO_DATA where WICO_BATCH_ID=" & paP.AP(batch_id) & " AND WICO_IMP_ALLOW_YN='Y' "

            selectDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

            If selectDT.Rows.Count > 0 Then
                Dim PreCOGP As String = ""
                Dim tot_pcs As Double = 0
                Dim tempCount As Integer = 0
                Dim CO_EDI_SIR_NO As String = ""
                Dim existsCount As Integer = 0

                Dim tempStr As String = ""
                Dim temp_COD_SEQ As Integer = 0

                Dim WIDO_FTRACK_NO, WIDO_CONF_DELI_DATE, WIDO_DELI_STATUS, WIDO_DELI_RMKS, WICO_CUS_CONT As String


                Dim newCreate As Boolean = False

                For i = 0 To selectDT.Rows.Count - 1
                    If PreCOGP <> selectDT.Rows(i).Item("CO_CODE").ToString.Trim Then
                        PreCOGP = selectDT.Rows(i).Item("CO_CODE").ToString.Trim

                        CO_EDI_SIR_NO = selectDT.Rows(i).Item("WICO_REF_NO").ToString.Trim

                        paP = New GlobalDBFunc.DBCmdPara
                        getvalueSQL = "Select count(*) as count from WMS_CUST_ORDER where imp_code=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                      " AND CO_STATUS <> 'CANCELLED' AND UPPER(CO_EDI_SIR_NO)=" & paP.AP(CO_EDI_SIR_NO.ToUpper)

                        existsCount = gU.decodeEmptyCInt(DB.getValueFromSQL(getvalueSQL, cnn, oTrans, paP), 0)

                        If existsCount = 0 Then newCreate = True Else newCreate = False

                        If newCreate Then
                            Dim co_track_no As String = ""
                            CO_CODE = DB.getDocNo("CO", cnn, oTrans)
                            co_track_no = DB.getDocNo("TRACKNO", cnn, oTrans)

                            temp_COD_SEQ = 0

                            Dim WICO_SHIP_MODE As String = ""
                            Dim WICO_REF_NO As String = selectDT.Rows(i).Item("WICO_REF_NO").ToString.Trim
                            Dim WICO_DELI_DATE As String = ""
                            Dim WICO_PROVINCE, WICO_CITY, WICO_SENDER, WICO_SENDER_COUNTRY, WICO_SENDER_PROVINCE, WICO_SENDER_REGION, WICO_SENDER_ADDR, WICO_SENDER_TEL, WICO_REGION, WICO_COUNTRY As String

                            WICO_PROVINCE = selectDT.Rows(i).Item("WICO_PROVINCE").ToString.Trim
                            WICO_CITY = selectDT.Rows(i).Item("WICO_CITY").ToString.Trim
                            WICO_SENDER = selectDT.Rows(i).Item("WICO_SENDER").ToString.Trim
                            WICO_SENDER_COUNTRY = selectDT.Rows(i).Item("WICO_SENDER_COUNTRY").ToString.Trim
                            WICO_SENDER_PROVINCE = selectDT.Rows(i).Item("WICO_SENDER_PROVINCE").ToString.Trim
                            WICO_SENDER_REGION = selectDT.Rows(i).Item("WICO_SENDER_REGION").ToString.Trim
                            WICO_SENDER_ADDR = selectDT.Rows(i).Item("WICO_SENDER_ADDR").ToString.Trim
                            WICO_SENDER_TEL = selectDT.Rows(i).Item("WICO_SENDER_TEL").ToString.Trim
                            WICO_REGION = selectDT.Rows(i).Item("WICO_REGION").ToString.Trim
                            WICO_COUNTRY = selectDT.Rows(i).Item("WICO_COUNTRY").ToString.Trim
                            WICO_CUS_CONT = selectDT.Rows(i).Item("WICO_CUS_CONT").ToString.Trim

                            paP = New GlobalDBFunc.DBCmdPara
                            insertSQL = "insert into wms_cust_order (" & _
                                        "co_code, imp_code, storer_code, co_status,  " & _
                                        "co_date, CO_TARGET_DELDATE, CO_EDI_SIR_NO,  co_track_no, CO_SHIP_MODE," & _
                                        "CO_PROVINCE,CO_CITY,CO_SENDER,CO_SENDER_COUNTRY,CO_SENDER_PROVINCE,CO_SENDER_REGION,CO_SENDER_ADDR,CO_SENDER_TEL,CO_REGION_DEL," & _
                                        "sys_cb, sys_cd, sys_lub, sys_lud) Values (" & _
                                        paP.AP(CO_CODE) & "," & paP.AP(Session("imp_code")) & "," & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & ",'NEW'," & _
                                        "Getdate(),Convert(datetime, " & paP.AP(WICO_DELI_DATE) & "," & DDFormat & ")," & paP.AP(WICO_REF_NO) & "," & paP.AP(co_track_no) & "," & paP.AP(WICO_SHIP_MODE) & "," & _
                                        paP.AP(WICO_PROVINCE, SqlDbType.NVarChar) & ", " & paP.AP(WICO_CITY, SqlDbType.NVarChar) & ", " & paP.AP(WICO_SENDER, SqlDbType.NVarChar) & ", " & paP.AP(WICO_SENDER_COUNTRY, SqlDbType.NVarChar) & ", " & paP.AP(WICO_SENDER_PROVINCE, SqlDbType.NVarChar) & ", " & _
                                        paP.AP(WICO_SENDER_REGION, SqlDbType.NVarChar) & ", " & paP.AP(WICO_SENDER_ADDR, SqlDbType.NVarChar) & ", " & paP.AP(WICO_SENDER_TEL) & ", " & _
                                        paP.AP(WICO_REGION, SqlDbType.NVarChar) & ", " & _
                                        "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                            gDB.amendData(insertSQL, cnn, oTrans, paP)


                            If selectDT.Rows(i).Item("WICO_CUS_CODE").ToString.Trim <> "" Then

                                getvalueSQL = "SELECT count(*) from wms_customer where UPPER(cus_code) = '" & selectDT.Rows(i).Item("WICO_CUS_CODE").ToString.Trim.ToUpper & "' " & _
                                            "and storer_code = '" & selectDT.Rows(i).Item("STORER_CODE").ToString.Trim & "' " & _
                                            "and imp_code = '" & Session("IMP_CODE") & "'"

                                existsCount = gU.decodeEmptyCInt(DB.getValueFromSQL(getvalueSQL, cnn, oTrans), 0)

                                If existsCount = 0 Then
                                    paP = New GlobalDBFunc.DBCmdPara
                                    updateSQL = "INSERT INTO wms_customer (IMP_CODE, STORER_CODE, CUS_CODE, CUS_STATUS, CUS_NAME, CUS_ADDR1_DEL,CUS_CONT_TEL_ORD," & _
                                                "cus_region_del, cus_country_del,CUS_CITY_DEL,CUS_PROVINCE_DEL,cus_cont_per_ord," & _
                                                " SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) VALUES (" & _
                                                paP.AP(Session("imp_code")) & "," & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WICO_CUS_CODE").ToString.Trim) & ",'ACTIVE'," & _
                                                paP.AP(selectDT.Rows(i).Item("WICO_CUS_NAME").ToString.Trim, SqlDbType.NVarChar) & "," & paP.AP(selectDT.Rows(i).Item("WICO_ADDR").ToString.Trim, SqlDbType.NVarChar) & "," & _
                                                paP.AP(selectDT.Rows(i).Item("WICO_CUS_CONT_TEL").ToString.Trim) & "," & _
                                                paP.AP(WICO_REGION, SqlDbType.NVarChar) & ", " & paP.AP(WICO_COUNTRY, SqlDbType.NVarChar) & ", " & paP.AP(WICO_CITY, SqlDbType.NVarChar) & ", " & paP.AP(WICO_PROVINCE, SqlDbType.NVarChar) & ", " & _
                                                paP.AP(selectDT.Rows(i).Item("WICO_CUS_CONT").ToString.Trim, SqlDbType.NVarChar) & "," & _
                                                "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                    gDB.amendData(updateSQL, cnn, oTrans, paP)
                                Else
                                    paP = New GlobalDBFunc.DBCmdPara
                                    updateSQL = " UPDATE wms_customer set " & _
                                                " CUS_NAME=" & paP.AP(selectDT.Rows(i).Item("WICO_CUS_NAME").ToString.Trim, SqlDbType.NVarChar) & "," & _
                                                " CUS_ADDR1_DEL=" & paP.AP(selectDT.Rows(i).Item("WICO_ADDR").ToString.Trim, SqlDbType.NVarChar) & "," & _
                                                " CUS_CONT_TEL_ORD=" & paP.AP(selectDT.Rows(i).Item("WICO_CUS_CONT_TEL").ToString.Trim) & "," & _
                                                " cus_region_del=" & paP.AP(WICO_REGION, SqlDbType.NVarChar) & ", " & _
                                                " cus_country_del=" & paP.AP(WICO_COUNTRY, SqlDbType.NVarChar) & "," & _
                                                " CUS_CITY_DEL=" & paP.AP(WICO_CITY, SqlDbType.NVarChar) & "," & _
                                                " CUS_PROVINCE_DEL=" & paP.AP(WICO_PROVINCE, SqlDbType.NVarChar) & "," & _
                                                " cus_cont_per_ord=" & paP.AP(selectDT.Rows(i).Item("WICO_CUS_CONT").ToString.Trim, SqlDbType.NVarChar) & "," & _
                                                " SYS_LUB='" & Session("usr_id") & "', " & _
                                                " SYS_LUD=getdate()" & _
                                                " WHERE UPPER(cus_code) = '" & selectDT.Rows(i).Item("WICO_CUS_CODE").ToString.Trim.ToUpper & "' " & _
                                                " and storer_code = '" & selectDT.Rows(i).Item("STORER_CODE").ToString.Trim & "' " & _
                                                " and imp_code = '" & Session("IMP_CODE") & "'"
                                    gDB.amendData(updateSQL, cnn, oTrans, paP)
                                End If

                                selectSQL = "select cus_name, cus_addr1_del, cus_addr2_del, cus_addr3_del, " & _
                                            "cus_area_del, cus_region_del, cus_country_del,CUS_CITY_DEL,CUS_PROVINCE_DEL, cus_cont_per_ord, cus_cont_tel_ord " & _
                                            "from wms_customer where upper(cus_code) = '" & selectDT.Rows(i).Item("WICO_CUS_CODE").ToString.Trim.ToUpper & "' " & _
                                            "and storer_code = '" & selectDT.Rows(i).Item("STORER_CODE").ToString.Trim & "' " & _
                                            "and imp_code = '" & Session("IMP_CODE") & "'"
                                tempDT = gDB.getDataTable(selectSQL, cnn, oTrans)

                                If tempDT.Rows.Count > 0 Then
                                    paP = New GlobalDBFunc.DBCmdPara
                                    updateSQL = " update wms_cust_order set " & _
                                                " CUS_CODE=" & paP.AP(selectDT.Rows(i).Item("WICO_CUS_CODE").ToString.Trim) & ", " & _
                                                " CUS_NAME =" & paP.AP(tempDT.Rows(0).Item("cus_name").ToString.Trim, SqlDbType.NVarChar) & ", " & _
                                                " CO_ADDR1 =" & paP.AP(tempDT.Rows(0).Item("cus_addr1_del").ToString.Trim, SqlDbType.NVarChar) & ", " & _
                                                " CO_ADDR2 =" & paP.AP(tempDT.Rows(0).Item("cus_addr2_del").ToString.Trim, SqlDbType.NVarChar) & ", " & _
                                                " CO_ADDR3 =" & paP.AP(tempDT.Rows(0).Item("cus_addr3_del").ToString.Trim, SqlDbType.NVarChar) & ", " & _
                                                " CO_AREA_DEL =" & paP.AP(tempDT.Rows(0).Item("cus_area_del").ToString.Trim, SqlDbType.NVarChar) & ", " & _
                                                " CO_REGION_DEL =" & paP.AP(tempDT.Rows(0).Item("cus_region_del").ToString.Trim, SqlDbType.NVarChar) & ", " & _
                                                " CO_COUNTRY_DEL =" & paP.AP(tempDT.Rows(0).Item("cus_country_del").ToString.Trim, SqlDbType.NVarChar) & ", " & _
                                                " CO_CITY =" & paP.AP(tempDT.Rows(0).Item("CUS_CITY_DEL").ToString.Trim, SqlDbType.NVarChar) & ", " & _
                                                " CO_PROVINCE =" & paP.AP(tempDT.Rows(0).Item("CUS_PROVINCE_DEL").ToString.Trim, SqlDbType.NVarChar) & ", " & _
                                                " CO_CUS_CONT =" & paP.AP(tempDT.Rows(0).Item("cus_cont_per_ord").ToString.Trim, SqlDbType.NVarChar) & ", " & _
                                                " CO_CUS_CONT_TEL =" & paP.AP(tempDT.Rows(0).Item("cus_cont_tel_ord").ToString.Trim) & _
                                                " where imp_code=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                                " AND co_code=" & paP.AP(CO_CODE)
                                    gDB.amendData(updateSQL, cnn, oTrans, paP)
                                End If

                            End If
                        Else
                            WIDO_FTRACK_NO = selectDT.Rows(i).Item("WIDO_FTRACK_NO").ToString.Trim
                            WIDO_CONF_DELI_DATE = selectDT.Rows(i).Item("WIDO_CONF_DELI_DATE").ToString.Trim

                            If WIDO_FTRACK_NO <> "" Then
                                paP = New GlobalDBFunc.DBCmdPara
                                updateSQL = " update wms_cust_order set " & _
                                            " CO_FTRACK_NO=" & paP.AP(WIDO_FTRACK_NO) & ", " & _
                                            "SYS_LUB='" & Session("usr_id") & "', SYS_LUD = getdate() " & _
                                            " where imp_code=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                            " AND UPPER(CO_EDI_SIR_NO)=" & paP.AP(CO_EDI_SIR_NO.ToUpper)
                                gDB.amendData(updateSQL, cnn, oTrans, paP)
                            End If

                            If WIDO_CONF_DELI_DATE <> "" Then
                                paP = New GlobalDBFunc.DBCmdPara
                                updateSQL = " update wms_cust_order set " & _
                                            " CO_TARGET_DELDATE=" & gU.convdbDate(gU.dbEncode(WIDO_CONF_DELI_DATE)) & ", " & _
                                            " SYS_LUB='" & Session("usr_id") & "', SYS_LUD = getdate() " & _
                                            " where imp_code=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                            " AND UPPER(CO_EDI_SIR_NO)=" & paP.AP(CO_EDI_SIR_NO.ToUpper)
                                gDB.amendData(updateSQL, cnn, oTrans, paP)
                            End If

                        End If

                        paP = New GlobalDBFunc.DBCmdPara
                        getvalueSQL = "Select count(*) as count from WMS_DELV_ORDER where imp_code=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                      " AND UPPER(DO_EDI_SIR_NO)=" & paP.AP(CO_EDI_SIR_NO.ToUpper)

                        existsCount = gU.decodeEmptyCInt(DB.getValueFromSQL(getvalueSQL, cnn, oTrans, paP), 0)

                        If existsCount > 0 Then
                            WIDO_FTRACK_NO = selectDT.Rows(i).Item("WIDO_FTRACK_NO").ToString.Trim
                            WIDO_CONF_DELI_DATE = selectDT.Rows(i).Item("WIDO_CONF_DELI_DATE").ToString.Trim
                            WIDO_DELI_STATUS = selectDT.Rows(i).Item("WIDO_DELI_STATUS").ToString.Trim
                            WIDO_DELI_RMKS = selectDT.Rows(i).Item("WIDO_DELI_RMKS").ToString.Trim

                            tempStr = ""
                            If WIDO_FTRACK_NO <> "" Then
                                tempStr = gU.appendToList(tempStr, "DO_FTRACK_NO='" & gU.dbEncode(WIDO_FTRACK_NO) & "'")
                            End If

                            If WIDO_CONF_DELI_DATE <> "" Then
                                tempStr = gU.appendToList(tempStr, "DO_CONF_DELDATE=convert(datetime, '" & gU.dbEncode(WIDO_CONF_DELI_DATE) & "'," & DDFormat & ")")
                            End If

                            If WIDO_DELI_STATUS <> "" Then
                                tempStr = gU.appendToList(tempStr, "DO_DELI_STATUS='" & gU.dbEncode(WIDO_DELI_STATUS) & "'")
                            End If

                            If WIDO_DELI_RMKS <> "" Then
                                tempStr = gU.appendToList(tempStr, "DO_DELI_RMKS=N'" & gU.dbEncode(WIDO_DELI_RMKS) & "'")
                                tempStr = gU.appendToList(tempStr, "DO_DELIVERY_RMKS=N'" & gU.dbEncode(WIDO_DELI_RMKS) & "'")
                            End If

                            If tempStr <> "" Then
                                paP = New GlobalDBFunc.DBCmdPara
                                updateSQL = "Update WMS_DELV_ORDER SET " & tempStr & _
                                            " where imp_code=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                            " AND UPPER(DO_EDI_SIR_NO)=" & paP.AP(CO_EDI_SIR_NO.ToUpper)
                                gDB.amendData(updateSQL, cnn, oTrans, paP)
                            End If

                        End If
                    End If

                    If newCreate Then
                        tot_pcs = gU.decodeEmptyCdbl(selectDT.Rows(i).Item("WICO_QTY").ToString.Trim, 0) * gU.decodeEmptyCdbl(selectDT.Rows(i).Item("WICO_PCS_PER_UOM").ToString.Trim, 1)
                        paP = New GlobalDBFunc.DBCmdPara
                        temp_COD_SEQ = gU.decodeEmptyCInt(DB.getValueFromSQL("Select ISNULL(max(convert(int,cod_seq)) + 1,1) as cod_seq from wms_cust_order_d where imp_code=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                                " AND co_code=" & paP.AP(CO_CODE), cnn, oTrans, paP), 1)

                        paP = New GlobalDBFunc.DBCmdPara
                        insertSQL = "insert into wms_cust_order_d ( " & _
                                    "CO_CODE, imp_code, storer_code, COD_seq, COD_disp_seq, " & _
                                    "COD_pallet_no, COD_batch_no, COD_itm_code, COD_pack_key, COD_ITM_DESC, " & _
                                    "COD_qty, COD_uom, COD_PCS_UOM, cod_totpcs, " & _
                                    " sys_cb, sys_cd, sys_lub, sys_lud) Values (" & _
                                    paP.AP(CO_CODE) & "," & paP.AP(Session("imp_code")) & "," & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "," & paP.AP(temp_COD_SEQ) & "," & paP.AP(temp_COD_SEQ) & "," & _
                                    paP.AP(gU.decodeNullOrEmpty(selectDT.Rows(i).Item("WICO_PALLET_NO").ToString.Trim, "000")) & "," & paP.AP(selectDT.Rows(i).Item("WICO_BATCH_NO").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("COD_ITM_CODE").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("COD_pack_key").ToString.Trim) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(selectDT.Rows(i).Item("WICO_ITM_NAME").ToString.Trim, ""))) & "," & _
                                    paP.AP(selectDT.Rows(i).Item("WICO_QTY").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WICO_UOM").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WICO_PCS_PER_UOM").ToString.Trim) & "," & paP.AP(tot_pcs) & "," & _
                                     "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "

                        'Dim showSQL As String = gDB.getCmdSql(insertSQL, paP)
                        gDB.amendData(insertSQL, cnn, oTrans, paP)

                        Dim COD_PCS_CARTON As String = ""
                        Dim COD_TOT_WGT As String = ""
                        Dim COD_TOT_CBM As String = ""

                        'If selectDT.Rows(i).Item("WICO_VND_CODE").ToString.Trim <> "" Then
                        paP = New GlobalDBFunc.DBCmdPara
                        selectSQL = "Select AITM_QTY_PER_CTN, AITM_VOL, CARTON_CBM from V_ALT_VEND_ITEM " & _
                                    "where imp_code=" & paP.AP(Session("imp_code")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                    " AND ITM_CODE=" & paP.AP(selectDT.Rows(i).Item("COD_ITM_CODE").ToString.Trim) & " AND pack_key=" & paP.AP(selectDT.Rows(i).Item("COD_PACK_KEY").ToString.Trim)
                        tempDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

                        If tempDT.Rows.Count > 0 Then
                            COD_PCS_CARTON = tempDT.Rows(0).Item("AITM_QTY_PER_CTN").ToString.Trim
                            COD_TOT_WGT = tempDT.Rows(0).Item("AITM_VOL").ToString.Trim
                            COD_TOT_CBM = tempDT.Rows(0).Item("CARTON_CBM").ToString.Trim

                            paP = New GlobalDBFunc.DBCmdPara
                            updateSQL = "Update wms_cust_order_d set " & _
                                        " COD_PCS_CARTON =" & paP.AP(COD_PCS_CARTON) & "," & _
                                        " COD_TOT_WGT =" & paP.AP(COD_TOT_WGT) & "," & _
                                        " COD_TOT_CBM =" & paP.AP(COD_TOT_CBM) & _
                                        " Where imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                                        " AND co_code=" & paP.AP(CO_CODE) & " AND cOD_SEQ=" & paP.AP(temp_COD_SEQ)


                            Dim temp1 As String = gDB.getCmdSql(updateSQL, paP)
                            gDB.amendData(updateSQL, cnn, oTrans, paP)
                        End If

                    End If
                    ' End If



                    updateSQL = "Update WMS_IMP_CO_DATA set WICO_IMP_YN='Y'" & _
                                " where WICO_BATCH_ID=" & gU.dbEncode(batch_id) & " AND WICO_LINE_NO='" & selectDT.Rows(i).Item("WICO_LINE_NO").ToString.Trim & "' "
                    gDB.amendData(updateSQL, cnn, oTrans)
                Next

                paP = New GlobalDBFunc.DBCmdPara
                updateSQL = "Update WMS_IMP_CO_LOG set WICO_END_DATE=Getdate(), WICO_STATUS='D', WICO_REMARKS='Success' where WICO_BATCH_ID=" & paP.AP(batch_id)
                gDB.amendData(updateSQL, cnn, oTrans, paP)

                successFlag = True
            Else
                paP = New GlobalDBFunc.DBCmdPara
                updateSQL = "Update WMS_IMP_CO_LOG set WICO_END_DATE=Getdate(), WICO_STATUS='F', WICO_REMARKS='No import rows' where WICO_BATCH_ID=" & paP.AP(batch_id)
                gDB.amendData(updateSQL, cnn, oTrans, paP)
                successFlag = False
            End If

            oTrans.Commit()

        Catch ex As Exception

            oTrans.Rollback()
            Return False
        Finally
            If cnn IsNot Nothing Then
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                    cnn.Dispose()
                End If
            End If
        End Try

        Return successFlag
    End Function
End Class
