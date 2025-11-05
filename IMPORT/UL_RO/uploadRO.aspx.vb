Imports Microsoft.VisualBasic
Imports System.IO
Imports NPOI.HSSF.UserModel
Imports NPOI.HPSF
Imports NPOI.POIFS.FileSystem
Imports System.Data.SqlClient
Imports System.Data
Imports System.Collections.Generic

Partial Class EXCEL_uploadRO
    Inherits System.Web.UI.Page

    Private SecUtil As New securityUtil
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Protected isSubmit As Boolean = False
    Dim DDFormat As String = ""
    Dim DDFormatNo As String = gU.getConfig("DDFORMATNO")
    Private batchID As String = ""
    Private SuccessCount As Integer = 0
    Private tempCount As Integer = 0
    Private RO_CODE As String = ""
    Private WIRO_BATCH_ID As Integer = 0
    Private nFun As New NPOIFunc
    'Private msgLog As PrgmLog = New PrgmLog(System.Configuration.ConfigurationManager.AppSettings.Item("LOG_PATH").ToString & "\AP_EXCEL", "uploadCFMaster.txt")


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        ar = New AccessRightUtils("UL_RO", Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If
        
        'DDFormat = gU.getConfig("DDFORMAT")
        DDFormat = "YYYY/MM/DD"

        user_id.Value = Session("usr_id")

        If Not IsPostBack Then
            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(RO_WH_CODE, "select WH_CODE, WH_NAME from WMS_WAREHOUSE where imp_code='" & Session("imp_code") & "' ORDER BY 2", "WH_CODE", "WH_NAME", , Session("gSelectLabel"))

            STORER_CODE.SelectedValue = Session("usr_pref_storer")
            Session("WIRO_BATCH_ID") = ""
        End If

    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim fileExt As String
       
        isSubmit = False

        If STORER_CODE.SelectedValue = "" Then
            uiFun.displayMsg(Me, "", "Please select storer!", Session("gLang"))
            Exit Sub
        End If

        If RO_WH_CODE.SelectedValue = "" Then
            uiFun.displayMsg(Me, "", "Please select warehouse!", Session("gLang"))
            Exit Sub
        End If


        If (FileUpload1.HasFile) Then
            fileExt = Path.GetExtension(FileUpload1.FileName)

            If (fileExt = ".xls" or fileExt = ".xlsx") And (FileUpload1.PostedFile.ContentType = "application/vnd.ms-excel"  orelse 
                                                            FileUpload1.PostedFile.ContentType = "application/excel" orelse _
                                                            FileUpload1.PostedFile.ContentType = "application/x-msexcel" orelse _
                                                            FileUpload1.PostedFile.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml")  Then
                If FileUpload1.PostedFile.ContentLength <= 5242880 Then
                    isSubmit = True
                    FileUpload1.Visible = False
                    btnSubmit.Visible = False
                    '                    cust_no.Visible = False
                Else
                    outputSpan.Text = "File size too large! Maximum upload file size is 5MB."
                    outputSpan.ForeColor = Drawing.Color.Red
                    Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "fail", "<script language=""JavaScript"">alert(""File size too large! Maximum upload file size is 5MB."");</script>")
                End If
            Else
                outputSpan.Text = "Incorrect file format! Excel format Only."
                outputSpan.ForeColor = Drawing.Color.Red
                Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "fail", "<script language=""JavaScript"">alert(""Incorrect file format! Excel format Only."");</script>")
            End If
        Else
            outputSpan.Text = "---"
            outputSpan.ForeColor = Drawing.Color.Black
            Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "fail", "<script language=""JavaScript"">alert(""Please select data file."");</script>")
        End If
    End Sub


    Private Function ReadFile() As Integer
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
            If Not IO.Directory.Exists(gU.getConfig("SYSP_UPLD_DIR") & "\TEMP") Then IO.Directory.CreateDirectory(gU.getConfig("SYSP_UPLD_DIR") & "\TEMP")
            errFilePath = gU.getConfig("SYSP_UPLD_DIR") & "\ERROR\" & sysFileName
            If Not IO.Directory.Exists(gU.getConfig("SYSP_UPLD_DIR") & "\ERROR") Then IO.Directory.CreateDirectory(gU.getConfig("SYSP_UPLD_DIR") & "\ERROR")
            FileUpload1.PostedFile.SaveAs(tmpFilePath)

            tFile = New FileStream(tmpFilePath, FileMode.Open, FileAccess.Read)

            ExcelWBObj = New HSSFWorkbook(tFile)
            ExcelWSObj = ExcelWBObj.GetSheetAt(0)
            'ParamWsobj = ExcelWBObj.GetSheet("PARAM")

            ExcelRow = ExcelWSObj.GetRow(0)

            'Check format - header
            If ExcelRow Is Nothing Then
                errMsg = "Invalid Excel format! Header row is missing."
                statusMsg = "Upload fail! Please correct data file."
                Throw New Exception()
            End If

            ExcelCell = ExcelRow.GetCell(1)
            If ExcelCell Is Nothing OrElse ExcelCell.ToString = "" Then
                errMsg = "Invalid Excel data! Missing Project Code."
                statusMsg = "Upload fail! Please correct data file."
                Throw New Exception()
            ElseIf ExcelCell.ToString.Length > 70 Then
                errMsg = "Invalid Excel format! Project Code length too long."
                statusMsg = "Upload fail! Please correct data file."
                Throw New Exception()
                'Else
                '    Dim prj_code As String = ExcelCell.ToString

                '    Dim projCount As Integer = 0

                '    paP = New GlobalDBFunc.DBCmdPara
                '    projCount = gU.decodeEmptyCInt(DB.getValueFromSQL("Select count(*) from WMS_REPLENISH where RO_STATUS <> 'CANCELLED' and imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(STORER_CODE.SelectedValue) & _
                '                " AND UPPER(RO_REF_NO)=UPPER(" & paP.AP(prj_code), , , paP), 0)

                '    If projCount > 0 Then
                '        errMsg = "Invalid Excel format! Project Code already Exists."
                '        statusMsg = "Upload fail! Please try another project code."
                '        Throw New Exception()
                '    End If

            End If

            ExcelRow = ExcelWSObj.GetRow(2)
            ''Check format - header 2
            For i = 0 To 5
                ExcelCell = ExcelRow.GetCell(i)
                If ExcelCell Is Nothing OrElse ExcelCell.ToString = "" Then
                    errMsg = "Invalid Excel format! Column Header " & CStr(i + 1) & " is missing."
                    statusMsg = "Upload fail! Please correct data file."
                    Throw New Exception()
                End If
            Next

            ExcelRow = ExcelWSObj.GetRow(3)

            isEmptyRow = True

            'Check if file is empty
            If Not ExcelRow Is Nothing Then
                For i = 0 To 5
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

            'If ParamWsobj Is Nothing OrElse nFun.getCell(ParamWsobj, 1, 5).ToString <> "RO" Then
            '    errMsg = "Template Type does not match. Please use correct template for upload RO."
            '    statusMsg = "Upload skipped! Please check your template file type."
            '    Throw New Exception()
            'End If

            StartRow = 3

            Dim MaxLengthCol() As Integer = {10, 1000, 14, 80, 10, 1}


            Dim cnn As SqlConnection
            cnn = gDB.getConnection()

            Dim oTrans As SqlTransaction
            oTrans = cnn.BeginTransaction()

            Try

                getValueSql = " select ISNULL(max(convert(int,WIRO_BATCH_ID)), 0) + 1 as max_batch_id " & _
                              " from WMS_IMP_RO_LOG "

                WIRO_BATCH_ID = gU.decodeEmptyCInt(DB.getValueFromSQL(getValueSql, cnn, oTrans), 0)
                Session("WIRO_BATCH_ID") = WIRO_BATCH_ID
                paP = New GlobalDBFunc.DBCmdPara
                updateSql = "Insert into WMS_IMP_RO_LOG (WIRO_BATCH_ID, WIRO_START_DATE, WIRO_STATUS, WIRO_FILE_NAME) Values (" & _
                            paP.AP(WIRO_BATCH_ID) & ",Getdate(),'F'," & paP.AP(sysFileName) & ")"

                gDB.amendData(updateSql, cnn, oTrans, paP)

                Dim tempSEQ As Integer = 0
                Dim str_code As String = STORER_CODE.SelectedValue
                Dim WIRO_SHIP_MODE, WIRO_ETA, WIRO_SKU, WIRO_ITM_NAME, WIRO_BATCH_NO, WIRO_PALLET_NO, WIRO_QTY, WIRO_EXPIRY_DATE, _
                    WIRO_MANU_DATE, WIRO_UOM, WIRO_REF_NO, WIRO_VND_CODE, WIRO_PCS_PER_UOM, WITM_CC As String

                Dim LastRownum As Integer = ExcelWSObj.LastRowNum

                Dim isEmptyCell As Boolean = True

                For i = StartRow To LastRownum
                    ExcelRow = ExcelWSObj.GetRow(i)
                    tempSEQ += 1

                    'Check if row is empty
                    isEmptyCell = True
                    For x = 0 To 5
                        If Not nFun.getCell(ExcelWSObj, i, x) Is Nothing AndAlso nFun.getCell(ExcelWSObj, i, x).ToString <> "" Then
                            isEmptyCell = False
                        End If
                    Next

                    If isEmptyCell Then
                        Exit For
                    End If

                    WIRO_REF_NO = Left(nFun.getCell(ExcelWSObj, 0, 1).ToString, 70)

                    WIRO_SKU = Left(nFun.getCell(ExcelWSObj, i, 0).ToString, MaxLengthCol(0))
                    WIRO_ITM_NAME = Left(nFun.getCell(ExcelWSObj, i, 1).ToString, MaxLengthCol(1))
                    WIRO_QTY = Left(nFun.getCell(ExcelWSObj, i, 2).ToString, MaxLengthCol(2))
                    WIRO_UOM = Left(nFun.getCell(ExcelWSObj, i, 3).ToString, MaxLengthCol(3))
                    WIRO_ETA = Left(nFun.getCell(ExcelWSObj, i, 4).ToString, MaxLengthCol(4))
                    WITM_CC = Left(nFun.getCell(ExcelWSObj, i, 5).ToString, MaxLengthCol(5))
                   
                    REM XXX
                    paP = New GlobalDBFunc.DBCmdPara
                    updateSql = "Insert into wms_imp_ro_data (WIRO_BATCH_ID,WIRO_LINE_NO,WIRO_IMP_ALLOW_YN,WIRO_IMP_YN,WIRO_IMP_MAP_CODE," & _
                                "WIRO_IMP_REMARKS, IMP_CODE, STORER_CODE, WIRO_ETA," & _
                                "WIRO_SKU,WIRO_ITM_NAME,WIRO_QTY,WIRO_UOM, WIRO_REF_NO, WITM_CC,WIRO_WH_CODE) Values(" & _
                                paP.AP(WIRO_BATCH_ID) & "," & paP.AP(tempSEQ) & ",'N','N','',''," & _
                                paP.AP(Session("imp_code")) & "," & paP.AP(STORER_CODE.SelectedValue) & ", " & paP.AP(WIRO_ETA) & "," & paP.AP(WIRO_SKU) & "," & _
                                paP.AP(WIRO_ITM_NAME) & "," & paP.AP(WIRO_QTY) & "," & paP.AP(WIRO_UOM) & "," & paP.AP(WIRO_REF_NO) & "," & paP.AP(WITM_CC) & "," & paP.AP(RO_WH_CODE.SelectedValue) & ")"

                    'Dim tempsql As String = gDB.getCmdSql(updateSql, paP)
                    gDB.amendData(updateSql, cnn, oTrans, paP)


                    If i Mod 5 = 0 Or i = LastRownum Then
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

            Return WIRO_BATCH_ID

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


        Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='black'>Preparing...</font>"";</script>")
        Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Processing</font>"";</script>")
        Response.Flush()

        Dim selectSQL As String = ""
        Dim tempDT As DataTable
        Dim Resultmsg As String = ""
        Dim updateSQL As String = ""

        Try
            WIRO_BATCH_ID = ReadFile()

            If WIRO_BATCH_ID = 0 Then
                Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload Failed.</font>"";</script>")
                '  Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='Upload fail! Please check your excel file data.';</script>")
                ' Response.Write("<script language=""JavaScript"">alert(""Upload fail! Please check your excel file data."");</script>")
                Response.Flush()
                Exit Sub

            Else
                If validateFile(WIRO_BATCH_ID) Then
                    If ImportData(WIRO_BATCH_ID) Then
                        'Response.Write("<script language=""JavaScript"">alert('File is uploaded successfully!');</script>")
                        'If Session("usr_type") = "C" Or Session("usr_type") = "T" Then
                        '    SendCustROMail(RO_CODE, STORER_CODE.SelectedValue)
                        'End If


                        Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='black'>-----</font>"";</script>")
                        Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='blue'>DONE.</font>"";</script>")
                        Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='File is uploaded successfully!';</script>")
                        Response.Write("<script language=""JavaScript"">document.getElementById('TR_FILE').style.display='none';</script>")
                        'TR_FILE.Visible = False
                        'TR_FILE.Style.Add("display", "none")
                        If Session("usr_type") = "S" Then
                            Response.Write("<script language=""JavaScript"">showLink('../../INBOUND/RO/ROMain.aspx?FrmUP=Y&storer_code=" & STORER_CODE.SelectedValue & "&ro_code=" & RO_CODE & "');</script>")
                        Else
                            Response.Write("<script language=""JavaScript"">showLink('../../INBOUND/RO/ROMain_C.aspx?FrmUP=Y&storer_code=" & STORER_CODE.SelectedValue & "&ro_code=" & RO_CODE & "');</script>")
                        End If


                        'Response.Write("<script language=""JavaScript"">showLogList();</script>")

                    Else
                        Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>Terminated.</font>"";</script>")
                        Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload Failed.</font>"";</script>")
                        Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & Resultmsg & "';</script>")
                        Response.Write("<script language=""JavaScript"">alert(""Data Insert Failed! Please check your excel file data."");</script>")
                        Response.Flush()
                    End If
                Else
                    Dim strBld As New StringBuilder

                    selectSQL = "Select WIRO_IMP_REMARKS from WMS_IMP_RO_DATA where WIRO_BATCH_ID='" & gU.dbEncode(WIRO_BATCH_ID) & "' and WIRO_IMP_ALLOW_YN='N'"
                    tempDT = gDB.getDataTable(selectSQL)

                    If tempDT.Rows.Count > 0 Then
                        For i = 0 To tempDT.Rows.Count - 1
                            strBld.Append(tempDT.Rows(i).Item("WIRO_IMP_REMARKS").ToString.Trim & "\n")
                        Next
                    End If

                    Resultmsg = strBld.ToString

                    updateSQL = "Update WMS_IMP_RO_LOG set WIRO_END_DATE=Getdate(), WIRO_STATUS='F', WIRO_REMARKS='Validation Failed'"
                    gDB.amendData(updateSQL)

                    Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>Terminated.</font>"";</script>")
                    Response.Write("<script language=""JavaScript"">statusSpan.innerHTML = ""<font color='red'>Upload Failed.</font>"";</script>")
                    If Resultmsg <> "" then Response.Write("<script language=""JavaScript"">document.getElementById('up_result').value='" & Resultmsg & "';</script>")
                    Response.Write("<script language=""JavaScript"">alert(""Data Validation failed! Please check your excel file data."");</script>")
                    Response.Flush()
                End If

            End If

        Catch ex As Exception
            Response.Write("<script language=""JavaScript"">outputSpan.innerHTML = ""<font color='red'>Terminated.</font>"";</script>")
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
            selectSQL = "Select * from wms_imp_ro_data where WIRO_BATCH_ID=" & paP.AP(batch_id) & " order by WIRO_LINE_NO"

            selectDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

            If selectDT.Rows.Count > 0 Then
                Dim hasError As Boolean = False
                Dim errMsg As String = ""

                Dim str_code, WIRO_SHIP_MODE, WIRO_ETA, WIRO_SKU, WIRO_ITM_NAME, WIRO_BATCH_NO, WIRO_PALLET_NO, WIRO_QTY, WIRO_EXPIRY_DATE, _
                    WIRO_MANU_DATE, WIRO_UOM, WIRO_REF_NO, WIRO_VND_CODE, WIRO_PCS_PER_UOM As String
                Dim ROD_ITM_CODE, ROD_PACK_KEY As String
                ROD_ITM_CODE = ""
                ROD_PACK_KEY = ""

                Dim WIRO_LINE_NO As Integer

                Dim tempcount As Integer = 0
                Dim tempUOM, tempPCSperUOM As String

                For i = 0 To selectDT.Rows.Count - 1
                    hasError = False
                    errMsg = ""


                    'WIRO_REF_NO = Left(nFun.getCell(ExcelWSObj, 0, 1).ToString, 70)

                    'WIRO_SKU = Left(nFun.getCell(ExcelWSObj, i, 0).ToString, MaxLengthCol(0))
                    'WIRO_ITM_NAME = Left(nFun.getCell(ExcelWSObj, i, 1).ToString, MaxLengthCol(1))
                    'WIRO_QTY = Left(nFun.getCell(ExcelWSObj, i, 2).ToString, MaxLengthCol(2))
                    'WIRO_UOM = Left(nFun.getCell(ExcelWSObj, i, 3).ToString, MaxLengthCol(3))
                    'WIRO_ETA = Left(nFun.getCell(ExcelWSObj, i, 4).ToString, MaxLengthCol(4))
                    'WITM_CC = Left(nFun.getCell(ExcelWSObj, i, 5).ToString, MaxLengthCol(5))

                    str_code = STORER_CODE.SelectedValue

                    WIRO_ETA = selectDT.Rows(i).Item("WIRO_ETA").ToString.Trim
                    WIRO_SKU = selectDT.Rows(i).Item("WIRO_SKU").ToString.Trim.ToUpper
                    WIRO_ITM_NAME = selectDT.Rows(i).Item("WIRO_ITM_NAME").ToString.Trim
                    WIRO_REF_NO = selectDT.Rows(i).Item("WIRO_REF_NO").ToString.Trim
                    WIRO_QTY = selectDT.Rows(i).Item("WIRO_QTY").ToString.Trim
                    WIRO_UOM = selectDT.Rows(i).Item("WIRO_UOM").ToString.Trim
                    tempUOM = selectDT.Rows(i).Item("WIRO_UOM").ToString.Trim

                    WIRO_LINE_NO = gU.decodeEmptyCInt(selectDT.Rows(i).Item("WIRO_LINE_NO").ToString.Trim, 0)

                    'If WIRO_SHIP_MODE <> "" Then

                    '    getvalueSQL = "select count(*) from wms_col_code where COLC_TABCOL='WMS_DELV_ORDER.DO_SHIP_MODE' and upper(COLC_CODE)=upper('" & gU.dbEncode(WIRO_SHIP_MODE) & "')"
                    '    tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    '    If tempcount <= 0 Then
                    '        hasError = True
                    '        errMsg &= "Invalid Ship Mode || "
                    '    End If

                    'End If

                    If WIRO_SKU = "" Then
                        hasError = True
                        errMsg &= "Empty Item No || "
                    End If

                    If WIRO_ITM_NAME = "" Then
                        hasError = True
                        errMsg &= "Empty Item Name || "
                    End If

                    'If WIRO_EXPIRY_DATE <> "" AndAlso Not gU.isValidDate(WIRO_EXPIRY_DATE, DDFormat) Then
                    '    hasError = True
                    '    errMsg &= "Invalid Expiry Date || "
                    'End If

                    'If WIRO_MANU_DATE <> "" AndAlso Not gU.isValidDate(WIRO_MANU_DATE, DDFormat) Then
                    '    hasError = True
                    '    errMsg &= "Invalid Manufactory Date || "
                    'End If

                    If WIRO_UOM <> "" Then
                        getvalueSQL = "select count(*) from wms_UOM where upper(UOM_CODE)='" & gU.dbEncode(WIRO_UOM.ToUpper) & "' AND IMP_CODE='" & gU.dbEncode(Session("imp_code")) & "'"
                        tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                        If tempcount <= 0 Then
                            hasError = True
                            errMsg &= "Invalid UOM || "
                        Else
                            getvalueSQL = "select UOM_CODE from wms_UOM where Upper(UOM_CODE)='" & gU.dbEncode(WIRO_UOM.ToUpper) & "' AND IMP_CODE='" & gU.dbEncode(Session("imp_code")) & "'"
                            tempUOM = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)
                        End If
                    Else
                        hasError = True
                        errMsg &= "Empty UOM || "
                    End If

                    If WIRO_QTY <> "" Then
                        If Not gU.IsWholeNumber(WIRO_QTY) Then
                            hasError = True
                            errMsg &= "Invalid Quantity || "
                        End If
                    Else
                        hasError = True
                        errMsg &= "Empty Quantity || "
                    End If

                    If WIRO_ETA <> "" AndAlso Not gU.isValidDate(WIRO_ETA, DDFormat) Then
                        hasError = True
                        errMsg &= "Invalid ETA || "
                    End If

                    'If WIRO_PCS_PER_UOM <> "" Then
                    '    If Not gU.IsWholeNumber(WIRO_PCS_PER_UOM) Then
                    '        hasError = True
                    '        errMsg &= "Invalid No. of UOM || "
                    '    End If
                    'End If

                    'If WIRO_SKU <> "" AndAlso WIRO_ITM_NAME <> "" Then
                    '    paP = New GlobalDBFunc.DBCmdPara

                    '    tempSQL = ""
                    '    'If WIRO_UOM <> "" Then
                    '    '    tempSQL = " AND ITM_UOM=" & paP.AP(WIRO_UOM.ToUpper)
                    '    'End If

                    '    'If WIRO_PCS_PER_UOM <> "" Then
                    '    '    If gU.IsWholeNumber(WIRO_PCS_PER_UOM) Then
                    '    '        tempSQL = " AND ITM_PCS_PER_UOM=" & paP.AP(WIRO_PCS_PER_UOM)
                    '    '    End If
                    '    'End If

                    '    selectSQL = "select itm_code,itm_name, pack_key, ITM_UOM, ITM_PCS_PER_UOM from wms_item where imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(str_code) & " and upper(ITM_SKU_NO)=" & paP.AP(WIRO_SKU.ToUpper) & " and upper(itm_name)=" & paP.AP(WIRO_ITM_NAME.ToUpper) & tempSQL
                    '    'Dim tempshow As String = gDB.getCmdSql(selectSQL, paP)

                    '    tempDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

                    '    If tempDT.Rows.Count > 0 Then
                    '        ROD_ITM_CODE = tempDT.Rows(0).Item("itm_code").ToString.Trim
                    '        ROD_PACK_KEY = tempDT.Rows(0).Item("pack_key").ToString.Trim
                    '        WIRO_PCS_PER_UOM = tempDT.Rows(0).Item("ITM_PCS_PER_UOM").ToString.Trim
                    '        WIRO_UOM = tempDT.Rows(0).Item("ITM_UOM").ToString.Trim
                    '        WIRO_ITM_NAME = tempDT.Rows(0).Item("ITM_NAME").ToString.Trim

                    '        If tempUOM <> "" AndAlso tempUOM.ToUpper <> WIRO_UOM.ToUpper Then
                    '            hasError = True
                    '            errMsg &= "UOM does not match the item record || "
                    '        End If

                    '        If tempPCSperUOM <> "" AndAlso tempPCSperUOM <> WIRO_PCS_PER_UOM Then
                    '            hasError = True
                    '            errMsg &= "No. per UOM does not match the item record || "
                    '        End If

                    '        If WIRO_VND_CODE <> "" Then
                    '            getvalueSQL = "Select count(*) from WMS_ALT_VEND_ITEM where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(str_code) & "' and itm_code='" & ROD_ITM_CODE & "' and pack_key='" & ROD_PACK_KEY & "' and upper(vnd_code)='" & gU.dbEncode(WIRO_VND_CODE.ToUpper) & "'"
                    '            tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    '            If tempcount <= 0 Then
                    '                hasError = True
                    '                errMsg &= "Vendor Code does not match the item || "
                    '            End If

                    '        Else

                    '            getvalueSQL = "Select count(*) from WMS_ALT_VEND_ITEM where imp_code='" & gU.dbEncode(Session("imp_code")) & "' and storer_code='" & gU.dbEncode(str_code) & "' and itm_code='" & ROD_ITM_CODE & "' and pack_key='" & ROD_PACK_KEY & "' and upper(vnd_code)='DEF_VEND' "
                    '            tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    '            If tempcount > 0 Then
                    '                WIRO_VND_CODE = "DEF_VEND"
                    '            End If

                    '        End If

                    '    Else
                    '        hasError = True
                    '        errMsg &= "Cannot Locate the item || "

                    '    End If
                    'Else
                    '    hasError = True
                    '    errMsg &= "Empty SKU No./ Item Name || "
                    'End If

                    If hasError Then
                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "update WMS_IMP_RO_DATA set WIRO_IMP_ALLOW_YN='N'," & _
                                    " WIRO_IMP_REMARKS=" & paP.AP("ROW No.(" & WIRO_LINE_NO & ") " & Left(errMsg, Len(errMsg) - 3)) & _
                                    " where WIRO_BATCH_ID=" & paP.AP(WIRO_BATCH_ID) & " AND WIRO_LINE_NO=" & paP.AP(WIRO_LINE_NO)
                        gDB.amendData(updateSQL, cnn, oTrans, paP)
                    Else
                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "update WMS_IMP_RO_DATA set WIRO_IMP_ALLOW_YN='Y', " & _
                                    " WIRO_UOM=" & paP.AP(tempUOM) & _
                                    " where WIRO_BATCH_ID=" & paP.AP(WIRO_BATCH_ID) & " AND WIRO_LINE_NO=" & paP.AP(WIRO_LINE_NO)
                        gDB.amendData(updateSQL, cnn, oTrans, paP)

                        'ROD_ITM_CODE
                    End If

                Next

                getvalueSQL = "Select Count(*) from WMS_IMP_RO_DATA where WIRO_BATCH_ID='" & gU.dbEncode(WIRO_BATCH_ID) & "' and WIRO_IMP_ALLOW_YN='N'"
                tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)


                If tempcount > 0 Then
                    successflag = False
                Else
                    hasError = False
                    errMsg = ""

                    'getvalueSQL = "Select Count(distinct upper(WIRO_REF_NO)) from wms_imp_ro_data where WIRO_BATCH_ID='" & gU.dbEncode(WIRO_BATCH_ID) & "' and WIRO_IMP_ALLOW_YN='Y' "
                    'tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    'If tempcount > 1 Then
                    '    hasError = True
                    '    errMsg &= "More than 1 REF# Exists || "
                    'End If

                    'getvalueSQL = "Select Count(distinct upper(WIRO_SHIP_MODE)) from wms_imp_ro_data where WIRO_BATCH_ID='" & gU.dbEncode(WIRO_BATCH_ID) & "' and WIRO_IMP_ALLOW_YN='Y' "
                    'tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    'If tempcount > 1 Then
                    '    hasError = True
                    '    errMsg &= "More than 1 Ship Mode Exists || "
                    'End If

                    'getvalueSQL = "Select Count(distinct upper(WIRO_ETA)) from wms_imp_ro_data where WIRO_BATCH_ID='" & gU.dbEncode(WIRO_BATCH_ID) & "' and WIRO_IMP_ALLOW_YN='Y' "
                    'tempcount = DB.getValueFromSQL(getvalueSQL, cnn, oTrans)

                    'If tempcount > 1 Then
                    '    hasError = True
                    '    errMsg &= "More than 1 ETA Exists || "
                    'End If

                    WIRO_REF_NO = selectDT.Rows(0).Item("WIRO_REF_NO").ToString.Trim.Trim

                    paP = New GlobalDBFunc.DBCmdPara
                    tempcount = gU.decodeEmptyCInt(DB.getValueFromSQL("Select count(*) from WMS_REPLENISH where RO_STATUS <> 'CANCELLED' and imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(STORER_CODE.SelectedValue) & _
                                " AND UPPER(RO_REF_NO)=UPPER(" & paP.AP(WIRO_REF_NO) & ")", , , paP), 0)

                    If tempcount > 0 Then

                        paP = New GlobalDBFunc.DBCmdPara
                        tempcount = gU.decodeEmptyCInt(DB.getValueFromSQL(" SELECT COUNT(*) as grCount FROM WMS_GOODSRCV INNER JOIN " & _
                                    " WMS_REPLENISH ON WMS_GOODSRCV.IMP_CODE = WMS_REPLENISH.IMP_CODE AND WMS_GOODSRCV.GR_DOC_NO = WMS_REPLENISH.RO_CODE AND  " & _
                                    " WMS_GOODSRCV.STORER_CODE = WMS_REPLENISH.STORER_CODE AND WMS_REPLENISH.RO_STATUS <> 'CANCELLED' " & _
                                    " WHERE WMS_GOODSRCV.GR_STATUS <> 'CANCELLED' AND UPPER(GR_DOC_NO)=UPPER('" & gU.dbEncode(WIRO_REF_NO) & "') "), 0)

                        If tempcount > 0 Then
                            errMsg = "Project Code already Exists in a Received Order.\n"
                            errMsg &= "Please upload with another project code.   "
                            WIRO_LINE_NO = Nothing
                            hasError = True
                        Else
                            errMsg = "Project Code already Exists in a Exisiting Order.\n"
                            WIRO_LINE_NO = Nothing
                            'TR_FILE.Visible = False
                            'EXP_TR.Visible = True
                            Response.Write("<script language=""JavaScript"">document.getElementById('TR_FILE').style.display='none';</script>")
                            Response.Write("<script language=""JavaScript"">document.getElementById('EXP_TR').style.display='';</script>")
                            Response.Write("<script language=""JavaScript"">document.getElementById('" & BTNOver.ClientID & "').value='Overwrite PO of contract No." & WIRO_REF_NO & "';</script>")
                            'BTNOver
                            hasError = True
                        End If

                    End If
                    

                    If hasError Then
                        paP = New GlobalDBFunc.DBCmdPara
                        updateSQL = "update WMS_IMP_RO_DATA set WIRO_IMP_ALLOW_YN='N'," & _
                                    " WIRO_IMP_REMARKS=" & paP.AP("ROW No.(" & WIRO_LINE_NO & ") " & Left(errMsg, Len(errMsg) - 3)) & _
                                    " where WIRO_BATCH_ID=" & paP.AP(WIRO_BATCH_ID) & " AND WIRO_LINE_NO=1"
                        gDB.amendData(updateSQL, cnn, oTrans, paP)

                        successflag = False
                    Else
                        successflag = True
                    End If



                End If



            Else
                paP = New GlobalDBFunc.DBCmdPara
                updateSQL = "update WMS_IMP_RO_LOG set WIRO_END_DATE=Getdate(), wiro_remarks='No Import detail Data is found' where WIRO_BATCH_ID=" & paP.AP(batch_id)
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

        Dim selectSQL, getvalueSQL, updateSQL, insertSQL As String
        Dim selectDT, tempDT As DataTable

        Dim paP As GlobalDBFunc.DBCmdPara
        Try
            paP = New GlobalDBFunc.DBCmdPara
            selectSQL = "Select * from WMS_IMP_RO_DATA where WIRO_BATCH_ID=" & paP.AP(batch_id) & " AND WIRO_IMP_ALLOW_YN='Y' "

            selectDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

            If selectDT.Rows.Count > 0 Then
                RO_CODE = ""
                Dim ro_track_no As String = ""

                ro_code = DB.getDocNo("RO", cnn, oTrans)
                ro_track_no = DB.getDocNo("TRACKNO", cnn, oTrans)

                Dim WIRO_SHIP_MODE As String = ""
                Dim WIRO_REF_NO As String = ""
                Dim WIRO_ETA As String = ""
                Dim WIRO_WH_CODE As String = ""

                selectSQL = "Select max(WIRO_REF_NO) as WIRO_REF_NO, max(WIRO_WH_CODE) as WIRO_WH_CODE " & _
                            " from WMS_IMP_RO_DATA where WIRO_BATCH_ID='" & gU.dbEncode(batch_id) & "' AND WIRO_IMP_ALLOW_YN='Y' "
                tempDT = gDB.getDataTable(selectSQL, cnn, oTrans)

                If tempDT.Rows.Count > 0 Then
                    'WIRO_SHIP_MODE = tempDT.Rows(0).Item("WIRO_SHIP_MODE").ToString.Trim
                    WIRO_REF_NO = tempDT.Rows(0).Item("WIRO_REF_NO").ToString.Trim
                    'WIRO_ETA = tempDT.Rows(0).Item("WIRO_ETA").ToString.Trim
                    WIRO_WH_CODE = tempDT.Rows(0).Item("WIRO_WH_CODE").ToString.Trim
                End If

                paP = New GlobalDBFunc.DBCmdPara
                insertSQL = "insert into wms_replenish (" & _
                            "ro_code, imp_code, storer_code, ro_status,RO_TYPE, ro_wh_code, " & _
                            "ro_date, ro_ref_no, RO_EDI_PO_NO,  ro_track_no," & _
                            "sys_cb, sys_cd, sys_lub, sys_lud) Values (" & _
                            paP.AP(RO_CODE) & "," & paP.AP(Session("imp_code")) & "," & paP.AP(selectDT.Rows(0).Item("STORER_CODE").ToString.Trim) & ",'NEW','NS'," & paP.AP(WIRO_WH_CODE) & "," & _
                            "cast(Getdate() as date)," & paP.AP(WIRO_REF_NO) & "," & paP.AP(WIRO_REF_NO) & "," & paP.AP(ro_track_no) & "," & _
                             "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                gDB.amendData(insertSQL, cnn, oTrans, paP)

                Dim tot_pcs As Double = 0
                Dim tempCount As Integer = 0

                Dim newItemCode As String = ""
                Dim newSKUNo As String = ""
                Dim itemCC As String = ""

                For i = 0 To selectDT.Rows.Count - 1

                    newItemCode = DB.getDocNo("NONSTOCKITM", cnn, oTrans)
                    newSKUNo = WIRO_REF_NO & "-" & selectDT.Rows(i).Item("WIRO_SKU").ToString.Trim

                    If selectDT.Rows(i).Item("WITM_CC").ToString.Trim.ToUpper = "Y" OrElse selectDT.Rows(i).Item("WITM_CC").ToString.Trim.ToUpper = "YES" Then itemCC = "Y" Else itemCC = "N"

                    insertSQL = " INSERT INTO WMS_ITEM " & _
                                " (IMP_CODE, STORER_CODE, PACK_KEY, ITM_CODE,  ITM_STATUS, ITM_NAME, ITM_DESC, ITM_VEND_CODE, ITM_SKU_NO,ITM_UOM, ITM_PCS_PER_UOM, ITM_NONSTOCK_YN, " & _
                                " PROJ_NO, ITM_CC, SYS_CD, SYS_CB,SYS_LUD, SYS_LUB) " & _
                                " VALUES('" & Session("imp_code") & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ",'1','" & gU.dbEncode(newItemCode) & "','NEW'," & _
                                gU.convdbNVCData(gU.dbEncode(selectDT.Rows(i).Item("WIRO_ITM_NAME").ToString.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(selectDT.Rows(i).Item("WIRO_ITM_NAME").ToString.Trim)) & ",'DEF_VEND'," & gU.convdbNVCData(gU.dbEncode(newSKUNo)) & "," & _
                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(selectDT.Rows(i).Item("WIRO_UOM").ToString.Trim, ""))) & ", 1,'Y', " & _
                                gU.convdbNVCData(gU.dbEncode(WIRO_REF_NO)) & ",'" & itemCC & "'," & _
                                "GetDate(),'" & Session("usr_id") & "',GetDate(),'" & Session("usr_id") & "') "

                    gDB.amendData(insertSQL, cnn, oTrans)

                    insertSQL = " INSERT INTO WMS_ALT_VEND_ITEM " & _
                                " (IMP_CODE, STORER_CODE, PACK_KEY, ITM_CODE, VND_CODE, VND_NAME, ALV_DATE_ADDED, AITM_STATUS," & _
                                " AITM_UOM, AITM_PCS_PER_PACK,AITM_QTY_PER_CTN," & _
                                " sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                " VALUES ('" & Session("imp_code") & "', " & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", '1','" & gU.dbEncode(newItemCode) & "', 'DEF_VEND', 'Default Vendor', getdate(), 'Active'," & _
                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(selectDT.Rows(i).Item("WIRO_UOM").ToString.Trim, ""))) & ", 1, 1," & _
                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                    gDB.amendData(insertSQL, cnn, oTrans)

                    tot_pcs = gU.decodeEmptyCdbl(selectDT.Rows(i).Item("WIRO_QTY").ToString.Trim, 0) '* gU.decodeEmptyCdbl(selectDT.Rows(i).Item("WIRO_PCS_PER_UOM").ToString.Trim, 1)

                    paP = New GlobalDBFunc.DBCmdPara
                    insertSQL = "insert into wms_replenish_d ( " & _
                                "ro_code, imp_code, storer_code, rod_seq, rod_disp_seq," & _
                                "rod_pallet_no, rod_batch_no, rod_itm_code, rod_pack_key, rod_itm_name," & _
                                "rod_qty, rod_uom, rod_pcs_per_uom, rod_tot_pcs,ROD_STATUS,ROD_ON_BEHALF,ROD_WH_CODE, " & _
                                " sys_cb, sys_cd, sys_lub, sys_lud) Values (" & _
                                paP.AP(RO_CODE) & "," & paP.AP(Session("imp_code")) & "," & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WIRO_LINE_NO").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WIRO_LINE_NO").ToString.Trim) & "," & _
                                paP.AP(gU.decodeNullOrEmpty(selectDT.Rows(i).Item("WIRO_PALLET_NO").ToString.Trim, "000")) & "," & paP.AP(selectDT.Rows(i).Item("WIRO_BATCH_NO").ToString.Trim) & "," & paP.AP(newItemCode) & ",1," & paP.AP(selectDT.Rows(i).Item("WIRO_ITM_NAME").ToString.Trim) & "," & _
                                paP.AP(selectDT.Rows(i).Item("WIRO_QTY").ToString.Trim) & "," & paP.AP(gU.decodeNullOrEmpty(selectDT.Rows(i).Item("WIRO_UOM").ToString.Trim, "PCS")) & ",1," & paP.AP(tot_pcs) & ",'NEW','N'," & paP.AP(WIRO_WH_CODE) & "," & _
                                "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                    'Dim showSQL As String = gDB.getCmdSql(insertSQL, paP) ROD_ON_BEHALF
                    gDB.amendData(insertSQL, cnn, oTrans, paP)

                    'Dim ROD_LENGTH As String = ""
                    'Dim ROD_WIDTH As String = ""
                    'Dim ROD_HEIGHT As String = ""
                    'Dim ROD_KG As String = ""
                    'Dim ROD_CBM As String = ""
                    'Dim ROD_VND_CODE As String = ""

                    'If selectDT.Rows(i).Item("WIRO_VND_CODE").ToString.Trim <> "" Then
                    '    paP = New GlobalDBFunc.DBCmdPara
                    '    selectSQL = "Select AITM_LENGTH, AITM_WIDTH, AITM_HIGHT, AITM_VOL, VND_CODE from WMS_ALT_VEND_ITEM " & _
                    '                "where imp_code=" & paP.AP(Session("imp_code")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                    '                " AND ITM_CODE=" & paP.AP(selectDT.Rows(i).Item("ROD_ITM_CODE").ToString.Trim) & " AND pack_key=" & paP.AP(selectDT.Rows(i).Item("ROD_PACK_KEY").ToString.Trim) & " AND upper(VND_CODE)=upper(" & paP.AP(selectDT.Rows(i).Item("WIRO_VND_CODE").ToString.Trim) & ")"
                    '    tempDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

                    '    If tempDT.Rows.Count > 0 Then
                    '        ROD_VND_CODE = tempDT.Rows(0).Item("VND_CODE").ToString.Trim
                    '        ROD_LENGTH = tempDT.Rows(0).Item("AITM_LENGTH").ToString.Trim
                    '        ROD_WIDTH = tempDT.Rows(0).Item("AITM_WIDTH").ToString.Trim
                    '        ROD_HEIGHT = tempDT.Rows(0).Item("AITM_HIGHT").ToString.Trim
                    '        ROD_KG = tempDT.Rows(0).Item("AITM_VOL").ToString.Trim
                    '        ROD_CBM = gU.decodeEmptyCdbl(tempDT.Rows(0).Item("AITM_LENGTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(tempDT.Rows(0).Item("AITM_WIDTH").ToString.Trim, 0) * gU.decodeEmptyCdbl(tempDT.Rows(0).Item("AITM_HIGHT").ToString.Trim, 0) / 1000000

                    '        paP = New GlobalDBFunc.DBCmdPara
                    '        updateSQL = "Update wms_replenish_d set " & _
                    '                    " ROD_VND_CODE=" & paP.AP(ROD_VND_CODE) & ", " & _
                    '                    " ROD_LENGTH =" & paP.AP(ROD_LENGTH) & "," & _
                    '                    " ROD_WIDTH =" & paP.AP(ROD_WIDTH) & "," & _
                    '                    " ROD_HEIGHT =" & paP.AP(ROD_HEIGHT) & "," & _
                    '                    " ROD_KG =" & paP.AP(ROD_KG) & "," & _
                    '                    " ROD_CBM =" & paP.AP(ROD_CBM) & _
                    '                    " Where imp_code=" & paP.AP(Session("imp_code")) & " AND storer_code=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & _
                    '                    " AND ro_code=" & paP.AP(RO_CODE) & " AND ROD_SEQ=" & paP.AP(selectDT.Rows(i).Item("WIRO_LINE_NO").ToString.Trim)

                    '        gDB.amendData(updateSQL, cnn, oTrans, paP)
                    '    End If
                    'End If

                    updateSQL = "Update WMS_IMP_RO_DATA set WIRO_IMP_YN='Y'" & _
                                " where WIRO_BATCH_ID=" & gU.dbEncode(batch_id) & " AND WIRO_LINE_NO='" & selectDT.Rows(i).Item("WIRO_LINE_NO").ToString.Trim & "' "
                    gDB.amendData(updateSQL, cnn, oTrans)


                    'If selectDT.Rows(i).Item("WIRO_BATCH_NO").ToString.Trim <> "" Then
                    '    paP = New GlobalDBFunc.DBCmdPara
                    '    getvalueSQL = "select Count(*) from WMS_DATE_CODE where DC_DATE_CODE=" & paP.AP(selectDT.Rows(i).Item("WIRO_BATCH_NO").ToString.Trim) & " AND imp_code=" & paP.AP(Session("imp_code")) & " AND STORER_CODE=" & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim)
                    '    tempCount = gU.decodeEmptyCdbl(DB.getValueFromSQL(getvalueSQL, cnn, oTrans, paP), 0)

                    '    If tempCount <= 0 Then
                    '        paP = New GlobalDBFunc.DBCmdPara
                    '        insertSQL = "Insert into WMS_DATE_CODE(IMP_CODE,STORER_CODE,DC_DATE_CODE,SYS_CB,SYS_CD,SYS_LUB,SYS_LUD) values (" & _
                    '                    paP.AP(Session("imp_code")) & "," & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WIRO_BATCH_NO").ToString.Trim) & "," & _
                    '                    "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                    '        gDB.amendData(insertSQL, cnn, oTrans, paP)
                    '    End If
                    'End If
                Next

                paP = New GlobalDBFunc.DBCmdPara
                updateSQL = "Update WMS_IMP_RO_LOG set WIRO_END_DATE=Getdate(), WIRO_STATUS='D', WIRO_REMARKS='Success' where WIRO_BATCH_ID=" & paP.AP(batch_id)
                gDB.amendData(updateSQL, cnn, oTrans, paP)

                successFlag = True
            Else
                paP = New GlobalDBFunc.DBCmdPara
                updateSQL = "Update WMS_IMP_RO_LOG set WIRO_END_DATE=Getdate(), WIRO_STATUS='F', WIRO_REMARKS='No import rows' where WIRO_BATCH_ID=" & paP.AP(batch_id)
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

        rmtpost.Url = "uploadRO.aspx"
        rmtpost.Target = "_self"
        rmtpost.Post()

    End Sub

    Protected Sub btnDL_Click(sender As Object, e As System.EventArgs) Handles btnDL.Click
        Dim sysPath As String = Server.MapPath("UPLOAD_NS_TMP.xls")
        Dim nFile As System.IO.FileInfo = New System.IO.FileInfo(sysPath)

        If nFile.Exists Then
            Response.Clear()
            Response.AddHeader("Content-Disposition", "attachment; filename=UPLOAD_NS_TMP" & nFile.Extension)
            Response.AddHeader("Content-Length", nFile.Length.ToString())
            Response.ContentType = "application/octet-stream"
            Response.WriteFile(nFile.FullName)
            Response.End()
        Else
            Response.Write("This file does not exist.")
        End If
    End Sub

    Protected Sub SendCustROMail(ByRef c_code As String, ByRef st_code As String)
        Dim app_email As String = ""
        Dim mail_title As String = "A New Replenishment Order has been created by customer."
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

        mail_body &= "The following Replenishment order has been created:" & vbNewLine
        mail_body &= "  RO Code: " & c_code & vbNewLine
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

    Protected Sub BTNOver_Click(sender As Object, e As System.EventArgs) Handles BTNOver.Click
        Dim batchID As String = ""
        isSubmit = False

        batchID = Session("WIRO_BATCH_ID").ToString

        If batchID <> "" Then
            Dim uRO_code As String = ""
            Dim selectDT As DataTable
            Dim selectSQL As String = ""
            Dim updateSQL As String = ""
            Dim insertSQL As String = ""
            Dim cnn As SqlConnection
            cnn = gDB.getConnection()

            Dim oTrans As SqlTransaction
            oTrans = cnn.BeginTransaction()

            Dim paP As GlobalDBFunc.DBCmdPara

            Try

                paP = New GlobalDBFunc.DBCmdPara
                selectSQL = "Select * from WMS_IMP_RO_DATA where WIRO_BATCH_ID=" & paP.AP(batchID)

                selectDT = gDB.getDataTable(selectSQL, cnn, oTrans, , paP)

                If selectDT.Rows.Count > 0 Then

                    Dim WIRO_REF_NO As String = ""
                    Dim WIRO_WH_CODE As String = ""

                    Dim limp_code, lstorer_code As String

                    WIRO_REF_NO = selectDT.Rows(0).Item("WIRO_REF_NO").ToString.Trim
                    WIRO_WH_CODE = selectDT.Rows(0).Item("WIRO_WH_CODE").ToString.Trim

                    limp_code = selectDT.Rows(0).Item("IMP_CODE").ToString.Trim
                    lstorer_code = selectDT.Rows(0).Item("STORER_CODE").ToString.Trim

                    paP = New GlobalDBFunc.DBCmdPara
                    selectSQL = "Select max(ro_code) as ro_code from WMS_REPLENISH where RO_STATUS <> 'CANCELLED' and imp_code=" & paP.AP(limp_code) & " AND storer_code=" & paP.AP(lstorer_code) & _
                                " AND UPPER(RO_REF_NO)=UPPER(" & paP.AP(WIRO_REF_NO) & ")"

                    uRO_code = DB.getValueFromSQL(selectSQL, cnn, oTrans, paP)

                    'paP = New GlobalDBFunc.DBCmdPara
                    'updateSQL = "delete from wms_item where concat(wms_item.imp_code,wms_item.storer_code,wms_item.itm_code, wms_item.PACK_KEY) in ( " & _
                    '            "select concat(WMS_REPLENISH_D.imp_code, WMS_REPLENISH_D.storer_code, WMS_REPLENISH_D.rod_itm_code, WMS_REPLENISH_D.ROD_PACK_KEY) from WMS_REPLENISH_D where WMS_REPLENISH_D.ro_code=" & paP.AP(uRO_code) & _
                    '            " AND WMS_REPLENISH_D.IMP_CODE=" & paP.AP(limp_code) & " AND WMS_REPLENISH_D.STORER_CODE=" & paP.AP(lstorer_code) & ") "
                    'gDB.amendData(updateSQL, cnn, oTrans, paP)

                    'paP = New GlobalDBFunc.DBCmdPara
                    'updateSQL = "delete from wms_alt_vend_item where concat(wms_alt_vend_item.imp_code,wms_alt_vend_item.storer_code,wms_alt_vend_item.itm_code, wms_alt_vend_item.PACK_KEY) in ( " & _
                    '            "select concat(WMS_REPLENISH_D.imp_code, WMS_REPLENISH_D.storer_code, WMS_REPLENISH_D.rod_itm_code, WMS_REPLENISH_D.ROD_PACK_KEY) from WMS_REPLENISH_D where ro_code=" & paP.AP(uRO_code) & _
                    '            " AND WMS_REPLENISH_D.IMP_CODE=" & paP.AP(limp_code) & " AND WMS_REPLENISH_D.STORER_CODE=" & paP.AP(lstorer_code) & ") "
                    'gDB.amendData(updateSQL, cnn, oTrans, paP)

                    paP = New GlobalDBFunc.DBCmdPara
                    updateSQL = "delete from wms_replenish_d where wms_replenish_d.ro_code=" & paP.AP(uRO_code) & " AND WMS_REPLENISH_D.IMP_CODE=" & paP.AP(limp_code) & " AND WMS_REPLENISH_D.STORER_CODE=" & paP.AP(lstorer_code)
                    gDB.amendData(updateSQL, cnn, oTrans, paP)


                    Dim newItemCode As String = ""
                    Dim newSKUNo As String = ""
                    Dim itemCC As String = ""
                    Dim tot_pcs As Double = 0
                    Dim existCount As Integer = 0

                    For i = 0 To selectDT.Rows.Count - 1
                        newSKUNo = WIRO_REF_NO & "-" & selectDT.Rows(i).Item("WIRO_SKU").ToString.Trim
                        existCount = gU.decodeEmptyCInt(DB.getValueFromSQL("Select count(*) from wms_item where itm_sku_no='" & gU.dbEncode(newSKUNo) & "' and upper(itm_name)='" & gU.dbEncode(selectDT.Rows(i).Item("WIRO_ITM_NAME").ToString.Trim.ToUpper) & "' and storer_code='" & lstorer_code & "' and imp_code='" & limp_code & "'"), 0)

                        If existCount = 0 Then
                            newItemCode = DB.getDocNo("NONSTOCKITM", cnn, oTrans)
                            If selectDT.Rows(i).Item("WITM_CC").ToString.Trim.ToUpper = "Y" OrElse selectDT.Rows(i).Item("WITM_CC").ToString.Trim.ToUpper = "YES" Then itemCC = "Y" Else itemCC = "N"

                            insertSQL = " INSERT INTO WMS_ITEM " & _
                                    " (IMP_CODE, STORER_CODE, PACK_KEY, ITM_CODE,  ITM_STATUS, ITM_NAME, ITM_DESC, ITM_VEND_CODE, ITM_SKU_NO,ITM_UOM, ITM_PCS_PER_UOM, ITM_NONSTOCK_YN, " & _
                                    " PROJ_NO, ITM_CC, SYS_CD, SYS_CB,SYS_LUD, SYS_LUB) " & _
                                    " VALUES('" & Session("imp_code") & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ",'1','" & gU.dbEncode(newItemCode) & "','NEW'," & _
                                    gU.convdbNVCData(gU.dbEncode(selectDT.Rows(i).Item("WIRO_ITM_NAME").ToString.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(selectDT.Rows(i).Item("WIRO_ITM_NAME").ToString.Trim)) & ",'DEF_VEND'," & gU.convdbNVCData(gU.dbEncode(newSKUNo)) & "," & _
                                    gU.convdbNVCData(gU.dbEncode(gU.decodeNull(selectDT.Rows(i).Item("WIRO_UOM").ToString.Trim, ""))) & ", 1,'Y', " & _
                                    gU.convdbNVCData(gU.dbEncode(WIRO_REF_NO)) & ",'" & itemCC & "'," & _
                                    "GetDate(),'" & Session("usr_id") & "',GetDate(),'" & Session("usr_id") & "') "

                            gDB.amendData(insertSQL, cnn, oTrans)

                            insertSQL = " INSERT INTO WMS_ALT_VEND_ITEM " & _
                                        " (IMP_CODE, STORER_CODE, PACK_KEY, ITM_CODE, VND_CODE, VND_NAME, ALV_DATE_ADDED, AITM_STATUS," & _
                                        " AITM_UOM, AITM_PCS_PER_PACK,AITM_QTY_PER_CTN," & _
                                        " sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                        " VALUES ('" & Session("imp_code") & "', " & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", '1','" & gU.dbEncode(newItemCode) & "', 'DEF_VEND', 'Default Vendor', getdate(), 'Active'," & _
                                        gU.convdbNVCData(gU.dbEncode(gU.decodeNull(selectDT.Rows(i).Item("WIRO_UOM").ToString.Trim, ""))) & ", 1, 1," & _
                                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                            gDB.amendData(insertSQL, cnn, oTrans)
                        Else
                            newItemCode = DB.getValueFromSQL("Select itm_code from wms_item where itm_sku_no='" & gU.dbEncode(newSKUNo) & "' and storer_code='" & lstorer_code & "' and imp_code='" & limp_code & "'")

                        End If
                        
                        tot_pcs = gU.decodeEmptyCdbl(selectDT.Rows(i).Item("WIRO_QTY").ToString.Trim, 0) '* gU.decodeEmptyCdbl(selectDT.Rows(i).Item("WIRO_PCS_PER_UOM").ToString.Trim, 1)

                        If newItemCode <> "" Then
                            paP = New GlobalDBFunc.DBCmdPara
                            insertSQL = "insert into wms_replenish_d ( " & _
                                        "ro_code, imp_code, storer_code, rod_seq, rod_disp_seq," & _
                                        "rod_pallet_no, rod_batch_no, rod_itm_code, rod_pack_key, rod_itm_name," & _
                                        "rod_qty, rod_uom, rod_pcs_per_uom, rod_tot_pcs,ROD_STATUS,ROD_ON_BEHALF,ROD_WH_CODE, " & _
                                        " sys_cb, sys_cd, sys_lub, sys_lud) Values (" & _
                                        paP.AP(uRO_code) & "," & paP.AP(Session("imp_code")) & "," & paP.AP(selectDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WIRO_LINE_NO").ToString.Trim) & "," & paP.AP(selectDT.Rows(i).Item("WIRO_LINE_NO").ToString.Trim) & "," & _
                                        paP.AP(gU.decodeNullOrEmpty(selectDT.Rows(i).Item("WIRO_PALLET_NO").ToString.Trim, "000")) & "," & paP.AP(selectDT.Rows(i).Item("WIRO_BATCH_NO").ToString.Trim) & "," & paP.AP(newItemCode) & ",1," & paP.AP(selectDT.Rows(i).Item("WIRO_ITM_NAME").ToString.Trim) & "," & _
                                        paP.AP(selectDT.Rows(i).Item("WIRO_QTY").ToString.Trim) & "," & paP.AP(gU.decodeNullOrEmpty(selectDT.Rows(i).Item("WIRO_UOM").ToString.Trim, "PCS")) & ",1," & paP.AP(tot_pcs) & ",'NEW','N'," & paP.AP(WIRO_WH_CODE) & "," & _
                                        "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "
                            'Dim showSQL As String = gDB.getCmdSql(insertSQL, paP) ROD_ON_BEHALF
                            gDB.amendData(insertSQL, cnn, oTrans, paP)

                            updateSQL = "Update WMS_IMP_RO_DATA set WIRO_IMP_YN='Y'" & _
                                        " where WIRO_BATCH_ID=" & gU.dbEncode(batchID) & " AND WIRO_LINE_NO='" & selectDT.Rows(i).Item("WIRO_LINE_NO").ToString.Trim & "' "
                            gDB.amendData(updateSQL, cnn, oTrans)
                        End If
                    Next

                    paP = New GlobalDBFunc.DBCmdPara
                    updateSQL = "Update WMS_IMP_RO_LOG set WIRO_END_DATE=Getdate(), WIRO_STATUS='D', WIRO_REMARKS='Success' where WIRO_BATCH_ID=" & paP.AP(batchID)
                    gDB.amendData(updateSQL, cnn, oTrans, paP)

                    outputSpan.Text = "-----"
                    outputSpan.ForeColor = Drawing.Color.Black
                    statusSpan.Text = "DONE."

                    up_result.Text = "PO Data Imported Successfully!"
                    TR_FILE.Visible = False
                    EXP_TR.Visible = False
                    Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", "showLink('../../INBOUND/RO/ROMain.aspx?FrmUP=Y&storer_code=" & lstorer_code & "&ro_code=" & uRO_code & "')", True)

                    oTrans.Commit()

                Else
                    oTrans.Rollback()
                    uiFun.displayMsg(Me, "", "No upload Record Found!", Session("gLang"))
                End If

                

            Catch ex As Exception
                oTrans.Rollback()
                uiFun.displayMsg(Me, "", "Update Error. Please try again.", Session("gLang"))
            Finally
                If cnn IsNot Nothing Then
                    If cnn.State = ConnectionState.Open Then
                        cnn.Close()
                        cnn.Dispose()
                    End If
                End If
            End Try

        Else
            uiFun.displayMsg(Me, "", "Update Error.", Session("gLang"))
        End If


    End Sub
End Class
