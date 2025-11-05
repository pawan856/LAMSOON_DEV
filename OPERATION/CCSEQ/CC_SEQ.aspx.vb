Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports NPOI.XSSF.UserModel
Imports NPOI.HPSF
Imports NPOI.POIFS.FileSystem
Imports NPOI.SS.Util
Imports NPOI.XSSF.Util


Partial Class OPERATION_CCSEQ_CC_SEQ
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private nFun As New NPOIFuncX

    Private DDFORMAT As String = "103"
    Private moduleAction As String = ""
    Private exceptionEditList As List(Of String)

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        DDFORMAT = gU.getConfig("DDFORMATNO")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        Server.ScriptTimeout = 7200

        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Not IsPostBack Then

            UP_YEAR.SelectedValue = Now.Year.ToString()
            DL_YEAR.SelectedValue = Now.Year.ToString()

            uiFun.load_dropdown(UP_CCLS_WH, "select distinct wh_main_wh from wms_warehouse order by wh_main_wh", "wh_main_wh", "wh_main_wh")
            uiFun.load_dropdown(DL_CCLS_WH, "select distinct wh_main_wh from wms_warehouse order by wh_main_wh", "wh_main_wh", "wh_main_wh")

        End If

    End Sub

    Protected Sub btnDL_Click(sender As Object, e As System.EventArgs) Handles btnDL.Click
        If DL_CCLS_WH.SelectedValue <> "" Then
            Call DownloadData()
        Else
            uiFun.displayMsg(Me, "", "Please select Main Warehouse!", Session("gLang"))
        End If
    End Sub

    Protected Sub DownloadData()
        Dim sqlString As String = ""
        Dim dt As DataTable

        sqlString = "SELECT WMS_CC_LOC_SEQ.CCLS_YEAR, WMS_CC_LOC_SEQ.CCLS_PERIOD,WMS_CC_LOC_SEQ.CCLS_WH, WMS_CC_LOC_SEQ.CCLS_DISPLAY_SEQ, WMS_CC_LOC_SEQ.CCLS_CUSTOM_SEQ, WMS_CC_LOC_SEQ.CCLS_LOC, WMS_WH_BIN.BN_CSMS_CODE " & _
                    "FROM WMS_CC_LOC_SEQ INNER JOIN " & _
                    "WMS_WH_BIN ON WMS_CC_LOC_SEQ.CCLS_LOC = WMS_WH_BIN.LOC_KEY " & _
                    "WHERE WMS_CC_LOC_SEQ.CCLS_YEAR='" & gU.dbEncode(DL_YEAR.SelectedValue) & "' and WMS_CC_LOC_SEQ.CCLS_PERIOD='" & gU.dbEncode(DL_PERIOD.SelectedValue) & "' " & _
                    "AND WMS_CC_LOC_SEQ.CCLS_WH='" & gU.dbEncode(DL_CCLS_WH.SelectedValue) & "'" & _
                    "ORDER BY CCLS_DISPLAY_SEQ "

        dt = gDB.getDataTable(sqlString)

        If dt.Rows.Count > 0 Then

            Dim ExcelWBObj As XSSFWorkbook
            Dim ReportSheet As XSSFSheet
            Dim tmpFileStream As FileStream
            Dim ReadFileStream As FileStream
            Dim tmpltFileName As String

            Dim cellstyle1 As XSSFCellStyle

            Dim i, currRow As Integer

            Dim fullTmplName As String = ""
            'Dim msgLog As New PrgmLog(Cache("SYSP_LOG_DIR"), "PrgmLog" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")

            'tmpltPath = System.Configuration.ConfigurationManager.AppSettings.Item("EXCEL_TMPLT_PATH").ToString
            tmpltFileName = "CC_SEQ_TMPL.xlsx"
            fullTmplName = Server.MapPath(tmpltFileName)

            ReadFileStream = New FileStream(fullTmplName, FileMode.Open, FileAccess.Read)

            Dim sysFileName, sysFilePath, fileFolder As String

            fileFolder = gU.getConfig("SYSP_TEMP_DIR")

            'SYSP_LOG_DIR

            If Not System.IO.Directory.Exists(fileFolder) Then
                System.IO.Directory.CreateDirectory(fileFolder)
            End If

            Try

                ExcelWBObj = New XSSFWorkbook(ReadFileStream)
                'TemplateSheet = ExcelWBObj.GetSheet("TEMPLATE")
                ReportSheet = ExcelWBObj.GetSheetAt(0)

                'Print Summary
                cellstyle1 = ExcelWBObj.CreateCellStyle

                cellstyle1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin
                cellstyle1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin
                cellstyle1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin
                cellstyle1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin

                Dim tempPath As String = gU.getConfig("SYSP_TEMP_DIR")

                nFun.setCellValue(ReportSheet, 0, 1, dt.Rows(0).Item("CCLS_YEAR").ToString.Trim)
                ReportSheet.GetRow(0).GetCell(1).CellStyle = cellstyle1

                nFun.setCellValue(ReportSheet, 1, 1, dt.Rows(0).Item("CCLS_PERIOD").ToString.Trim)
                ReportSheet.GetRow(1).GetCell(1).CellStyle = cellstyle1


                nFun.setCellValue(ReportSheet, 2, 1, dt.Rows(0).Item("CCLS_WH").ToString.Trim)
                ReportSheet.GetRow(2).GetCell(1).CellStyle = cellstyle1

                currRow = 5

                Dim xCell As XSSFCell

                For i = 0 To dt.Rows.Count - 1

                    nFun.setCellValue(ReportSheet, currRow, 0, dt.Rows(i).Item("CCLS_CUSTOM_SEQ").ToString.Trim)
                    nFun.setCellValue(ReportSheet, currRow, 1, dt.Rows(i).Item("CCLS_LOC").ToString.Trim)
                    nFun.setCellValue(ReportSheet, currRow, 2, dt.Rows(i).Item("BN_CSMS_CODE").ToString.Trim)

                    For z = 0 To 2
                        xCell = nFun.getCell(ReportSheet, currRow, z)
                        xCell.CellStyle = cellstyle1
                    Next

                    currRow += 1
                Next

                'ExcelWBObj.RemoveSheetAt(ExcelWBObj.GetSheetIndex("TEMPLATE"))
                ''ExcelWBObj.RemoveName(0)

                sysFileName = DL_YEAR.SelectedValue & "_" & DL_PERIOD.SelectedValue & "_LOCSEQ_" & Now.Date.ToString("ddMMyyyyHHmmssff")
                sysFileName = sysFileName & "." & LCase("XLSX")
                sysFilePath = fileFolder & "\" & sysFileName

                tmpFileStream = New FileStream(sysFilePath, FileMode.Create)
                ExcelWBObj.Write(tmpFileStream)

                tmpFileStream.Close()
                tmpFileStream.Dispose()



                Dim nFile As System.IO.FileInfo = New System.IO.FileInfo(sysFilePath)
                If File.Exists(sysFilePath) Then
                    'Response.Clear()
                    'Response.Buffer = True
                    'Response.AddHeader("Content-Disposition", "attachment; filename=" & sysFileName)
                    'Response.AddHeader("Content-Length", nFile.Length.ToString())
                    'Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    'Response.WriteFile(nFile.FullName)
                    'HttpContext.Current.Response.Flush()
                    'Response.End()
                    ClientScript.RegisterStartupScript(Me.GetType(), "OpenWindow", "window.open('../../FileHandler.ashx?filePath=" & Server.UrlEncode(sysFilePath) & "&fileName=" & Server.UrlEncode(sysFileName) & "', '_blank');", True)
                End If

            Catch ex As Exception
                Throw ex
            Finally
                
                ReadFileStream.Close()
                ReadFileStream.Dispose()
                tmpFileStream = Nothing
                ReadFileStream = Nothing
            End Try
        Else
            uiFun.displayMsg(Me, "", "No data is found for selected Year/Period!\nPlease select other criteria.", Session("gLang"))
        End If

        load_ModalPopupExtender.Hide()
    End Sub

    Protected Sub btnUP_Click(sender As Object, e As System.EventArgs) Handles btnUP.Click
        If ValidateFile() Then
            Dim returnFilePath As String = ""
            Dim returnMsg As String = ""
            If ReadFile(returnFilePath, returnMsg) Then
                returnMsg = ""
                If ImportFile(returnFilePath, returnMsg) Then

                    txtUP_RESULT.Text = "Data Uploaded successfully."
                Else
                    txtUP_RESULT.Text = returnMsg
                End If
            Else
                txtUP_RESULT.Text = returnMsg
            End If
        End If

        load_ModalPopupExtender.Hide()
    End Sub

    Protected Function ValidateFile() As Boolean
        If UP_CCLS_WH.SelectedValue = "" Then
            uiFun.displayMsg(Me, "", "Please select Main Warehouse!", Session("gLang"))
            Return False
        End If

        If Not UP_FILE.HasFile Then
            uiFun.displayMsg(Me, "", "Please select a file to upload!", Session("gLang"))
            Return False

        Else

            Dim fileExt As String = Path.GetExtension(UP_FILE.FileName)

            If fileExt = ".xlsx" And UP_FILE.PostedFile.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" Then

                If UP_FILE.PostedFile.ContentLength <= 5242880 Then


                Else

                    uiFun.displayMsg(Me, "", "File size too large! Maximum upload file size is 5MB.", Session("gLang"))
                    Return False
                End If

            Else

                uiFun.displayMsg(Me, "", "Only excel file is accepted!", Session("gLang"))
                Return False
            End If

        End If

        Return True
    End Function


    Protected Function ReadFile(ByRef returnFilePath As String, ByRef returnMsg As String) As Boolean
        Dim ExcelWBObj As XSSFWorkbook
        Dim ExcelWSObj As XSSFSheet
        Dim ExcelRow As XSSFRow
        Dim ExcelCell As XSSFCell
        'Dim uStream As Stream
        Dim StartRow As Integer = 0
        Dim errMsg, statusMsg As String
        Dim isEmptyRow As Boolean = True
        Dim getValueSql, updateSql As String
        Dim sysFileName, sysFilePath, tmpFilePath, errFilePath As String

        Dim errCount As Integer = 0

        Path.GetExtension(UP_FILE.FileName)
        sysFileName = Session("usr_id") & Now.ToString("yyyyMMddHHmmssfff") & Path.GetExtension(UP_FILE.FileName)
        sysFilePath = gU.getConfig("SYSP_UPLD_DIR") & "\" & sysFileName
        tmpFilePath = gU.getConfig("SYSP_UPLD_DIR") & "\TEMP\" & sysFileName
        errFilePath = gU.getConfig("SYSP_UPLD_DIR") & "\ERROR\" & sysFileName

        errMsg = ""

        Try
            UP_FILE.PostedFile.SaveAs(tmpFilePath)

            Using tFile As New FileStream(tmpFilePath, FileMode.Open, FileAccess.Read)

                ExcelWBObj = New XSSFWorkbook(tFile)
                ExcelWSObj = ExcelWBObj.GetSheetAt(0)

                StartRow = 4

                ExcelRow = ExcelWSObj.GetRow(StartRow)

                If ExcelRow Is Nothing Then
                    errMsg = "Invalid Excel format! Header row is missing."
                    statusMsg = "Upload fail! Please correct data file."

                    returnMsg = errMsg & vbCrLf & statusMsg
                    Return False
                End If

                ''Check format - header 2
                For i = 0 To 2
                    ExcelCell = ExcelRow.GetCell(i)
                    If ExcelCell Is Nothing OrElse ExcelCell.ToString = "" Then
                        errMsg = "Invalid Excel format! Column Header " & CStr(i + 1) & " is missing."
                        statusMsg = "Upload fail! Please correct data file."

                        returnMsg = errMsg & vbCrLf & statusMsg
                        Return False
                    End If
                Next

                ExcelRow = ExcelWSObj.GetRow(StartRow + 1)

                isEmptyRow = True

                'Check if file is empty
                If Not ExcelRow Is Nothing Then
                    For i = 0 To 1
                        ExcelCell = ExcelRow.GetCell(i)
                        If Not ExcelCell Is Nothing AndAlso Trim(ExcelCell.ToString) <> "" Then
                            isEmptyRow = False
                        End If
                    Next
                End If

                If isEmptyRow Then
                    errMsg = "Empty Excel file! No data is found in first row."
                    statusMsg = "Upload skipped! Please input data."

                    returnMsg = errMsg & vbCrLf & statusMsg
                    Return False
                End If

                Dim LastRownum As Integer = ExcelWSObj.LastRowNum


                Dim locSeq As String = ""
                Dim locCode As String = ""

                Dim keyCount As Integer = 0
                StartRow = 5
                Dim readCount As Integer = 0

                For i = StartRow To LastRownum
                    ExcelRow = ExcelWSObj.GetRow(i)

                    locSeq = Left(nFun.getCell(ExcelWSObj, i, 0).ToString, 10)
                    locCode = Left(nFun.getCell(ExcelWSObj, i, 1).ToString, 80)

                    If String.IsNullOrWhiteSpace(locSeq) Then
                        errMsg &= "Row " & i + 1 & ": Empty Location Seq!" & vbCrLf
                        errCount += 1
                    Else
                        If Not gU.isDecimal(locSeq) Then
                            errMsg &= "Row " & i + 1 & ": Location Seq is not numeric!" & vbCrLf
                            errCount += 1

                        End If
                    End If

                    If String.IsNullOrWhiteSpace(locCode) Then
                        errMsg &= "Row " & i + 1 & ": Empty Location Code!" & vbCrLf
                        errCount += 1
                    Else

                        keyCount = gU.decodeEmptyCInt(DB.getValueFromSQL("Select count(*) from WMS_WH_BIN where loc_key='" & gU.dbEncode(locCode) & "'"), 0)

                        If keyCount = 0 Then
                            errMsg &= "Row " & i + 1 & ": Location Code(" & locCode & ") does not Exists!" & vbCrLf
                            errCount += 1
                        End If

                    End If

                    readCount += 1
                 
                Next

            End Using

            If errCount > 0 Then
                File.Move(tmpFilePath, errFilePath)
                returnFilePath = errFilePath

                returnMsg = "Upload Failed!" & vbCrLf & errMsg & vbCrLf & "Please check your data file."
                Return False
            Else
                File.Move(tmpFilePath, sysFilePath)

                returnFilePath = sysFilePath
                returnMsg = ""
            End If


        Catch ex As Exception
            If File.Exists(tmpFilePath) Then
                File.Move(tmpFilePath, errFilePath)

                returnFilePath = errFilePath
            End If

            Return False

        Finally
            ExcelWSObj = Nothing
            ExcelWBObj = Nothing

        End Try
        'Save the uploaded file

        Return True
    End Function

    Protected Function ImportFile(ByVal filePath As String, ByRef returnMsg As String) As Boolean

        If File.Exists(filePath) Then
            Dim ExcelWBObj As XSSFWorkbook
            Dim ExcelWSObj As XSSFSheet
            Dim ExcelRow As XSSFRow
            Dim ExcelCell As XSSFCell
            'Dim uStream As Stream
            Dim StartRow As Integer = 0
            Dim errMsg, statusMsg As String
            Dim isEmptyRow As Boolean = True
            Dim getValueSql, updateSql As String


            Dim gConn As SqlConnection
            gConn = gDB.getConnection()

            Dim oTrans As SqlTransaction
            oTrans = gConn.BeginTransaction()

            Try

                Using tFile As New FileStream(filePath, FileMode.Open, FileAccess.Read)

                    ExcelWBObj = New XSSFWorkbook(tFile)
                    ExcelWSObj = ExcelWBObj.GetSheetAt(0)

                    Dim LastRownum As Integer = ExcelWSObj.LastRowNum


                    Dim locSeq As String = ""
                    Dim locCode As String = ""

                    Dim UpPeriod As String = UP_PERIOD.SelectedValue
                    Dim UpYear As String = UP_YEAR.SelectedValue
                    Dim upMainWH As String = UP_CCLS_WH.SelectedValue
                    Dim keyCount As Integer = 0
                    Dim LocSeqCount As Integer = 0
                    Dim CustomSeq As Integer = 0
                    StartRow = 5

                    updateSql = "Delete from wms_cc_loc_seq where CCLS_YEAR='" & gU.dbEncode(UpYear) & "' AND CCLS_PERIOD='" & gU.dbEncode(UpPeriod) & "' AND CCLS_WH='" & gU.dbEncode(upMainWH) & "'"
                    gDB.amendData(updateSql, gConn, oTrans)

                    For i = StartRow To LastRownum
                        LocSeqCount += 1
                        ExcelRow = ExcelWSObj.GetRow(i)

                        CustomSeq = Left(nFun.getCell(ExcelWSObj, i, 0).ToString, 10)
                        locSeq = LocSeqCount
                        locCode = Left(nFun.getCell(ExcelWSObj, i, 1).ToString, 80)


                        keyCount = gU.decodeEmptyCInt(DB.getValueFromSQL("Select count(*) from wms_cc_loc_seq where CCLS_YEAR='" & gU.dbEncode(UpYear) & "' AND CCLS_PERIOD='" & gU.dbEncode(UpPeriod) & "' AND CCLS_WH='" & gU.dbEncode(upMainWH) & "' and CCLS_DISPLAY_SEQ=" & locSeq, gConn, oTrans), 0)
                        If keyCount = 0 Then
                            updateSql = "Insert into wms_cc_loc_seq (CCLS_YEAR, CCLS_PERIOD, CCLS_WH, CCLS_DISPLAY_SEQ, CCLS_CUSTOM_SEQ, CCLS_LOC, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB) values (" & _
                                        gU.convdbVCData(gU.dbEncode(UpYear)) & "," & gU.convdbVCData(gU.dbEncode(UpPeriod)) & "," & gU.convdbNVCData(gU.dbEncode(upMainWH)) & "," & gU.convdbVCData(gU.dbEncode(locSeq)) & "," & _
                                        gU.convdbVCData(gU.dbEncode(CustomSeq)) & "," & gU.convdbVCData(gU.dbEncode(locCode)) & "," & _
                                        gU.convdbVCData(gU.dbEncode(Session("usr_id"))) & ",getdate(), getdate()," & gU.convdbVCData(gU.dbEncode(Session("usr_id"))) & ")"
                            gDB.amendData(updateSql, gConn, oTrans)

                        Else
                            returnMsg = "Duplicate Sequence is found. Please check ur upload data."
                            oTrans.Rollback()
                            Return False
                        End If

                    Next

                End Using

                oTrans.Commit()
            Catch ex As Exception
                oTrans.Rollback()

                returnMsg = "ERROR:" & ex.Message
                Return False
            Finally
                ExcelWSObj = Nothing
                ExcelWBObj = Nothing

                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try



        Else
            returnMsg = "Uploaded File Not Exists!"
            Return False
        End If


        Return True
    End Function


End Class
