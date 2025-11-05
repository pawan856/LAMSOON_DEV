Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports NPOI.XSSF.UserModel
Imports NPOI.HPSF
Imports NPOI.POIFS.FileSystem
Imports NPOI.SS.Util
Imports NPOI.XSSF.Util


Partial Class RPT_SLDTL_main
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private nf As New NPOIFuncX

    Protected Const IMG_PATH As String = "../../images/"
    Protected Const ROOT_PATH As String = "../../"
    Protected Const FUN_CODE As String = "RPT_MAST_CCPLAN"

    Protected isGenDownload As Boolean = False
    Private ExcelRow_num As Long = 0

    Dim customDateTimeFormat As DateTimeFormatInfo = New DateTimeFormatInfo()


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils(FUN_CODE, Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        'Response.Write(Session("GENERIC_SESSION_SRCH_COMPLETE_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_WHERE_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_GROUPBY_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_ORDERBY_SQL"))
        'Response.Write("<br><br>")
        'Response.End()
        'Session("GENERIC_SESSION_SRCH_WHERE_SQL")


        If Not IsPostBack Then
            'Response.Write(Session("SEARCH_SESSION_PAGE_FR_DATE_RANGE"))
            'Response.Write(Session("SEARCH_SESSION_PAGE_TO_DATE_RANGE"))
            'Response.Write(Session("SEARCH_SESSION_PAGE_DAY_MONTH"))
            'Response.Write(Session("SEARCH_SESSION_PAGE_TEAM_CODE"))

            isGenDownload = True
            tr_download.Style.Add("display", "none")
            ibtn_download.Visible = True


        End If

        ar.hideForm(Me)
    End Sub

    Protected Sub genDownloadFile()

        Dim result_dt As New DataTable
        Dim total_dt As New DataTable

        Dim ExcelWBObj As XSSFWorkbook
        Dim TemplateSheet, ReportSheet As XSSFSheet


        Dim sql_string As String = ""
        Dim addSQL As String = ""


        Dim tmpltPath, tmpltFileName As String
        Dim tmpFileStream As FileStream
        Dim ReadFileStream As FileStream

        customDateTimeFormat.DateSeparator = "/"
        customDateTimeFormat.TimeSeparator = ":"
        customDateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
        customDateTimeFormat.LongDatePattern = "dd/MM/yyyy"
        customDateTimeFormat.ShortTimePattern = "HH:mm"
        customDateTimeFormat.LongTimePattern = "HH:mm"
        customDateTimeFormat.FullDateTimePattern = "dd/MM/yyyy HH:mm"

        'tmpltPath = System.Configuration.ConfigurationManager.AppSettings.Item("EXCEL_TMPLT_PATH").ToString

        tmpltFileName = "RPT_MAST_CCPLAN.xlsx"
        tmpltPath = Server.MapPath(tmpltFileName)
        ReadFileStream = New FileStream(tmpltPath, FileMode.Open, FileAccess.Read)

        Dim sysFileName, sysFilePath, fileFolder As String

        fileFolder = HttpContext.Current.Cache("SYSP_TEMP_DIR")

        If Not System.IO.Directory.Exists(fileFolder) Then
            System.IO.Directory.CreateDirectory(fileFolder)
        End If


        Try

            sql_string = Session("GENERIC_SESSION_SRCH_COMPLETE_SQL")


            'Dim tempSQL As String = gDB.getCmdSql(sql_string, pa1)
            result_dt = gDB.getDataTable(sql_string)

            If Not String.IsNullOrWhiteSpace(sql_string) AndAlso result_dt.Rows.Count > 0 Then

                ExcelWBObj = New XSSFWorkbook(ReadFileStream)
                'TemplateSheet = ExcelWBObj.GetSheet("TEMPLATE")
                ReportSheet = ExcelWBObj.GetSheetAt(0)

                'XXXXXXXXXXXXXXXXXX Excel Cell StyleXXXXXXXXXXXXXXXXXXXXX
                'Dim warpStyle As XSSFCellStyle
                'warpStyle = ExcelWBObj.CreateCellStyle()
                'warpStyle.WrapText = True

                'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


                'XXXXXXXXXXX Set Param Sheet XXXXXXXXXXXXXXXXXXXXXX


                'nf.setCellValue(ParamSheet, 13, 3, gU.decodeNullOrEmpty(FR_DO_SCH_DATE, Now.Date.ToString("dd/MM/yyyy")))
                'nf.setCellValue(ParamSheet, 13, 8, gU.decodeNullOrEmpty(TO_DO_SCH_DATE, Now.Date.ToString("dd/MM/yyyy")))

                'XXXXXXXXXXXXX TITLE XXXXXXXXXXXXXXXXXXXXXXX

                'nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(0, 0, 0, 17), 0, 0)
                'For i = 0 To 17
                '    ReportSheet.SetColumnWidth(i, TemplateSheet.GetColumnWidth(i))
                'Next


                ''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

                'XXXXXXXXXXXXXXXXXX Column Header XXXXXXXXXXXXXXXXXXXXXXXXXXXXX


                'REM XXXXXXXXXXXXXXXXXXX CONTENT BLOCK XXXXXXXXXXXXXXXXXXX


                Dim colName As String = ""
                Dim PreLoc As String = ""
                Dim TotalItem As Integer = 0
                Dim locArr() As String
                Dim startRow As Integer = 0

                ExcelRow_num = 6

                For i = 0 To result_dt.Rows.Count - 1
                    If PreLoc <> result_dt.Rows(i).Item("locCode").ToString.Trim Then
                        locArr = Split(PreLoc, "||")
                        If locArr.Length = 4 Then

                            nf.setCellValue(ReportSheet, ExcelRow_num, 0, locArr(0))
                            nf.setCellValue(ReportSheet, ExcelRow_num, 1, locArr(1))
                            nf.setCellValue(ReportSheet, ExcelRow_num, 2, locArr(2))
                            nf.setCellValue(ReportSheet, ExcelRow_num, 3, locArr(3))

                            nf.setCellValue(ReportSheet, ExcelRow_num, 4, "No of Item: " & TotalItem)
                            ReportSheet.AddMergedRegion(New CellRangeAddress(ExcelRow_num, ExcelRow_num, 4, 7))

                            ReportSheet.GroupRow(startRow, ExcelRow_num - 1)
                            ReportSheet.SetRowGroupCollapsed(startRow, True)

                            nf.InsertRows(ReportSheet, ExcelRow_num + 1, 1)
                            ExcelRow_num += 1

                        End If
                        startRow = ExcelRow_num
                        PreLoc = result_dt.Rows(i).Item("locCode").ToString.Trim
                        TotalItem = 0
                    End If

                    'WH_MAIN_WH
                    'ILOC_WH
                    'ILOC_FLOOR
                    'ILOC_AREA
                    'ITM_GP_CODE
                    'ITM_SKU_NO
                    'ITM_NAME
                    'ITEM_PRICE_CLASS

                    nf.setCellValue(ReportSheet, ExcelRow_num, 0, result_dt.Rows(i).Item("WH_MAIN_WH").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 1, result_dt.Rows(i).Item("ILOC_WH").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 2, result_dt.Rows(i).Item("ILOC_FLOOR").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 3, result_dt.Rows(i).Item("ILOC_AREA").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 4, result_dt.Rows(i).Item("ITM_GP_CODE").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 5, result_dt.Rows(i).Item("ITM_SKU_NO").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 6, result_dt.Rows(i).Item("ITM_NAME").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 7, result_dt.Rows(i).Item("ITEM_PRICE_CLASS").ToString.Trim)

                    nf.InsertRows(ReportSheet, ExcelRow_num + 1, 1)
                    ExcelRow_num += 1
                    TotalItem += 1

                Next

                locArr = Split(PreLoc, "||")
                If locArr.Length = 4 Then

                    nf.setCellValue(ReportSheet, ExcelRow_num, 0, locArr(0))
                    nf.setCellValue(ReportSheet, ExcelRow_num, 1, locArr(1))
                    nf.setCellValue(ReportSheet, ExcelRow_num, 2, locArr(2))
                    nf.setCellValue(ReportSheet, ExcelRow_num, 3, locArr(3))

                    nf.setCellValue(ReportSheet, ExcelRow_num, 4, "No of Item: " & TotalItem)
                    ReportSheet.AddMergedRegion(New CellRangeAddress(ExcelRow_num, ExcelRow_num, 4, 7))

                    ReportSheet.GroupRow(startRow, ExcelRow_num - 1)
                    ReportSheet.SetRowGroupCollapsed(startRow, True)

                    nf.InsertRows(ReportSheet, ExcelRow_num + 1, 1)
                    ExcelRow_num += 1

                End If

                nf.setCellValue(ReportSheet, 3, 0, "Print Date: " & Now.Date.ToString("dd/MM/yyyy"))

                Dim backColorStyle As XSSFCellStyle
                Dim greyBack As New XSSFColor(Drawing.Color.Silver)
                Dim ifont As New XSSFFont
                backColorStyle = ExcelWBObj.CreateCellStyle
                backColorStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground
                backColorStyle.SetFillForegroundColor(greyBack)

                ifont = ExcelWBObj.CreateFont
                ifont.FontHeightInPoints = 9
                ifont.Boldweight = NPOI.SS.UserModel.FontBoldWeight.Bold
                ifont.FontName = "Arial Narrow"
                backColorStyle.SetFont(ifont)
                backColorStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin
                backColorStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin
                backColorStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin
                backColorStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin

                backColorStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center


                For i = 0 To 7
                    ReportSheet.GetRow(5).GetCell(i).CellStyle = backColorStyle
                Next


                'nf.removeRow(ReportSheet, ExcelRow_num)
                'nf.removeRow(ReportSheet, ExcelRow_num)
                ReportSheet.ShiftRows(ExcelRow_num + 1, ReportSheet.LastRowNum, -1, True, False)
                ReportSheet.ShiftRows(ExcelRow_num + 1, ReportSheet.LastRowNum, -1, True, False)

                ReportSheet.GetRow(0).HeightInPoints = 50.25


                sysFileName = "RPT_MAST_CCPLAN" & Now.ToString("yyyyMMddHHmmssfff")
                sysFileName = sysFileName & "." & LCase("XLSX")
                sysFilePath = fileFolder & "\" & sysFileName


                tmpFileStream = New FileStream(sysFilePath, FileMode.Create)
                ExcelWBObj.Write(tmpFileStream)

                tmpFileStream.Close()
                tmpFileStream.Dispose()

                Response.Write("<script language=""JavaScript"">dsp_status.innerHTML = ""DONE"";</script>")
                Response.Write("<script language=""JavaScript"">document.forms[0].hdf_file_path.value = """ & gU.jsString(sysFilePath) & """;</script>")
                Response.Write("<script language=""JavaScript"">document.getElementById(""tr_download"").style.display = """";</script>")
                Response.Flush()

            Else

                Response.Write("<script language=""JavaScript"">dsp_status.innerHTML = ""No Record Found."";</script>")

            End If

            REM XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

        Catch ex As Exception
            Throw ex
        Finally
            'tmpStreamWriter.Close()
            'tmpStreamWriter.Dispose()
            'tmpStreamWriter = Nothing

            ReadFileStream.Close()
            ReadFileStream.Dispose()
            tmpFileStream = Nothing
            ReadFileStream = Nothing
        End Try



    End Sub


    Protected Sub ibtn_download_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ibtn_download.Click
        Dim fullFilePath As String = hdf_file_path.Value
        Dim userFileName As String = "RPT_MAST_CCPLAN" & Now.ToString("yyyyMMdd")

        If fullFilePath <> "" Then
            Dim nFile As System.IO.FileInfo = New System.IO.FileInfo(fullFilePath)

            If nFile.Exists Then
                Response.Clear()
                Response.AddHeader("Content-Disposition", "attachment; filename=" & userFileName & nFile.Extension)
                Response.AddHeader("Content-Length", nFile.Length.ToString())
                Response.ContentType = "application/octet-stream"
                Response.WriteFile(nFile.FullName)
                Response.End()
            Else
                Response.Write("This file does not exist.")
            End If
        End If
    End Sub

End Class
