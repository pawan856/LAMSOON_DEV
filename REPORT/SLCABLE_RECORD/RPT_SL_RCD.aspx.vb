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


Partial Class RPT_SLCABLE_RCD_main
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
    Protected Const FUN_CODE As String = "RPT_SL_RCD"

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

        'Response.Write(Session("GENERIC_SESSION_COMPLETE_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_WHERE_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_GROUPBY_SQL"))
        'Response.Write("<br><br>")
        'Response.Write(Session("GENERIC_SESSION_SRCH_ORDERBY_SQL"))
        'Response.Write("<br><br>")

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
        Dim TemplateSheet, ReportSheet, NewSheet As XSSFSheet


        Dim pa1 As GlobalDBFunc.DBCmdPara


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

        tmpltFileName = "SLCABLE_RECORD_TMP.xlsx"
        tmpltPath = Server.MapPath(tmpltFileName)
        ReadFileStream = New FileStream(tmpltPath, FileMode.Open, FileAccess.Read)

        Dim sysFileName, sysFilePath, fileFolder As String

        fileFolder = HttpContext.Current.Cache("SYSP_TEMP_DIR")

        If Not System.IO.Directory.Exists(fileFolder) Then
            System.IO.Directory.CreateDirectory(fileFolder)
        End If


        Try
            Dim itm_sku_no As String

            pa1 = New GlobalDBFunc.DBCmdPara

            itm_sku_no = Session("SEARCH_SESSION_PAGE_WMS_ITEM_ITM_SKU_NO")

            If itm_sku_no.Trim <> "" Then
                'If itm_sku_no.Contains(",") Then
                '    itm_sku_no = gU.formatList(itm_sku_no.Trim)

                '    Dim skuList() As String = gU.listToArray(itm_sku_no)
                '    Dim tempList As String = ""

                '    For i = 0 To skuList.Length - 1
                '        tempList &= "'" & gU.dbEncode(skuList(i)) & "',"
                '    Next

                '    If tempList <> "" Then
                '        tempList = Left(tempList, Len(tempList) - 1)
                '        addSQL &= " AND WMS_ITEM.ITM_SKU_NO in (" & tempList & ") "
                '    End If

                'Else
                addSQL &= " AND WMS_ITEM.ITM_SKU_NO='" & gU.dbEncode(itm_sku_no.Trim) & "' "
                'End If
            End If


            sql_string = " SELECT WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM.ITM_NAME, " & _
                         " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO " & _
                         " FROM WMS_ITEM_LOC_BAL_S INNER JOIN " & _
                         " WMS_ITEM ON WMS_ITEM_LOC_BAL_S.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL_S.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                         " WMS_ITEM_LOC_BAL_S.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL_S.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                         " where ILBS_QTY2 > 0 and isnULL(ILBS_SL,'N') = 'Y' " & addSQL & _
                         " order by ITM_SKU_NO, ILBS_DRUM_ID, ILBS_QTY2 "


            'Dim tempSQL As String = gDB.getCmdSql(sql_string, pa1)
            result_dt = gDB.getDataTable(sql_string)

            If result_dt.Rows.Count > 0 Then

                ExcelWBObj = New XSSFWorkbook(ReadFileStream)
                TemplateSheet = ExcelWBObj.GetSheet("TEMPLATE")
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

                Dim totalLength As Double = 0
                Dim title As String = ""
                'XXXXXXXXXXXXXXXXXX Column Header XXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                title = "SHORT LENGTH CABLE RECORD FOR LOT NO." & result_dt.Rows(0).Item("itm_sku_no").ToString.Trim & result_dt.Rows(0).Item("itm_name").ToString.Trim

                nf.setCellValue(ReportSheet, 1, 0, title)
                nf.setCellValue(ReportSheet, 8, 3, result_dt.Rows(0).Item("itm_sku_no").ToString.Trim)
                'REM XXXXXXXXXXXXXXXXXXX CONTENT BLOCK XXXXXXXXXXXXXXXXXXX


                ExcelRow_num = 9

                For i = 0 To result_dt.Rows.Count - 1

                    'If PrevSKU <> result_dt.Rows(i).Item("itm_sku_no").ToString.Trim Then

                    'NewSheet = ExcelWBObj.CloneSheet(0)
                    'sheetIDX = ExcelWBObj.GetSheetIndex(NewSheet)
                    '    ExcelWBObj.SetSheetName(sheetIDX, "Sort " & result_dt.Rows(i).Item("itm_sku_no").ToString.Trim)
                    'ReportSheet = NewSheet

                    '    PrevSKU = result_dt.Rows(i).Item("itm_sku_no").ToString.Trim
                    'End If

                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(4, 4, 0, 1), ExcelRow_num, 0)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 0, result_dt.Rows(i).Item("ILBS_DRUM_ID").ToString.Trim)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 1, result_dt.Rows(i).Item("ILBS_QTY2").ToString.Trim, "DEC")

                    totalLength += gU.decodeEmptyCdbl(result_dt.Rows(i).Item("ILBS_QTY2").ToString.Trim, 0)

                    ExcelRow_num += 1

                Next

                nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(5, 5, 0, 1), ExcelRow_num, 0)
                nf.setCellValue(ReportSheet, 6, 9, totalLength, "DEC")

                'ExcelWBObj.SetSheetName(0, "Sort " & result_dt.Rows(0).Item("itm_sku_no").ToString.Trim)
                ExcelWBObj.RemoveSheetAt(ExcelWBObj.GetSheetIndex("TEMPLATE"))

                sysFileName = "RPT_SL_RCD" & Now.ToString("yyyyMMddHHmmssfff")
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
        Dim userFileName As String = "Short_Length_RECORD_RPT" & Now.ToString("yyyyMMdd")

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
