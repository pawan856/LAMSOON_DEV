Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports NPOI.HSSF.UserModel
Imports NPOI.HPSF
Imports NPOI.POIFS.FileSystem
Imports NPOI.SS.Util
Imports NPOI.HSSF.Util


Partial Class RPT_SHORTLEN_main
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private nf As New NPOIFunc

    Protected Const IMG_PATH As String = "../../images/"
    Protected Const ROOT_PATH As String = "../../"
    Protected Const FUN_CODE As String = "RPT_SLCABLE"

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

        Dim ExcelWBObj As HSSFWorkbook
        Dim TemplateSheet, ReportSheet As HSSFSheet


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

        tmpltFileName = "RPT_SECL_TMP.xls"
        tmpltPath = Server.MapPath(tmpltFileName)
        ReadFileStream = New FileStream(tmpltPath, FileMode.Open, FileAccess.Read)

        Dim sysFileName, sysFilePath, fileFolder As String

        fileFolder = HttpContext.Current.Cache("SYSP_TEMP_DIR")

        If Not System.IO.Directory.Exists(fileFolder) Then
            System.IO.Directory.CreateDirectory(fileFolder)
        End If


        Try
            Dim itm_sku_no, yrRange As String

            pa1 = New GlobalDBFunc.DBCmdPara

            itm_sku_no = Session("SEARCH_SESSION_PAGE_WMS_ITEM_ITM_SKU_NO")
            yrRange = Session("SEARCH_SESSION_PAGE_YrRange")

            Dim yrArr() As String
            Dim selectCol As String = ""
            Dim Pivotcol As String = ""
            Dim frYear, toYear As Integer

            If yrRange <> "" Then
                yrArr = Split(yrRange, "-")
                frYear = CInt(yrArr(0))
                toYear = CInt(yrArr(1))

                For x = frYear To toYear
                    For i = 1 To 12
                        selectCol &= "[" & x & i & "] as c" & x & i & ","
                        Pivotcol &= "[" & x & i & "],"
                    Next
                Next

                'If selectCol <> "" Then selectCol = Left(selectCol, Len(selectCol) - 1)
                If Pivotcol <> "" Then Pivotcol = Left(Pivotcol, Len(Pivotcol) - 1)
            End If

            If itm_sku_no.Trim <> "" Then

                addSQL &= " AND WMS_ITEM.ITM_SKU_NO='" & gU.dbEncode(itm_sku_no.Trim) & "' "

            End If
            

            'sql_string = " select pvt.IMP_CODE, pvt.STORER_CODE, pvt.ITM_CODE, pvt.PACK_KEY, ITM_NAME, ITM_DESC," & selectCol & " ITM_SKU_NO" & _
            '             " from " & _
            '             " (select maxRs.IMP_CODE, maxRs.STORER_CODE, maxRs.ITM_CODE, " & _
            '             " maxRs.PACK_KEY, maxRs.ITM_NAME, maxRs.ITM_DESC, maxRs.ITM_SKU_NO, maxRs.ITSX_SERIAL_NO, yrMon " & _
            '             " from ( " & _
            '             " SELECT WMS_ITEM_LOC_BAL_S_TX.IMP_CODE, WMS_ITEM_LOC_BAL_S_TX.STORER_CODE, WMS_ITEM_LOC_BAL_S_TX.ITM_CODE,  " & _
            '             " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL_S_TX.ITSX_SERIAL_NO,  " & _
            '             " concat(year(WMS_ITEM_LOC_BAL_S_TX.ITX_DATE),month(WMS_ITEM_LOC_BAL_S_TX.ITX_DATE)) as yrMon, " & _
            '             " max(WMS_ITEM_LOC_BAL_S_TX.itx_date) as maxDate " & _
            '             " FROM WMS_ITEM_LOC_BAL_S_TX INNER JOIN " & _
            '             " WMS_ITEM ON WMS_ITEM_LOC_BAL_S_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND  " & _
            '             " WMS_ITEM_LOC_BAL_S_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND WMS_ITEM_LOC_BAL_S_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND  " & _
            '             " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY = WMS_ITEM.PACK_KEY		 " & _
            '             " WHERE WMS_ITEM.ITM_TYPE = 'CABLE' " & addSQL & _
            '             " group by WMS_ITEM_LOC_BAL_S_TX.IMP_CODE, WMS_ITEM_LOC_BAL_S_TX.STORER_CODE, WMS_ITEM_LOC_BAL_S_TX.ITM_CODE,  " & _
            '             " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL_S_TX.ITSX_SERIAL_NO,  " & _
            '             " concat(year(WMS_ITEM_LOC_BAL_S_TX.ITX_DATE),month(WMS_ITEM_LOC_BAL_S_TX.ITX_DATE)) " & _
            '             " ) maxRs " & _
            '             " inner join WMS_ITEM_LOC_BAL_S_TX as balRslt on  " & _
            '             " maxRs.IMP_CODE = BalRslt.IMP_CODE AND  " & _
            '             " maxRs.STORER_CODE = BalRslt.STORER_CODE AND  " & _
            '             " maxRs.ITM_CODE = BalRslt.ITM_CODE AND  " & _
            '             " maxRs.PACK_KEY = BalRslt.PACK_KEY AND  " & _
            '             " maxRs.ITSX_SERIAL_NO = BalRslt.ITSX_SERIAL_NO AND  " & _
            '             " maxRs.maxDate = BalRslt.ITX_DATE " & _
            '             " where balRslt.ITSX_BAL_AFTER2 > 0 and isnull(balRslt.ITSX_SL,'N') = 'Y' ) P  " & _
            '             " PIVOT (  " & _
            '             " COUNT (ITSX_SERIAL_NO) " & _
            '             " FOR yrMon in " & _
            '             " (" & Pivotcol & ") " & _
            '             " ) AS pvt " & _
            '             " ORDER BY pvt.IMP_CODE, pvt.STORER_CODE, pvt.itm_sku_no "
            sql_string = " select isnull(FDRs.IMP_CODE,pvt.imp_code) as IMP_CODE, isnull(FDRs.STORER_CODE,pvt.storer_code) as storer_code, isnull(FDRs.ITM_CODE,pvt.itm_code) as itm_code, isnull(FDRs.PACK_KEY,pvt.pack_key) as pack_key, " & _
                         " isNull(itm.ITM_NAME,pvt.ITM_NAME) as itm_name, isnull(itm.ITM_DESC,pvt.ITM_DESC) as itm_desc, isnull(itm.ITM_SKU_NO ,pvt.ITM_SKU_NO) as itm_sku_no, " & selectCol & " isNull(cast(FD_COUNT as varchar) + ' x ' + Cast(cast(CLength as int)as varchar) + ILBS_UOM2,'-') as FD_COUNT  " & _
                         " from  (SELECT        WMS_ITEM_LOC_BAL_S.IMP_CODE, WMS_ITEM_LOC_BAL_S.STORER_CODE, WMS_ITEM_LOC_BAL_S.ITM_CODE, PACK_KEY, count(WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO) as FD_COUNT, max(WMS_ITEM_LOC_BAL_S.ILBS_QTY2) as Clength, max(WMS_ITEM_LOC_BAL_S.ILBS_UOM2) as ILBS_UOM2 " & _
                         " FROM            WMS_ITEM_LOC_BAL_S " & _
                         " where  isnull(WMS_ITEM_LOC_BAL_S.ILBS_SL,'N') <> 'Y' and WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0 " & _
                         " group by WMS_ITEM_LOC_BAL_S.IMP_CODE, WMS_ITEM_LOC_BAL_S.STORER_CODE, WMS_ITEM_LOC_BAL_S.ITM_CODE, WMS_ITEM_LOC_BAL_S.PACK_KEY) as FDRs " & _
                         " full outer join  " & _
                         " (select maxRs.IMP_CODE, maxRs.STORER_CODE, maxRs.ITM_CODE, maxRs.PACK_KEY, maxRs.ITM_NAME, maxRs.ITM_DESC, maxRs.ITM_SKU_NO, maxRs.ITSX_SERIAL_NO, yrMon " & _
                         " from ( " & _
                         " SELECT        WMS_ITEM_LOC_BAL_S_TX.IMP_CODE, WMS_ITEM_LOC_BAL_S_TX.STORER_CODE, WMS_ITEM_LOC_BAL_S_TX.ITM_CODE,  " & _
                         " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL_S_TX.ITSX_SERIAL_NO,  " & _
                         " concat(year(WMS_ITEM_LOC_BAL_S_TX.ITX_DATE),month(WMS_ITEM_LOC_BAL_S_TX.ITX_DATE)) as yrMon, " & _
                         " max(WMS_ITEM_LOC_BAL_S_TX.itx_date) as maxDate " & _
                         " FROM            WMS_ITEM_LOC_BAL_S_TX INNER JOIN " & _
                         " WMS_ITEM ON WMS_ITEM_LOC_BAL_S_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND  " & _
                         " WMS_ITEM_LOC_BAL_S_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND WMS_ITEM_LOC_BAL_S_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND  " & _
                         " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                         " WHERE WMS_ITEM.ITM_TYPE = 'CABLE' " & addSQL & _
                         " group by WMS_ITEM_LOC_BAL_S_TX.IMP_CODE, WMS_ITEM_LOC_BAL_S_TX.STORER_CODE, WMS_ITEM_LOC_BAL_S_TX.ITM_CODE, " & _
                         "         WMS_ITEM_LOC_BAL_S_TX.PACK_KEY, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL_S_TX.ITSX_SERIAL_NO,  " & _
                         " 		 concat(year(WMS_ITEM_LOC_BAL_S_TX.ITX_DATE),month(WMS_ITEM_LOC_BAL_S_TX.ITX_DATE)) " & _
                         " ) maxRs " & _
                         " inner join WMS_ITEM_LOC_BAL_S_TX as balRslt on  " & _
                         " 			maxRs.IMP_CODE = BalRslt.IMP_CODE AND  " & _
                         " 			maxRs.STORER_CODE = BalRslt.STORER_CODE AND  " & _
                         " 			maxRs.ITM_CODE = BalRslt.ITM_CODE AND  " & _
                         " 			maxRs.PACK_KEY = BalRslt.PACK_KEY AND  " & _
                         " 			maxRs.ITSX_SERIAL_NO = BalRslt.ITSX_SERIAL_NO AND  " & _
                         " 			maxRs.maxDate = BalRslt.ITX_DATE " & _
                         " where balRslt.ITSX_BAL_AFTER2 > 0 and isnull(balRslt.ITSX_SL,'N') = 'Y' ) P " & _
                         " PIVOT " & _
                         " ( COUNT (ITSX_SERIAL_NO)  " & _
                         " FOR yrMon in " & _
                         " (" & Pivotcol & ") " & _
                         " ) AS pvt " & _
                         " on pvt.IMP_CODE = FDRs.IMP_CODE AND pvt.STORER_CODE = FDRs.STORER_CODE AND pvt.ITM_CODE=FDRs.ITM_CODE AND pvt.PACK_KEY=FDRs.PACK_KEY " & _
                         " left outer join wms_item  itm on  " & _
                         " FDRs.IMP_CODE = itm.IMP_CODE AND  " & _
                         " FDRs.STORER_CODE = itm.STORER_CODE AND FDRs.ITM_CODE = itm.ITM_CODE AND  " & _
                         " FDRs.PACK_KEY = itm.PACK_KEY " & _
                         " WHERE 1=1 " & Replace(addSQL, "WMS_ITEM.", "itm.") & _
                         " Order by itm_sku_no "
            'Dim tempSQL As String = gDB.getCmdSql(sql_string, pa1)
            result_dt = gDB.getDataTable(sql_string)

            If result_dt.Rows.Count > 0 Then

                ExcelWBObj = New HSSFWorkbook(ReadFileStream)
                TemplateSheet = ExcelWBObj.GetSheet("TEMPLATE")
                ReportSheet = ExcelWBObj.GetSheetAt(0)


                'XXXXXXXXXXXXXXXXXX Excel Cell StyleXXXXXXXXXXXXXXXXXXXXX
                'Dim warpStyle As HSSFCellStyle
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

                Dim startCol As Integer = 2

                For x = frYear To toYear
                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(1, 1, 2, 2), 1, startCol)
                    nf.setCellValue(ReportSheet, 1, startCol, x)

                    For i = 1 To 12
                        nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(2, 3, 2, 2), 2, startCol)
                        'nf.setCellValue(ReportSheet, 2, startCol, "1/" & CStr(i).PadLeft(2, "0") & "/" & x)
                        nf.setCellValue(ReportSheet, 2, startCol, MonthName(i, True))
                        nf.setCellValue(ReportSheet, 3, startCol, "S.L. Qty")
                        startCol += 1
                    Next
                Next

                nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(2, 2, 3, 3), 2, startCol)

                'REM XXXXXXXXXXXXXXXXXXX CONTENT BLOCK XXXXXXXXXXXXXXXXXXX

                ExcelRow_num = 4
                Dim colName As String = ""

                For i = 0 To result_dt.Rows.Count - 1

                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(4, 4, 0, 1), ExcelRow_num, 0)

                    nf.setCellValue(ReportSheet, ExcelRow_num, 0, result_dt.Rows(i).Item("ITM_SKU_NO").ToString)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 1, result_dt.Rows(i).Item("ITM_DESC").ToString)

                    startCol = 2
                    For x = frYear To toYear
                        For y = 1 To 12
                            colName = "c" & x & y
                            If result_dt.Rows(i).Item(colName).ToString <> "0" AndAlso result_dt.Rows(i).Item(colName).ToString <> "" Then
                                nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(4, 4, 2, 2), ExcelRow_num, startCol)

                                nf.setCellValue(ReportSheet, ExcelRow_num, startCol, result_dt.Rows(i).Item(colName).ToString, "INT")
                            End If
                            startCol += 1
                        Next
                    Next

                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(3, 3, 2, 2), ExcelRow_num, startCol)
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, result_dt.Rows(i).Item("FD_COUNT").ToString)
                    'FD_COUNT
                    ExcelRow_num += 1

                Next

                ExcelWBObj.RemoveSheetAt(ExcelWBObj.GetSheetIndex("TEMPLATE"))

                sysFileName = "Short_Length_RPT" & Now.ToString("yyyyMMddHHmmssfff")
                sysFileName = sysFileName & "." & LCase("XLS")
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
        Dim userFileName As String = "Short_Length_RPT" & Now.ToString("yyyyMMdd")

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
