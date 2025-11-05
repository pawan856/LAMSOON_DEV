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


Partial Class REPORT_SHORT_LEN_RPT_SHORTLEN_W
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

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils(FUN_CODE, Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        Server.ScriptTimeout = 7200

        If Not IsPostBack Then


        End If

        ar.hideForm(Me)
    End Sub

    Protected Sub btnGen_Click(sender As Object, e As System.EventArgs) Handles btnGen.Click
        If String.IsNullOrWhiteSpace(START_DATE.Text) Then
            uiFun.displayMsgNew(UDP1, "", "Please enter a start date for generation report.", Session("gLang"))
        Else
            Dim returnMsg As String = ""

            If GenReport(returnMsg) Then
                uiFun.displayMsgNew(UDP1, "", "Report has been generated.", Session("gLang"))
                btnDL.Visible = True
            Else
                errMsg.Text = returnMsg
            End If
        End If
    End Sub

    Protected Function GenReport(Optional ByRef ReturnMsg As String = "") As Boolean
        Dim successFlag As Boolean = False

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction


        transaction = gConn.BeginTransaction()

        Try
            Dim s_date As String = START_DATE.Text.Trim
            Dim selectSQL, updateSQL, insertSQL As String
            Dim FR_DATE, temp_Date As Date
            Dim tempDateStr As String = ""
            Dim resultDT As DataTable

            customDateTimeFormat.DateSeparator = "/"
            customDateTimeFormat.TimeSeparator = ":"
            customDateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
            customDateTimeFormat.LongDatePattern = "dd/MM/yyyy"
            customDateTimeFormat.ShortTimePattern = "HH:mm"
            customDateTimeFormat.LongTimePattern = "HH:mm"
            customDateTimeFormat.FullDateTimePattern = "dd/MM/yyyy HH:mm"


            FR_DATE = Convert.ToDateTime(s_date, customDateTimeFormat)

            updateSQL = "DELETE from WMS_SL_SUM_TMP"
            gDB.amendData(updateSQL, gConn, transaction)

            Dim selectCol As String = ""
            Dim Pivotcol As String = ""


            For i = 0 To 23
                temp_Date = FR_DATE.AddMonths(i)
                tempDateStr = temp_Date.ToString("dd/MM/yyyy")

                selectCol &= "[" & tempDateStr & "] as c" & i & ","
                Pivotcol &= "[" & tempDateStr & "],"



                'insertSQL = " insert into WMS_SL_SUM_TMP(IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ITX_DATE, SL_QTY) " & _
                '            " (select IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, convert(date,'" & tempDateStr & "',103), count(distinct ITSX_SERIAL_NO) from " & _
                '            " (select MaxBal.IMP_CODE, MaxBal.STORER_CODE, MaxBal.ITM_CODE, MaxBal.PACK_KEY,maxbal.ITSX_SERIAL_NO,  MaxBal.ITSX_BAL_AFTER2 from  " & _
                '            " (SELECT WMS_ITEM_LOC_BAL_S_TX.IMP_CODE, WMS_ITEM_LOC_BAL_S_TX.STORER_CODE, WMS_ITEM_LOC_BAL_S_TX.ITM_CODE,  " & _
                '            " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY, MAX(WMS_ITEM_LOC_BAL_S_TX.ITX_DATE) AS maxDate, WMS_ITEM_LOC_BAL_S_TX.ITSX_SERIAL_NO " & _
                '            " FROM WMS_ITEM_LOC_BAL_S_TX INNER JOIN " & _
                '            " WMS_ITEM ON WMS_ITEM_LOC_BAL_S_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND  " & _
                '            " WMS_ITEM_LOC_BAL_S_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND WMS_ITEM_LOC_BAL_S_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND  " & _
                '            " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                '            " WHERE (WMS_ITEM_LOC_BAL_S_TX.ITSX_SL = 'Y') AND (WMS_ITEM.ITM_TYPE = 'CABLE') AND (WMS_ITEM_LOC_BAL_S_TX.ITX_DATE < DATEADD(DAY, 1,  " & _
                '            " CONVERT(date, '" & tempDateStr & "', 103))) " & _
                '            " GROUP BY WMS_ITEM_LOC_BAL_S_TX.IMP_CODE, WMS_ITEM_LOC_BAL_S_TX.STORER_CODE, WMS_ITEM_LOC_BAL_S_TX.ITM_CODE, " & _
                '            " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY, WMS_ITEM_LOC_BAL_S_TX.ITSX_SERIAL_NO) maxDate " & _
                '            " inner join WMS_ITEM_LOC_BAL_S_TX MaxBal on " & _
                '            " MaxBal.IMP_CODE = maxDate.IMP_CODE and MaxBal.STORER_CODE= maxDate.STORER_CODE " & _
                '            " and MaxBal.ITM_CODE = maxDate.ITM_CODE and  MaxBal.PACK_KEY = maxDate.PACK_KEY and maxbal.ITX_DATE = maxDate.maxDate and maxbal.ITSX_SERIAL_NO = maxDate.ITSX_SERIAL_NO) qty " & _
                '            " where ITSX_BAL_AFTER2 > 0 group by  IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY) "

                insertSQL = " insert into WMS_SL_SUM_TMP(IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ITX_DATE, SL_QTY) " & _
                            " (select MaxBal.IMP_CODE, MaxBal.STORER_CODE, MaxBal.ITM_CODE,maxbal.PACK_KEY, convert(date,'" & tempDateStr & "',103), sum(MaxBal.ITSX_BAL_AFTER2) as SL_QTY from " & _
                            " (SELECT WMS_ITEM_LOC_BAL_S_TX.IMP_CODE, WMS_ITEM_LOC_BAL_S_TX.STORER_CODE, WMS_ITEM_LOC_BAL_S_TX.ITM_CODE, " & _
                            " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY, MAX(WMS_ITEM_LOC_BAL_S_TX.ITX_DATE) AS maxDate, WMS_ITEM_LOC_BAL_S_TX.ITSX_SERIAL_NO " & _
                            " FROM WMS_ITEM_LOC_BAL_S_TX INNER JOIN " & _
                            " WMS_ITEM ON WMS_ITEM_LOC_BAL_S_TX.IMP_CODE = WMS_ITEM.IMP_CODE AND  " & _
                            " WMS_ITEM_LOC_BAL_S_TX.STORER_CODE = WMS_ITEM.STORER_CODE AND WMS_ITEM_LOC_BAL_S_TX.ITM_CODE = WMS_ITEM.ITM_CODE AND  " & _
                            " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                            " WHERE (WMS_ITEM_LOC_BAL_S_TX.ITSX_SL = 'Y') AND (WMS_ITEM.ITM_TYPE = 'CABLE') AND (WMS_ITEM_LOC_BAL_S_TX.ITX_DATE < DATEADD(DAY, 1,CONVERT(date, '" & tempDateStr & "', 103))) " & _
                            " GROUP BY WMS_ITEM_LOC_BAL_S_TX.IMP_CODE, WMS_ITEM_LOC_BAL_S_TX.STORER_CODE, WMS_ITEM_LOC_BAL_S_TX.ITM_CODE,  " & _
                            " WMS_ITEM_LOC_BAL_S_TX.PACK_KEY, WMS_ITEM_LOC_BAL_S_TX.ITSX_SERIAL_NO) maxDate " & _
                            " inner join WMS_ITEM_LOC_BAL_S_TX MaxBal on  " & _
                            " MaxBal.IMP_CODE = maxDate.IMP_CODE and MaxBal.STORER_CODE= maxDate.STORER_CODE  " & _
                            " and MaxBal.ITM_CODE = maxDate.ITM_CODE and  MaxBal.PACK_KEY = maxDate.PACK_KEY and maxbal.ITX_DATE = maxDate.maxDate and maxbal.ITSX_SERIAL_NO = maxDate.ITSX_SERIAL_NO " & _
                            " group by MaxBal.IMP_CODE, MaxBal.STORER_CODE, MaxBal.ITM_CODE,maxbal.PACK_KEY) "


                gDB.amendData(insertSQL, gConn, transaction)

            Next

            If Pivotcol <> "" Then Pivotcol = Left(Pivotcol, Len(Pivotcol) - 1)

            selectSQL = " select  isnull(FDRs.IMP_CODE,pvt.imp_code) as IMP_CODE, isnull(FDRs.STORER_CODE,pvt.storer_code) as storer_code, isnull(FDRs.ITM_CODE,pvt.itm_code) as itm_code, isnull(FDRs.PACK_KEY,pvt.pack_key) as pack_key, " & _
                        " isNull(itm2.ITM_NAME,itm.ITM_NAME) as itm_name, isnull(itm2.ITM_DESC,itm.ITM_DESC) as itm_desc, isnull(itm2.ITM_SKU_NO ,itm.ITM_SKU_NO) as itm_sku_no," & selectCol & _
                        " isNull(cast(FD_COUNT as varchar) + ' x ' + Cast(cast(CLength as int)as varchar) + ILBS_UOM2,'-') as FD_COUNT " & _
                        " from (SELECT IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, convert(varchar,ITX_DATE,103) as sITX_DATE, SL_QTY " & _
                        " FROM WMS_SL_SUM_TMP ) as p " & _
                        " PIVOT (Max (SL_QTY) FOR sITX_DATE IN " & _
                        " (" & Pivotcol & " ) " & _
                        " ) AS pvt  inner join wms_item itm on  " & _
                        " pvt.IMP_CODE = itm.IMP_CODE and pvt.STORER_CODE = itm.STORER_CODE and pvt.ITM_CODE=itm.ITM_CODE and  pvt.PACK_KEY= itm.PACK_KEY " & _
                        " full outer join " & _
                        " (SELECT WMS_ITEM_LOC_BAL_S.IMP_CODE, WMS_ITEM_LOC_BAL_S.STORER_CODE, WMS_ITEM_LOC_BAL_S.ITM_CODE, PACK_KEY, count(WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO) as FD_COUNT, max(WMS_ITEM_LOC_BAL_S.ILBS_QTY2) as Clength, max(WMS_ITEM_LOC_BAL_S.ILBS_UOM2) as ILBS_UOM2 " & _
                        " FROM WMS_ITEM_LOC_BAL_S " & _
                        " where  isnull(WMS_ITEM_LOC_BAL_S.ILBS_SL,'N') <> 'Y' and WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0 " & _
                        " group by WMS_ITEM_LOC_BAL_S.IMP_CODE, WMS_ITEM_LOC_BAL_S.STORER_CODE, WMS_ITEM_LOC_BAL_S.ITM_CODE, WMS_ITEM_LOC_BAL_S.PACK_KEY) as FDRs  " & _
                        " on pvt.IMP_CODE = FDRs.IMP_CODE AND pvt.STORER_CODE = FDRs.STORER_CODE AND pvt.ITM_CODE=FDRs.ITM_CODE AND pvt.PACK_KEY=FDRs.PACK_KEY " & _
                        " left outer join wms_item  itm2 on  " & _
                        " FDRs.IMP_CODE = itm2.IMP_CODE AND " & _
                        " FDRs.STORER_CODE = itm2.STORER_CODE AND FDRs.ITM_CODE = itm2.ITM_CODE AND " & _
                        " FDRs.PACK_KEY = itm2.PACK_KEY " & _
                        " Order by itm_sku_no "

            resultDT = gDB.getDataTable(selectSQL, gConn, transaction)
            If resultDT.Rows.Count > 0 Then
                genDownloadFile(resultDT)
                updateSQL = "DELETE from WMS_SL_SUM_TMP"
                gDB.amendData(updateSQL, gConn, transaction)
            Else
                ReturnMsg = "No Record Found"
                Return False
            End If



            transaction.Commit()


            successFlag = True


        Catch ex As Exception
            transaction.Rollback()
            'Response.Write(ex.Message)
            uiFun.displayMsgNew(UDP1, "1008", "", Session("gLang"))
            successFlag = False
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

        Return successFlag
    End Function

    Protected Function genDownloadFile(ByRef result_DT As DataTable) As Boolean
        Dim successFlag As Boolean = False

        Dim ExcelWBObj As HSSFWorkbook
        Dim TemplateSheet, ReportSheet As HSSFSheet

        Dim pa1 As GlobalDBFunc.DBCmdPara

        Dim sql_string As String = ""
        Dim addSQL As String = ""

        Dim tmpltPath, tmpltFileName As String
        Dim tmpFileStream As FileStream
        Dim ReadFileStream As FileStream

        Dim s_date As String = START_DATE.Text.Trim
        Dim FR_DATE, temp_Date As Date
        Dim tempDateStr As String = ""

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

            pa1 = New GlobalDBFunc.DBCmdPara

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
                FR_DATE = Convert.ToDateTime(s_date, customDateTimeFormat)
                Dim tempYr As String = ""

                For x = 0 To 23
                    temp_Date = FR_DATE.AddMonths(x)
                    tempDateStr = temp_Date.ToString("d-MMM")

                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(1, 3, 2, 2), 1, startCol)
                    'nf.setCellValue(ReportSheet, 2, startCol, "1/" & CStr(i).PadLeft(2, "0") & "/" & x)

                    If tempYr <> temp_Date.Year.ToString Then
                        nf.setCellValue(ReportSheet, 1, startCol, temp_Date.Year.ToString)
                        tempYr = temp_Date.Year.ToString
                    End If

                    nf.setCellValue(ReportSheet, 2, startCol, tempDateStr)
                    nf.setCellValue(ReportSheet, 3, startCol, "S.L. Qty")
                    startCol += 1

                Next

                nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(2, 2, 3, 3), 2, startCol)

                'REM XXXXXXXXXXXXXXXXXXX CONTENT BLOCK XXXXXXXXXXXXXXXXXXX

                ExcelRow_num = 4
                Dim colName As String = ""

                For i = 0 To result_DT.Rows.Count - 1

                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(4, 4, 0, 1), ExcelRow_num, 0)

                    nf.setCellValue(ReportSheet, ExcelRow_num, 0, result_DT.Rows(i).Item("ITM_SKU_NO").ToString)
                    nf.setCellValue(ReportSheet, ExcelRow_num, 1, result_DT.Rows(i).Item("ITM_DESC").ToString)

                    startCol = 2
                    For x = 0 To 23
                        colName = "c" & x
                        If result_DT.Rows(i).Item(colName).ToString <> "0" AndAlso result_DT.Rows(i).Item(colName).ToString <> "" Then
                            nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(4, 4, 2, 2), ExcelRow_num, startCol)

                            nf.setCellValue(ReportSheet, ExcelRow_num, startCol, result_DT.Rows(i).Item(colName).ToString, "INT")
                        End If
                        startCol += 1
                    Next

                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(3, 3, 2, 2), ExcelRow_num, startCol)
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, result_DT.Rows(i).Item("FD_COUNT").ToString)
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

                hdf_file_path.Value = sysFilePath
                successFlag = True
                'Response.Write("<script language=""JavaScript"">dsp_status.innerHTML = ""DONE"";</script>")
                'Response.Write("<script language=""JavaScript"">document.forms[0].hdf_file_path.value = """ & gU.jsString(sysFilePath) & """;</script>")
                'Response.Write("<script language=""JavaScript"">document.getElementById(""tr_download"").style.display = """";</script>")
                'Response.Flush()


            Else
                successFlag = False
                'Response.Write("<script language=""JavaScript"">dsp_status.innerHTML = ""No Record Found."";</script>")

            End If
            Return successFlag
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



    End Function


    Protected Sub btnDL_Click(sender As Object, e As System.EventArgs) Handles btnDL.Click
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
                'Response.Write("This file does not exist.")
            End If
        End If
    End Sub
End Class
