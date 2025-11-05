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
Imports System.Drawing
Imports System.Drawing.Drawing2D

Partial Class REPORT_RPT_LOCATION
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
    Const pageAbb As String = "exportLOCEXL"

    Protected isGenDownload As Boolean = False


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

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

            'Response.Write(Session("SEARCH_SESSION_PAGE_DAY_MONTH"))
            'Response.Write(Session("SEARCH_SESSION_PAGE_TEAM_CODE"))
            Dim dt As New DataTable

            ViewState("WH_CODE") = ""
            ViewState("WH_CODE") = Server.UrlDecode(Request("wh_code"))
            ViewState("havebarcode") = Server.UrlDecode(Request("havebarcode"))

            If ViewState("WH_CODE") Is Nothing OrElse ViewState("WH_CODE") = "" Then
                uiFun.displayMsg(Me, "", "No Location is Found!", Session("gLang"))
                dsp_status.Text = "No Location is Found!"
                dsp_status.ForeColor = Drawing.Color.Red

            Else
                isGenDownload = True
                tr_download.Style.Add("display", "none")
                ibtn_download.Visible = True
            End If

        End If

        ar.hideForm(Me)
    End Sub

    Protected Sub genDownloadFile()
        Dim ExcelWBObj As HSSFWorkbook
        Dim ReportSheet As HSSFSheet
        Dim tmpFileStream As FileStream
        Dim ReadFileStream As FileStream
        Dim tmpltPath, tmpltFileName As String
        Dim tmpDt As DataTable
        Dim i, currRow As Integer
        Dim w_code As String = ""
        Dim fullTmplName As String = ""
        'Dim msgLog As New PrgmLog(Cache("SYSP_LOG_DIR"), "PrgmLog" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")
        
        'tmpltPath = System.Configuration.ConfigurationManager.AppSettings.Item("EXCEL_TMPLT_PATH").ToString
        tmpltFileName = "location_barcode.xls"
        fullTmplName = Server.MapPath(tmpltFileName)

        ReadFileStream = New FileStream(fullTmplName, FileMode.Open, FileAccess.Read)

        Dim sysFileName, sysFilePath, fileFolder As String

        fileFolder = gU.getConfig("SYSP_TEMP_DIR")

        'SYSP_LOG_DIR

        If Not System.IO.Directory.Exists(fileFolder) Then
            System.IO.Directory.CreateDirectory(fileFolder)
        End If

        Try
            Dim oBmp As Bitmap
            Dim anchor As HSSFClientAnchor

            Dim picbuff As Byte()
            Dim picNo As Integer
            Dim EXLPIC As HSSFPicture
            Dim patriarch As HSSFPatriarch
            Dim fs As FileStream

            ExcelWBObj = New HSSFWorkbook(ReadFileStream)
            'TemplateSheet = ExcelWBObj.GetSheet("TEMPLATE")
            ReportSheet = ExcelWBObj.GetSheetAt(0)

            'Print Summary

            tmpDt = gU.getSessionTempData(pageAbb, "codeDT", Nothing)
            Dim tempPath As String = gU.getConfig("SYSP_TEMP_DIR")

            currRow = 1

            For i = 0 To tmpDt.Rows.Count - 1

                nf.setCellValue(ReportSheet, currRow, 0, tmpDt.Rows(i).Item("WH_NAME").ToString.Trim)
                nf.setCellValue(ReportSheet, currRow, 1, tmpDt.Rows(i).Item("FL_NAME").ToString.Trim)
                nf.setCellValue(ReportSheet, currRow, 2, tmpDt.Rows(i).Item("AR_NAME").ToString.Trim)
                nf.setCellValue(ReportSheet, currRow, 3, tmpDt.Rows(i).Item("RK_NAME").ToString.Trim)
                nf.setCellValue(ReportSheet, currRow, 4, tmpDt.Rows(i).Item("BN_CODE").ToString.Trim)
                nf.setCellValue(ReportSheet, currRow, 5, tmpDt.Rows(i).Item("BN_CSMS_CODE").ToString.Trim)
                nf.setCellValue(ReportSheet, currRow, 6, tmpDt.Rows(i).Item("LCN_CODE").ToString.Trim)
                ReportSheet.GetRow(currRow).HeightInPoints = 50

                If ViewState("havebarcode") = "N" Then
                Else
                    oBmp = Code128Rendering.MakeBarcodeImage(tmpDt.Rows(i).Item("LCN_CODE").ToString.Trim, 2, True)
                    'oBmp = ResizeImage(oBmp, 40)
                    oBmp.Save(tempPath & "/Barcode" & tmpDt.Rows(i).Item("LCN_CODE").ToString.Trim & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
                    patriarch = TryCast(ReportSheet.CreateDrawingPatriarch(), HSSFPatriarch)


                    anchor = New HSSFClientAnchor(0, 0, 0, 0, 7, currRow, 7, currRow)
                    anchor.AnchorType = 2
                    fs = New FileStream(tempPath & "/Barcode" & tmpDt.Rows(i).Item("LCN_CODE").ToString.Trim & ".jpg", FileMode.Open, FileAccess.Read)

                    ReDim picbuff(fs.Length)

                    fs.Read(picbuff, 0, CInt(fs.Length))

                    picNo = ExcelWBObj.AddPicture(picbuff, NPOI.SS.UserModel.PictureType.JPEG)

                    EXLPIC = TryCast(patriarch.CreatePicture(anchor, picNo), HSSFPicture)

                    'Reset the image to the original size.
                    EXLPIC.Resize()
                    EXLPIC.LineStyle = NPOI.SS.UserModel.LineStyle.None
                End If
                currRow += 1
            Next

            'ExcelWBObj.RemoveSheetAt(ExcelWBObj.GetSheetIndex("TEMPLATE"))
            ''ExcelWBObj.RemoveName(0)

            sysFileName = "LocExl" & Now.Date.ToString("ddMMyyyyHHmmssff")
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
        Dim userFileName As String = "LocExl" & Now.Date.ToString("ddMMyyyy")

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

    Private Function ResizeImage(image As System.Drawing.Bitmap, percent As Integer) As System.Drawing.Bitmap
        ' percent is the actual integer percent of the original size

        Dim imgThumb As New System.Drawing.Bitmap(image.Width * percent \ 100, image.Height * percent \ 100)

        Dim sourceRect As New Rectangle(0, 0, image.Width, image.Height)
        Dim destRect As New Rectangle(0, 0, imgThumb.Width, imgThumb.Height)


        Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(imgThumb)
        g.CompositingQuality = CompositingQuality.HighQuality
        g.SmoothingMode = SmoothingMode.HighQuality
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.DrawImage(image, destRect, sourceRect, GraphicsUnit.Pixel)
        Return imgThumb
    End Function
End Class
