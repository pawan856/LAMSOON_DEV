Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Data.SqlClient
Imports System.Globalization
Imports NPOI.HSSF.UserModel
Imports NPOI.HPSF
Imports NPOI.POIFS.FileSystem
Imports NPOI.SS.Util
Imports NPOI.HSSF.Util


Partial Class CARGO_RCP_MAIN
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private nf As New NPOIFunc
    Private rptU As New ReportUtils

    Protected Const IMG_PATH As String = "../../../images/"
    Protected Const ROOT_PATH As String = "../../../"
    Protected Const FUN_CODE As String = "IB_GR"

    Protected isGenDownload As Boolean = False
    Private ExcelRow_num As Long = 0


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

            'Response.Write(Session("SEARCH_SESSION_PAGE_DAY_MONTH"))
            'Response.Write(Session("SEARCH_SESSION_PAGE_TEAM_CODE"))
            ViewState("STORER_CODE") = ""
            ViewState("GR_CODE") = ""
            ViewState("STORER_CODE") = Request("STORER_CODE")
            ViewState("GR_CODE") = Request("GR_CODE")

            If ViewState("STORER_CODE") = "" AndAlso ViewState("GR_CODE") = "" Then
                uiFun.displayMsg(Me, "", "Missing Key Field!", Session("gLang"))
                dsp_status.Text = "Missing Key Field for generating report!"
                dsp_status.ForeColor = Drawing.Color.Red
            Else
                isGenDownload = True
                tr_download.Style.Add("display", "none")
                ibtn_download.Visible = True
            End If

        End If

        ' ar.hideForm(Me)
    End Sub

    Protected Sub genDownloadFile()

        Dim ExcelWBObj As HSSFWorkbook

        Dim tmpltPath, tmpltFileName As String
        Dim tmpFileStream As FileStream
        Dim ReadFileStream As FileStream


        'tmpltPath = System.Configuration.ConfigurationManager.AppSettings.Item("EXCEL_TMPLT_PATH").ToString
        tmpltFileName = "CARGO_RCP.xls"

        ReadFileStream = New FileStream(Server.MapPath(tmpltFileName), FileMode.Open, FileAccess.Read)



        Dim sysFileName, sysFilePath, fileFolder As String

        fileFolder = HttpContext.Current.Cache("SYSP_TEMP_DIR")

        If Not System.IO.Directory.Exists(fileFolder) Then
            System.IO.Directory.CreateDirectory(fileFolder)
        End If


        Try
            ExcelWBObj = New HSSFWorkbook(ReadFileStream)


            'XXXXXXXXXXXXXXXXXX Excel Cell StyleXXXXXXXXXXXXXXXXXXXXX


            'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

            Dim SelectSQL As String = ""
            Dim itemDT As DataTable

            Dim colSQL As String = ""
            Dim colDT As DataTable

            Dim paP As GlobalDBFunc.DBCmdPara

            SelectSQL = " SELECT gm.GR_RCV_BY,  gm.GR_HAWB,  gm.GR_INV_NO, gm.STORER_CODE,  " & _
                        "   ISNULL(  CASE " & _
                        "     WHEN wms_storer.sto_shortname = '' " & _
                        "     THEN wms_storer.sto_name " & _
                        "     ELSE wms_storer.sto_shortname " & _
                        "   END, WMS_STORER.STO_NAME)       AS storer_messer, " & _
                        "   WMS_STORER.STO_CONT_PER_ORD     AS storer_attn, " & _
                        "   ''                            AS storer_cc, " & _
                        "   TO_CHAR(gm.GR_DATE, 'DD/MM/YYYY') AS report_date, " & _
                        "   gm.GR_REM,  gd.GRD_DISP_SEQ,  gd.GRD_REF_NO,decode(substr( gd.GRD_BATCH_NO,0,6),'@B#_E_','','@B#_M_','', gd.GRD_BATCH_NO) as GRD_BATCH_NO,  gd.GRD_PALLET_NO,  gd.GRD_CARTON_NO,  gd.GRD_ITM_NAME, " & _
                        "   gd.GRD_ITM_CODE,  gd.GRD_UOM,  gd.GRD_PCS_PER_UOM, ISNULL(gd.GRD_PCS_PER_UOM,0) * ISNULL(ga.GRA_PA_QTY,0) AS GRD_TOT_PCS,DECODE(VLOC.AR_DAMAGE_YN,'Y',0, ga.GRA_PA_QTY) as GRA_PA_QTY, DECODE(VLOC.AR_DAMAGE_YN,'Y', ga.GRA_PA_QTY,0) as GRA_DEP, " & _
                        "   (  CASE " & _
                        "     WHEN ISNULL(gd.grd_no_of_carton, 0) > 0 " & _
                        "     THEN round(gd.grd_cbm,3) " & _
                        "     ELSE 0 " & _
                        "   END) AS tot_cbm, " & _
                        "   (  CASE " & _
                        "     WHEN ISNULL(gd.grd_no_of_carton, 0) > 0 " & _
                        "     THEN gd.grd_kg " & _
                        "     ELSE 0 " & _
                        "   END) AS tot_kg, " & _
                        "   gd.GRD_LENGTH,  gd.GRD_WIDTH,  gd.GRD_HEIGHT,  gd.GRD_CBM,  gd.GRD_KG, gd.grd_po_qty, " & _
                        "   TO_CHAR(GRA_EXPIRY_DATE,'DD/MM/YYYY') as GRA_EXPIRY_DATE, TO_CHAR(GRA_MANU_DATE,'DD/MM/YYYY') as GRA_MANU_DATE, " & _
                        "   ISNULL(gd.GRD_REF_SEQ, '-1') AS grd_ref_seq, " & _
                        "   gd.GRD_NO_OF_CARTON,  gd.GRD_PCS_PER_CARTON,  '' AS total_carton, Round((gd.grd_length * gd.grd_width * gd.grd_height / 1000000),4) as ctn_cbm," & _
                        "   ISNULL(gm.GR_TOT_PALLET,0)  AS total_pallet, Round((ga.gra_pa_qty/gd.GRD_RCV_QTY) * gd.GRD_NO_OF_CARTON),4) as no_of_carton_per_pa, " & _
                        "   Round(((ga.gra_pa_qty/gd.GRD_RCV_QTY) * gd.GRD_NO_OF_CARTON) * (gd.grd_length * gd.grd_width * gd.grd_height / 1000000),4) as pa_cbm, Round((ga.gra_pa_qty/gd.GRD_RCV_QTY) * gd.GRD_NO_OF_CARTON) * gd.GRD_KG,4) as pa_kg, " & _
                        "   gm.GR_REF_NO AS ref_no,  'From ' || ISNULL(WMS_REPLENISH.RO_DEST, N'') AS from_dest, GA.GRA_LOC, vloc.ar_code, vloc.rk_code, vloc.bn_code, " & _
                        "itm.ITM_DESC, itm.ITM_CAT,  itm.ITM_TYPE,  itm.ITM_BRAND,  itm.ITM_SERIES,  itm.ITM_SERIES_NO,  itm.ITM_MODEL," & _
                        "itm.ITM_PARENT,  itm.ITM_FLAGS,  itm.ITM_REMARKS,  itm.ITM_SPEC,  itm.ITM_VEND_ITM_NO," & _
                        "itm.ITM_VEND_CODE,  itm.ITM_PROD_GROUP,  itm.ITM_SIZE,  itm.ITM_TEMP_YN,  itm.ITM_TEMP_FR_CHAR," & _
                        "itm.ITM_TEMP_TO_CHAR,  itm.ITM_TEMP_TO,  itm.ITM_SKU_NO,  itm.ITM_COLOR_CODE,  itm.ITM_SERIAL_NO," & _
                        "itm.ITM_PART_NO,  itm.ITM_PROD_NO,  itm.ITM_QTY_OF_UNIT,  itm.ITM_PCS_PER_UOM,  itm.ITM_TEMP_FR" & _
                        " FROM WMS_GOODSRCV     gm,     WMS_GOODSRCV_D   gd, " & _
                        "      WMS_STORER,     WMS_REPLENISH, wms_item itm, WMS_GOODSRCV_PA GA, V_LOCATION VLOC " & _
                        " Where gm.IMP_CODE = gd.IMP_CODE AND gm.STORER_CODE = gd.STORER_CODE " & _
                        " AND gm.GR_CODE = gd.GR_CODE  " & _
                        " AND WMS_STORER.STORER_CODE(+) = gm.STORER_CODE AND WMS_STORER.IMP_CODE(+)   = gm.IMP_CODE " & _
                        " AND WMS_REPLENISH.RO_CODE(+)   = gm.GR_DOC_NO AND WMS_REPLENISH.IMP_CODE(+) = gm.IMP_CODE " & _
                        " and gd.imp_code = itm.imp_code(+) and gd.storer_code=itm.storer_code(+) and gd.grd_itm_code=itm.itm_code(+) and gd.GRD_PACK_KEY = itm.pack_key(+) " & _
                        " AND gd.IMP_CODE = GA.IMP_CODE" & _
                        " AND gd.STORER_CODE  = GA.STORER_CODE AND gd.GR_CODE      = GA.GR_CODE " & _
                        " AND gd.GRD_ITM_CODE = GA.GRA_ITM_CODE AND gd.GRD_PACK_KEY = GA.GRA_PACK_KEY " & _
                        " AND ISNULL(gd.GRD_BATCH_NO, '')  = ISNULL(GA.GRA_BATCH_NO, '') " & _
                        " AND GA.GRA_LOC = VLOC.LOC(+) " & _
                        " AND gm.imp_code = '" & Session("IMP_CODE") & "' " & _
                        " and gm.gr_code='" & gU.dbEncode(ViewState("GR_CODE")) & "' and gm.storer_code='" & gU.dbEncode(ViewState("STORER_CODE")) & "' " & _
                        " order by gd.grd_disp_seq, to_number(gd.grd_seq) "

            itemDT = gDB.getDataTable(SelectSQL)

            If itemDT IsNot Nothing AndAlso itemDT.Rows.Count > 0 Then

                Dim TemplateSheet, ReportSheet As HSSFSheet
                Dim TemplateRow As HSSFRow

                TemplateSheet = ExcelWBObj.GetSheet("TEMPLATE")
                ReportSheet = ExcelWBObj.GetSheetAt(0)

                Dim startCol As Integer = 0

                Dim oBmp As Bitmap
                Dim tempPath As String = gU.getConfig("SYSP_TEMP_DIR")

                oBmp = Code128Rendering.MakeBarcodeImage(ViewState("GR_CODE"), 2, True)
                oBmp.Save(tempPath & "/Barcode" & ViewState("GR_CODE") & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

                Dim anchor As HSSFClientAnchor
                Dim fs As FileStream
                Dim picbuff As Byte()
                Dim picNo As Integer
                Dim EXLPIC As HSSFPicture
                Dim patriarch As HSSFPatriarch
                patriarch = TryCast(ReportSheet.CreateDrawingPatriarch(), HSSFPatriarch)

                Dim maxcount As Integer = 0

                anchor = New HSSFClientAnchor(0, 0, 0, 0, 12, 1, 0, 0)
                fs = New FileStream(tempPath & "/Barcode" & ViewState("GR_CODE") & ".jpg", FileMode.Open, FileAccess.Read)

                ReDim picbuff(fs.Length)

                fs.Read(picbuff, 0, CInt(fs.Length))

                fs.Close()
                fs.Dispose()
                fs = Nothing

                picNo = ExcelWBObj.AddPicture(picbuff, NPOI.SS.UserModel.PictureType.JPEG)

                EXLPIC = TryCast(patriarch.CreatePicture(anchor, picNo), HSSFPicture)

                'Reset the image to the original size.
                EXLPIC.Resize()
                EXLPIC.LineStyle = NPOI.SS.UserModel.LineStyle.None

                nf.setCellValue(ReportSheet, 3, 14, ViewState("GR_CODE"))
                nf.setCellValue(ReportSheet, 7, 2, itemDT.Rows(0).Item("storer_messer").ToString.Trim)
                nf.setCellValue(ReportSheet, 9, 2, itemDT.Rows(0).Item("storer_attn").ToString.Trim)
                nf.setCellValue(ReportSheet, 11, 2, itemDT.Rows(0).Item("storer_cc").ToString.Trim)
                nf.setCellValue(ReportSheet, 13, 2, itemDT.Rows(0).Item("GR_RCV_BY").ToString.Trim)

                nf.setCellValue(ReportSheet, 7, 9, itemDT.Rows(0).Item("ref_no").ToString.Trim)
                nf.setCellValue(ReportSheet, 9, 9, itemDT.Rows(0).Item("GR_HAWB").ToString.Trim)
                nf.setCellValue(ReportSheet, 11, 9, itemDT.Rows(0).Item("GR_INV_NO").ToString.Trim)
                nf.setCellValue(ReportSheet, 13, 9, itemDT.Rows(0).Item("report_date").ToString.Trim)



                paP = New GlobalDBFunc.DBCmdPara
                colSQL = "Select FLDO_FIELD_NAME, FLDO_FIELD_DESC from WMS_FIELD_OPTION Where imp_code=" & paP.AP(Session("imp_code")) & " AND STORER_CODE=" & paP.AP(ViewState("STORER_CODE")) & " AND FUN_CODE='IB_GR_CRP' AND FLDO_FIELD_OPTION='Y' and FLDO_FIELD_TYPE is null order by FLDO_DISPLAY_SEQ"
                colDT = gDB.getDataTable(colSQL, , , , paP)

                ExcelRow_num = 16

                nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(16, 16, 0, 0), ExcelRow_num, 0)
                nf.setCellValue(ReportSheet, ExcelRow_num, 0, "No.")
                startCol += 1

                nf.setCellValue(ReportSheet, ExcelRow_num, 0, "Stock No.")
                startCol += 1

                nf.setCellValue(ReportSheet, ExcelRow_num, 0, "Item Desc.")
                startCol += 1


                If colDT.Rows.Count > 0 Then
                    If colDT.Rows.Count >= 4 Then maxcount = 4 Else maxcount = colDT.Rows.Count
                    For i = 0 To maxcount - 1
                        nf.setCellValue(ReportSheet, ExcelRow_num, startCol, colDT.Rows(i).Item("FLDO_FIELD_DESC").ToString.Trim)
                        startCol += 1
                    Next
                End If


                TemplateRow = TemplateSheet.GetRow(16)
                ReportSheet.GetRow(16).Height = TemplateRow.Height


                ExcelRow_num += 1

                startCol = 0

                Dim colName As String = ""
                For i = 0 To itemDT.Rows.Count - 1
                    startCol = 0

                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(18, 18, 0, 0), ExcelRow_num, startCol)
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("grd_disp_seq").ToString.Trim)
                    startCol += 1

                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(18, 18, 1, 1), ExcelRow_num, startCol)
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("ITM_SKU_NO").ToString.Trim)
                    startCol += 1

                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(18, 18, 2, 2), ExcelRow_num, startCol)
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("ITM_DESC").ToString.Trim)
                    startCol += 1

                    If colDT.Rows.Count > 0 Then

                        colName = ""
                        For X = 0 To maxcount - 1
                            colName = colDT.Rows(X).Item("FLDO_FIELD_NAME").ToString.Trim
                            nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(18, 18, 3, 3), ExcelRow_num, startCol)
                            nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item(colName).ToString.Trim)
                            startCol += 1
                        Next
                    End If

                    startCol = 7
                    nf.CopyRange(TemplateSheet, ReportSheet, New CellRangeAddress(18, 18, 7, 17), ExcelRow_num, startCol)
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("grd_po_qty").ToString.Trim, "DEC")
                    startCol += 1
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("GRA_PA_QTY").ToString.Trim, "DEC")
                    startCol += 1
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("GRA_DEP").ToString.Trim, "DEC")
                    startCol += 1
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("no_of_carton_per_pa").ToString.Trim, "DEC")
                    startCol += 1
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("ctn_cbm").ToString.Trim, "DEC")
                    startCol += 1
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("grd_length").ToString.Trim, "DEC")
                    startCol += 1
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("grd_width").ToString.Trim, "DEC")
                    startCol += 1
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("grd_height").ToString.Trim, "DEC")
                    startCol += 1
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("pa_cbm").ToString.Trim, "DEC")
                    startCol += 1
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("pa_kg").ToString.Trim, "DEC")
                    startCol += 1
                    nf.setCellValue(ReportSheet, ExcelRow_num, startCol, itemDT.Rows(i).Item("GRA_LOC").ToString.Trim)
                    startCol += 1

                    'ReportSheet.GetRow(ExcelRow_num).HeightInPoints = ReportSheet.DefaultRowHeightInPoints

                    ExcelRow_num += 1
                    'ReportSheet.ShiftRows(ExcelRow_num, ReportSheet.LastRowNum, 1, True, False, True)
                    nf.InsertRows(ReportSheet, ExcelRow_num, 1)


                Next

                nf.removeRow(ReportSheet, ExcelRow_num)

                Dim total_cbm, total_kg As Double
                Dim total_ctn, total_rcv_qty, total_desc_qty As Integer

                total_cbm = itemDT.Compute("Sum(pa_cbm)", "")
                total_kg = itemDT.Compute("Sum(pa_kg)", "")

                total_ctn = itemDT.Compute("Sum(no_of_carton_per_pa)", "")
                total_rcv_qty = itemDT.Compute("Sum(GRA_PA_QTY)", "")
                total_desc_qty = itemDT.Compute("Sum(GRA_DEP)", "")


                nf.setCellValue(ReportSheet, ExcelRow_num, 4, itemDT.Rows(0).Item("total_pallet").ToString.Trim, "DEC")

                nf.setCellValue(ReportSheet, ExcelRow_num, 8, total_rcv_qty, "DEC")
                nf.setCellValue(ReportSheet, ExcelRow_num, 9, total_desc_qty, "DEC")
                nf.setCellValue(ReportSheet, ExcelRow_num, 10, total_ctn, "DEC")

                nf.setCellValue(ReportSheet, ExcelRow_num, 15, total_cbm, "DEC")
                nf.setCellValue(ReportSheet, ExcelRow_num, 16, total_kg, "DEC")
                nf.setCellValue(ReportSheet, ExcelRow_num + 1, 2, itemDT.Rows(0).Item("gr_rem").ToString.Trim)

                ExcelWBObj.RemoveSheetAt(ExcelWBObj.GetSheetIndex("TEMPLATE"))


                sysFileName = "Cargo Receipt" & ViewState("GR_CODE")
                sysFileName = sysFileName & "." & LCase("XLS")
                sysFilePath = fileFolder & "\" & sysFileName


                tmpFileStream = New FileStream(sysFilePath, FileMode.Create)
                ExcelWBObj.Write(tmpFileStream)

                tmpFileStream.Close()
                tmpFileStream.Dispose()

                ExcelWBObj = Nothing

                Response.Write("<script language=""JavaScript"">dsp_status.innerHTML = ""DONE"";</script>")
                Response.Write("<script language=""JavaScript"">document.forms[0].hdf_file_path.value = """ & gU.jsString(sysFilePath) & """;</script>")
                Response.Write("<script language=""JavaScript"">document.getElementById(""tr_download"").style.display = """";</script>")
                Response.Flush()

            Else
                dsp_status.Text = "No GR Record is found!"
                dsp_status.ForeColor = Drawing.Color.Red
            End If

            REM XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



            'Catch ex As Exception
            '    Throw ex
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
        Dim userFileName As String = "CARGO_RECEIPT" & Now.Date.ToString("yyyyMMdd")

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
