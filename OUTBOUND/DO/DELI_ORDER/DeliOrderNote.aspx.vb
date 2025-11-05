Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms

Partial Class OUTBOUND_DO_DELI_ORDER_DeliOrderNote
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private rptU As New ReportUtils
    Private DB As New DBfunc

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim do_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""

        Dim Paras(3) As ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            ViewState("do_code") = ""
            do_code = Request.Form("do_code")
            storer_code = Request.Form("storer_code")
            ViewState("do_code") = do_code
            ViewState("storer_code") = storer_code

        End If
        
        storer_code = ViewState("storer_code")
        do_code = ViewState("do_code")

        If Not IsPostBack Then
            sqlString = " SELECT     dm.STORER_CODE, dm.DO_CODE, dm.DO_ISSUED_BY, convert(varchar, dm.DO_DATE, 103) AS do_date, " & _
                     " Convert(varchar, dm.DO_CONF_DELDATE, 103) AS do_conf_deldate, Convert(varchar(5), dm.DO_CONF_DELTIME, 108) AS do_conf_deltime, dm.CUS_NAME, " & _
                     " (CASE WHEN not dm.do_addr1 is null THEN dm.do_addr1 ELSE '' END + CASE WHEN not dm.do_addr2 is null THEN char(13) + char(10) " & _
                     " + dm.do_addr2 ELSE '' END + CASE WHEN not dm.do_addr3 is null THEN char(13) + char(10) " & _
                     " + dm.do_addr3 ELSE '' END + CASE WHEN not dm.do_area_del is null THEN char(13) + char(10) " & _
                     " + dm.do_area_del ELSE '' END + CASE WHEN not dm.do_region_del is null  " & _
                     " THEN ' ' + dm.do_region_del ELSE '' END + CASE WHEN not dm.do_country_del is null THEN char(13) + char(10) " & _
                     " + dm.do_country_del ELSE '' END) AS do_addr, dm.DO_CUS_CONT, dm.DO_CUS_CONT_TEL, dm.DO_CUS_REF_NO, dd.DOD_DISP_SEQ, " & _
                     " dd.DOD_ITM_CODE, itm.ITM_DESC as DOD_ITM_DESC, dd.DOD_QTY, dd.DOD_REM, convert(decimal, ISNULL(dm.DO_TOTL_PALLETS, 0)) AS total_pallet, " & _
                     " convert(decimal, ISNULL(dm.DO_TOTL_CARTONS, 0)) AS total_carton, dm.DO_DRIVER, dm.DO_INV_NO,dd.DOD_BATCH_NO, " & _
                     " WMS_VENDOR.VND_NAME,itm.ITM_SKU_NO,dd.DOD_TOT_CBM, dd.DOD_TOT_WGT, Convert(varchar, DOD_EXPIRY_DATE,103) as DOD_EXPIRY_DATE, Convert(varchar, DOD_MANU_DATE,103) as DOD_MANU_DATE, " & _
                     " DO_CONSIGNEE, (CASE WHEN not dm.DO_CONSIGNEE_ADDR1 is null THEN dm.DO_CONSIGNEE_ADDR1 ELSE '' END + CASE WHEN not dm.DO_CONSIGNEE_ADDR2 is null THEN char(13) + char(10) " & _
                     " + dm.DO_CONSIGNEE_ADDR2 ELSE '' END + CASE WHEN not dm.DO_CONSIGNEE_ADDR3 is null THEN char(13) + char(10) " & _
                     " + dm.DO_CONSIGNEE_ADDR3 ELSE '' END) as DO_CONSIGNEE_ADDR1, itm.ITM_MODEL, dm.DO_FTRACK_NO,DO_DELIVERY_RMKS,Convert(varchar, DO_ARRIVAL_DATE,103) as DO_ARRIVAL_DATE, DO_POST_CODE,DO_REM,DO_EDI_SIR_NO,DO_EDI_WIT_NO, " & _
                     " CONVERT(varchar, DO_TARGET_DELDATE, 103) as DO_TARGET_DELDATE " & _
                     " FROM WMS_DELV_ORDER dm INNER JOIN " & _
                     " WMS_DELV_ORDER_D dd ON dm.IMP_CODE = dd.IMP_CODE AND dm.STORER_CODE = dd.STORER_CODE AND " & _
                     " dm.DO_CODE = dd.DO_CODE Left outer join " & _
                     " WMS_VENDOR ON dd.IMP_CODE = WMS_VENDOR.IMP_CODE AND dd.STORER_CODE = WMS_VENDOR.STORER_CODE AND " & _
                     " dd.DOD_VND_CODE = WMS_VENDOR.VND_CODE " & _
                     " LEFT OUTER JOIN WMS_ITEM itm " & _
                     " ON dd.IMP_CODE = itm.IMP_CODE AND dd.STORER_CODE=itm.STORER_CODE AND dd.DOD_ITM_CODE = itm.ITM_CODE AND dd.DOD_PACK_KEY = itm.PACK_KEY " & _
                    "Where dm.imp_code = '" & imp_code & "' " & _
                    "and dm.do_code='" & gU.dbEncode(do_code) & "' and dm.storer_code='" & gU.dbEncode(storer_code) & "' " & _
                    "order by dd.dod_disp_seq "

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then

                Dim oBmp As Bitmap
                Dim iBCWriter As New ZXing.BarcodeWriter

                iBCWriter.Format = ZXing.BarcodeFormat.QR_CODE

                Dim ImgByte() As Byte
                Dim ByteString As String

                oBmp = iBCWriter.Write(do_code)

                Using IOStream As New MemoryStream
                    oBmp.Save(IOStream, Imaging.ImageFormat.Png)
                    ImgByte = IOStream.ToArray

                    ByteString = Convert.ToBase64String(ImgByte)
                    Paras(0) = New ReportParameter("QRImg", ByteString)

                End Using

                iBCWriter.Format = ZXing.BarcodeFormat.CODE_128
                iBCWriter.Options.PureBarcode = True

                Dim FtrackNo As String = ""
                FtrackNo = DB.getValueFromSQL("SELECT MAX(DOB_FTRACK_NO) as DOB_FTRACK_NO from WMS_DO_BOX_TRACK WHERE IMP_CODE='" & Session("IMP_CODE") & "' AND STORER_CODE='" & gU.dbEncode(nDataSource.Rows(0).Item("STORER_CODE").ToString.Trim) & "' AND DO_CODE='" & gU.dbEncode(nDataSource.Rows(0).Item("DO_CODE").ToString.Trim) & "'")
                Paras(3) = New ReportParameter("DO_FTRACK_NO", FtrackNo)

                oBmp = iBCWriter.Write(gU.decodeNullOrEmpty(FtrackNo, "-"))

                Using IOStream As New MemoryStream
                    oBmp.Save(IOStream, Imaging.ImageFormat.Png)
                    ImgByte = IOStream.ToArray

                    ByteString = Convert.ToBase64String(ImgByte)
                    Paras(1) = New ReportParameter("Code128Img", ByteString)
                End Using

                ViewState("dt") = nDataSource
                ViewState("Paras") = Paras
                Dim Mwp_no As String = ""
                Mwp_no = DB.getValueFromSQL("SELECT MAX(PLD_WAVE_PICK_NO) as PLD_WAVE_PICK_NO from WMS_DO_PICKLIST_D WHERE IMP_CODE='" & Session("IMP_CODE") & "' AND STORER_CODE='" & gU.dbEncode(nDataSource.Rows(0).Item("STORER_CODE").ToString.Trim) & "' AND DO_CODE='" & gU.dbEncode(nDataSource.Rows(0).Item("DO_CODE").ToString.Trim) & "'")
                Paras(2) = New ReportParameter("PLD_WAVE_PICK_NO", Mwp_no)

                reportSource(nDataSource, Paras)

            Else
                Response.Write("No Delivery Order Note Found.")
            End If
        End If
        

    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "OUTBOUND\DO\DELI_ORDER\AE_delinote.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))
            ReportViewer1.LocalReport.EnableExternalImages = True

            If Not IsNothing(paraarray) Then
                ReportViewer1.LocalReport.SetParameters(paraarray)
                ReportViewer1.ZoomMode = ZoomMode.PageWidth
            End If

            Try
                ReportViewer1.LocalReport.Refresh()

            Catch ex As OutOfMemoryException
                GC.Collect()
                ReportViewer1.LocalReport.Refresh()
            End Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Function ReturnShowFlag(ByVal field_name As String, ByVal storer_code As String, ByVal fun_code As String) As String
        Dim selectSQL As String = ""
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim showFlag As String = "Y"

        'paP = New GlobalDBFunc.DBCmdPara
        'selectSQL = "Select ISNULL(FLDO_FIELD_OPTION,'Y') as FLDO_FIELD_OPTION from wms_field_option where imp_code=" & paP.AP(Session("imp_code")) & " AND STORER_CODE=" & paP.AP(storer_code) & " AND fun_code=" & paP.AP(fun_code) & " AND FLDO_FIELD_NAME=" & paP.AP(field_name)

        'showFlag = gU.decodeNullOrEmpty(db.getValueFromSQL(selectSQL, , , paP), "Y")

        Return showFlag
    End Function
End Class
