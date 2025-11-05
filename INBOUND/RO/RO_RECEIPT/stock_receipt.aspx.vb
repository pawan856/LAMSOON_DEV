Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms

Partial Class INBOUND_GR_GR_STOCK_REP_stock_receipt
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private rptU As New ReportUtils
    Private db As New DBfunc

    Private wmsFun As New WMSFunc

    Dim ro_code As String = ""
    Dim storer_code As String = ""
    Dim imp_code As String = ""
    Dim gr_doc_type As String = ""
    Dim EBS_PO_NO As String = ""
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("IB_RO", Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        ar.hideForm(Me)

        If Not IsPostBack Then
            ro_code = Request.Form("ro_code")
            storer_code = Request.Form("storer_code")
            gr_doc_type = Request.Form("gr_doc_type")
            EBS_PO_NO = Request.Form("EBS_PO_NO")
            ViewState("EBS_PO_NO") = EBS_PO_NO
            ViewState("ro_code") = ro_code
            ViewState("storer_code") = storer_code
            'ViewState("gr_doc_type") = gr_doc_type

            BindGV(IsPostBack)
        Else
            ro_code = ViewState("ro_code")
            storer_code = ViewState("storer_code")
            'gr_doc_type = ViewState("gr_doc_type")
        End If

    End Sub

    Private Sub BindGV(ByVal isPostBack As Boolean)
        Dim sqlString As String
        Dim nDataSource As DataTable

        Dim Paras(10) As ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        'If gr_doc_type <> "RO" Then
        '    'Response.Write("Incorrect Document Type.")
        '    'Exit Sub
        'End If

        sqlString = "SELECT RO_RCV_BY, RO_CUST_INV_NO,  WMS_REPLENISH.STORER_CODE, itm.itm_sku_no,WMS_REPLENISH.RO_EDI_PO_NO,WMS_REPLENISH.RO_TRACK_NO,ISNULL(WMS_REPLENISH.RO_CONTAINER_NO,'')RO_CONTAINER_NO, " &
                    " case when v_alt.aitm_qty_per_ctn > 0 then ceiling(WMS_REPLENISH_D.ROD_QTY / v_alt.aitm_qty_per_ctn) else 0 end as C_QTY, " &
                    "   ISNULL(CASE  " &
                    "     WHEN wms_storer.sto_shortname = ''  " &
                    "     THEN wms_storer.sto_name  " &
                    "     ELSE wms_storer.sto_shortname  " &
                    "   END, WMS_STORER.STO_NAME)         AS storer_messer, " &
                    "   WMS_STORER.STO_CONT_PER_ORD       AS storer_attn, " &
                    "   ''                                AS storer_cc, " &
                    "   Convert(varchar,RO_DATE, 103) AS report_date, " &
                    "   RO_REM,  ROD_DISP_SEQ,  ROD_REF_NO,  ROD_BATCH_NO AS old_batch_no,  ROD_PALLET_NO, " &
                    "  ROD_CARTON_NO,  ROD_ITM_NAME,  ROD_ITM_CODE,  ROD_UOM,  ROD_PCS_PER_UOM,  ROD_TOT_PCS AS old_total, " &
                    "   ROD_cbm AS tot_cbm,  ROD_kg AS tot_kg,  ROD_LENGTH,  ROD_WIDTH,  ROD_HEIGHT,  ROD_CBM,  ROD_KG, " &
                    "   ''                        AS total_carton, " &
                    "   RO_REF_NO              AS ref_no,  'From '  + ISNULL(WMS_REPLENISH.RO_DEST, N'') AS from_dest, " &
                    "   '' AS product_type,  ISNULL(ROD_PCS_PER_UOM,0) * ISNULL(ROD_TOT_PCS,0) AS ROD_TOT_PCS, " &
                    " Convert(varchar, ROD_MANU_DATE,103) as ROD_MANU_DATE, convert(varchar, ROD_EXPIRY_DATE,103) as ROD_EXPIRY_DATE " &
                    " FROM  WMS_REPLENISH inner join WMS_REPLENISH_D" &
                    " on WMS_REPLENISH_D.IMP_CODE             = WMS_REPLENISH.IMP_CODE " &
                    " AND WMS_REPLENISH_D.STORER_CODE            = WMS_REPLENISH.STORER_CODE " &
                    " AND WMS_REPLENISH_D.RO_CODE                = WMS_REPLENISH.RO_CODE " &
                    " left outer join WMS_STORER " &
                    " ON WMS_STORER.STORER_CODE = WMS_REPLENISH.STORER_CODE " &
                    " AND WMS_STORER.IMP_CODE    = WMS_REPLENISH.IMP_CODE " &
                    " LEFT OUTER JOIN  wms_item itm " &
                    " ON WMS_REPLENISH_D.imp_code = itm.imp_code AND WMS_REPLENISH_D.storer_code = itm.storer_code AND WMS_REPLENISH_D.ROD_ITM_CODE = itm.itm_code and WMS_REPLENISH_D.ROD_PACK_KEY = itm.pack_key " &
                    " LEFT OUTER JOIN V_ALT_VEND_ITEM v_alt " &
                    " ON WMS_REPLENISH_D.imp_code = v_alt.imp_code AND WMS_REPLENISH_D.storer_code = v_alt.storer_code AND WMS_REPLENISH_D.ROD_ITM_CODE = v_alt.itm_code and WMS_REPLENISH_D.ROD_PACK_KEY = v_alt.pack_key " &
                    " Where WMS_REPLENISH.imp_code = '" & imp_code & "' " &
                    " and WMS_REPLENISH.ro_code='" & gU.dbEncode(ro_code) & "' and WMS_REPLENISH.storer_code='" & gU.dbEncode(storer_code) & "' "

        sqlString = sqlString & "order by rod_disp_seq, convert(int, rod_seq) "


        '"(gd.grd_pcs_per_uom * gd.grd_cbm) as tot_cbm, " & _
        '            "(gd.grd_pcs_per_uom * gd.grd_kg) as tot_kg, " & _

        nDataSource = gDB.getDataTable(sqlString)

        Dim totalCarton As Integer = 0

        Dim tot_cbm As Double = 0
        Dim tot_kg As Double = 0

        Dim oBmp As Bitmap
        Dim tempPath As String = gU.getConfig("SYSP_TEMP_DIR") & "\RO\" & EBS_PO_NO
        If nDataSource.Rows.Count > 0 Then
            Dim RO_CONTAINER_NO As String = nDataSource.Rows(0)("RO_CONTAINER_NO")
            If Not IO.Directory.Exists(tempPath) Then
                IO.Directory.CreateDirectory(tempPath)
            End If

            'oBmp = rptU.GetCode39(ro_code)
            'oBmp.Save(tempPath & "/Barcode.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

            'oBmp = Code128Rendering.MakeBarcodeImage(ro_code, 2, True)
            'oBmp.Save(tempPath & "\" & ro_code & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

            oBmp = Code128Rendering.MakeQRImage(EBS_PO_NO, tempPath & "\" & EBS_PO_NO & ".jpg")
            oBmp.Save(tempPath & "\" & EBS_PO_NO & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
            Paras(0) = New ReportParameter("BarCodePath", "file:///" & tempPath & "\" & EBS_PO_NO & ".jpg")
            If RO_CONTAINER_NO <> "" Then
                oBmp = Code128Rendering.MakeQRImage(RO_CONTAINER_NO, tempPath & "\" & RO_CONTAINER_NO & ".jpg")
                oBmp.Save(tempPath & "\" & RO_CONTAINER_NO & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

            End If
            Paras(10) = New ReportParameter("BarCodeCNTRNoPath", "file:///" & tempPath & "\" & RO_CONTAINER_NO & ".jpg")
            Paras(1) = New ReportParameter("RO_CODE", ro_code)
                Paras(2) = New ReportParameter("SHOW_ROD_ITM_CODE", ReturnShowFlag("ROD_ITM_CODE", storer_code, "IB_RO_RCP"))
                Paras(3) = New ReportParameter("SHOW_ROD_ITM_NAME", ReturnShowFlag("ROD_ITM_NAME", storer_code, "IB_RO_RCP"))
                Paras(4) = New ReportParameter("SHOW_ITM_SKU_NO", ReturnShowFlag("ITM_SKU_NO", storer_code, "IB_RO_RCP"))
                Paras(5) = New ReportParameter("SHOW_ROD_BATCH_NO", ReturnShowFlag("ROD_BATCH_NO", storer_code, "IB_RO_RCP"))
                Paras(6) = New ReportParameter("SHOW_ROD_PALLET_NO", ReturnShowFlag("ROD_PALLET_NO", storer_code, "IB_RO_RCP"))
                Paras(7) = New ReportParameter("SHOW_ROD_CARTON_NO", ReturnShowFlag("ROD_CARTON_NO", storer_code, "IB_RO_RCP"))
                Paras(8) = New ReportParameter("SHOW_ROD_MANU_DATE", ReturnShowFlag("ROD_MANU_DATE", storer_code, "IB_RO_RCP"))
                Paras(9) = New ReportParameter("SHOW_ROD_EXPIRY_DATE", ReturnShowFlag("ROD_EXPIRY_DATE", storer_code, "IB_RO_RCP"))


                'For i As Integer = 0 To nDataSource.Rows.Count - 1
                'Dim cbm As Double = gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("grd_length").ToString, 0) * _
                '                            gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("grd_width").ToString, 0) * _
                '                            gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("grd_height").ToString, 0)

                'If cbm <> 0 Then cbm = cbm / 1000000

                'nDataSource.Rows(i).Item("tot_cbm") = cbm * gU.decodeEmptyCInt(nDataSource.Rows(i).Item("grd_no_of_carton").ToString, 1)
                'nDataSource.Rows(i).Item("tot_cbm") = gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("tot_cbm").ToString, 0) * gU.decodeEmptyCInt(nDataSource.Rows(i).Item("grd_no_of_carton").ToString, 1)
                'nDataSource.Rows(i).Item("tot_kg") = gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("tot_kg").ToString, 0) * gU.decodeEmptyCInt(nDataSource.Rows(i).Item("grd_no_of_carton").ToString, 1)
                'nDataSource.Rows(i).Item("grd_kg") = gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("grd_kg").ToString, 0) * gU.decodeEmptyCInt(nDataSource.Rows(i).Item("grd_no_of_carton").ToString, 1)
                'totalCarton = totalCarton + gU.decodeEmptyCInt(nDataSource.Rows(i).Item("grd_no_of_carton").ToString, 0)
                'Next

                'nDataSource.Rows(0).Item("total_carton") = totalCarton

                'nDataSource.Rows(0).Item("tot_kg") = tot_kg / nDataSource.Rows.Count * totalCarton
                'nDataSource.Rows(0).Item("tot_cbm") = Math.Round(tot_cbm / nDataSource.Rows.Count * totalCarton, 4)
            End If

            nDataSource.AcceptChanges()

        If nDataSource.Rows.Count > 0 Then
            'For i As Integer = 0 To nDataSource.Rows.Count - 1
            '    Dim rowCartonQty As Integer = 0

            '    wmsFun.getCartonQty(rowCartonQty, "rod_carton_no", Nothing, False, nDataSource.Rows(i))
            '    nDataSource.Rows(i).Item("C_QTY") = rowCartonQty
            'Next

            'nDataSource.AcceptChanges()

            reportSource(nDataSource, Paras)

        Else
            Response.Write("No Stock Receipt Found.")

        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "INBOUND\RO\RO_RECEIPT\stock_receipt.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))
            ReportViewer1.LocalReport.EnableExternalImages = True

            If Not IsNothing(paraarray) Then
                ReportViewer1.LocalReport.SetParameters(paraarray)
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

    Protected Sub btnSort_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSort.Click
        BindGV(True)
    End Sub

    Protected Function ReturnShowFlag(ByVal field_name As String, ByVal storer_code As String, ByVal fun_code As String) As String
        Dim selectSQL As String = ""
        Dim paP As GlobalDBFunc.DBCmdPara
        Dim showFlag As String = "Y"

        paP = New GlobalDBFunc.DBCmdPara
        selectSQL = "Select ISNULL(FLDO_FIELD_OPTION,'Y') as FLDO_FIELD_OPTION from wms_field_option where imp_code=" & paP.AP(Session("imp_code")) & " AND STORER_CODE=" & paP.AP(storer_code) & " AND fun_code=" & paP.AP(fun_code) & " AND FLDO_FIELD_NAME=" & paP.AP(field_name)

        showFlag = gU.decodeNullOrEmpty(db.getValueFromSQL(selectSQL, , , paP), "Y")

        Return showFlag
    End Function
End Class
