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
    Private DB As New DBfunc
    Private wmsFun As New WMSFunc

    Dim gr_code As String = ""
    Dim storer_code As String = ""
    Dim imp_code As String = ""
    Dim gr_doc_type As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("IB_GR", Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        ar.hideForm(Me)

        If Not IsPostBack Then
            gr_code = Request.Form("gr_code")
            storer_code = Request.Form("storer_code")
            gr_doc_type = Request.Form("gr_doc_type")

            ViewState("gr_code") = gr_code
            ViewState("storer_code") = storer_code
            ViewState("gr_doc_type") = gr_doc_type

            BindGV(IsPostBack)
        Else
            gr_code = ViewState("gr_code")
            storer_code = ViewState("storer_code")
            gr_doc_type = ViewState("gr_doc_type")
        End If

        If storer_code = "BLP01" Then
            btnSort.Visible = True
        End If
    End Sub

    Private Sub BindGV(ByVal isPostBack As Boolean)
        Dim sqlString As String
        Dim nDataSource As DataTable

        Dim Paras(9) As ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

      
        If gr_doc_type <> "RO" Then
            'Response.Write("Incorrect Document Type.")
            'Exit Sub
        End If

        sqlString = " SELECT gm.GR_RCV_BY,  gm.GR_HAWB,  gm.GR_INV_NO,  1 AS C_QTY,  gm.STORER_CODE,  " & _
                    " ISNULL(  CASE " & _
                    "   WHEN wms_storer.sto_shortname = '' " & _
                    "   THEN wms_storer.sto_name " & _
                    "   ELSE wms_storer.sto_shortname " & _
                    " END, WMS_STORER.STO_NAME)       AS storer_messer, " & _
                    " WMS_STORER.STO_CONT_PER_ORD     AS storer_attn, " & _
                    " ''                            AS storer_cc, " & _
                    " convert(varchar,gm.GR_DATE, 103) AS report_date, " & _
                    " gm.GR_REM,  gd.GRD_DISP_SEQ,  gd.GRD_REF_NO,itm.itm_sku_no, case substring( gd.GRD_BATCH_NO,0,6) when '@B#_E_' then '' when '@B#_M_' then '' else gd.GRD_BATCH_NO end as GRD_BATCH_NO,  gd.GRD_PALLET_NO,  gd.GRD_CARTON_NO,  gd.GRD_ITM_NAME, " & _
                    " gd.GRD_ITM_CODE,  gd.GRD_UOM,  gd.GRD_PCS_PER_UOM, ISNULL(gd.GRD_PCS_PER_UOM,0) * ISNULL(gd.GRD_TOT_PCS,0) AS GRD_TOT_PCS,  gd.GRD_RCV_QTY AS GRD_RCV_QTY, " & _
                    " (  CASE " & _
                    "   WHEN ISNULL(gd.grd_no_of_carton, 0) > 0 " & _
                    "   THEN round(gd.grd_cbm,3) " & _
                    "   ELSE 0 " & _
                    " END) AS tot_cbm, " & _
                    " (  CASE " & _
                    "   WHEN ISNULL(gd.grd_no_of_carton, 0) > 0 " & _
                    "   THEN gd.grd_kg " & _
                    "   ELSE 0 " & _
                    " END) AS tot_kg, " & _
                    " gd.GRD_LENGTH,  gd.GRD_WIDTH,  gd.GRD_HEIGHT,  gd.GRD_CBM,  gd.GRD_KG, " & _
                    " convert(varchar, GRD_EXPIRY_DATE,103 ) as GRD_EXPIRY_DATE, convert(varchar, GRD_MANU_DATE,103) as GRD_MANU_DATE, " & _
                    " ISNULL(gd.GRD_REF_SEQ, '-1') AS grd_ref_seq, " & _
                    " gd.GRD_NO_OF_CARTON,  gd.GRD_PCS_PER_CARTON,  '' AS total_carton, " & _
                    " ISNULL(gm.GR_TOT_PALLET,0)  AS total_pallet, " & _
                    " gm.GR_REF_NO AS ref_no,  'From ' + ISNULL(WMS_REPLENISH.RO_DEST, N'') AS from_dest,  '' AS product_type " & _
                    " FROM WMS_GOODSRCV     gm inner join WMS_GOODSRCV_D   gd " & _
                    " on gm.IMP_CODE     = gd.IMP_CODE AND gm.STORER_CODE = gd.STORER_CODE " & _
                    " AND gm.GR_CODE     = gd.GR_CODE  " & _
                    " left outer join WMS_STORER on WMS_STORER.STORER_CODE = gm.STORER_CODE AND WMS_STORER.IMP_CODE   = gm.IMP_CODE " & _
                    " left outer join WMS_REPLENISH on WMS_REPLENISH.RO_CODE   = gm.GR_DOC_NO AND WMS_REPLENISH.IMP_CODE = gm.IMP_CODE " & _
                    " left outer join wms_item itm on gd.imp_code = itm.imp_code and gd.storer_code=itm.storer_code and gd.grd_itm_code=itm.itm_code and gd.GRD_PACK_KEY = itm.pack_key " & _
                    " Where gm.imp_code = '" & imp_code & "' " & _
                    " and gm.gr_code='" & gU.dbEncode(gr_code) & "' and gm.storer_code='" & gU.dbEncode(storer_code) & "' "


        If isPostBack AndAlso storer_code = "BLP01" Then
            sqlString = sqlString & "order by " & _
                                    "case when ISNUMERIC(gd.grd_carton_no) = 0 then " & _
                                     "case when isnumeric(SUBSTRING(gd.grd_carton_no, 0, CHARINDEX('-', gd.GRD_CARTON_NO,0))) = 0 then " & _
                                      "case when isnumeric(SUBSTRING(gd.grd_carton_no, 0, CHARINDEX(',', gd.GRD_CARTON_NO,0))) = 0 then " & _
                                       "gd.grd_disp_seq " & _
                                      "end " & _
                                     "end " & _
                                    "end, " & _
                                    "case when ISNUMERIC(gd.grd_carton_no) = 1 then to_number(gd.grd_carton_no) else " & _
                                     "case when isnumeric(SUBSTRING(gd.grd_carton_no, 0, CHARINDEX('-', gd.GRD_CARTON_NO,0))) = 1 then " & _
                                      "to_number(SUBSTRING(gd.grd_carton_no, 0, CHARINDEX('-', gd.GRD_CARTON_NO,0))) " & _
                                     "else " & _
                                      "case when isnumeric(SUBSTRING(gd.grd_carton_no, 0, CHARINDEX(',', gd.GRD_CARTON_NO,0))) = 1 then " & _
                                       "to_number(SUBSTRING(gd.grd_carton_no, 0, CHARINDEX(',', gd.GRD_CARTON_NO,0))) " & _
                                      "else " & _
                                       "to_number(gd.grd_seq) " & _
                                      "end " & _
                                     "end " & _
                                    "end"
        Else
            sqlString = sqlString & "order by gd.grd_disp_seq, convert(int,gd.grd_seq) "
        End If

        '"(gd.grd_pcs_per_uom * gd.grd_cbm) as tot_cbm, " & _
        '            "(gd.grd_pcs_per_uom * gd.grd_kg) as tot_kg, " & _

        nDataSource = gDB.getDataTable(sqlString)

        Dim totalCarton As Integer = 0

        Dim tot_cbm As Double = 0
        Dim tot_kg As Double = 0

        Dim oBmp As bitmap
        Dim tempPath As String = gU.getConfig("SYSP_TEMP_DIR") & "\GR\" & gr_code
        If nDataSource.Rows.Count > 0 Then

            If Not IO.Directory.Exists(tempPath) Then
                IO.Directory.CreateDirectory(tempPath)
            End If

            'oBmp = rptU.GetCode39(gr_code)
            'oBmp.Save(tempPath & "\Barcode.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

            oBmp = Code128Rendering.MakeBarcodeImage(gr_code, 2, True)
            oBmp.Save(tempPath & "\Barcode.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)


            Paras(0) = New ReportParameter("BarCodePath", "file:///" & tempPath & "\Barcode.jpg")
            Paras(1) = New ReportParameter("GR_CODE", gr_code)

            Paras(2) = New ReportParameter("SHOW_GRD_ITM_CODE", ReturnShowFlag("GRD_ITM_CODE", storer_code, "IB_GR_RCP"))
            Paras(3) = New ReportParameter("SHOW_GRD_ITM_NAME", ReturnShowFlag("GRD_ITM_NAME", storer_code, "IB_GR_RCP"))
            Paras(4) = New ReportParameter("SHOW_ITM_SKU_NO", ReturnShowFlag("ITM_SKU_NO", storer_code, "IB_GR_RCP"))
            Paras(5) = New ReportParameter("SHOW_GRD_BATCH_NO", ReturnShowFlag("GRD_BATCH_NO", storer_code, "IB_GR_RCP"))
            Paras(6) = New ReportParameter("SHOW_GRD_PALLET_NO", ReturnShowFlag("GRD_PALLET_NO", storer_code, "IB_GR_RCP"))
            Paras(7) = New ReportParameter("SHOW_GRD_CARTON_NO", ReturnShowFlag("GRD_CARTON_NO", storer_code, "IB_GR_RCP"))
            Paras(8) = New ReportParameter("SHOW_GRD_EXPIRY_DATE", ReturnShowFlag("GRD_EXPIRY_DATE", storer_code, "IB_GR_RCP"))
            Paras(9) = New ReportParameter("SHOW_GRD_MANU_DATE", ReturnShowFlag("GRD_MANU_DATE", storer_code, "IB_GR_RCP"))

            For i As Integer = 0 To nDataSource.Rows.Count - 1
                'Dim cbm As Double = gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("grd_length").ToString, 0) * _
                '                            gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("grd_width").ToString, 0) * _
                '                            gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("grd_height").ToString, 0)

                'If cbm <> 0 Then cbm = cbm / 1000000

                'nDataSource.Rows(i).Item("tot_cbm") = cbm * gU.decodeEmptyCInt(nDataSource.Rows(i).Item("grd_no_of_carton").ToString, 1)
                'nDataSource.Rows(i).Item("tot_cbm") = gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("tot_cbm").ToString, 0) * gU.decodeEmptyCInt(nDataSource.Rows(i).Item("grd_no_of_carton").ToString, 1)
                nDataSource.Rows(i).Item("tot_kg") = gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("tot_kg").ToString, 0) * gU.decodeEmptyCInt(nDataSource.Rows(i).Item("grd_no_of_carton").ToString, 1)
                nDataSource.Rows(i).Item("grd_kg") = gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("grd_kg").ToString, 0) * gU.decodeEmptyCInt(nDataSource.Rows(i).Item("grd_no_of_carton").ToString, 1)
                totalCarton = totalCarton + gU.decodeEmptyCInt(nDataSource.Rows(i).Item("grd_no_of_carton").ToString, 0)
            Next

            nDataSource.Rows(0).Item("total_carton") = totalCarton

            'nDataSource.Rows(0).Item("tot_kg") = tot_kg / nDataSource.Rows.Count * totalCarton
            'nDataSource.Rows(0).Item("tot_cbm") = Math.Round(tot_cbm / nDataSource.Rows.Count * totalCarton, 4)
        End If

        nDataSource.AcceptChanges()

        If nDataSource.Rows.Count > 0 Then
            For i As Integer = 0 To nDataSource.Rows.Count - 1
                Dim rowCartonQty As Integer = 0

                wmsFun.getCartonQty(rowCartonQty, "grd_carton_no", Nothing, False, nDataSource.Rows(i))
                nDataSource.Rows(i).Item("C_QTY") = rowCartonQty
            Next

            nDataSource.AcceptChanges()

            reportSource(nDataSource, Paras)

        Else
            If storer_code = "BLP01" Then
                Response.Write("No Stock Receipt Found Or Item's Weigth and CBM is zero.")
            Else
                Response.Write("No Stock Receipt Found.")
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "INBOUND\GR\GR_STOCK_REP\stock_receipt.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("rptDataSet_gr_report", sourcetbl))
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

        'paP = New GlobalDBFunc.DBCmdPara
        'selectSQL = "Select ISNULL(FLDO_FIELD_OPTION,'Y') as FLDO_FIELD_OPTION from wms_field_option where imp_code=" & paP.AP(Session("imp_code")) & " AND STORER_CODE=" & paP.AP(storer_code) & " AND fun_code=" & paP.AP(fun_code) & " AND FLDO_FIELD_NAME=" & paP.AP(field_name)

        'showFlag = gU.decodeNullOrEmpty(db.getValueFromSQL(selectSQL, , , paP), "Y")

        Return showFlag
    End Function
End Class
