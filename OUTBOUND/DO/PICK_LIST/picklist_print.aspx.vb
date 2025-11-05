Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms

Partial Class OUTBOUND_DO_PICKLIST_PRINT
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private rptU As New ReportUtils
    Const formatStr As String = "V"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim do_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""

        Dim Paras(4) As ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        ViewState("do_code") = ""

        do_code = Request.Form("do_code")
        storer_code = Request.Form("storer_code")
        If Not IsPostBack Then
            ViewState("do_code") = do_code

            sqlString = " SELECT wms_do_picklist_d.PLD_SEQ,  wms_do_picklist_d.PLD_ITEM_NO, " & _
                        " wms_do_picklist_d.PLD_PACK_KEY,  wms_do_picklist_d.PLD_LOC,  wms_do_picklist_d.PLD_FOI_QTY, wms_do_picklist_d.PLD_SERIAL_NO, wms_do_picklist_d.PLD_QTY2," & _
                        " wms_do_picklist_d.PLD_PALLET_NO,  wms_do_picklist_d.PLD_BATCH_NO,  wms_item.itm_name, wms_item.ITM_SKU_NO, dod.DOD_REM,ISNULL(Convert(varchar, wms_do_picklist_d.PLD_EXPIRY_DATE,103), convert(varchar, loc.ILOC_EXPIRY_DATE,103))as ILOC_EXPIRY_DATE, " & _
                        " ISNULL(Convert(varchar, wms_do_picklist_d.PLD_MANU_DATE,103), convert(varchar, loc.ILOC_MANU_DATE,103))as ILOC_MANU_DATE,isNULL(wms_do_picklist_d.PLD_PL_LIST_NO,1) as PLD_PL_LIST_NO, dod.DOD_DISP_SEQ, WMS_WH_BIN.BN_CSMS_CODE, WMS_ITEM_WH.IW_PICK_LOC, B2.BN_CSMS_CODE as BN_CSMS_CODE2,  " & _
                        " CONVERT(varchar, dod.DOD_DATE_REQ,103) as CO_DATE_REQ " & _
                        " FROM wms_do_picklist_d inner join wms_item " & _
                        " on wms_do_picklist_d.imp_code = wms_item.imp_code " & _
                        " AND wms_do_picklist_d.STORER_CODE = wms_item.STORER_CODE " & _
                        " AND  wms_do_picklist_d.PLD_ITEM_NO = wms_item.itm_code " & _
                        " and wms_do_picklist_d.PLD_PACK_KEY= wms_item.PACK_KEY " & _
                        " left outer join (select imp_code, storer_code, do_code, dod_itm_code, DOD_PALLET_NO, DOD_PACK_KEY, DOD_BATCH_NO, max(DOD_REM) as DOD_REM, max(DOD_DISP_SEQ) as DOD_DISP_SEQ,min(DOD_DATE_REQ) as DOD_DATE_REQ from wms_delv_order_d " & _
                        " group by imp_code, storer_code, do_code, DOD_ITM_CODE, DOD_PALLET_NO, DOD_PACK_KEY, DOD_BATCH_NO) dod " & _
                        " on wms_do_picklist_d.imp_code  = dod.imp_code  " & _
                        " AND wms_do_picklist_d.STORER_CODE = dod.STORER_CODE " & _
                        " AND wms_do_picklist_d.PLD_ITEM_NO = dod.DOD_ITM_CODE " & _
                        " AND wms_do_picklist_d.PLD_PACK_KEY= dod.DOD_PACK_KEY " & _
                        " AND ISNULL(wms_do_picklist_d.PLD_BATCH_NO,'') = ISNULL(dod.DOD_BATCH_NO,'') " & _
                        " AND ISNULL(wms_do_picklist_d.PLD_PALLET_NO,'000') = ISNULL(dod.DOD_PALLET_NO,'000')" & _
                        " AND wms_do_picklist_d.do_code= dod.do_code " & _
                        " left outer join wms_item_loc_bal loc " & _
                        " ON wms_do_picklist_d.imp_code = loc.imp_code and wms_do_picklist_d.STORER_CODE = loc.STORER_CODE " & _
                        " AND wms_do_picklist_d.PLD_ITEM_NO = loc.itm_code AND wms_do_picklist_d.PLD_PACK_KEY = loc.pack_key " & _
                        " AND ISNULL(wms_do_picklist_d.PLD_BATCH_NO,'') = ISNULL(loc.ILOC_BATCH_NO,'') " & _
                        " AND ISNULL(wms_do_picklist_d.PLD_PALLET_NO,'000') = ISNULL(loc.ILOC_PALLET_NO,'000') AND wms_do_picklist_d.PLD_LOC = loc.ILOC_LOC  " & _
                        " left outer join WMS_WH_BIN " & _
                        " ON WMS_WH_BIN.LOC_KEY = loc.ILOC_LOC " & _
                        " left outer join WMS_ITEM_WH " & _
                        " ON WMS_ITEM_WH.IMP_CODE = WMS_ITEM.IMP_CODE " & _
                        " AND WMS_ITEM_WH.STORER_CODE = WMS_ITEM.STORER_CODE " & _
                        " AND WMS_ITEM_WH.ITM_CODE = WMS_ITEM.ITM_CODE " & _
                        " AND WMS_ITEM_WH.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                        " AND WMS_ITEM_WH.WH_CODE = WMS_DO_PICKLIST_D.PLD_WH " & _
                        " left outer join WMS_WH_BIN B2 " & _
                        " ON WMS_ITEM_WH.IW_PICK_LOC = B2.LOC_KEY " & _
                        " left outer join wms_delv_order " & _
                        " ON WMS_DO_PICKLIST_D.IMP_CODE = WMS_DELV_ORDER.IMP_CODE AND " & _
                        " WMS_DO_PICKLIST_D.STORER_CODE = WMS_DELV_ORDER.STORER_CODE AND WMS_DO_PICKLIST_D.DO_CODE = WMS_DELV_ORDER.DO_CODE  " & _
                        " LEFT OUTER JOIN WMS_CUST_ORDER  " & _
                        " ON WMS_DELV_ORDER.IMP_CODE = WMS_CUST_ORDER.IMP_CODE AND  " & _
                        " WMS_DELV_ORDER.STORER_CODE = WMS_CUST_ORDER.STORER_CODE AND WMS_DELV_ORDER.DO_CO_CODE = WMS_CUST_ORDER.CO_CODE " & _
                        " Where wms_do_picklist_d.imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                        " and wms_do_picklist_d.do_code='" & gU.dbEncode(do_code) & "' and wms_do_picklist_d.storer_code='" & gU.dbEncode(storer_code) & "' " & _
                        " order by PLD_PL_LIST_NO, WMS_ITEM_WH.IW_PICK_LOC,dod.DOD_DISP_SEQ, pld_loc,pld_item_no,pld_pack_key,pld_pallet_no,pld_batch_no,pld_seq "

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then
                Dim iBCWriter As New ZXing.BarcodeWriter
                iBCWriter.Format = ZXing.BarcodeFormat.QR_CODE

                Dim oBmp, oBmp2, DOBmp As Bitmap
                Dim tempPath As String = gU.getConfig("PKLIST_BARCODE") & "\" & do_code

                Dim ItemCode As String = ""
                Dim LocCode As String = ""
                Dim PLListNo As String = ""

                Dim tempDT As DataTable

                If Not IO.Directory.Exists(tempPath) Then
                    IO.Directory.CreateDirectory(tempPath)
                End If

                'DOBmp = Code128Rendering.MakeBarcodeImage(do_code, 2, True)
                'DOBmp.Save(tempPath & "\" & do_code & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

                For i = 0 To nDataSource.Rows.Count - 1

                    ItemCode = nDataSource.Rows(i).Item("PLD_ITEM_NO").ToString.Trim + ";" + nDataSource.Rows(i).Item("PLD_PACK_KEY").ToString.Trim + ";" + nDataSource.Rows(i).Item("PLD_BATCH_NO").ToString.Trim + ";" + nDataSource.Rows(i).Item("PLD_SERIAL_NO").ToString.Trim

                    'DOBmp = rptU.GetCode39(do_code)
                    'oBmp = rptU.GetCode39(ItemCode)
                    'oBmp = Code128Rendering.MakeBarcodeImage(ItemCode, 2, True)
                    oBmp = iBCWriter.Write(ItemCode)
                    'oBmp2 = rptU.GetCode39(nDataSource.Rows(i).Item("PLD_LOC").ToString.Trim)
                    'oBmp2 = Code128Rendering.MakeBarcodeImage(nDataSource.Rows(i).Item("PLD_LOC").ToString.Trim, 2, True)
                    'oBmp2 = iBCWriter.Write(gU.decodeNullOrEmpty(nDataSource.Rows(i).Item("PLD_LOC").ToString.Trim, " "))
                    oBmp.Save(tempPath & "\" & "OB" & do_code & "_" & nDataSource.Rows(i).Item("PLD_SEQ").ToString.Trim & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
                    'oBmp2.Save(tempPath & "\" & "LOC" & do_code & "_" & nDataSource.Rows(i).Item("PLD_SEQ").ToString.Trim & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
                Next

                sqlString = "select distinct ISNULL(PLD_PL_LIST_NO,1) as PLD_PL_LIST_NO  from WMS_DO_PICKLIST_D Where WMS_DO_PICKLIST_D.imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                            " AND WMS_DO_PICKLIST_D.do_code='" & gU.dbEncode(do_code) & "' and WMS_DO_PICKLIST_D.storer_code='" & gU.dbEncode(storer_code) & "' order by PLD_PL_LIST_NO "

                tempDT = gDB.getDataTable(sqlString)

                If tempDT.Rows.Count > 0 Then
                    For i = 0 To tempDT.Rows.Count - 1
                        PLListNo = do_code & "-" & tempDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim

                        DOBmp = Code128Rendering.MakeBarcodeImage(PLListNo, 2, True)
                        DOBmp.Save(tempPath & "\" & "PL" & do_code & "_" & tempDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
                    Next
                End If

                Dim do_rem As String = ""

                sqlString = "Select do_rem from wms_delv_order where " & _
                            "wms_delv_order.imp_code = '" & gU.dbEncode(imp_code) & "' and wms_delv_order.do_code='" & gU.dbEncode(do_code) & "' and wms_delv_order.storer_code='" & gU.dbEncode(storer_code) & "' "

                do_rem = gU.decodeNullOrEmpty(DB.getValueFromSQL(sqlString), "")

                Paras(0) = New ReportParameter("BarCodePath", "file:///" & tempPath & "\")
                Paras(1) = New ReportParameter("DO_CODE", do_code)
                Paras(2) = New ReportParameter("DO_REM", do_rem)

                Dim DO_EDI_WIT_NO As String = ""

                sqlString = "Select DO_EDI_WIT_NO from wms_delv_order where " & _
                            "wms_delv_order.imp_code = '" & gU.dbEncode(imp_code) & "' and wms_delv_order.do_code='" & gU.dbEncode(do_code) & "' and wms_delv_order.storer_code='" & gU.dbEncode(storer_code) & "' "

                DO_EDI_WIT_NO = gU.decodeNullOrEmpty(DB.getValueFromSQL(sqlString), "")

                Paras(3) = New ReportParameter("DO_EDI_WIT_NO", DO_EDI_WIT_NO)

                Dim DO_EDI_SIR_NO As String = ""

                sqlString = "Select DO_EDI_SIR_NO from wms_delv_order where " & _
                            "wms_delv_order.imp_code = '" & gU.dbEncode(imp_code) & "' and wms_delv_order.do_code='" & gU.dbEncode(do_code) & "' and wms_delv_order.storer_code='" & gU.dbEncode(storer_code) & "' "

                DO_EDI_SIR_NO = gU.decodeNullOrEmpty(DB.getValueFromSQL(sqlString), "")

                Paras(4) = New ReportParameter("DO_EDI_SIR_NO", DO_EDI_SIR_NO)

                reportSource(nDataSource, Paras)
            Else
                Response.Write("No Picking List has been Found.")
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            If formatStr = "V" Then
                ReportViewer1.LocalReport.ReportPath = "OUTBOUND\DO\PICK_LIST\picklist_v.rdlc"
                ReportViewer1.Width = 800
            Else
                ReportViewer1.LocalReport.ReportPath = "OUTBOUND\DO\PICK_LIST\picklist.rdlc"
                ReportViewer1.Width = 1100
            End If
            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))
            ReportViewer1.LocalReport.EnableExternalImages = True
            ReportViewer1.ShowPrintButton = True

            If Not IsNothing(paraarray) Then
                ReportViewer1.LocalReport.SetParameters(paraarray)
            End If

            Try

                ReportViewer1.LocalReport.Refresh()

                Dim formatName As String = "word"

                For Each extension As RenderingExtension In ReportViewer1.LocalReport.ListRenderingExtensions

                    If extension.Name.ToLower = formatName Then

                        Dim m_isVisible As System.Reflection.FieldInfo = extension.GetType.GetField("m_isVisible", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)

                        m_isVisible.SetValue(extension, False)

                        Exit For

                    End If

                Next



            Catch ex As OutOfMemoryException
                GC.Collect()
                ReportViewer1.LocalReport.Refresh()
            End Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
