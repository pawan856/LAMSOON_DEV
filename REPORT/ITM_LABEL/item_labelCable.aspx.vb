Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms


Partial Class REPORT_ITM_LABEL_item_label
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private wmsFun As New WMSFunc
    Private rptU As New ReportUtils
    Private cU As New CommonUtils

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils("MAST_IM", Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        ar.hideForm(Me)

        If Not IsPostBack Then

            ViewState("lblType") = ""
            ViewState("itemStr") = ""
            ViewState("GR_CODE") = ""
            ViewState("RO_CODE") = ""
            ViewState("RT_CODE") = ""
            ViewState("DOC_NO") = ""
            ViewState("STORER_CODE") = ""

            ViewState("lblType") = Server.UrlDecode(Request("lblType"))
            ViewState("itemStr") = Server.UrlDecode(Request("itemStr"))
            ViewState("GR_CODE") = Server.UrlDecode(Request("GR_CODE"))
            ViewState("RO_CODE") = Server.UrlDecode(Request("RO_CODE"))
            ViewState("RT_CODE") = Server.UrlDecode(Request("RT_CODE"))
            ViewState("DOC_NO") = Server.UrlDecode(Request("DOC_NO"))
            ViewState("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))

            REM testing
            'ViewState("itemStr") = "000111"
            'ViewState("lblType") = "GR"
            'ViewState("GR_CODE") = "000156"
            REM testing 


            If ViewState("itemStr") = "" Then
                Response.Write("Please Select Item to print label!")
                Response.End()

            Else
                Call BindGV()
            End If

        End If

    End Sub

    Private Sub BindGV()
        Dim sqlString As String = ""
        Dim nDataSource As New DataTable

        Dim Paras(2) As ReportParameter
        Dim SYS_TEMP_FOLDER As String = gU.getConfig("SYSP_TEMP_DIR")

        Dim itemCode() As String
        Dim itemKey() As String


        Dim itemStr As String = ViewState("itemStr")
        Dim storer_code As String = ViewState("STORER_CODE")
        Dim gr_code As String = ViewState("GR_CODE")
        Dim ro_code As String = ViewState("RO_CODE")
        Dim rt_code As String = ViewState("RT_CODE")

        Dim doc_no As String = ViewState("DOC_NO")

        Dim itemSQL As String = ""
        Select Case ViewState("lblType")
            Case "GR"
                itemCode = Split(itemStr, "||")
                itemSQL = ""

                For x = 0 To itemCode.Length - 1
                    itemSQL &= "'" & itemCode(x) & "',"
                Next

                itemSQL = " AND wms_goodsrcv_d.GRD_SEQ in (" & Left(itemSQL, Len(itemSQL) - 1) & ")"

                sqlString = " SELECT WMS_ITEM.ITM_CODE,  WMS_ITEM.PACK_KEY,  WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_SKU_NO," & _
                            " wms_goodsrcv_d.GRD_BATCH_NO as BATCH_NO,  wms_goodsrcv_d.GRD_PALLET_NO as PALLET_NO, " & _
                            " WMS_GOODSRCV_D_S.GRS_SERIAL_NO AS SERIAL_NO, WMS_GOODSRCV_D.GRD_QTY2 as ITM_QTY2, WMS_GOODSRCV.GR_EDI_PO_NO as EDI_NO, convert(varchar,WMS_GOODSRCV.GR_DATE,103) as rcv_date " &
                            " FROM wms_goodsrcv_d INNER JOIN WMS_ITEM " & _
                            " ON WMS_ITEM.IMP_CODE = wms_goodsrcv_d.IMP_CODE " & _
                            " AND WMS_ITEM.STORER_CODE = wms_goodsrcv_d.STORER_CODE " & _
                            " AND WMS_ITEM.ITM_CODE = wms_goodsrcv_d.GRD_ITM_CODE " & _
                            " AND WMS_ITEM.PACK_KEY = wms_goodsrcv_d.GRD_PACK_KEY " & _
                            " LEFT OUTER JOIN WMS_GOODSRCV_D_S ON WMS_GOODSRCV_D.IMP_CODE = WMS_GOODSRCV_D_S.IMP_CODE AND " & _
                            " WMS_GOODSRCV_D.STORER_CODE = WMS_GOODSRCV_D_S.STORER_CODE AND WMS_GOODSRCV_D.GR_CODE = WMS_GOODSRCV_D_S.GR_CODE AND " & _
                            " WMS_GOODSRCV_D.GRD_ITM_CODE = WMS_GOODSRCV_D_S.GRS_ITM_CODE AND " & _
                            " WMS_GOODSRCV_D.GRD_PACK_KEY = WMS_GOODSRCV_D_S.GRS_PACK_KEY AND " & _
                            " isnull(WMS_GOODSRCV_D.GRD_PALLET_NO,'000') = isnull(WMS_GOODSRCV_D_S.GRS_PALLET_NO,'000') AND " & _
                            " isnull(WMS_GOODSRCV_D.GRD_BATCH_NO, '') =isnull(WMS_GOODSRCV_D_S.GRS_BATCH_NO, '') " & _
                            " INNER JOIN WMS_GOODSRCV ON WMS_GOODSRCV_D.IMP_CODE = WMS_GOODSRCV.IMP_CODE AND " & _
                            " WMS_GOODSRCV_D.STORER_CODE = WMS_GOODSRCV.STORER_CODE AND WMS_GOODSRCV_D.GR_CODE = WMS_GOODSRCV.GR_CODE " & _
                            " WHERE WMS_ITEM.IMP_CODE = '" & Session("IMP_CODE") & "' AND wms_item.storer_code = '" & storer_code & "' " & _
                            " AND wms_goodsrcv_d.GR_CODE='" & gU.dbEncode(gr_code) & "' " & itemSQL

            Case "ITEM"

                itemKey = Split(itemStr, "||")

                sqlString = " SELECT ITM_CODE, PACK_KEY, ITM_NAME, ITM_DESC, '' AS PALLET_NO, '' AS BATCH_NO, ITM_SKU_NO, '' AS SERIAL_NO, '' AS EDI_NO, ITM_QTY2,'' as rcv_date " & _
                            " FROM WMS_ITEM " & _
                            " Where imp_code='" & Session("IMP_CODE") & "' and storer_code='" & storer_code & "'" & _
                            " and itm_code='" & itemKey(0) & "' and pack_key='" & itemKey(1) & "'"
            Case "RO"

                itemCode = Split(itemStr, "||")

                itemSQL = ""

                For x = 0 To itemCode.Length - 1

                    itemSQL &= "'" & itemCode(x) & "',"

                Next

                itemSQL = " AND wms_replenish_d.ROD_SEQ in (" & Left(itemSQL, Len(itemSQL) - 1) & ")"

                sqlString = " SELECT WMS_ITEM.ITM_CODE,  WMS_ITEM.PACK_KEY,  WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_SKU_NO, '' as SERIAL_NO," & _
                            " WMS_REPLENISH_D.ROD_PALLET_NO as PALLET_NO, WMS_REPLENISH_D.ROD_BATCH_NO as BATCH_NO, WMS_REPLENISH.RO_EDI_PO_NO as EDI_NO, '' as rcv_date, WMS_REPLENISH_D.ROD_QTY2 as ITM_QTY2 " &
                            " FROM wms_replenish_d " & _
                            " INNER JOIN WMS_ITEM " & _
                            " ON WMS_ITEM.IMP_CODE     = wms_replenish_d.IMP_CODE " & _
                            " AND WMS_ITEM.STORER_CODE = wms_replenish_d.STORER_CODE " & _
                            " AND WMS_ITEM.ITM_CODE    = wms_replenish_d.ROD_ITM_CODE " & _
                            " AND WMS_ITEM.PACK_KEY    = wms_replenish_d.ROD_PACK_KEY " & _
                            " INNER JOIN WMS_REPLENISH ON WMS_REPLENISH_D.IMP_CODE = WMS_REPLENISH.IMP_CODE AND " & _
                            " WMS_REPLENISH_D.STORER_CODE = WMS_REPLENISH.STORER_CODE AND WMS_REPLENISH_D.RO_CODE = WMS_REPLENISH.RO_CODE " & _
                            " WHERE WMS_ITEM.IMP_CODE = '" & Session("IMP_CODE") & "' AND wms_item.storer_code = '" & storer_code & "' " & _
                            " AND wms_replenish_d.RO_CODE='" & gU.dbEncode(ro_code) & "' " & itemSQL


            Case "SR"
                itemCode = Split(itemStr, "||")
                itemSQL = ""

                For x = 0 To itemCode.Length - 1

                    itemSQL &= "'" & itemCode(x) & "',"

                Next
                itemSQL = " AND WMS_STOCK_RETURN_D.RTD_SEQ in (" & Left(itemSQL, Len(itemSQL) - 1) & ")"

                sqlString = " SELECT WMS_ITEM.ITM_CODE,  WMS_ITEM.PACK_KEY,  WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_SKU_NO," & _
                            " WMS_STOCK_RETURN_D.RTD_BATCH_NO as BATCH_NO,  WMS_STOCK_RETURN_D.RTD_PALLET_NO as PALLET_NO, " & _
                            " WMS_STOCK_RETURN_D.RTD_SERIAL_NO AS SERIAL_NO, WMS_STOCK_RETURN.RT_REF_NO2 as EDI_NO, convert(varchar, WMS_STOCK_RETURN.RT_DATE,103) as rcv_date, WMS_STOCK_RETURN_D.RTD_QTY2 as ITM_QTY2 " &
                            " FROM WMS_STOCK_RETURN_D " & _
                            " INNER JOIN WMS_ITEM " & _
                            " ON WMS_ITEM.IMP_CODE     = WMS_STOCK_RETURN_D.IMP_CODE " & _
                            " AND WMS_ITEM.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE " & _
                            " AND WMS_ITEM.ITM_CODE    = WMS_STOCK_RETURN_D.RTD_ITM_CODE " & _
                            " AND WMS_ITEM.PACK_KEY    = WMS_STOCK_RETURN_D.RTD_PACK_KEY " & _
                            " INNER JOIN WMS_STOCK_RETURN ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND " & _
                            " WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE " & _
                            " WHERE WMS_ITEM.IMP_CODE = '" & Session("IMP_CODE") & "' AND wms_item.storer_code = '" & storer_code & "' " & _
                            " AND WMS_STOCK_RETURN_D.RT_CODE='" & gU.dbEncode(rt_code) & "' " & itemSQL
            Case "CO"

                itemCode = Split(itemStr, "||")
                itemSQL = ""

                For x = 0 To itemCode.Length - 1

                    itemSQL &= "'" & itemCode(x) & "',"

                Next
                itemSQL = " AND WMS_CUST_ORDER_D.COD_SEQ in (" & Left(itemSQL, Len(itemSQL) - 1) & ")"

                sqlString = " SELECT WMS_ITEM.ITM_CODE,  WMS_ITEM.PACK_KEY,  WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_SKU_NO," & _
                            " WMS_CUST_ORDER_D.COD_BATCH_NO as BATCH_NO,  WMS_CUST_ORDER_D.COD_PALLET_NO as PALLET_NO, '' AS EDI_NO, WMS_CUST_ORDER_D.COD_QTY2 as ITM_QTY2, '' as rcv_date, " & _
                            " '' AS SERIAL_NO " & _
                            " FROM WMS_CUST_ORDER_D INNER JOIN " & _
                            " WMS_ITEM ON WMS_CUST_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_CUST_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                            " WMS_CUST_ORDER_D.COD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_CUST_ORDER_D.COD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                            " AND WMS_ITEM.IMP_CODE = '" & Session("IMP_CODE") & "' AND wms_item.storer_code = '" & storer_code & "' " & _
                            " AND WMS_CUST_ORDER_D.CO_CODE='" & gU.dbEncode(doc_no) & "' " & itemSQL
            Case "DO"
                itemCode = Split(itemStr, "||")
                itemSQL = ""

                For x = 0 To itemCode.Length - 1

                    itemSQL &= "'" & itemCode(x) & "',"

                Next
                itemSQL = " AND WMS_DO_PICKLIST_D.PLD_SEQ in (" & Left(itemSQL, Len(itemSQL) - 1) & ")"

                sqlString = " SELECT WMS_ITEM.ITM_CODE,  WMS_ITEM.PACK_KEY,  WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_SKU_NO," & _
                            " WMS_DO_PICKLIST_D.PLD_BATCH_NO as BATCH_NO,  WMS_DO_PICKLIST_D.PLD_PALLET_NO as PALLET_NO, " & _
                            " WMS_DO_PICKLIST_D.PLD_SERIAL_NO AS SERIAL_NO, WMS_DO_PICKLIST_D.PLD_QTY2 as ITM_QTY2, WMS_GOODSRCV.GR_EDI_PO_NO as EDI_NO, convert(varchar,WMS_GOODSRCV.GR_DATE,103) as rcv_date  " &
                            " FROM WMS_DO_PICKLIST_D LEFT OUTER JOIN " & _
                            " WMS_GOODSRCV_D_S ON WMS_GOODSRCV_D_S.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND  " & _
                            " WMS_GOODSRCV_D_S.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND  " & _
                            " WMS_GOODSRCV_D_S.GRS_SERIAL_NO = WMS_DO_PICKLIST_D.PLD_SERIAL_NO AND  " & _
                            " WMS_GOODSRCV_D_S.GRS_ITM_CODE = WMS_DO_PICKLIST_D.PLD_ITEM_NO AND  " & _
                            " WMS_GOODSRCV_D_S.GRS_PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY AND " & _
                            " WMS_GOODSRCV_D_S.GRS_BATCH_NO = WMS_DO_PICKLIST_D.PLD_BATCH_NO AND  " & _
                            " WMS_GOODSRCV_D_S.GRS_PALLET_NO = WMS_DO_PICKLIST_D.PLD_PALLET_NO LEFT OUTER JOIN " & _
                            " WMS_GOODSRCV ON WMS_GOODSRCV_D_S.IMP_CODE = WMS_GOODSRCV.IMP_CODE AND  " & _
                            " WMS_GOODSRCV_D_S.STORER_CODE = WMS_GOODSRCV.STORER_CODE AND WMS_GOODSRCV_D_S.GR_CODE = WMS_GOODSRCV.GR_CODE " & _
                            " INNER JOIN WMS_ITEM ON WMS_DO_PICKLIST_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_DO_PICKLIST_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                            " WMS_DO_PICKLIST_D.PLD_ITEM_NO = WMS_ITEM.ITM_CODE AND WMS_DO_PICKLIST_D.PLD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                            " WHERE WMS_ITEM.IMP_CODE = '" & Session("IMP_CODE") & "' AND wms_item.storer_code = '" & storer_code & "' " & _
                            " AND WMS_DO_PICKLIST_D.DO_CODE='" & gU.dbEncode(doc_no) & "' " & itemSQL

        End Select

        nDataSource = gDB.getDataTable(sqlString)
        Dim lblType As String = gU.decodeNullOrEmpty(ViewState("lblType"), "")

        If lblType = "GR" Then
            Paras(0) = New ReportParameter("GR_CODE", gr_code)
        ElseIf lblType = "RO" Then
            Paras(0) = New ReportParameter("GR_CODE", ro_code)
        Else
            Paras(0) = New ReportParameter("GR_CODE", "")
        End If
        Paras(1) = New ReportParameter("SYS_TEMP_FOLDER", SYS_TEMP_FOLDER)
        Paras(2) = New ReportParameter("lblType", lblType)

        Dim barcodeCode As String = ""
        Dim oBMP As Drawing.Bitmap
        Dim iBCWriter As New ZXing.BarcodeWriter

        iBCWriter.Format = ZXing.BarcodeFormat.QR_CODE
        iBCWriter.Options.Margin = 0

        Dim WarningMsg As String = ""
        Dim nDataRow As DataRow
        If nDataSource.Rows.Count > 0 Then

            For i = 0 To nDataSource.Rows.Count - 1

                barcodeCode = nDataSource.Rows(i).Item("ITM_SKU_NO").ToString.Trim + ";" + nDataSource.Rows(i).Item("PACK_KEY").ToString.Trim + ";" + nDataSource.Rows(i).Item("BATCH_NO").ToString.Trim + ";" + nDataSource.Rows(i).Item("SERIAL_NO").ToString.Trim

                'If nDataSource.Rows(i).Item("BATCH_NO").ToString.Trim <> "" Then barcodeCode &= "+" + nDataSource.Rows(i).Item("BATCH_NO").ToString.Trim

                'oBMP = rptU.GetCode39(barcodeCode.Trim)
                'oBMP = Code128Rendering.MakeBarcodeImage(barcodeCode.Trim, 2, True)

                oBMP = iBCWriter.Write(barcodeCode)

                oBMP.Save(SYS_TEMP_FOLDER & "/Barcode" & i + 1 & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

                nDataRow = nDataSource.Rows(i)

                'WarningMsg = returnWarningText(nDataRow)
                'nDataSource.Rows(i).Item("warningStr") = WarningMsg

                'Dim cheClass As String = nDataSource.Rows(i).Item("ITM_CHE_CLASS").ToString.Trim
                'Dim classArr() As String = gU.listToArray(cheClass)

                'Dim arrUpper As Integer

                'If classArr.Length > 4 Then arrUpper = 4 Else arrUpper = classArr.Length

                'For x = 0 To arrUpper - 1
                '    nDataSource.Rows(i).Item("IMG_CHE_CLASS" & x + 1) = returnCheIMG(classArr(x))
                'Next


            Next

            nDataSource.AcceptChanges()
            reportSource(nDataSource, Paras)

        Else

            Response.Write("No Item Label is found.")

        End If




    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "REPORT\itm_label\itm_labelCable.rdlc"
            ReportViewer1.LocalReport.EnableExternalImages = True
            ReportViewer1.LocalReport.DataSources.Clear()

            'ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("rptDataset", sourcetbl))
            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))

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

    Private Function returnWarningText(ByVal iRow As DataRow) As String
        Dim returnStr As String = ""

        If iRow.Item("ITM_DG_YN").ToString.Trim = "Y" Then
            returnStr &= "Dangerous Good - Category " & iRow.Item("ITM_DG_CAT").ToString.Trim
        End If

        If iRow.Item("ITM_REQ_STORE_HUM_YN").ToString.Trim = "Y" Then
            If returnStr <> "" Then returnStr &= vbCrLf
            returnStr &= "Required to store at relative humidity " & iRow.Item("ITM_REQ_STORE_HUM_DESC").ToString.Trim
        End If

        If iRow.Item("ITM_REQ_STORE_AIRCON_YN").ToString.Trim = "Y" Then
            If returnStr <> "" Then returnStr &= vbCrLf
            returnStr &= "Store at Air-Con Warehouse"
        End If

        If iRow.Item("ITM_REQ_STORE_COLD_YN").ToString.Trim = "Y" Then
            If returnStr <> "" Then returnStr &= vbCrLf
            returnStr &= "Store at Cold Store"
        End If

        If iRow.Item("ITM_REQ_MSDS_YN").ToString.Trim = "Y" Then
            If returnStr <> "" Then returnStr &= vbCrLf
            returnStr &= "MSDS is required"
        End If

        Select Case iRow.Item("ITM_RESTRICTED_ITEM").ToString.Trim
            Case "1"
                If returnStr <> "" Then returnStr &= vbCrLf
                returnStr &= "General Restricted Item"

            Case "2"
                If returnStr <> "" Then returnStr &= vbCrLf
                returnStr &= "T&D Restricted Item"

            Case "3"
                If returnStr <> "" Then returnStr &= vbCrLf
                returnStr &= "GEN Restricted Item"

        End Select

        Return returnStr
    End Function

    Private Function returnCheIMG(ByVal cheClass As String) As String
        Dim returnName As String = ""

        Select Case cheClass
            Case "1F"
                returnName = "ID1F"
            Case "2E"
                returnName = "ID2E"
            Case "3O"
                returnName = "ID3O"
            Case "4H"
                returnName = "ID4H"
            Case "5T"
                returnName = "ID5T"
            Case "6C"
                returnName = "ID6C"
            Case "7I"
                returnName = "ID7I"
            Case "8C"
                returnName = "ID8C"
        End Select

        Return returnName
    End Function
End Class
