Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms

Partial Class INBOUND_GR_PALIST_PRINT
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
        Dim gr_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""
        Dim tempDT As DataTable
        Dim Paras(2) As ReportParameter

        Dim iBCWriter As New ZXing.BarcodeWriter

        iBCWriter.Format = ZXing.BarcodeFormat.QR_CODE
        'iBCWriter.Options.Margin = 0

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        ViewState("gr_code") = ""

        gr_code = Request.Form("gr_code")
        storer_code = Request.Form("storer_code")
        If Not IsPostBack Then
            ViewState("gr_code") = gr_code

            sqlString = " SELECT WMS_GOODSRCV_PA.GR_CODE, WMS_ITEM.ITM_NAME, WMS_GOODSRCV_PA.GRA_PA_LIST_NO, WMS_GOODSRCV_PA.GRA_LOC, GRA_ITM_CODE,GRA_PACK_KEY,GRA_PALLET_NO," & _
                        " WMS_GOODSRCV_PA.GRA_SUG_QTY as GRA_PA_QTY, WMS_GOODSRCV_PA.GRA_BATCH_NO, WMS_ITEM.ITM_PREF_LOC, WMS_GOODSRCV_PA.GRA_SEQ, " & _
                        " WMS_GOODSRCV_PA.GRA_DISP_SEQ, WMS_GOODSRCV_PA.GRA_ITM_CODE,wms_item.itm_sku_no, WMS_GOODSRCV_D_S.GRS_SERIAL_NO,WMS_ITEM_WH.IW_PREF_LOC1, A.BN_CSMS_CODE as CSMS_IW_PREF_LOC1, B.BN_CSMS_CODE as CSMS_GRA_LOC " & _
                        " FROM WMS_GOODSRCV_PA LEFT OUTER JOIN " & _
                        " WMS_ITEM ON WMS_GOODSRCV_PA.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_GOODSRCV_PA.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                        " WMS_GOODSRCV_PA.GRA_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_GOODSRCV_PA.GRA_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                        " left outer JOIN WMS_GOODSRCV_D_S ON " & _
                        " WMS_GOODSRCV_PA.IMP_CODE = WMS_GOODSRCV_D_S.IMP_CODE AND  " & _
                        " WMS_GOODSRCV_PA.STORER_CODE = WMS_GOODSRCV_D_S.STORER_CODE AND WMS_GOODSRCV_PA.GR_CODE = WMS_GOODSRCV_D_S.GR_CODE AND  " & _
                        " WMS_GOODSRCV_PA.GRA_ITM_CODE = WMS_GOODSRCV_D_S.GRS_ITM_CODE AND  " & _
                        " WMS_GOODSRCV_PA.GRA_PACK_KEY = WMS_GOODSRCV_D_S.GRS_PACK_KEY AND WMS_GOODSRCV_PA.GRA_LOC = WMS_GOODSRCV_D_S.GRS_LOC " & _
                        " LEFT OUTER JOIN WMS_ITEM_WH ON WMS_GOODSRCV_PA.IMP_CODE = WMS_ITEM_WH.IMP_CODE AND " & _
                        " WMS_GOODSRCV_PA.STORER_CODE = WMS_ITEM_WH.STORER_CODE AND WMS_GOODSRCV_PA.GRA_ITM_CODE = WMS_ITEM_WH.ITM_CODE AND  " & _
                        " WMS_GOODSRCV_PA.GRA_PACK_KEY = WMS_ITEM_WH.PACK_KEY AND WMS_GOODSRCV_PA.GRA_WH = WMS_ITEM_WH.WH_CODE " & _
                        " LEFT OUTER JOIN WMS_WH_BIN A ON A.LOC_KEY = WMS_ITEM_WH.IW_PREF_LOC1 " & _
                        " LEFT OUTER JOIN WMS_WH_BIN B ON B.LOC_KEY = WMS_GOODSRCV_PA.GRA_LOC " & _
                        " Where WMS_GOODSRCV_PA.imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                        " AND WMS_GOODSRCV_PA.gr_code='" & gU.dbEncode(gr_code) & "' and WMS_GOODSRCV_PA.storer_code='" & gU.dbEncode(storer_code) & "' " & _
                        " ORDER BY WMS_GOODSRCV_PA.GRA_PA_LIST_NO, WMS_GOODSRCV_PA.GRA_DISP_SEQ, WMS_GOODSRCV_PA.GRA_SEQ, WMS_GOODSRCV_PA.GRA_ITM_CODE "

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then

                Dim oBmp, oBmp2 As Bitmap
                Dim tempPath As String = gU.getConfig("PALIST_BARCODE") & "\" & gr_code

                Dim ItemCode As String = ""
                Dim LocCode As String = ""
                Dim PAListNo As String = ""

                If Not IO.Directory.Exists(tempPath & "\" & gr_code) Then
                    IO.Directory.CreateDirectory(tempPath)
                End If

                For i = 0 To nDataSource.Rows.Count - 1

                    ItemCode = nDataSource.Rows(i).Item("ITM_SKU_NO").ToString.Trim + ";" + nDataSource.Rows(i).Item("GRA_PACK_KEY").ToString.Trim + ";" + nDataSource.Rows(i).Item("GRA_BATCH_NO").ToString.Trim + ";" + nDataSource.Rows(i).Item("GRS_SERIAL_NO").ToString.Trim

                    'oBmp = Code128Rendering.MakeBarcodeImage(ItemCode, 2, True)
                    oBmp = iBCWriter.Write(ItemCode)
                    oBmp.Save(tempPath & "\" & "IB" & gr_code & "_" & nDataSource.Rows(i).Item("GRA_SEQ").ToString.Trim & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
                    'oBmp2 = Code128Rendering.MakeBarcodeImage(nDataSource.Rows(i).Item("PLD_LOC").ToString.Trim, 2, True)
                    'oBmp2.Save(tempPath & "\" & "LOC" & do_code & "_" & nDataSource.Rows(i).Item("PLD_SEQ").ToString.Trim & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
                Next

                sqlString = "select distinct GRA_PA_LIST_NO  from WMS_GOODSRCV_PA Where WMS_GOODSRCV_PA.imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                            " AND WMS_GOODSRCV_PA.gr_code='" & gU.dbEncode(gr_code) & "' and WMS_GOODSRCV_PA.storer_code='" & gU.dbEncode(storer_code) & "' order by GRA_PA_LIST_NO "

                tempDT = gDB.getDataTable(sqlString)

                If tempDT.Rows.Count > 0 Then
                    For i = 0 To tempDT.Rows.Count - 1

                        PAListNo = gr_code & "-" & tempDT.Rows(i).Item("GRA_PA_LIST_NO").ToString.Trim

                        oBmp2 = Code128Rendering.MakeBarcodeImage(PAListNo, 2, True)
                        oBmp2.Save(tempPath & "\" & "PA" & gr_code & "_" & tempDT.Rows(i).Item("GRA_PA_LIST_NO").ToString.Trim & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

                    Next
                End If


                Paras(0) = New ReportParameter("BarCodePath", "file:///" & tempPath & "\")
                Paras(1) = New ReportParameter("GR_CODE", gr_code)

                Dim GR_EDI_PO_NO As String = ""

                sqlString = "Select GR_EDI_PO_NO from WMS_GOODSRCV where " & _
                            "WMS_GOODSRCV.imp_code = '" & gU.dbEncode(imp_code) & "' and WMS_GOODSRCV.GR_CODE='" & gU.dbEncode(gr_code) & "' and WMS_GOODSRCV.storer_code='" & gU.dbEncode(storer_code) & "' "

                GR_EDI_PO_NO = gU.decodeNullOrEmpty(DB.getValueFromSQL(sqlString), "")

                Paras(2) = New ReportParameter("GR_EDI_PO_NO", GR_EDI_PO_NO)

                reportSource(nDataSource, Paras)
            Else
                Response.Write("No Put Away has been Found.")
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            If formatStr = "V" Then
                ReportViewer1.LocalReport.ReportPath = "INBOUND\GR\GR_LIST\palist_v.rdlc"
                ReportViewer1.Width = 800
            Else
                ReportViewer1.LocalReport.ReportPath = "INBOUND\GR\GR_LIST\palist.rdlc"
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
