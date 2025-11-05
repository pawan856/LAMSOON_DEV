Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.Collections.Generic

Partial Class OUTBOUND_DO_PACK_LABEL_Default
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private rptU As New ReportUtils
    Private DB As New DBfunc
    Private m_currentPageIndex As Integer
    Private m_streams As IList(Of Stream)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim do_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""
        Dim dov_seq As String = ""

        Dim Paras(2) As ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            ViewState("do_code") = ""
            do_code = Request.Form("do_code")
            storer_code = Request.Form("storer_code")
            dov_seq = Request.Form("dov_seq")

            ViewState("do_code") = do_code
            ViewState("storer_code") = storer_code
            ViewState("dov_seq") = dov_seq

        End If

        storer_code = ViewState("storer_code")
        do_code = ViewState("do_code")
        dov_seq = ViewState("dov_seq")

        If Not IsPostBack Then
            sqlString = " SELECT WMS_DELV_ORDER.DO_FTRACK_NO as DO_TRACK_NO,WMS_DELV_ORDER.DO_REM, WMS_DELV_ORDER.DO_ADDR1, WMS_DELV_ORDER.DO_ADDR2, WMS_DELV_ORDER.DO_ADDR3, WMS_DELV_ORDER.DO_CUS_CONT, WMS_DELV_ORDER.DO_CUS_CONT_TEL, " & _
                        " WMS_DELV_ORDER.DO_EDI_WIT_NO, WMS_DELV_ORDER_VIDEO.DOV_WEIGHT, WMS_DELV_ORDER_VIDEO.DOV_BOX_NO " & _
                        " FROM WMS_DELV_ORDER INNER JOIN " & _
                        " WMS_DELV_ORDER_VIDEO ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_VIDEO.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_VIDEO.STORER_CODE AND " & _
                        " WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_VIDEO.DO_CODE " & _
                        " Where WMS_DELV_ORDER.imp_code = '" & imp_code & "' " & _
                        " and WMS_DELV_ORDER.do_code='" & gU.dbEncode(do_code) & "' and WMS_DELV_ORDER.storer_code='" & gU.dbEncode(storer_code) & "' " & _
                        " and WMS_DELV_ORDER_VIDEO.DOV_SEQ='" & gU.dbEncode(dov_seq) & "'"

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then

                Dim oBmp As Bitmap
                Dim iBCWriter As New ZXing.BarcodeWriter

                iBCWriter.Format = ZXing.BarcodeFormat.CODE_128
                iBCWriter.Options.PureBarcode = True

                Dim ImgByte() As Byte
                Dim ByteString As String

                oBmp = iBCWriter.Write(gU.decodeNullOrEmpty(nDataSource.Rows(0).Item("DO_TRACK_NO").ToString.Trim, "0"))

                Using IOStream As New MemoryStream
                    oBmp.Save(IOStream, Imaging.ImageFormat.Png)
                    ImgByte = IOStream.ToArray

                    ByteString = Convert.ToBase64String(ImgByte)
                    Paras(0) = New ReportParameter("TRACK_BARCODE", ByteString)

                End Using

                ViewState("dt") = nDataSource
                ViewState("Paras") = Paras

                Dim itmNameStr As String = ""
                Dim itmDT As DataTable

                sqlString = " SELECT distinct WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC " & _
                            " FROM WMS_DELV_ORDER INNER JOIN " & _
                            " WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND " & _
                            " WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE INNER JOIN " & _
                            " WMS_ITEM ON WMS_DELV_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_DELV_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                            " WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                            " Where WMS_DELV_ORDER.imp_code = '" & imp_code & "' " & _
                            " and WMS_DELV_ORDER.do_code='" & gU.dbEncode(do_code) & "' and WMS_DELV_ORDER.storer_code='" & gU.dbEncode(storer_code) & "'"

                itmDT = gDB.getDataTable(sqlString)
                If itmDT.Rows.Count > 0 Then
                    For i = 0 To itmDT.Rows.Count - 1
                        itmNameStr &= itmDT.Rows(i).Item("ITM_NAME").ToString.Trim
                        If i <> itmDT.Rows.Count - 1 Then
                            itmNameStr &= vbCrLf
                        End If
                    Next
                End If

                Paras(1) = New ReportParameter("ITEM_NAME", itmNameStr)

                iBCWriter.Format = ZXing.BarcodeFormat.QR_CODE
                iBCWriter.Options.PureBarcode = True

                oBmp = iBCWriter.Write("http://210.177.23.43/AE_PROD/eoffice_Login.aspx")

                Using IOStream As New MemoryStream
                    oBmp.Save(IOStream, Imaging.ImageFormat.Png)
                    ImgByte = IOStream.ToArray

                    ByteString = Convert.ToBase64String(ImgByte)
                    Paras(2) = New ReportParameter("QRBarcode", ByteString)

                End Using


                reportSource(nDataSource, Paras)

            Else
                Response.Write("No Packing Label Found.")
            End If
        End If


    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "OUTBOUND\DO\PACK_LABEL\PACKLABEL.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))
            ReportViewer1.LocalReport.EnableExternalImages = True

            If Not IsNothing(paraarray) Then
                ReportViewer1.LocalReport.SetParameters(paraarray)
                ReportViewer1.ZoomMode = ZoomMode.PageWidth
            End If

            Try
                ReportViewer1.LocalReport.Refresh()

                'Export(ReportViewer1.LocalReport)
                'Print()


            Catch ex As OutOfMemoryException
                GC.Collect()
                ReportViewer1.LocalReport.Refresh()
            End Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Print()
        If m_streams Is Nothing OrElse m_streams.Count = 0 Then
            Throw New Exception("Error: no stream to print.")
        End If
        Dim printDoc As New PrintDocument()
        'printDoc.PrinterSettings.PrinterName = "HP LaserJet Professional M1212nf MFP_1"
        If Not printDoc.PrinterSettings.IsValid Then
            Throw New Exception("Error: cannot find the default printer.")
        Else
            AddHandler printDoc.PrintPage, AddressOf PrintPage
            m_currentPageIndex = 0
            printDoc.Print()
        End If
    End Sub

    Private Sub PrintPage(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        Dim pageImage As New Metafile(m_streams(m_currentPageIndex))

        ' Adjust rectangular area with printer margins.
        Dim adjustedRect As New Rectangle(ev.PageBounds.Left - CInt(ev.PageSettings.HardMarginX), _
                                          ev.PageBounds.Top - CInt(ev.PageSettings.HardMarginY), _
                                          ev.PageBounds.Width, _
                                          ev.PageBounds.Height)

        ' Draw a white background for the report
        ev.Graphics.FillRectangle(Brushes.White, adjustedRect)

        ' Draw the report content
        ev.Graphics.DrawImage(pageImage, adjustedRect)

        ' Prepare for the next page. Make sure we haven't hit the end.
        m_currentPageIndex += 1
        ev.HasMorePages = (m_currentPageIndex < m_streams.Count)
    End Sub

    Private Sub Export(ByVal report As LocalReport)
        Dim deviceInfo As String = "<DeviceInfo>" & _
            "<OutputFormat>EMF</OutputFormat>" & _
            "<PageWidth>8.5in</PageWidth>" & _
            "<PageHeight>11in</PageHeight>" & _
            "<MarginTop>0.25in</MarginTop>" & _
            "<MarginLeft>0.25in</MarginLeft>" & _
            "<MarginRight>0.25in</MarginRight>" & _
            "<MarginBottom>0.25in</MarginBottom>" & _
            "</DeviceInfo>"
        Dim warnings As Warning()
        m_streams = New List(Of Stream)()
        report.Render("Image", deviceInfo, AddressOf CreateStream, warnings)
        For Each stream As Stream In m_streams
            stream.Position = 0
        Next
    End Sub

    Private Function CreateStream(ByVal name As String, ByVal fileNameExtension As String, ByVal encoding As Encoding, ByVal mimeType As String, ByVal willSeek As Boolean) As Stream
        Dim stream As Stream = New MemoryStream()
        m_streams.Add(stream)
        Return stream
    End Function


    Private Function LoadSalesData() As DataTable
        ' Create a new DataSet and read sales data file 
        ' data.xml into the first DataTable.
        Dim dataSet As New DataSet()
        dataSet.ReadXml("..\..\data.xml")
        Return dataSet.Tables(0)
    End Function

End Class
