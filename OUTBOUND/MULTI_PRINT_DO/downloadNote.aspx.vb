Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms

Partial Class OUTBOUND_MULTI_PRINT_DO_downloadNote
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private rptU As New ReportUtils
    Private DB As New DBfunc

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 1800

        If Not IsPostBack Then
            Dim nDataSource As DataTable

            nDatasource = Session("nDataSource")
            If nDataSource IsNot Nothing AndAlso nDataSource.Rows.Count > 0 Then
                downRPT(nDataSource, Nothing)
            End If

        End If
    End Sub


    Public Sub downRPT(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Dim warn As Warning() = Nothing
        Dim streamids As String()
        Dim mimeType As String = String.Empty
        Dim encoding As String = String.Empty
        Dim extension As String = String.Empty
        Dim byteviewer As Byte()

        Dim deviceInfo As String = _
          "<DeviceInfo>" + _
          "  <OutputFormat>PDF</OutputFormat>" + _
          "  <PageWidth>21cm</PageWidth>" + _
          "  <PageHeight>29.7cm</PageHeight>" + _
          "  <MarginTop>0.5cm</MarginTop>" + _
          "  <MarginLeft>0.5cm</MarginLeft>" + _
          "  <MarginRight>0.2cm</MarginRight>" + _
          "  <MarginBottom>0.5cm</MarginBottom>" + _
          "</DeviceInfo>"


        Dim rptViewer As New ReportViewer

        rptViewer.LocalReport.ReportPath = "OUTBOUND\MULTI_PRINT_DO\MDeliNote.rdlc"
        rptViewer.LocalReport.DataSources.Clear()
        rptViewer.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))
        rptViewer.LocalReport.EnableExternalImages = True

        If Not IsNothing(paraarray) Then
            rptViewer.LocalReport.SetParameters(paraarray)
        End If

        Try

            byteviewer = rptViewer.LocalReport.Render("PDF", deviceInfo, mimeType, encoding, extension, streamids, warn)
            Response.Buffer = True
            Response.Clear()
            Response.ContentType = mimeType
            Response.AddHeader("content-disposition", "attachment; filename=DeliNote" & Now.Date.ToString("ddMMyyyy") & ".pdf")
            Response.BinaryWrite(byteviewer)
            Response.Flush()

        Catch ex As Exception
            Dim str As String = ex.Message
        End Try

    End Sub
End Class
