Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.Reporting.WebForms

Partial Class cms_preview 
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private RptU As New ReportUtils
    Private ar As AccessRightUtils

    Private menu_code As String = ""
    Private NoAccess As Boolean = False
    Private rdlcPhyPath As String = ""
    Private schemaPhyPath As String = ""
    Private rpt_file_type As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Buffer = True
        Response.ExpiresAbsolute = Now
        Response.Expires = 0
        Response.CacheControl = "no-cache"

        If Not IsPostBack Then
            Dim DataSource As New ReportDataSource()

            Me.Form.Focus()

            Try
                menu_code = Request("menu_code")

                ar = New AccessRightUtils(menu_code, Session("usr_id"), Me)

                If menu_code Is Nothing Or menu_code = "" Then
                    ar.Force_PageEndCtrlClear(Me)
                    Exit Sub
                End If

                NoAccess = ar.hideForm(Me)

                If Not Session("rdlc_DataSource") Is Nothing Then
                    DataSource = Session("rdlc_DataSource")
                Else
                    Exit Sub
                End If

                If Not Session("rdlc_Path") Is Nothing Then
                    rdlcPhyPath = Session("rdlc_Path")
                Else
                    Exit Sub
                End If

                If Not Session("schema_Path") Is Nothing Then
                    schemaPhyPath = Session("schema_Path")
                Else
                    Exit Sub
                End If

                If Not Session("rpt_file_type") Is Nothing Then
                    rpt_file_type = Session("rpt_file_type")
                End If

                If Not DataSource Is Nothing Then
                    reportSource(DataSource, rdlcPhyPath)
                End If

                If rpt_file_type <> 0 Then
                    Select Case rpt_file_type
                        Case 1

                    End Select
                End If

                ' Call RptU.deleteTempFiles(rdlcPhyPath, schemaPhyPath)

            Catch ex As Exception
                Me.Controls.Clear()
                Response.Write(ex.Message)
            End Try

        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As ReportDataSource, ByVal rdlc As String, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = rdlc

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(sourcetbl)

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

    <System.Web.Services.WebMethod()> _
    Public Shared Function deleteTempFiles() As Boolean
        Try
            Dim rdlcPhyPath As String = HttpContext.Current.Session("rdlc_Path")
            Dim schemaPhyPath As String = HttpContext.Current.Session("schema_Path")

            If File.Exists(rdlcPhyPath) Then
                File.Delete(rdlcPhyPath)
            End If

            If File.Exists(schemaPhyPath) Then
                File.Delete(schemaPhyPath)
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try        
    End Function
End Class
