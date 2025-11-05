Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.Reporting.WebForms

Partial Class REPORT_ob_co_summary
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gFunc As New DBfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils("RPT_COS", Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim storer_code As String = ""
        Dim imp_code As String = ""

        Dim vStorer As String = ""

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        If Not IsPostBack Then


            sqlString = Request.Form("GENERIC_REQUEST_COMPLETE_SQL")

            nDataSource = gDB.getDataTable(sqlString)

            vStorer = Request.Form("WMS_CUST_ORDER_STORER_CODE")

            If vStorer = "" Then
                vStorer = "All"
            Else
                Dim STORERSQL = "SELECT STO_NAME FROM WMS_STORER WHERE STORER_CODE = '" & vStorer & "'"

                vStorer = gFunc.getValueFromSQL(STORERSQL)
            End If

            If nDataSource.Rows.Count > 0 Then
                Dim paras(0) As ReportParameter

                paras(0) = New ReportParameter("storer", vStorer)

                reportSource(nDataSource, paras)
            Else
                Dim strScript As String = "alert(""No CO Summary is Found."");window.open('','_self');window.close();"

                If Not Me.ClientScript Is Nothing Then

                    If Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType(), "") Then
                        Me.ClientScript.RegisterStartupScript(Me.GetType(), "", strScript, True)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "REPORT\ob_co_summary.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("rptDataSet_co_rpt", sourcetbl))

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
End Class
