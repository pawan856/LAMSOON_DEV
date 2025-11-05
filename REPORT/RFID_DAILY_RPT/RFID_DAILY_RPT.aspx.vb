Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms

Partial Class RFID_DAILY_RPT
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private rptU As New ReportUtils
    Const formatStr As String = "V"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("RFID_DAILY_RPT", Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim whereSQL As String = ""

        If Not IsPostBack Then
            whereSQL = Session("GENERIC_SESSION_SRCH_WHERE_SQL")
            If whereSQL <> "" Then whereSQL = " AND " & whereSQL

            sqlString = " SELECT WRST_CODE, WRAL_KEY, WRMV_BATCH_ID, WRMV_TAG_ID, cast(WRMV_BATCH_TIME as time(0)) as WRMV_BATCH_TIME, WRAL_IMP_CODE, WRAL_STORER_CODE, WRAL_ITM_CODE, WRAL_PACK_KEY, WRAL_ITM_SKU_NO, WRAL_ITM_NAME, " & _
                        " WRAL_DSP_ALERT_TYPE AS WRAL_ALERT_TYPE, WRMV_STOCK_OUT_TIME, WRAL_STATUS, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, WRAL_BATCH_NO, WRAL_PALLET_NO, WRWL_BATCH_ID_CALL, WRWL_BATCH_ID_WMS, " & _
                        " convert(varchar,WRMV_BATCH_TIME,103) as WRMV_BATCH_DATE_1,convert(varchar,WRMV_BATCH_TIME,112) as WRMV_BATCH_DATE " & _
                        " FROM WMS_RFID_ALERT_WMS " & _
                        " WHERE 1=1 " & whereSQL & " and LEN(WRMV_TAG_ID) <= 24 ORDER BY WRMV_BATCH_TIME "

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then

                reportSource(nDataSource, Nothing)
            Else
                Response.Write("No Alert Data Found.")
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "REPORT\RFID_DAILY_RPT\RFID_DAILY_RPT.rdlc"
            ReportViewer1.Width = 800

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))
            ReportViewer1.LocalReport.EnableExternalImages = True
            ReportViewer1.ShowPrintButton = True

            If Not IsNothing(paraarray) Then
                ReportViewer1.LocalReport.SetParameters(paraarray)
            End If

            Try

                ReportViewer1.LocalReport.Refresh()

                'Dim formatName As String = "word"

                'For Each extension As RenderingExtension In ReportViewer1.LocalReport.ListRenderingExtensions

                '    If extension.Name.ToLower = formatName Then

                '        Dim m_isVisible As System.Reflection.FieldInfo = extension.GetType.GetField("m_isVisible", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)

                '        m_isVisible.SetValue(extension, False)

                '        Exit For

                '    End If

                'Next



            Catch ex As OutOfMemoryException
                GC.Collect()
                ReportViewer1.LocalReport.Refresh()
            End Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
