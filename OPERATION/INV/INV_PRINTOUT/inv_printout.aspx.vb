Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms

Partial Class OPERATION_INV_INV_PRINTOUT_inv_printout
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("OP_INV", Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim INVH_NO As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""
        Dim listmode As String = ""

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        INVH_NO = Request.Form("INVH_NO")
        storer_code = Request.Form("storer_code")
        listmode = Request.Form("listmode")

        If listmode = "" Then
            listmode = "DEFAULT"
        Else
            Select Case listmode
                Case "D"
                    listmode = "DEFAULT"
                Case "I"
                    listmode = "ITEM"
                Case "L"
                    listmode = "LOCATION"
                Case Else
                    listmode = "DEFAULT"
            End Select
        End If

        sqlString = "SELECT WMS_INVOICE_HD.IMP_CODE, WMS_INVOICE_HD.STORER_CODE, WMS_INVOICE_HD.INVH_NO, " & _
                    "convert(varchar,WMS_INVOICE_HD.INVH_DATE,103) AS INVH_DATE, " & _
                    "convert(varchar,WMS_INVOICE_HD.INVH_FR_MONTH,103) AS INVH_FR_MONTH, " & _
                    "convert(varchar,WMS_INVOICE_HD.INVH_TO_MONTH,103) AS INVH_TO_MONTH, " & _
                    "WMS_INVOICE_HD.INVH_GEN_BY, " & _
                    "(ISNULL(CASE WHEN WMS_INVOICE_HD.CUS_NAME = '' THEN WMS_INVOICE_HD.CUS_CODE ELSE WMS_INVOICE_HD.CUS_NAME END, '') + " & _
                    "CASE WHEN ISNULL(WMS_INVOICE_HD.INVH_ADDR1,'') <> '' THEN CHAR(13)+CHAR(10)+WMS_INVOICE_HD.INVH_ADDR1 ELSE '' END + " & _
                    "CASE WHEN ISNULL(WMS_INVOICE_HD.INVH_ADDR2,'') <> '' THEN CHAR(13)+CHAR(10)+WMS_INVOICE_HD.INVH_ADDR2 ELSE '' END + " & _
                    "CASE WHEN ISNULL(WMS_INVOICE_HD.INVH_ADDR3,'') <> '' THEN CHAR(13)+CHAR(10)+WMS_INVOICE_HD.INVH_ADDR3 ELSE '' END) AS INVH_MESSER, " & _
                    "WMS_INVOICE_HD.INVH_ATTN, WMS_INVOICE_HD.INVH_TEL, " & _
                    "WMS_INVOICE_HD.INVH_FAX, " & _
                    "WMS_INVOICE_DTL.INVD_SEQ, WMS_INVOICE_DTL.INVD_DISPLAY_SEQ, " & _
                    "WMS_INVOICE_DTL.INVD_ITM_CODE, WMS_INVOICE_DTL.INVD_PACK_KEY, WMS_INVOICE_DTL.INVD_LOC, " & _
                    "WMS_INVOICE_DTL.INVD_TYPE, WMS_INVOICE_DTL.INVD_DESC, " & _
                    "convert(varchar,WMS_INVOICE_DTL.INVD_TX_DATE,103) AS INVD_TX_DATE, " & _
                    "WMS_INVOICE_DTL.INVD_UNIT, " & _
                    "WMS_INVOICE_DTL.INVD_RATE, WMS_INVOICE_DTL.INVD_CURR, " & _
                    "WMS_INVOICE_DTL.INVD_QTY, WMS_INVOICE_DTL.INVD_AMOUNT, " & _
                    "WMS_INVOICE_HD.SYS_CB, WMS_INVOICE_HD.SYS_CD, " & _
                    "WMS_INVOICE_HD.SYS_LUB, WMS_INVOICE_HD.SYS_LUD FROM " & _
                    "WMS_INVOICE_HD INNER JOIN WMS_INVOICE_DTL ON " & _
                    "WMS_INVOICE_HD.IMP_CODE = WMS_INVOICE_DTL.IMP_CODE AND " & _
                    "WMS_INVOICE_HD.STORER_CODE = WMS_INVOICE_DTL.STORER_CODE AND " & _
                    "WMS_INVOICE_HD.INVH_NO = WMS_INVOICE_DTL.INVH_NO " & _
                    "AND WMS_INVOICE_HD.IMP_CODE = '" & imp_code & "' " & _
                    "AND WMS_INVOICE_HD.STORER_CODE='" & gU.dbEncode(storer_code) & "' " & _
                    "AND WMS_INVOICE_HD.INVH_NO='" & gU.dbEncode(INVH_NO) & "' " & _
                    "AND (WMS_INVOICE_DTL.INVD_DISPLAY_GRP = '" & gU.dbEncode(listmode) & "' OR " & _
                        "WMS_INVOICE_DTL.INVD_DISPLAY_GRP = 'ANY') " & _
                    "ORDER BY WMS_INVOICE_DTL.INVD_DISPLAY_SEQ "

        nDataSource = gDB.getDataTable(sqlString)
        If nDataSource.Rows.Count > 0 Then
            reportSource(nDataSource)
        Else
            Response.Write("No Invoice Found or No Invoice has been generated during the invoice date range.")
        End If

    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "OPERATION\INV\INV_PRINTOUT\invoice_printing.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("rptDataSet_inv_printout", sourcetbl))

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
