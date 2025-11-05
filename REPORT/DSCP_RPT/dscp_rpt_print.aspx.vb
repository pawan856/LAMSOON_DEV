Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.Reporting.WebForms

Partial Class DSCP_RPT_PRINT
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private rptU As New ReportUtils
    Const formatStr As String = "V"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("IB_RO", Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim ro_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""
        Dim tempDT As DataTable
        Dim Paras(3) As ReportParameter

        
        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            ViewState("ro_code") = ""
            ViewState("storer_code") = ""

            ro_code = Request.Form("ro_code")
            storer_code = Request.Form("storer_code")

            ViewState("ro_code") = ro_code
            ViewState("storer_code") = storer_code

            sqlString = " SELECT WMS_GOODSRCV_D.GR_CODE, WMS_GOODSRCV_D.GRD_SEQ, WMS_GOODSRCV_D.GRD_PO_QTY, WMS_GOODSRCV_D.GRD_REJ_QTY, " & _
                        " WMS_GOODSRCV_D.GRD_REJ_REASON,WMS_ITEM.ITM_SKU_NO, WMS_GOODSRCV_D.GRD_REJ_PHOTO1,wms_item.ITM_NAME, wms_item.itm_desc, 'GR' as doc_type " & _
                        " FROM WMS_GOODSRCV_D INNER JOIN  " & _
                        " WMS_GOODSRCV ON WMS_GOODSRCV_D.IMP_CODE = WMS_GOODSRCV.IMP_CODE AND   " & _
                        " WMS_GOODSRCV_D.STORER_CODE = WMS_GOODSRCV.STORER_CODE AND WMS_GOODSRCV_D.GR_CODE = WMS_GOODSRCV.GR_CODE  " & _
                        " INNER JOIN WMS_ITEM ON WMS_GOODSRCV_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_GOODSRCV_D.STORER_CODE = WMS_ITEM.STORER_CODE AND   " & _
                        " WMS_GOODSRCV_D.GRD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_GOODSRCV_D.GRD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                        " where isNull(WMS_GOODSRCV_D.GRD_REJ_QTY,0) > 0 AND WMS_GOODSRCV.GR_DOC_NO='" & gU.dbEncode(ro_code) & "'" & _
                        " AND WMS_GOODSRCV_D.IMP_CODE='" & imp_code & "' AND WMS_GOODSRCV_D.STORER_CODE='" & gU.dbEncode(storer_code) & "' " & _
                        " union  " & _
                        " SELECT WMS_GOODSRCV_INSP.GR_CODE, WMS_GOODSRCV_INSP.GRI_SEQ, WMS_GOODSRCV_INSP.GRI_INSP_QTY, WMS_GOODSRCV_INSP.GRI_REJ_QTY,  " & _
                        " WMS_GOODSRCV_INSP.GRI_REJ_REASON,WMS_ITEM.ITM_SKU_NO, WMS_GOODSRCV_INSP.GRI_PHOTO,wms_item.ITM_NAME, wms_item.itm_desc, 'INSP' as doc_type " & _
                        " FROM WMS_GOODSRCV_INSP INNER JOIN " & _
                        " WMS_GOODSRCV ON WMS_GOODSRCV_INSP.IMP_CODE = WMS_GOODSRCV.IMP_CODE AND  " & _
                        " WMS_GOODSRCV_INSP.STORER_CODE = WMS_GOODSRCV.STORER_CODE AND WMS_GOODSRCV_INSP.GR_CODE = WMS_GOODSRCV.GR_CODE " & _
                        " INNER JOIN WMS_ITEM ON WMS_GOODSRCV_INSP.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_GOODSRCV_INSP.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                        " WMS_GOODSRCV_INSP.GRI_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_GOODSRCV_INSP.GRI_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                        " where isNull(WMS_GOODSRCV_INSP.GRI_REJ_QTY,0) > 0  AND WMS_GOODSRCV.GR_DOC_NO='" & gU.dbEncode(ro_code) & "'" & _
                        " AND WMS_GOODSRCV_INSP.IMP_CODE='" & imp_code & "' AND WMS_GOODSRCV_INSP.STORER_CODE='" & gU.dbEncode(storer_code) & "' " & _
                        " union  " & _
                        " SELECT WMS_GOODSRCV_PA.GR_CODE, WMS_GOODSRCV_PA.GRA_SEQ,  WMS_GOODSRCV_PA.GRA_PA_QTY,  " & _
                        " WMS_GOODSRCV_PA.GRA_REJ_QTY, WMS_GOODSRCV_PA.GRA_REJ_REASON, WMS_ITEM.ITM_SKU_NO, WMS_GOODSRCV_PA.GRA_REJ_PHOTO1,wms_item.ITM_NAME, wms_item.itm_desc, 'PA' as doc_type " & _
                        " FROM WMS_GOODSRCV_PA INNER JOIN " & _
                        " WMS_GOODSRCV ON WMS_GOODSRCV_PA.IMP_CODE = WMS_GOODSRCV.IMP_CODE AND  " & _
                        " WMS_GOODSRCV_PA.STORER_CODE = WMS_GOODSRCV.STORER_CODE AND WMS_GOODSRCV_PA.GR_CODE = WMS_GOODSRCV.GR_CODE INNER JOIN " & _
                        " WMS_ITEM ON WMS_GOODSRCV_PA.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_GOODSRCV_PA.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                        " WMS_GOODSRCV_PA.GRA_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_GOODSRCV_PA.GRA_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                        " where isNull(WMS_GOODSRCV_PA.GRA_REJ_QTY,0) > 0  AND WMS_GOODSRCV.GR_DOC_NO='" & gU.dbEncode(ro_code) & "'" & _
                        " AND WMS_GOODSRCV_PA.IMP_CODE='" & imp_code & "' AND WMS_GOODSRCV_PA.STORER_CODE='" & gU.dbEncode(storer_code) & "' "

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource.Rows.Count > 0 Then

                Dim vnd_name As String = ""
                Dim edi_no As String = ""

                sqlString = " SELECT WMS_REPLENISH.RO_VND_CODE, WMS_REPLENISH.RO_EDI_PO_NO, " & _
                            " isNull(isnUll(WMS_VENDOR.VND_SHORTNAME,WMS_VENDOR.VND_NAME),WMS_REPLENISH.RO_VND_CODE) as vnd_name " & _
                            " FROM WMS_REPLENISH LEFT OUTER JOIN " & _
                            " WMS_VENDOR ON WMS_REPLENISH.IMP_CODE = WMS_VENDOR.IMP_CODE AND WMS_REPLENISH.STORER_CODE = WMS_VENDOR.STORER_CODE AND " & _
                            " WMS_REPLENISH.RO_VND_CODE = WMS_VENDOR.VND_CODE " & _
                            " Where WMS_REPLENISH.IMP_CODE='" & gU.dbEncode(imp_code) & "' AND WMS_REPLENISH.STORER_CODE='" & gU.dbEncode(storer_code) & "' " & _
                            " AND WMS_REPLENISH.RO_CODE='" & gU.dbEncode(ro_code) & "'"

                tempDT = gDB.getDataTable(sqlString)

                Paras(0) = New ReportParameter("ParaTo", vnd_name)
                Paras(1) = New ReportParameter("RO_EDI_PO_NO", edi_no)
                Paras(2) = New ReportParameter("MANAGER_NAME", "S. T. Au-Yeung")
                Paras(3) = New ReportParameter("MANAGER_TITLE", "GROUP SUPPLIES MANAGE")

                reportSource(nDataSource, Paras)
            Else
                Response.Write("No Discrepancy item Found.")
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "REPORT\DSCP_RPT\dscp_rpt.rdlc"
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
