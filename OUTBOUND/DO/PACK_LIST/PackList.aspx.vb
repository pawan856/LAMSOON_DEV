Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.Reporting.WebForms

Partial Class OUTBOUND_DO_Pack_list
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private wmsFun As New WMSFunc

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)

        'ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim do_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""

        Dim Total_crt As Integer = 0

        Dim Paras(0) As ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then


            do_code = Request("do_code")
            storer_code = Request("storer_code")

            'imp_code = "WMS"
            'do_code = "000001"
            'storer_code = "001"

            sqlString = " SELECT  1 as C_QTY, WMS_DO_PACKING_D.PAD_CARTON_NO, WMS_DO_PACKING_D.PAD_ITM_CODE, WMS_ALT_VEND_ITEM.AITM_NAME, WMS_ALT_VEND_ITEM.AITM_ORIGIN, " & _
                        "  WMS_DO_PACKING_D.PAD_GROSS_WEIGHT as AITM_GROSS_WEIGHT, WMS_DO_PACKING_D.PAD_NET_WEIGHT as AITM_NET_WEIGHT,ISNULL(WMS_DO_PACKING_D.PAD_QTY,0) as pad_qty, " & _
                        "  WMS_DELV_ORDER.DO_TRADE_TERMS, WMS_DELV_ORDER.DO_SHIP_MODE, WMS_DELV_ORDER.DO_INV_NO, " & _
                        "  WMS_DELV_ORDER.DO_SHIP_TO, WMS_DELV_ORDER.DO_SHIP_ADDR1 + char(13) + char(10) " & _
                        "  + WMS_DELV_ORDER.DO_SHIP_ADDR2 + char(13) + char(10) + WMS_DELV_ORDER.DO_SHIP_ADDR3 AS DO_SHIP_ADDR, " & _
                        "  case when CHARINDEX('-',WMS_DO_PACKING_D.PAD_CARTON_NO)> 0 then right('00000000000000000000' + replace(SUBSTRING(WMS_DO_PACKING_D.PAD_CARTON_NO,1,CHARINDEX('-',WMS_DO_PACKING_D.PAD_CARTON_NO)),'-',''),20) else right('00000000000000000000' + ISNULL(WMS_DO_PACKING_D.PAD_CARTON_NO,0),20) end  as sort_col," & _
                        "  WMS_DELV_ORDER.DO_CONSIGNEE, WMS_DELV_ORDER.DO_CONSIGNEE_ADDR1 + char(13) + char(10) " & _
                        "  + WMS_DELV_ORDER.DO_CONSIGNEE_ADDR2 + char(13) + char(10) + WMS_DELV_ORDER.DO_CONSIGNEE_ADDR3 AS DO_CONSIGNEE_ADDR,  " & _
                        "  WMS_DO_PACKING_D.PAD_PACK_SIZE as PAD_PACK_KEY, WMS_ALT_VEND_ITEM.AITM_UOM, WMS_DO_PACKING_D.PAD_REF_NO, WMS_DELV_ORDER.CUS_NAME, " & _
                        "  WMS_DELV_ORDER.DO_ADDR1 + char(13) + char(10) + WMS_DELV_ORDER.DO_ADDR2 + char(13) + char(10) + WMS_DELV_ORDER.DO_ADDR3 as do_addr , WMS_DELV_ORDER.DO_PAY_TERMS,WMS_DELV_ORDER.DO_DATE, WMS_DO_PACKING_D.PAD_QTY_PER_CTN " & _
                        " FROM WMS_DELV_ORDER INNER JOIN " & _
                          " WMS_DO_PACKING_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DO_PACKING_D.IMP_CODE AND " & _
                         " WMS_DELV_ORDER.STORER_CODE = WMS_DO_PACKING_D.STORER_CODE AND " & _
                         " WMS_DELV_ORDER.DO_CODE = WMS_DO_PACKING_D.DO_CODE LEFT OUTER JOIN " & _
                         " WMS_ALT_VEND_ITEM ON WMS_DO_PACKING_D.PAD_PACK_KEY = WMS_ALT_VEND_ITEM.PACK_KEY AND " & _
                         " WMS_DO_PACKING_D.PAD_ITM_CODE = WMS_ALT_VEND_ITEM.ITM_CODE AND " & _
                         " WMS_DO_PACKING_D.PAD_VND_CODE = WMS_ALT_VEND_ITEM.VND_CODE AND  " & _
                         " WMS_DO_PACKING_D.IMP_CODE = WMS_ALT_VEND_ITEM.IMP_CODE AND " & _
                         " WMS_DO_PACKING_D.STORER_CODE = WMS_ALT_VEND_ITEM.STORER_CODE " & _
                         " where WMS_DELV_ORDER.IMP_CODE='" & gU.dbEncode(imp_code) & "' and WMS_DELV_ORDER.STORER_CODE='" & gU.dbEncode(storer_code) & "' and WMS_DELV_ORDER.DO_CODE='" & gU.dbEncode(do_code) & "'" & _
                         " Order by sort_col "

            nDataSource = gDB.getDataTable(sqlString)

            'Dim keyList As New ArrayList
            'wmsFun.getCartonQty(Total_crt, "PAD_CARTON_NO", keyList, True, , nDataSource)

            If nDataSource.Rows.Count > 0 Then
                For i As Integer = 0 To nDataSource.Rows.Count - 1
                    Dim rowCartonQty As Integer = 0

                    wmsFun.getCartonQty(rowCartonQty, "PAD_CARTON_NO", Nothing, False, nDataSource.Rows(i))


                    nDataSource.Rows(i).Item("C_QTY") = rowCartonQty
                Next

                nDataSource.AcceptChanges()
                Paras(0) = New ReportParameter("DO_CODE", do_code)
                reportSource(nDataSource, Paras)
            Else
                Response.Write("No Pack List Found.")
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "OUTBOUND\DO\Pack_List\pklst.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))

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
