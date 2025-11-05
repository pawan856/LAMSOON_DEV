Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Globalization
Imports Microsoft.Reporting.WebForms.Internal.Soap.ReportingServices2005.Execution
Imports Microsoft.Reporting.WebForms
Imports System.Drawing

Partial Class OUTBOUND_DO_PICKING_LIST_picking_list
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private wmsFun As New WMSFunc
    Private UiFun As New UIfunc

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim do_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""
        Dim tempSQL As String = ""
        Dim tempSQL2 As String = ""

        Dim Paras(6) As Microsoft.Reporting.WebForms.ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        ViewState("STOERE_CODE") = ""
        ViewState("DO_CODE") = ""

        ViewState("DO_CODE") = Server.UrlDecode(Request("DO_CODE"))
        ViewState("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))

        If Not IsPostBack Then

            storer_code = ViewState("STORER_CODE")
            do_code = ViewState("DO_CODE")

            sqlString = "SELECT WMS_DELV_ORDER_D.DOD_REM,WMS_DELV_ORDER_D.DOD_FL_CODE,WMS_DELV_ORDER_D.DOD_WH_CODE,WMS_DELV_ORDER_D.DOD_LOC_WH,WMS_DELV_ORDER_D.DOD_SEQ, WMS_ITEM.ITM_SKU_NO,WMS_DELV_ORDER_D.DOD_EXPIRY_DATE, WMS_DELV_ORDER_D.DOD_QTY, WMS_DELV_ORDER_D.DOD_ITM_DESC, WMS_DELV_ORDER_D.DOD_SS_QTY,WMS_CUST_ORDER.CUS_NAME" &
                    " FROM WMS_DELV_ORDER_D INNER JOIN " &
                            " WMS_ITEM On WMS_DELV_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE And WMS_DELV_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE And " &
                            " WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE And WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " &
                            " WMS_CUST_ORDER On WMS_DELV_ORDER_D.IMP_CODE=WMS_CUST_ORDER.IMP_CODE And " &
                            " WMS_DELV_ORDER_D.STORER_CODE= WMS_CUST_ORDER.STORER_CODE And " &
                            " WMS_DELV_ORDER_D.DOD_CO_CODE=WMS_CUST_ORDER.CO_CODE" &
                            " Where WMS_DELV_ORDER_D.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "' AND WMS_DELV_ORDER_D.STORER_CODE='" & gU.dbEncode(storer_code) & "' " &
                            " And WMS_DELV_ORDER_D.DO_CODE ='" & gU.dbEncode(do_code) & "' order by WMS_DELV_ORDER_D.DOD_ITM_DESC "
            nDataSource = gDB.getDataTable(sqlString)


            If nDataSource.Rows.Count > 0 Then

                'Dim seqNo As String = ""
                Dim sku_no As String = ""
                Dim lot_no As String = ""
                Dim qty As Integer = 0
                Dim ss_qty As Integer = 0
                Dim itm_desc As String = ""
                Dim cus_name As String = ""
                Dim route As String = ""
                Dim dd As String = ""
                Dim wh As String = ""
                Dim loc As String = ""
                Dim floor As String = ""

                Dim dtl As DataTable

                sqlString = "SELECT WMS_DELV_ORDER.ROUTE_ID,WMS_DELV_ORDER.DO_DATE from WMS_DELV_ORDER"
                dtl = gDB.getDataTable(sqlString)

                route = dtl.Rows(0)("ROUTE_ID").ToString
                dd = dtl.Rows(0)("DO_DATE").ToString
                Dim ddate As DateTime = DateTime.Parse(dd, CultureInfo.InvariantCulture)
                Dim reformatted As String = ddate.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture)
                dd = reformatted

                floor = nDataSource.Rows(0)("DOD_FL_CODE").ToString
                wh = nDataSource.Rows(0)("DOD_WH_CODE").ToString
                loc = nDataSource.Rows(0)("DOD_LOC_WH").ToString
                Dim wh_loc As String = wh + loc
                sku_no = nDataSource.Rows(0)("ITM_SKU_NO").ToString
                qty = nDataSource.Rows(0)("DOD_QTY").ToString
                itm_desc = nDataSource.Rows(0)("DOD_ITM_DESC").ToString

                lot_no = nDataSource.Rows(0)("DOD_EXPIRY_DATE").ToString
                If lot_no = "" Then
                    lot_no = ""
                Else
                    Dim ddatetime As DateTime = DateTime.Parse(lot_no, CultureInfo.InvariantCulture)
                    Dim reformattedDate As String = ddatetime.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                    lot_no = reformattedDate
                End If

                Dim sku_lot_no As String = sku_no + lot_no

                Dim tempPath As String = gU.getConfig("SYSP_TEMP_DIR") & "\DO\" & sku_lot_no
                If Not IO.Directory.Exists(tempPath) Then
                    IO.Directory.CreateDirectory(tempPath)
                End If

                Dim tempPathLoc As String = gU.getConfig("SYSP_TEMP_DIR") & "\DO\" & loc
                If Not IO.Directory.Exists(tempPathLoc) Then
                    IO.Directory.CreateDirectory(tempPathLoc)
                End If

                'QR for SKU_LOT_NO'
                Dim oBmp As Bitmap
                oBmp = Code128Rendering.MakeQRImage(sku_lot_no, tempPath & "\" & sku_lot_no & ".jpg")
                oBmp.Save(tempPath & "\" & sku_lot_no & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

                'QR for location'
                Dim oBmpLoc As Bitmap
                oBmpLoc = Code128Rendering.MakeQRImage(loc, tempPathLoc & "\" & loc & ".jpg")
                oBmpLoc.Save(tempPathLoc & "\" & loc & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

                'Parameters'
                Paras(0) = New Microsoft.Reporting.WebForms.ReportParameter("BarCodePath", "file:///" & tempPath & "\" & sku_lot_no & ".jpg")
                Paras(1) = New Microsoft.Reporting.WebForms.ReportParameter("ROUTE_ID", route)
                Paras(2) = New Microsoft.Reporting.WebForms.ReportParameter("DO_DATE", dd)
                Paras(3) = New Microsoft.Reporting.WebForms.ReportParameter("DOD_LOC_WH", wh_loc)
                Paras(4) = New Microsoft.Reporting.WebForms.ReportParameter("BarCodePathLoc", "file:///" & tempPathLoc & "\" & loc & ".jpg")
                Paras(5) = New Microsoft.Reporting.WebForms.ReportParameter("DOD_EXPIRY_DATE", lot_no)
                Paras(6) = New Microsoft.Reporting.WebForms.ReportParameter("DOD_FL_CODE", floor)

                reportSource(nDataSource, Paras)
            Else

                Response.Write("No Check List record has been Found.")
                Response.End()
            End If
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As Microsoft.Reporting.WebForms.ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "OUTBOUND\DO\PICKING_LIST\picking_list.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", sourcetbl))
            ReportViewer1.LocalReport.EnableExternalImages = True

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
