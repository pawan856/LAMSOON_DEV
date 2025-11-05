Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.Reporting.WebForms

Partial Class INBOUND_GR_GR_LABELS_gr_labels
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("IB_GR", Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim gr_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then


            gr_code = Request.Form("gr_code")
            storer_code = Request.Form("storer_code")

            sqlString = " SELECT gm.storer_code,  gm.gr_track_no,  gd.grd_pallet_no,  gd.grd_ref_no,  convert(varchar,gm.gr_date, 103) AS gr_date, " & _
                        "  gm.gr_destination,  gd.grd_batch_no,  gd.grd_stocktype,  gd.grd_itm_code,  gd.grd_itm_name,  ISNULL(gd.grd_rcv_qty, 0) AS grd_rcv_qty, " & _
                        "  ISNULL(gd.GRD_CARTON_NO, 0) as total_carton " & _
                        " FROM wms_goodsrcv gm,  wms_goodsrcv_d gd " & _
                        " WHERE gm.imp_code  = gd.imp_code " & _
                        " AND gm.storer_code = gd.storer_code " & _
                        " AND gm.gr_code     = gd.gr_code " & _
                        "and gm.imp_code = '" & imp_code & "' " & _
                        "and gm.gr_code='" & gU.dbEncode(gr_code) & "' and gm.storer_code='" & gU.dbEncode(storer_code) & "' " & _
                        "order by gd.grd_pallet_no "


            nDataSource = gDB.getDataTable(sqlString)

            reportSource(nDataSource)

        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "INBOUND\GR\GR_LABELS\gr_pallet_lbl.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("rptDataSet_gr_pallet_label", sourcetbl))

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
