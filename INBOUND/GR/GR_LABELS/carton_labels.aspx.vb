Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.Reporting.WebForms

Partial Class INBOUND_GR_GR_LABELS_carton_labels
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Private wmsFun As New WMSFunc

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ar = New AccessRightUtils("IB_GR", Session("usr_id"), Me)

        ar.hideForm(Me)

        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim gr_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack() Then


            gr_code = Request.Form("gr_code")
            storer_code = Request.Form("storer_code")

            sqlString = "select gm.storer_code, gm.gr_track_no, gd.grd_ref_no, " & _
                        "convert(varchar,gm.gr_date, 103) as gr_date, " & _
                        "gm.gr_destination, gd.grd_carton_no, " & _
                        "gd.grd_itm_code, gd.grd_itm_name, " & _
                        "case when ISNULL(gd.GRD_PCS_PER_CARTON, 0) = 0 then 1 else gd.GRD_PCS_PER_CARTON end as grd_rcv_qty, grd_rcv_qty as total_pcs " & _
                        "from wms_goodsrcv gm, wms_goodsrcv_d gd where " & _
                        "gm.imp_code = gd.imp_code and " & _
                        "gm.storer_code = gd.storer_code and " & _
                        "gm.gr_code = gd.gr_code " & _
                        "and gm.imp_code = '" & imp_code & "' " & _
                        "and gm.gr_code='" & gU.dbEncode(gr_code) & "' and gm.storer_code='" & gU.dbEncode(storer_code) & "' " & _
                        "order by gd.grd_carton_no "

            nDataSource = gDB.getDataTable(sqlString)

            Dim distinctCarton As New ArrayList

            Dim rowsArray As New ArrayList

            Dim keyArray As Object()
            Dim keyList As New ArrayList
            ReDim keyArray(nDataSource.Rows.Count - 1)

            wmsFun.getCartonQty(Nothing, "grd_carton_no", keyList, True, , nDataSource)

            If keyList IsNot Nothing Then keyArray = keyList.ToArray
            Dim remain_pcs As Double = 0

            For i As Integer = 0 To nDataSource.Rows.Count - 1
                distinctCarton = New ArrayList

                wmsFun.getCartonQty(Nothing, "grd_carton_no", distinctCarton, False, nDataSource.Rows(i))

                If distinctCarton.Count = 0 Then
                    rowsArray.Add(nDataSource.Rows(i))
                Else
                    For x As Integer = 0 To distinctCarton.Count - 1
                        Dim tempRow As DataRow = nDataSource.NewRow
                        tempRow.Item("storer_code") = nDataSource.Rows(i).Item("storer_code").ToString.Trim
                        tempRow.Item("gr_track_no") = nDataSource.Rows(i).Item("gr_track_no").ToString.Trim
                        tempRow.Item("grd_ref_no") = nDataSource.Rows(i).Item("grd_ref_no").ToString.Trim
                        tempRow.Item("gr_date") = nDataSource.Rows(i).Item("gr_date").ToString.Trim
                        tempRow.Item("gr_destination") = nDataSource.Rows(i).Item("gr_destination").ToString.Trim
                        tempRow.Item("grd_itm_code") = nDataSource.Rows(i).Item("grd_itm_code").ToString.Trim
                        tempRow.Item("grd_itm_name") = nDataSource.Rows(i).Item("grd_itm_name").ToString.Trim

                        tempRow.Item("total_pcs") = nDataSource.Rows(i).Item("total_pcs")

                        If x = distinctCarton.Count - 1 Then
                            remain_pcs = gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("total_pcs").ToString.Trim, 0) Mod gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("grd_rcv_qty").ToString.Trim, 1)
                            If remain_pcs = 0 Then
                                tempRow.Item("grd_rcv_qty") = gU.decodeEmptyCdbl(nDataSource.Rows(i).Item("grd_rcv_qty").ToString.Trim, 1)
                            Else
                                tempRow.Item("grd_rcv_qty") = remain_pcs
                            End If
                        Else
                            tempRow.Item("grd_rcv_qty") = nDataSource.Rows(i).Item("grd_rcv_qty").ToString.Trim
                        End If

                        tempRow.Item("grd_carton_no") = gU.decodeEmptyCInt(distinctCarton(x), 1)
                        rowsArray.Add(tempRow)
                    Next
                End If
            Next

            Dim allRows As Object() = rowsArray.ToArray

            Dim collection = From c In allRows _
                                 Order By _
                                 c.item("storer_code"), _
                                 CInt(c.item("grd_carton_no")), _
                                 c.item("grd_ref_no"), _
                                 c.item("gr_track_no"), _
                                 c.item("gr_destination"), _
                                 c.item("gr_date") _
                                 Ascending _
                                Select c

            Dim sortRows As Object()

            sortRows = collection.ToArray

            Dim cartonTable As New DataTable

            cartonTable = nDataSource.Clone

            For t As Integer = 0 To sortRows.Count - 1
                Dim nRow As DataRow = cartonTable.NewRow

                nRow.Item("storer_code") = DirectCast(sortRows(t), DataRow).Item("storer_code").ToString.Trim
                nRow.Item("gr_track_no") = DirectCast(sortRows(t), DataRow).Item("gr_track_no").ToString.Trim
                nRow.Item("grd_ref_no") = DirectCast(sortRows(t), DataRow).Item("grd_ref_no").ToString.Trim
                nRow.Item("gr_date") = DirectCast(sortRows(t), DataRow).Item("gr_date").ToString.Trim
                nRow.Item("gr_destination") = DirectCast(sortRows(t), DataRow).Item("gr_destination").ToString.Trim
                nRow.Item("grd_itm_code") = DirectCast(sortRows(t), DataRow).Item("grd_itm_code").ToString.Trim
                nRow.Item("grd_itm_name") = DirectCast(sortRows(t), DataRow).Item("grd_itm_name").ToString.Trim
                nRow.Item("grd_rcv_qty") = DirectCast(sortRows(t), DataRow).Item("grd_rcv_qty").ToString.Trim
                nRow.Item("total_pcs") = DirectCast(sortRows(t), DataRow).Item("total_pcs").ToString.Trim
                nRow.Item("grd_carton_no") = DirectCast(sortRows(t), DataRow).Item("grd_carton_no").ToString.Trim
                cartonTable.Rows.Add(nRow)

            Next

            cartonTable.AcceptChanges()

            reportSource(cartonTable)
        End If
    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "INBOUND\GR\GR_LABELS\gr_carton_lbl.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("rptDataSet_gr_carton_label", sourcetbl))

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
