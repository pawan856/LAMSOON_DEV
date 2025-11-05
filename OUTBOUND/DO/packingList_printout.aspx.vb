Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports Microsoft.Reporting.WebForms

Partial Class OUTBOUND_DO_packingList_printout
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Private wmsFun As New WMSFunc

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)

        'ar.hideForm(Me)

        Dim sqlString As String = ""
        Dim nDataSource As DataTable
        Dim do_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then

            do_code = Request("do_code")
            storer_code = Request("storer_code")

            sqlString += "select pd.imp_code, pd.storer_code, pd.do_code, pd.pad_pack_no, "
            sqlString += "pd.pad_display_seq, pd.pad_pallet_no, pd.pad_pack_type, pd.pad_totl_packs, "
            sqlString += "pd.pad_uom, pd.pad_carton_no,case when CHARINDEX('-',pd.PAD_CARTON_NO)> 0 then right(replicate('0',20) + replace(substring(pd.PAD_CARTON_NO,1,CHARINDEX('-',pd.PAD_CARTON_NO)),'-',''),20) else right(replicate('0',20) + ISNULL(pd.PAD_CARTON_NO,0),20) end  as sort_col, pd.pad_itm_code, pd.pad_vnd_code, "
            sqlString += "pd.pad_batch_no, pd.pad_qty, pd.pad_length, pd.pad_width, pd.pad_height, "
            sqlString += "pd.pad_ref_no, pd.pad_pack_by, pd.pad_net_weight, pd.pad_gross_weight, pd.pad_min_packing, "
            sqlString += "pd.pad_origin, pd.sys_cb, pd.sys_cd, pd.sys_lub, pd.sys_lud, "
            sqlString += "do.cus_name, do.do_date as do_date, case when ISNULL(pd.pad_ref_no, '') <> '' then pd.pad_ref_no + ' / ' else '' end + pd.pad_batch_no as c_pad_batch_no, "
            sqlString += "ISNULL(pd.pad_net_weight,0) as pad_net_weight, ISNULL(pd.pad_gross_weight,0) as pad_gross_weight, d.aitm_origin, "
            sqlString += "0 as tot_qty, 0 as ctn_qty, wms_storer.sto_name as storer_name, do.do_inv_no, "
            sqlString += "do.do_pack_label_1, do.do_pack_label_2, do.do_pack_label_3, "
            sqlString += "do.do_pack_label_qty_1, do.do_pack_label_qty_2, do.do_pack_label_qty_3, "
            'sqlString += "ISNULL(d.aitm_pcs_per_pack, pd.pad_pack_key) as pad_pack_key "
            sqlString += "pd.pad_pack_size as pad_pack_key,pd.PAD_QTY_PER_CTN "
            sqlString += "from wms_delv_order do inner join wms_storer on "
            sqlString += "do.imp_code = wms_storer.imp_code "
            sqlString += "and do.storer_code = wms_storer.storer_code "
            sqlString += "inner join wms_do_packing_d pd on "
            sqlString += "do.imp_code = pd.imp_code "
            sqlString += "and do.storer_code = pd.storer_code "
            sqlString += "and do.do_code = pd.do_code "
            sqlString += "left outer join wms_item m on "
            sqlString += "pd.imp_code = m.imp_code and pd.storer_code = m.storer_code "
            sqlString += "and pd.pad_pack_key = m.pack_key "
            sqlString += "and pd.pad_itm_code = m.itm_code "
            sqlString += "left outer join wms_alt_vend_item d on m.imp_code = d.imp_code "
            sqlString += "and m.storer_code = d.storer_code and m.itm_code = d.itm_code "
            sqlString += "and m.pack_key = d.pack_key and pd.pad_vnd_code = d.vnd_code "
            sqlString += "where pd.do_code = '" & gU.dbEncode(do_code) & "' "
            sqlString += "and pd.storer_code = '" & gU.dbEncode(storer_code) & "' "
            sqlString += "and pd.imp_code = '" & imp_code & "' "
            sqlString += "order by sort_col, pd.pad_pack_no"

            nDataSource = gDB.getDataTable(sqlString)

            Dim paras(0) As ReportParameter

            If nDataSource.Rows.Count > 0 Then
                Dim rowCtnQty As Integer = 0

                wmsFun.getCartonQty(rowCtnQty, "pad_carton_no", Nothing, True, , nDataSource)
                paras(0) = New ReportParameter("totalCtn", rowCtnQty)

                For i As Integer = 0 To nDataSource.Rows.Count - 1
                    rowCtnQty = 0

                    'wmsFun.getCartonQty(rowCtnQty, "pad_carton_no", Nothing, False, nDataSource.Rows(i))

                    'If rowCtnQty = 0 Then
                    '    rowCtnQty = 1
                    'End If

                    Dim packSize As Integer = 0

                    'packSize = rowCtnQty * gU.decodeEmptyCInt(nDataSource.Rows(i).Item("pad_pack_key").ToString, 0)
                    packSize = gU.decodeEmptyCInt(nDataSource.Rows(i).Item("pad_pack_key").ToString, 0)

                    nDataSource.Rows(i).Item("pad_pack_key") = packSize

                    nDataSource.Rows(i).Item("tot_qty") = packSize * gU.decodeEmptyCInt(nDataSource.Rows(i).Item("pad_qty").ToString, 0)

                    nDataSource.Rows(i).Item("ctn_qty") = rowCtnQty
                Next

                nDataSource.AcceptChanges()

                reportSource(nDataSource, paras)
            Else
                Response.Write("No Packing Items Found.")
            End If
        End If

    End Sub

    Private Sub reportSource(ByVal sourcetbl As DataTable, Optional ByVal paraarray() As ReportParameter = Nothing)
        Try
            ReportViewer1.LocalReport.ReportPath = "OUTBOUND\DO\packinglist.rdlc"

            ReportViewer1.LocalReport.DataSources.Clear()

            ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("rptDataSet_packing_list", sourcetbl))

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
