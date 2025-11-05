Imports System
Imports System.Configuration
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Xml.Serialization
Imports System.Data
Imports System.ServiceModel
Imports System.Data.SqlClient
Imports System.Web.Caching

Public Class WMSService
    Implements IWMSService

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private gU As New GeneralUtils

    Public Function LoginWMS(ByVal user As String, ByVal pwd As String) As Boolean Implements IWMSService.LoginWMS
        If user <> "" And pwd <> "" Then
            'Return Membership.ValidateUser(user, pwd)
        End If
    End Function

    Public Function getROList(ByVal storer As String) As DataTable Implements IWMSService.getROList
        Dim nSQL As String = ""
        Dim datatbl As New DataTable

        nSQL += "select m.storer_code, m.imp_code, m.ro_code, d.rod_seq, d.rod_disp_seq, d.rod_pallet_no, d.rod_carton_no, d.rod_batch_no,  "
        nSQL += "d.rod_ref_no, d.rod_itm_parent, d.rod_itm_code, ISNULL(d.rod_pack_key,1) as rod_pack_key, d.rod_itm_name,  "
        nSQL += "d.rod_qty, d.rod_uom, d.rod_pcs_per_uom, d.rod_tot_pcs, d.rod_length, d.rod_width, d.rod_height,  "
        nSQL += "d.rod_kg, d.rod_cbm, m.ro_ref_no, d.rod_vnd_code,  "
        nSQL += "m.ro_dest, to_char(m.ro_date, 'yy/mm/dd') as ro_date,  "
        nSQL += "(d.rod_qty - ISNULL(g.grd_po_qty, 0)) as avail_qty, m.ro_track_no, wms_storer.sto_shortname, 'N' saveFlag, '' as gr_date "
        nSQL += "from wms_replenish m inner join wms_storer on "
        nSQL += "m.imp_code = wms_storer.imp_code  "
        nSQL += "and m.storer_code = wms_storer.storer_code  "
        nSQL += "left outer join wms_replenish_d d on "
        nSQL += "m.imp_code = d.imp_code  "
        nSQL += "and m.storer_code = d.storer_code  "
        nSQL += "and m.ro_code = d.ro_code left outer join "
        nSQL += "( "
        nSQL += "select gd.storer_code, gd.imp_code, gd.grd_itm_code, gd.grd_vnd_code, "
        nSQL += "gd.grd_pack_key, gd.grd_pallet_no, gr_doc_no, sum(gd.grd_po_qty) grd_po_qty "
        nSQL += "from wms_goodsrcv gh, wms_goodsrcv_d gd  "
        nSQL += "where gh.imp_code = gd.imp_code  "
        nSQL += "and gh.storer_code = gd.storer_code  "
        nSQL += "and gh.gr_code = gd.gr_code  "
        nSQL += "and gh.imp_code = 'WMS'  "

        If storer <> "" Then
            nSQL += "and gh.storer_code = '" & storer & "'  "
        End If

        nSQL += "and ISNULL(gh.gr_status, '') <> 'CANCELLED'  "
        nSQL += "group by  gd.storer_code, gd.imp_code, gd.grd_itm_code, gd.grd_vnd_code, "
        nSQL += "gd.grd_pack_key, gd.grd_pallet_no, gr_doc_no "
        nSQL += ") g on "
        nSQL += "d.imp_code = g.imp_code "
        nSQL += "and d.storer_code = g.storer_code "
        nSQL += "and d.rod_itm_code = g.grd_itm_code  "
        nSQL += "and d.rod_pack_key = g.grd_pack_key  "
        nSQL += "and ISNULL(d.rod_vnd_code,'000') = ISNULL(g.grd_vnd_code, '000') "
        nSQL += "and ISNULL(d.rod_pallet_no, '000') = ISNULL(g.grd_pallet_no, '000') "
        nSQL += "and d.ro_code = g.gr_doc_no "
        nSQL += "where m.imp_code = 'WMS'  "

        If storer <> "" Then
            nSQL += "and m.storer_code = '" & storer & "' "
        End If

        nSQL += "and m.ro_status not in ('CLOSED','CONFIRMED') "
        nSQL += "and (d.rod_qty - ISNULL(g.grd_po_qty, 0)) > 0 "
        nSQL += "order by m.ro_date desc, m.ro_code desc, d.rod_disp_seq"


        Try
            datatbl = gDB.getDataTable(nSQL)

            Return datatbl

        Catch ex As Exception
            Console.Write(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function getROItemList(ByVal storer As String) As DataTable Implements IWMSService.getROItemList
        Dim nSQL As String = ""
        Dim datatbl As New DataTable

        nSQL += "select m.storer_code, m.imp_code, m.ro_code, d.rod_seq, d.rod_disp_seq, d.rod_pallet_no, d.rod_carton_no, d.rod_batch_no,  "
        nSQL += "d.rod_ref_no, d.rod_itm_parent, d.rod_itm_code, ISNULL(d.rod_pack_key,1) as rod_pack_key, d.rod_itm_name,  "
        nSQL += "d.rod_qty, d.rod_uom, d.rod_pcs_per_uom, d.rod_tot_pcs, d.rod_length, d.rod_width, d.rod_height,  "
        nSQL += "d.rod_kg, d.rod_cbm, m.ro_ref_no, d.rod_vnd_code,  "
        nSQL += "m.ro_dest, to_char(varchar, m.ro_date, 'yy/mm/dd') as ro_date,  "
        nSQL += "(d.rod_qty - ISNULL(g.grd_po_qty, 0)) as avail_qty, m.ro_track_no  "
        nSQL += "from wms_replenish m left outer join wms_replenish_d d on "
        nSQL += "m.imp_code = d.imp_code  "
        nSQL += "and m.storer_code = d.storer_code  "
        nSQL += "and m.ro_code = d.ro_code left outer join "
        nSQL += "( "
        nSQL += "select gd.storer_code, gd.imp_code, gd.grd_itm_code, gd.grd_vnd_code, "
        nSQL += "gd.grd_pack_key, gd.grd_pallet_no, gr_doc_no, sum(gd.grd_po_qty) grd_po_qty "
        nSQL += "from wms_goodsrcv gh, wms_goodsrcv_d gd  "
        nSQL += "where gh.imp_code = gd.imp_code  "
        nSQL += "and gh.storer_code = gd.storer_code  "
        nSQL += "and gh.gr_code = gd.gr_code  "
        nSQL += "and gh.imp_code = 'WMS'  "

        If storer <> "" Then
            nSQL += "and gh.storer_code = '" & storer & "'  "
        End If

        nSQL += "and ISNULL(gh.gr_status, '') <> 'CANCELLED'  "
        nSQL += "group by  gd.storer_code, gd.imp_code, gd.grd_itm_code, gd.grd_vnd_code, "
        nSQL += "gd.grd_pack_key, gd.grd_pallet_no, gr_doc_no "
        nSQL += ") g on "
        nSQL += "d.imp_code = g.imp_code "
        nSQL += "and d.storer_code = g.storer_code "
        nSQL += "and d.rod_itm_code = g.grd_itm_code  "
        nSQL += "and d.rod_pack_key = g.grd_pack_key  "
        nSQL += "and ISNULL(d.rod_vnd_code,'000') = ISNULL(g.grd_vnd_code, '000') "
        nSQL += "and ISNULL(d.rod_pallet_no, '000') = ISNULL(g.grd_pallet_no, '000') "
        nSQL += "and d.ro_code = g.gr_doc_no "
        nSQL += "where m.imp_code = 'WMS'  "

        If storer <> "" Then
            nSQL += "and m.storer_code = '" & storer & "' "
        End If

        nSQL += "and m.ro_status not in ('CLOSED','CONFIRMED') "
        nSQL += "and (d.rod_qty - ISNULL(g.grd_po_qty, 0)) > 0 "
        nSQL += "order by m.ro_code desc, d.rod_disp_seq"

        Try
            datatbl = gDB.getDataTable(nSQL)

            Return datatbl

        Catch ex As Exception
            Console.Write(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function createGR(ByVal grTable As DataTable, ByVal roTable As DataTable, _
                             ByVal usrID As String, ByVal gr_date As String, _
                             ByVal paTable As DataTable) As Boolean Implements IWMSService.createGR
        Dim gConn As SqlConnection

        SetCache()

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction

        Try
            If grTable IsNot Nothing Then
                Dim nextNo As String = ""
                Dim sql_string As String = ""
                Dim itemSQL As String = ""

                nextNo = DB.getDocNo("GR", gConn, transaction)

                sql_string += "insert into wms_goodsrcv ("
                sql_string += "imp_code, storer_code, gr_code,"
                sql_string += "gr_status, gr_date, gr_rcv_by, gr_doc_no,"
                sql_string += "po_date, gr_destination, gr_track_no, gr_ref_no,"
                sql_string += "sys_cb, sys_cd, sys_lub, sys_lud"
                sql_string += ") values ( "
                sql_string += gU.convdbNVCData("WMS") & "," & gU.convdbNVCData(gU.dbEncode(roTable.Rows(0).Item("storer_code").ToString)) & "," & gU.convdbNVCData(nextNo) & ","
                sql_string += gU.convdbNVCData("NEW") & "," & gU.convdbDate(gr_date) & "," & gU.convdbNVCData(usrID) & "," & gU.convdbNVCData(gU.dbEncode(roTable.Rows(0).Item("ro_code").ToString)) & ","
                sql_string += gU.convdbDate(roTable.Rows(0).Item("ro_date").ToString) & "," & gU.convdbNVCData(gU.dbEncode(roTable.Rows(0).Item("ro_dest").ToString)) & ","
                sql_string += gU.convdbNVCData(gU.dbEncode(roTable.Rows(0).Item("ro_track_no").ToString)) & "," & gU.convdbNVCData(gU.dbEncode(roTable.Rows(0).Item("ro_ref_no").ToString)) & ","
                sql_string += gU.convdbNVCData(usrID) & ", Getdate()," & gU.convdbNVCData(usrID) & ",Getdate()) "

                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                Dim grdSEQ As Integer = 1

                For Each rows As DataRow In grTable.Select("gr_code = '" & gU.dbEncode(roTable.Rows(0).Item("ro_code").ToString) & "'")
                    itemSQL = ""
                    itemSQL += "insert into WMS_GOODSRCV_D "
                    itemSQL += "(imp_code, storer_code, gr_code, grd_seq, grd_disp_seq, grd_pack, grd_pack_no, grd_pallet_no, "
                    itemSQL += "grd_carton_no, grd_batch_no, grd_batch_cd, grd_ref_no, grd_itm_code, grd_brand, grd_series, "
                    itemSQL += "grd_model, grd_itm_name, grd_po_qty, grd_rcv_qty, grd_os_qty, "
                    itemSQL += "grd_uom, grd_width, grd_length, grd_height, grd_kg, grd_cbm, "
                    itemSQL += "grd_destination, grd_stocktype, grd_pack_key, grd_pcs_per_uom, grd_tot_pcs, grd_ref_seq, "
                    itemSQL += "grd_no_of_carton, grd_pcs_per_carton, grd_vnd_code, "
                    itemSQL += "sys_cb, sys_cd, sys_lub, sys_lud) "
                    itemSQL += "values ("
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("imp_code").ToString.Trim, ""))) & "," & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("storer_code").ToString.Trim, ""))) & "," & gU.convdbNVCData(gU.dbEncode(nextNo)) & ", "
                    itemSQL += gU.convdbNVCData(grdSEQ) & ","
                    itemSQL += gU.convdbNVCData(grdSEQ) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_pack").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_pack_no").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_pallet_no").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_carton_no").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_batch_no").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_batch_cd").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_ref_no").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_itm_code").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_brand").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_series").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_model").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_itm_name").ToString.Trim, ""))) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_po_qty").ToString.Trim, "0")) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_rcv_qty").ToString.Trim, "0")) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_os_qty").ToString.Trim, "0")) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_uom").ToString.Trim, ""))) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_width").ToString.Trim, "0")) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_length").ToString.Trim, "0")) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_height").ToString.Trim, "0")) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_kg").ToString.Trim, "0")) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_cbm").ToString.Trim, "0")) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_destination").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_stocktype").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_pack_key").ToString.Trim, ""))) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pcs_per_uom").ToString.Trim, "0")) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_tot_pcs").ToString.Trim, "0")) & ","
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_ref_seq").ToString.Trim, "null")) & ", "
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_no_of_carton").ToString.Trim, "0")) & ", "
                    itemSQL += gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("grd_pcs_per_carton").ToString.Trim, "0")) & ", "
                    itemSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("grd_vnd_code").ToString.Trim, ""))) & ","
                    itemSQL += gU.convdbNVCData(usrID) & ", Getdate()," & gU.convdbNVCData(usrID) & ",Getdate()) "

                    If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                    grdSEQ += 1
                Next

                Dim paSQL As String = ""
                Dim graSEQ As Integer = 1

                For Each paRow As DataRow In paTable.Select("gr_code = '" & gU.dbEncode(roTable.Rows(0).Item("ro_code").ToString) & "'")
                    paSQL = ""
                    paSQL += "insert into wms_goodsrcv_pa "
                    paSQL += "(imp_code, storer_code, gr_code, gra_seq, gra_disp_seq, gra_itm_code, gra_pack_key, "
                    paSQL += "gra_pa_qty, gra_batch_no, gra_pallet_no, gra_ref_no, "
                    paSQL += "sys_cb, sys_cd, sys_lub, sys_lud) "
                    paSQL += "values "
                    paSQL += "(" & gU.convdbNVCData("WMS") & ","
                    paSQL += gU.convdbNVCData(gU.dbEncode(roTable.Rows(0).Item("storer_code").ToString)) & ","
                    paSQL += gU.convdbNVCData(nextNo) & ","
                    paSQL += gU.convdbNVCData(graSEQ) & ","
                    paSQL += graSEQ & ","
                    paSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(paRow.Item("gra_itm_code").ToString.Trim, ""))) & ","
                    paSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(paRow.Item("gra_pack_key").ToString.Trim, ""))) & ","
                    paSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(paRow.Item("gra_pa_qty").ToString.Trim, "0"))) & ","
                    paSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(paRow.Item("gra_batch_no").ToString.Trim, "N/A"))) & ","
                    paSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(paRow.Item("gra_pallet_no").ToString.Trim, ""))) & ","
                    paSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(paRow.Item("gra_ref_no").ToString.Trim, ""))) & ","
                    paSQL += gU.convdbNVCData(usrID) & ", Getdate()," & gU.convdbNVCData(usrID) & ",Getdate()) "

                    If paSQL <> "" Then gDB.amendData(paSQL, gConn, transaction)

                    graSEQ += 1

                    If gU.decodeNullOrEmpty(paRow.Item("gra_batch_no").ToString.Trim, "") <> "" Then
                        Dim gotDC As String = DB.getValueFromSQL("select 1 from wms_date_code where " & _
                                                                 "dc_date_code = '" & gU.dbEncode(gU.decodeNullOrEmpty(paRow.Item("gra_batch_no").ToString.Trim, "")) & "' " & _
                                                                 "and storer_code = '" & gU.dbEncode(gU.decodeNullOrEmpty(paRow.Item("storer_code").ToString.Trim, "")) & "' " & _
                                                                 "and imp_code = '" & gU.dbEncode("WMS") & "' ")

                        If gotDC Is Nothing Then
                            Dim dcSQL As String = ""

                            dcSQL += "insert into wms_date_code ("
                            dcSQL += "dc_date_code, storer_code, imp_code, "
                            dcSQL += "sys_cb, sys_cd, sys_lub, sys_lud) values ( "
                            dcSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(paRow.Item("gra_batch_no").ToString.Trim, ""))) & ","
                            dcSQL += gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(paRow.Item("storer_code").ToString.Trim, ""))) & ","
                            dcSQL += gU.convdbNVCData(gU.dbEncode("WMS")) & ","
                            dcSQL += gU.convdbNVCData(usrID) & ", Getdate()," & gU.convdbNVCData(usrID) & ",Getdate()) "

                            If dcSQL <> "" Then gDB.amendData(dcSQL, gConn, transaction)
                        End If
                    End If
                Next

                transaction.Commit()
            End If

            Return True

        Catch ex As Exception
            transaction.Rollback()
            Console.Write(ex.Message)

            Return False
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Function

    Public Function getPickList(ByVal storer As String) As DataTable Implements IWMSService.getPickList
        Dim nSQL As String = ""
        Dim datatbl As New DataTable

        nSQL += "select wms_delv_order.imp_code, "
        nSQL += "wms_delv_order.storer_code, wms_delv_order.do_code, wms_storer.sto_shortname, wms_delv_order.do_inv_no, "
        nSQL += "pld_seq, pld_item_no, pld_pack_key, pld_pallet_no,  ISNULL(DOD_QTY,pld_do_qty) as dod_qty, DOD_DISP_SEQ, dod_pcs_uom, dod_seq, pld_batch_no, "
        nSQL += "wms_do_picklist_d.sys_cb, wms_do_picklist_d.sys_cd, wms_do_picklist_d.sys_lub, wms_do_picklist_d.sys_lud, "
        nSQL += "'N' saveFlag, 0 as pl_actual_qty "
        nSQL += "from wms_do_picklist_d inner join wms_delv_order on "
        nSQL += "wms_do_picklist_d.imp_code = wms_delv_order.imp_code "
        nSQL += "and wms_do_picklist_d.storer_code = wms_delv_order.storer_code "
        nSQL += "and wms_do_picklist_d.do_code = wms_delv_order.do_code "
        nSQL += "inner join wms_storer on "
        nSQL += "wms_delv_order.imp_code = wms_storer.imp_code "
        nSQL += "and wms_delv_order.storer_code = wms_storer.storer_code "
        nSQL += "left outer join wms_delv_order_d on "
        nSQL += "wms_do_picklist_d.imp_code = wms_delv_order_d.imp_code "
        nSQL += "and wms_do_picklist_d.storer_code = wms_delv_order_d.storer_code "
        nSQL += "and wms_do_picklist_d.do_code = wms_delv_order_d.do_code "
        nSQL += "and wms_do_picklist_d.pld_item_no = wms_delv_order_d.dod_itm_code "
        nSQL += "and wms_do_picklist_d.pld_pack_key = wms_delv_order_d.dod_pack_key "
        nSQL += "and wms_do_picklist_d.pld_batch_no = (case "
        nSQL += "when wms_delv_order_d.DOD_BATCH_NO is null then 'N/A' "
        nSQL += "when wms_delv_order_d.DOD_BATCH_NO = '' then 'N/A' "
        nSQL += "else wms_delv_order_d.DOD_BATCH_NO end) "
        nSQL += "where wms_delv_order.imp_code = 'WMS' "

        'nSQL += "(select sum(pld_item_qty) from wms_do_picklist_d pd where "
        'nSQL += "pd.imp_code = wms_do_picklist_d.imp_code "
        'nSQL += "and pd.storer_code = wms_do_picklist_d.storer_code "
        'nSQL += "and pd.do_code = wms_do_picklist_d.do_code "
        'nSQL += "and pd.pld_pack_key = wms_do_picklist_d.pld_pack_key "
        'nSQL += "and pd.pld_item_no = wms_do_picklist_d.pld_item_no) as pl_actual_qty "

        If storer <> "" Then
            nSQL += "and wms_delv_order.storer_code = '" & storer & "'  "
        End If

        nSQL += "and wms_delv_order.do_status not in ('CANCELLED','POSTED') "
        nSQL += "order by to_number(wms_delv_order.do_code) desc "

        Try
            datatbl = gDB.getDataTable(nSQL)

            Return datatbl

        Catch ex As Exception
            Console.Write(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function getPickItemList(ByVal storer As String) As DataTable Implements IWMSService.getPickItemList
        Dim nSQL As String = ""
        Dim datatbl As New DataTable


        nSQL += "select wms_do_picklist_d.imp_code, "
        nSQL += "wms_do_picklist_d.storer_code, wms_do_picklist_d.do_code, wms_storer.sto_shortname, wms_delv_order.do_inv_no, "
        nSQL += "pld_seq, pld_picked_by, pld_item_no, pld_pack_key, pld_pallet_no, 0 as pld_item_qty, pld_wh, pld_loc, "
        nSQL += "pld_floor, pld_area, pld_rack, pld_bin, pld_is_loan, pld_loan_date, PLD_BATCH_NO, l.ILOC_BAL_QTY, pld_do_qty, ISNULL(DOD_QTY,pld_do_qty) as dod_qty, wms_do_picklist_d.sys_cb, "
        nSQL += "wms_do_picklist_d.sys_cd, wms_do_picklist_d.sys_lub, wms_do_picklist_d.sys_lud, 'N' saveFlag "
        nSQL += "from wms_delv_order inner join wms_storer on "
        nSQL += "wms_delv_order.imp_code = wms_storer.imp_code  "
        nSQL += "and wms_delv_order.storer_code = wms_storer.storer_code  "
        nSQL += "inner join wms_do_picklist_d on "
        nSQL += "wms_do_picklist_d.imp_code = wms_delv_order.imp_code "
        nSQL += "and wms_do_picklist_d.storer_code = wms_delv_order.storer_code "
        nSQL += "and wms_do_picklist_d.do_code = wms_delv_order.do_code "
        nSQL += "left outer join WMS_ITEM_LOC_BAL l on "
        nSQL += "wms_do_picklist_d.STORER_CODE = l.STORER_CODE "
        nSQL += "and wms_do_picklist_d.PLD_ITEM_NO = l.ITM_CODE "
        nSQL += "and wms_do_picklist_d.PLD_PACK_KEY = l.PACK_KEY "
        nSQL += "and wms_do_picklist_d.PLD_PALLET_NO = l.ILOC_PALLET_NO "
        nSQL += "and wms_do_picklist_d.PLD_LOC = l.ILOC_LOC "
        nSQL += "and wms_do_picklist_d.PLD_BATCH_NO = l.ILOC_BATCH_NO "
        nSQL += "left outer join WMS_DATE_CODE DC on "
        nSQL += "l.ILOC_BATCH_NO = DC.DC_DATE_CODE "
        nSQL += "and l.STORER_CODE = DC.STORER_CODE "
        nSQL += "and l.IMP_CODE = DC.IMP_CODE "

        nSQL += "left outer join wms_delv_order_d on "
        nSQL += "wms_do_picklist_d.imp_code = wms_delv_order_d.imp_code "
        nSQL += "and wms_do_picklist_d.storer_code = wms_delv_order_d.storer_code "
        nSQL += "and wms_do_picklist_d.do_code = wms_delv_order_d.do_code "
        nSQL += "and wms_do_picklist_d.pld_item_no = wms_delv_order_d.dod_itm_code "
        nSQL += "and wms_do_picklist_d.pld_pack_key = wms_delv_order_d.dod_pack_key "
        nSQL += "and wms_do_picklist_d.pld_batch_no = (case "
        nSQL += "when wms_delv_order_d.DOD_BATCH_NO is null then 'N/A' "
        nSQL += "when wms_delv_order_d.DOD_BATCH_NO = '' then 'N/A' "
        nSQL += "else wms_delv_order_d.DOD_BATCH_NO end) "

        nSQL += "where wms_do_picklist_d.imp_code = 'WMS'  "

        If storer <> "" Then
            nSQL += "and wms_do_picklist_d.storer_code = '" & storer & "'  "
        End If

        nSQL += "and wms_delv_order.do_status not in ('CANCELLED','POSTED') "
        nSQL += "order by DC.DC_DATE_CODE, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_BATCH_NO, PLD_WH, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN "

        Try
            datatbl = gDB.getDataTable(nSQL)

            Return datatbl

        Catch ex As Exception
            Console.Write(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function getLocation() As DataTable Implements IWMSService.getLocation
        Dim nSQL As String = ""
        Dim datatbl As New DataTable

        nSQL += "select ar_code, rk_code,"
        nSQL += "wh_code+replace(str(fl_num,2),' ','0')+"
        nSQL += "replicate('0', 3 - len(ar_code))+ar_code+"
        nSQL += "replicate('0', 4 - len(rk_code))+rk_code+"
        nSQL += "replicate('0', 3 - len(bn_code))+bn_code as loc from wms_wh_bin "

        Try
            datatbl = gDB.getDataTable(nSQL)

            Return datatbl

        Catch ex As Exception
            Console.Write(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function updatePickList(ByVal usrID As String, ByVal dt As DataTable) As Boolean Implements IWMSService.updatePickList
        Dim gConn As SqlConnection

        SetCache()

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction

        Try
            For Each rows As DataRow In dt.Rows
                If rows.Item("saveFlag").ToString = "Y" Then
                    Dim itemSQL As String = ""

                    itemSQL += "update wms_do_picklist_d set "
                    itemSQL += "pld_picked_by = " & gU.convdbNVCData(usrID) & ", "
                    itemSQL += "pld_item_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, ""))) & ", "
                    itemSQL += "pld_pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, ""))) & ", "
                    itemSQL += "pld_item_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_item_qty").ToString.Trim, "0")) & ", "

                    If rows.Item("pld_loc").ToString.Trim <> "" Then
                        Dim locSrch As String = ""
                        Dim locDt As New DataTable

                        locSrch += "select wh_code, fl_num, ar_code, rk_code, bn_code "
                        locSrch += "from wms_wh_bin "
                        locSrch += "where wh_code+replace(str(fl_num,2),' ','0')+"
                        locSrch += "replicate('0', 3 - len(ar_code))+ar_code+"
                        locSrch += "replicate('0', 4 - len(rk_code))+rk_code+"
                        locSrch += "replicate('0', 3 - len(bn_code))+bn_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "'"

                        locDt = gDB.getDataTable(locSrch)

                        If locDt.Rows.Count > 0 Then
                            itemSQL += "pld_loc = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, ""))) & ", "
                            itemSQL += "pld_wh = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(locDt.Rows(0).Item("wh_code").ToString.Trim, ""))) & ", "
                            itemSQL += "pld_floor = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(locDt.Rows(0).Item("fl_num").ToString.Trim, ""))) & ", "
                            itemSQL += "pld_area = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(locDt.Rows(0).Item("ar_code").ToString.Trim, ""))) & ", "
                            itemSQL += "pld_rack = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(locDt.Rows(0).Item("rk_code").ToString.Trim, ""))) & ", "
                            itemSQL += "pld_bin = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(locDt.Rows(0).Item("bn_code").ToString.Trim, ""))) & ", "
                        End If
                    End If


                    itemSQL += "sys_lub = " & gU.convdbNVCData(usrID) & ", "
                    itemSQL += "sys_lud = Getdate() "
                    itemSQL += "where imp_code = 'WMS' "
                    itemSQL += "and storer_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("storer_code").ToString.Trim, "")) & "' "
                    itemSQL += "and do_code = '" & gU.dbEncode(gU.decodeNull(rows.Item("do_code").ToString.Trim, "")) & "' "
                    itemSQL += "and pld_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_seq").ToString.Trim, "")) & "'"

                    If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                End If
            Next

            transaction.Commit()

            Return True

        Catch ex As Exception
            transaction.Rollback()
            Console.Write(ex.Message)
            Return False
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Function

    Public Function getGRList(ByVal storer As String) As DataTable Implements IWMSService.getGRList

    End Function

    Public Function getGRItemList(ByVal storer As String) As DataTable Implements IWMSService.getGRItemList

    End Function

    Public Function getDataTableFrmService(ByVal as_sql As String) As DataTable Implements IWMSService.getDataTableFrmService
        Try

            Dim dt As DataTable = gDB.getDataTable(as_sql)

            Return dt

        Catch ex As Exception
            Console.Write(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function getTableSchema(ByVal as_sql As String) As DataTable Implements IWMSService.getTableSchema
        Try

            Dim dt As New DataTable

            gDB.getDataTable(as_sql, , , dt)

            Return dt

        Catch ex As Exception
            Console.Write(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function getValueFromSQL(ByVal as_sql As String) As String Implements IWMSService.getValueFromSQL
        Try
            Dim value As String = ""


            value = DB.getValueFromSQL(as_sql)

            Return value

        Catch ex As Exception
            Console.Write(ex.Message)
            Return Nothing
        End Try
    End Function

    Private Sub SetCache()
        Dim value As String
        Dim settingKeys As String()
        Dim onRemove As CacheItemRemovedCallback
        Dim PageCache As Cache = HttpRuntime.Cache

        onRemove = New CacheItemRemovedCallback(AddressOf Me.RemovedCallback)

        settingKeys = System.Configuration.ConfigurationManager.AppSettings.AllKeys

        For x As Integer = 0 To settingKeys.Length - 1
            If PageCache(settingKeys(x)) Is Nothing Then
                value = System.Configuration.ConfigurationManager.AppSettings.Item(settingKeys(x))
                PageCache.Insert(settingKeys(x), value)
            End If
        Next

    End Sub

    Private itemRemoved As Boolean = False
    Private reason As CacheItemRemovedReason

    Public Sub RemovedCallback(ByVal k As String, ByVal v As Object, ByVal r As CacheItemRemovedReason)
        itemRemoved = True
        reason = r
    End Sub

End Class
