Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

Public Module GenericAppMod
    Private ctrlValueList As New Dictionary(Of String, String())
    Private gU As New GeneralUtils
    Private gDB As New GlobalDBFunc

    Private DDFORMAT As String = "DD/MM/YYYY"

    Enum ModuleType As Integer
        Search = 0
        Maintainence = 1
    End Enum

    Public Property ctrlValue() As Dictionary(Of String, String())
        Get
            Return ctrlValueList
        End Get

        Set(ByVal value As Dictionary(Of String, String()))
            ctrlValueList = value
        End Set
    End Property

    Public Function get_ReturnSQL(ByVal module_type As ModuleType, ByVal fun_code As String, _
                                    Optional ByRef srchSQL As String = "", _
                                    Optional ByRef whereSQL As String = "", _
                                    Optional ByRef criteriaSQL As String = "", _
                                    Optional ByRef groupbySQL As String = "", _
                                    Optional ByRef orderbySQL As String = "", _
                                    Optional ByRef FieldList As String = "") As String

        Dim value As String = ""
        Dim having_value As String = ""

        Select Case module_type
            Case 0
                REM========Search========
                Select Case fun_code

                    Case "RPT_SB"
                        Dim itx_date As String = HttpContext.Current.Session("SEARCH_SESSION_PAGE_S_ITX_DATE")

                        If itx_date <> "" Then
                            Dim appSQL As String = ""
                            Dim sto_code, itm_code, sku_no, p_key, itm_name, i_wh, i_floor, i_area, i_rack, i_bin As String

                            sto_code = HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_LOC_BAL_STORER_CODE")
                            itm_code = HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_LOC_BAL_ITM_CODE")
                            sku_no = HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_ITM_SKU_NO")
                            p_key = HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_LOC_BAL_PACK_KEY")
                            itm_name = HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_ITM_NAME")
                            i_wh = HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_LOC_BAL_ILOC_WH")
                            i_floor = HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_LOC_BAL_ILOC_FLOOR")
                            i_area = HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_LOC_BAL_ILOC_AREA")
                            i_rack = HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_LOC_BAL_ILOC_RACK")
                            i_bin = HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_LOC_BAL_ILOC_BIN")


                            If sto_code <> "" Then appSQL &= " AND tx1.storer_code='" & gU.dbEncode(sto_code) & "' "
                            If itm_code.Trim <> "" Then appSQL &= " AND upper(tx1.itm_code) like upper('%" & gU.dbEncode(itm_code) & "%')"
                            If sku_no.Trim <> "" Then appSQL &= " AND upper(wms_item.itm_sku_no) like upper('%" & gU.dbEncode(sku_no) & "%')"
                            If p_key.Trim <> "" Then appSQL &= " AND tx1.pack_key='" & gU.dbEncode(p_key) & "' "
                            If itm_name.Trim <> "" Then appSQL &= " AND upper(wms_item.itm_name) like upper('%" & gU.dbEncode(itm_name) & "%')"
                            If i_wh.Trim <> "" Then appSQL &= " AND upper(tx1.Itx_WH) like upper('%" & gU.dbEncode(i_wh) & "%')"
                            If i_floor.Trim <> "" Then appSQL &= " AND upper(tx1.Itx_FLOOR) like upper('%" & gU.dbEncode(i_floor) & "%')"
                            If i_area.Trim <> "" Then appSQL &= " AND upper(tx1.Itx_AREA) like upper('%" & gU.dbEncode(i_area) & "%')"
                            If i_rack.Trim <> "" Then appSQL &= " AND upper(tx1.Itx_RACK) like upper('%" & gU.dbEncode(i_rack) & "%')"
                            If i_bin.Trim <> "" Then appSQL &= " AND upper(tx1.ILOC_BIN) like upper('%" & gU.dbEncode(i_bin) & "%')"

                            value = " select tx1.STORER_CODE,tx1.ITM_CODE,tx1.PACK_KEY,tx1.ITX_PALLET_NO as ILOC_PALLET_NO, tx1.ITX_LOC,tx1.ITX_BAL_AFTER as ILOC_BAL_QTY, tx1.ITX_BATCH_NO,tx1.ITX_BATCH_NO as ILOC_BATCH_NO,  " & _
                                    " Convert(varchar, tx1.ITX_EXPIRY_DATE,103) as ILOC_EXPIRY_DATE,Convert(varchar, tx1.ITX_MANU_DATE,103) as ILOC_MANU_DATE, '" & gU.dbEncode(itx_date) & "' as itx_date," & _
                                    " tx1.Itx_WH  as ILOC_WH, tx1.Itx_FLOOR as ILOC_FLOOR,tx1.Itx_AREA as iloc_area, tx1.Itx_RACK as ILOC_RACK, tx1.Itx_BIN as ILOC_BIN,  " & _
                                    " wms_item.itm_name, wms_item.itm_desc, wms_item.itm_sku_no, " & _
                                    " t.aitm_qty_per_ctn, " & _
                                    " case when t.aitm_qty_per_ctn > 0 then ceiling(tx1.ITX_BAL_AFTER / t.aitm_qty_per_ctn) else 0 end as NO_OF_CARTON, " & _
                                    " tx1.ITX_BAL_AFTER * t.carton_cbm as TOTAL_carton_cbm " & _
                                    " from wms_item_loc_bal_tx tx1 inner join " & _
                                    " (select imp_code, storer_code, ITM_CODE,pack_key, ITX_PALLET_NO, ITX_LOC,ITX_BATCH_NO, max(ITX_DATE) as datekey from wms_item_loc_bal_tx where itx_date < Convert(datetime, '" & itx_date & "',103) + 1 " & _
                                    " group by imp_code, storer_code, ITM_CODE,pack_key, ITX_PALLET_NO, ITX_LOC,ITX_BATCH_NO) max1 " & _
                                    " on tx1.imp_code = max1.imp_code and tx1.storer_code = max1.storer_code and tx1.itm_code = max1.itm_code and tx1.pack_key = max1.pack_key  " & _
                                    " and ISNULL(tx1.itx_pallet_no,'000') = ISNULL(max1.ITX_PALLET_NO,'000') and ISNULL(tx1.ITX_BATCH_NO, '') = ISNULL(max1.ITX_BATCH_NO, '')  " & _
                                    " and tx1.itx_loc= max1.itx_loc and tx1.ITX_DATE = max1.datekey " & _
                                    " left outer join wms_item " & _
                                    " on wms_item.imp_code = tx1.imp_code and wms_item.storer_code = tx1.storer_code and wms_item.itm_code = tx1.itm_code and wms_item.pack_key = tx1.pack_key " & _
                                    " left outer join  V_ALT_VEND_ITEM t " & _
                                    " on WMS_ITEM.IMP_CODE = t.IMP_CODE  AND WMS_ITEM.STORER_CODE = t.STORER_CODE AND WMS_ITEM.ITM_CODE = t.ITM_CODE AND WMS_ITEM.PACK_KEY = t.PACK_KEY  " & _
                                    " Where tx1.ITX_BAL_AFTER > 0" & _
                                    appSQL

                            value &= " " & orderbySQL.Replace("WMS_ITEM_LOC_BAL.", "")
                        Else

                            value = " SELECT WMS_STORER.STO_SHORTNAME,Convert(varchar, Getdate(),103) as itx_date, " & _
                                    " WMS_ITEM_LOC_BAL.ILOC_FLOOR,WMS_ITEM_LOC_BAL.ILOC_AREA,WMS_ITEM_LOC_BAL.ILOC_RACK, WMS_ITEM_LOC_BAL.ILOC_BIN, " & _
                                    " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, " & _
                                    " WMS_ITEM_LOC_BAL.ILOC_BAL_CBM,WMS_ITEM_LOC_BAL.ILOC_BAL_KG, " & _
                                    " WMS_ITEM_LOC_BAL.ITM_CODE, " & _
                                    " WMS_ITEM.ITM_DESC,WMS_ITEM.itm_name, " & _
                                    " WMS_ITEM.ITM_SKU_NO, " & _
                                    " WMS_ITEM_LOC_BAL.PACK_KEY, " & _
                                    " WMS_ITEM_LOC_BAL.ILOC_BAL_QTY, " & _
                                    " WMS_ITEM_LOC_BAL.ILOC_BAL_QTY * ISNULL(WMS_ITEM.ITM_PCS_PER_UOM,1) as TOTAL_PCS, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, " & _
                                    " WMS_ITEM_LOC_BAL.ILOC_WH, " & _
                                    " t.aitm_qty_per_ctn, " & _
                                    " case when t.aitm_qty_per_ctn > 0 then ceiling(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY / t.aitm_qty_per_ctn) else 0 end as NO_OF_CARTON, " & _
                                    " WMS_ITEM_LOC_BAL.ILOC_BAL_QTY  * t.carton_cbm as TOTAL_carton_cbm, " & _
                                    " Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE,103) as ILOC_EXPIRY_DATE, Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE,103) as ILOC_MANU_DATE " & _
                                    " FROM WMS_ITEM_LOC_BAL " & _
                                    " INNER JOIN WMS_STORER ON " & _
                                    " WMS_ITEM_LOC_BAL.IMP_CODE = WMS_STORER.IMP_CODE " & _
                                    " AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_STORER.STORER_CODE " & _
                                    " LEFT OUTER JOIN WMS_ITEM ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE " & _
                                    " AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE " & _
                                    " AND WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE " & _
                                    " AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                    " LEFT OUTER JOIN V_ALT_VEND_ITEM t ON " & _
                                    " WMS_ITEM.IMP_CODE = t.IMP_CODE " & _
                                    " AND WMS_ITEM.STORER_CODE = t.STORER_CODE " & _
                                    " AND WMS_ITEM.ITM_CODE = t.ITM_CODE " & _
                                    " AND WMS_ITEM.PACK_KEY = t.PACK_KEY  " & _
                                    " WHERE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0 "

                            If whereSQL <> "" Then value &= " and " & whereSQL
                            value &= groupbySQL & " " & orderbySQL

                        End If

                    Case "LOOKUP_IM_BAL"
                        Dim selectYN As String = HttpContext.Current.Session("SEARCH_SESSION_PAGE_SELECT_ALL")
                        Dim tempStr As String = ""

                        If selectYN <> "Y" Then
                            tempStr = " AND (ISNULL(WMS_ITEM_LOC_BAL.TOTAL_BAL_QTY, 0) - ISNULL(HOLD_STOCK.HOLD_QTY, 0) - ISNULL(TOTAL_PICK.TOTAL_PICKED_QTY, 0)) > 0 "
                        End If

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("IS_Cable")) AndAlso HttpContext.Current.Session("IS_Cable") = "Y" OrElse HttpContext.Current.Session("SEARCH_SESSION_PAGE_ISCABLE") = "Y" Then
                            tempStr &= " AND WMS_ITEM.itm_type='CABLE' "
                        End If

                        Dim batchSQL As String = "WMS_ITEM_LOC_BAL.ILOC_BATCH_NO"

                        'If HttpContext.Current.Session("usr_type") = "C" OrElse HttpContext.Current.Session("usr_type") = "T" Then
                        '    batchSQL = " decode(substr(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,0,6),'@B#_E_','','@B#_M_','',WMS_ITEM_LOC_BAL.ILOC_BATCH_NO) as ILOC_BATCH_NO"
                        'End If
                        DDFORMAT = "103"

                        value = "SELECT WMS_ITEM.STORER_CODE as value, " & _
                                    "WMS_ITEM.ITM_CODE + '#_#' + WMS_ITEM.PACK_KEY AS VALUE1, " & _
                                    "isnull(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'') + '#_#' + isnull(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') AS VALUE2, " & _
                                    "WMS_ITEM_LOC_BAL.ILOC_WH, " & _
                                    "WMS_ITEM.IMP_CODE, " & _
                                    "WMS_ITEM.STORER_CODE, " & _
                                    "WMS_ITEM.ITM_CODE, " & _
                                    "WMS_ITEM.ITM_SKU_NO, " & _
                                    "WMS_ITEM.ITM_NAME,WMS_ITEM.ITM_UOM, " & _
                                    "WMS_ITEM.ITM_SERIES_NO, " & _
                                    "Convert(varchar,WMS_ITEM.SYS_CD, " & gU.dbEncode(DDFORMAT) & ") AS SYS_CD, " & _
                                    "WMS_ITEM.PACK_KEY, " & _
                                    "WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, " & _
                                    batchSQL & ", " & _
                                    "WMS_ITEM.ITM_DESC, " & _
                                    "WMS_ITEM_LOC_BAL.TOTAL_BAL_QTY, " & _
                                    "Convert(varchar,WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, " & gU.dbEncode(DDFORMAT) & ") AS ILOC_EXPIRY_DATE, " & _
                                    "Convert(varchar,WMS_ITEM_LOC_BAL.ILOC_MANU_DATE, " & gU.dbEncode(DDFORMAT) & ") AS ILOC_MANU_DATE, " & _
                                    "HOLD_STOCK.HOLD_QTY, " & _
                                    "TOTAL_PICK.TOTAL_PICKED_QTY, CO_ITEM.total_co_qty, " & _
                                    "case when ISNULL(WMS_ITEM_LOC_BAL.TOTAL_BAL_QTY, 0) - ISNULL(HOLD_STOCK.HOLD_QTY, 0) - ISNULL(TOTAL_PICK.TOTAL_PICKED_QTY, 0) < 0 then 0 else " & _
                                        "ISNULL(WMS_ITEM_LOC_BAL.TOTAL_BAL_QTY, 0) - ISNULL(HOLD_STOCK.HOLD_QTY, 0) - ISNULL(TOTAL_PICK.TOTAL_PICKED_QTY, 0) end as AVAIL_QTY " & _
                                "FROM WMS_ITEM " & _
                                "LEFT OUTER JOIN ( " & _
                                        "select imp_code, " & _
                                        "storer_code, " & _
                                        "ITM_CODE, " & _
                                        "PACK_KEY, " & _
                                        "ILOC_WH, " & _
                                        "ISNULL(ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, " & _
                                        "ISNULL(ILOC_BATCH_NO, '') AS ILOC_BATCH_NO, " & _
                                        "ILOC_EXPIRY_DATE as ILOC_EXPIRY_DATE, " & _
                                        "ILOC_MANU_DATE as ILOC_MANU_DATE, " & _
                                        "sum(ILOC_BAL_QTY) as TOTAL_BAL_QTY " & _
                                        "from WMS_ITEM_LOC_BAL " & _
                                        "where ISNULL(ILOC_BAL_QTY, 0) > 0 " & _
                                        "group by imp_code, storer_code, ITM_CODE, PACK_KEY, ILOC_WH, ISNULL(ILOC_PALLET_NO, '000'), ISNULL(ILOC_BATCH_NO, ''), ILOC_EXPIRY_DATE, ILOC_MANU_DATE " & _
                                        ") WMS_ITEM_LOC_BAL " & _
                                    "on WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE " & _
                                    "AND WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE " & _
                                    "AND WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE " & _
                                    "AND WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " & _
                                "LEFT OUTER JOIN ( " & _
                                        "select hold.imp_code, " & _
                                            "hold.storer_code, " & _
                                            "hold.COD_ITM_CODE, " & _
                                            "hold.COD_PACK_KEY, " & _
                                            "hold.COD_PALLET_NO, " & _
                                            "hold.COD_BATCH_NO, " & _
                                            "sum(hold.hold_qty - ISNULL(pick.picked_qty, 0)) as hold_qty " & _
                                        "from ( " & _
                                            "select h.imp_code, " & _
                                                "h.storer_code, " & _
                                                "h.co_code, " & _
                                                "h.COH_ITM_CODE AS COD_ITM_CODE, " & _
                                                "h.COH_PACK_KEY AS COD_PACK_KEY, " & _
                                                "ISNULL(h.COH_PALLET_NO, '000') as COD_PALLET_NO, " & _
                                                "isnull(h.COH_BATCH_NO, '') AS COD_BATCH_NO, " & _
                                                "sum(ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0)) as hold_qty " & _
                                            "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
                                            "where h.imp_code = d.imp_code " & _
                                            "and h.storer_code = d.storer_code " & _
                                            "and h.co_code = d.co_code " & _
                                            "and h.cod_seq = d.cod_seq " & _
                                            "and c.imp_code = d.imp_code " & _
                                            "and c.storer_code = d.storer_code " & _
                                            "and c.co_code = d.co_code " & _
                                            "and h.coh_status <> 'RELEASE' " & _
                                            "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
                                            "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " & _
                                            "group by h.imp_code, h.storer_code, h.co_code, h.COH_ITM_CODE, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000'), isnull(h.COH_BATCH_NO, '')) hold " & _
                                        "left outer join " & _
                                            "( " & _
                                            "select p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, sum(p.pld_item_qty) as picked_qty " & _
                                            "from wms_do_picklist_d p, wms_delv_order d " & _
                                            "where d.imp_code = p.imp_code " & _
                                            "and d.storer_code = p.storer_code " & _
                                            "and d.do_code = p.do_code " & _
                                            "and d.do_status = 'PICKED' " & _
                                            "group by p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, '') " & _
                                            ") pick " & _
                                        "on hold.imp_code = pick.imp_code " & _
                                        "and hold.storer_code = pick.storer_code " & _
                                        "and hold.co_code = pick.do_co_code " & _
                                        "and hold.COD_ITM_CODE = pick.pld_item_no " & _
                                        "and hold.COD_PACK_KEY = pick.pld_pack_key " & _
                                        "and hold.COD_PALLET_NO = pick.pld_pallet_no " & _
                                        "and hold.COD_BATCH_NO = pick.pld_batch_no " & _
                                        "group by hold.imp_code, hold.storer_code, hold.COD_ITM_CODE, hold.COD_PACK_KEY, hold.COD_PALLET_NO, hold.COD_BATCH_NO " & _
                                        ") HOLD_STOCK " & _
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE = HOLD_STOCK.IMP_CODE " & _
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = HOLD_STOCK.STORER_CODE " & _
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = HOLD_STOCK.COD_ITM_CODE " & _
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = HOLD_STOCK.COD_PACK_KEY " & _
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = HOLD_STOCK.COD_PALLET_NO " & _
                                    "AND WMS_ITEM_LOC_BAL.ILOC_BATCH_NO = HOLD_STOCK.COD_BATCH_NO " & _
                                "LEFT OUTER JOIN ( " & _
                                        "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, sum(p.pld_item_qty) as total_picked_qty " & _
                                        "from wms_do_picklist_d p, wms_delv_order d " & _
                                        "where d.imp_code = p.imp_code " & _
                                        "and d.storer_code = p.storer_code " & _
                                        "and d.do_code = p.do_code " & _
                                        "and d.do_status = 'PICKED' " & _
                                        "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, '')) TOTAL_PICK " & _
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE = TOTAL_PICK.IMP_CODE " & _
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = TOTAL_PICK.STORER_CODE " & _
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = TOTAL_PICK.PLD_ITEM_NO " & _
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = TOTAL_PICK.PLD_PACK_KEY " & _
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = TOTAL_PICK.PLD_PALLET_NO " & _
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = TOTAL_PICK.PLD_BATCH_NO " & _
                                "LEFT OUTER JOIN ( " & _
                                        "select imp_code, storer_code, cod_itm_code, cod_pack_key, cod_pallet_no, cod_batch_no, sum(cod_os_qty) as total_co_qty " & _
                                        "from ( " & _
                                        "select c2.imp_code, c2.storer_code, c2.co_code, c2.cod_itm_code, c2.cod_pack_key, ISNULL(c2.cod_pallet_no, '000') as cod_pallet_no, ISNULL(c2.cod_batch_no, '') as cod_batch_no, " & _
                                        "ISNULL(sum(c2.cod_qty), 0) - ISNULL(max(c2.cod_post_qty), 0) as cod_os_qty " & _
                                        "from WMS_CUST_ORDER_D c2, WMS_CUST_ORDER c1 " & _
                                        "where c1.imp_code = c2.imp_code " & _
                                        "and c1.storer_code = c2.storer_code " & _
                                        "and c1.co_code = c2.co_code " & _
                                        "and c1.co_status not in ('CANCELLED', 'CLOSED') " & _
                                        "group by c2.imp_code, c2.storer_code, c2.co_code, c2.cod_itm_code, c2.cod_pack_key, ISNULL(c2.cod_pallet_no, '000'), ISNULL(c2.cod_batch_no, '') " & _
                                        "having ISNULL(sum(c2.cod_qty), 0) - ISNULL(max(c2.cod_post_qty), 0) > 0 " & _
                                        ") subCO " & _
                                        "group by imp_code, storer_code, cod_itm_code, cod_pack_key, cod_pallet_no, cod_batch_no " & _
                                        ") CO_ITEM " & _
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE = CO_ITEM.IMP_CODE " & _
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = CO_ITEM.STORER_CODE " & _
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = CO_ITEM.COD_ITM_CODE " & _
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = CO_ITEM.COD_PACK_KEY " & _
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = CO_ITEM.COD_PALLET_NO " & _
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = CO_ITEM.COD_BATCH_NO " & _
                                "WHERE 2=2 " & tempStr


                        '"group by h.imp_code, h.storer_code, h.co_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO) hold " & _


                        '                "LEFT OUTER JOIN ( " & _
                        '"select h.imp_code, " & _
                        '"h.storer_code, " & _
                        '"d.COD_ITM_CODE, " & _
                        '"d.COD_PACK_KEY, " & _
                        '"ISNULL(d.COD_PALLET_NO, '000') as COD_PALLET_NO, " & _
                        '"d.COD_BATCH_NO, " & _
                        '"sum(ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0)) as hold_qty " & _
                        '"from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
                        '"where h.imp_code = d.imp_code " & _
                        '"and h.storer_code = d.storer_code " & _
                        '"and h.co_code = d.co_code " & _
                        '"and h.cod_seq = d.cod_seq " & _
                        '"and c.imp_code = d.imp_code " & _
                        '"and c.storer_code = d.storer_code " & _
                        '"and c.co_code = d.co_code " & _
                        '"and h.coh_status <> 'RELEASE' " & _
                        '"and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
                        '"and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " & _
                        '"group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO " & _
                        '") HOLD_STOCK " & _

                        If whereSQL <> "" Then value &= " and " & whereSQL

                        value &= groupbySQL & " " & orderbySQL



                    Case "INQ_001", "INQ_001R"
                        Dim batchSQL As String = "WMS_ITEM_LOC_BAL.ILOC_BATCH_NO"

                        'If HttpContext.Current.Session("usr_type") = "C" OrElse HttpContext.Current.Session("usr_type") = "T" Then
                        '    batchSQL = " decode(substr(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,0,6),'@B#_E_','','@B#_M_','',WMS_ITEM_LOC_BAL.ILOC_BATCH_NO) as ILOC_BATCH_NO"
                        'End If

                        Dim unionSQL As String = ""

                        value = "SELECT WMS_STORER.STO_SHORTNAME,  WMS_ITEM.ITM_SKU_NO, WMS_WAREHOUSE.WH_MAIN_WH, " &
                                    "WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, " &
                                    "WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC, " &
                                    "WMS_ITEM_LOC_BAL.ILOC_WH, WMS_ITEM_LOC_BAL.ILOC_FLOOR, WMS_ITEM_LOC_BAL.ILOC_AREA, " &
                                    "WMS_ITEM_LOC_BAL.ILOC_RACK, WMS_ITEM_LOC_BAL.ILOC_BIN, WMS_ITEM_LOC_BAL.ILOC_BAL_QTY, " &
                                    "ISNULL(WMS_ITEM.ITM_PCS_PER_UOM ,1) as ITM_PCS_PER_UOM, (ISNULL(WMS_ITEM.ITM_PCS_PER_UOM ,1) * WMS_ITEM_LOC_BAL.ILOC_BAL_QTY ) as TOTAL_NO, " &
                                    "WMS_ITEM_LOC_BAL.ILOC_BAL_CBM, WMS_ITEM_LOC_BAL.ILOC_BAL_KG, " &
                                    batchSQL & ", " &
                                    "Convert(varchar, ILOC_EXPIRY_DATE, 103) AS ILOC_EXPIRY_DATE, " &
                                    "Convert(varchar, ILOC_MANU_DATE, 103) AS ILOC_MANU_DATE, " &
                                    "WMS_ITEM_LOC_BAL.VND_CODE, " &
                                    "WMS_item.itm_name, WMS_ITEM.ITM_TEMP_FR, WMS_ITEM.ITM_TEMP_TO, " &
                                    "ISNULL(WMS_WAREHOUSE.WH_NAME, WMS_ITEM_LOC_BAL.ILOC_WH) as WH_NAME, " &
                                    "ISNULL(WMS_WH_FL.FL_NAME, WMS_ITEM_LOC_BAL.ILOC_FLOOR) as FL_NAME, " &
                                    "ISNULL(WMS_WH_AREA.AR_NAME, WMS_ITEM_LOC_BAL.ILOC_AREA) as AR_NAME, " &
                                    "ISNULL(WMS_WH_RACK.RK_NAME, WMS_ITEM_LOC_BAL.ILOC_RACK) as RK_NAME, " &
                                    "TOTAL_BAL.TOTAL_BAL_QTY, HOLD_STOCK.HOLD_QTY, " &
                                    "PICK_ITEM.PICKED_QTY, TOTAL_PICK.TOTAL_PICKED_QTY, CO_ITEM.total_co_qty, " &
                                    "ISNULL(TOTAL_BAL.TOTAL_BAL_QTY, 0) - ISNULL(HOLD_STOCK.HOLD_QTY, 0) as STOCK_AVAIL_QTY, " &
                                    "ISNULL(TOTAL_BAL.TOTAL_BAL_QTY, 0) - ISNULL(HOLD_STOCK.HOLD_QTY, 0) - ISNULL(TOTAL_PICK.TOTAL_PICKED_QTY, 0) as AVAIL_QTY, " &
                                    "t.aitm_qty_per_ctn, " &
                                    "case when t.aitm_qty_per_ctn > 0 then ceiling(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY / t.aitm_qty_per_ctn) else 0 end as NO_OF_CARTON, " &
                                    "WMS_ITEM_LOC_BAL.ILOC_BAL_QTY * t.carton_cbm as TOTAL_carton_cbm, " &
                                    "WMS_ITEM_LOC_BAL.ILOC_BAL_QTY * t.AITM_VOL as TOTAL_carton_KG, " &
                                    "'' as ILBS_SERIAL_NO, NULL as ILBS_QTY2, '' as ILBS_DRUM_ID, " &
                                    " WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_EMB, " &
                                    " CONVERT(varchar, WMS_ITEM.ITM_PD_RCV_DATE, 103) as ITM_PD_RCV_DATE, WMS_ITEM.ITM_PD_CONTRACT_NO, WMS_ITEM.ITM_PD_ARR_NOTICE_NO, WMS_ITEM.ITM_PD_COND_OF_SPARES, WMS_ITEM.ITM_PD_ST_1_YEAR, WMS_ITEM.ITM_PD_ST_OVER_1_YEAR, " &
                                    " WMS_ITEM.ITM_DRAWING_NO, WMS_ITEM.ITM_SERIAL_NO_YN, WMS_ITEM.ITM_STACKABLE_YN, WMS_ITEM.ITM_INSP_YN, WMS_ITEM.ITM_SCRAP_YN,  " &
                                    " WMS_ITEM.ITM_NONSTOCK_YN, WMS_ITEM.ITM_DG_YN, WMS_ITEM.ITM_REQ_STORE_HUM_YN, WMS_ITEM.ITM_REQ_STORE_AIRCON_YN,  " &
                                    " WMS_ITEM.ITM_CHE_CLASS, WMS_ITEM.ITM_REQ_MSDS_YN, WMS_ITEM.ITM_WEIGHT_TYPE, WMS_ITEM.ITM_ORO_YN, WMS_ITEM.ITEM_PRICE_CLASS,  " &
                                    " WMS_ITEM_LOC_BAL.SYS_LUB, WMS_ITEM_LOC_BAL.SYS_LUD,WMS_ITEM_LOC_BAL.SYS_LUD, WMS_ITEM.ITM_ROP_APL, WMS_ITEM.ITM_ROP_LAM, WMS_ITEM.ITM_ROP_CABLE,WMS_ITEM.ITM_ROP_NP, " &
                                    " replace(replace(replace(replace(replace(replace(replace(replace(replace(WMS_ITEM.ITM_CHE_CLASS,'1F','Flammable'),'2E','Explosive'),'3O', 'Oxidizing'),'4H','Harmful'),'5T','Toxic'),'6C','Corrosive'),'7I','Irritant'),'8C','Carcinogen'),'NA','NA') as che_class, " &
                                    " gr_itm.total_rcv_qty,WMS_ITEM_LOC_BAL.ILOC_PO_NO, NUll as BORD_QTY, WMS_ITEM.ITM_CRITICAL_YN, WMS_ITEM.ITM_RESTRICTED_ITEM,'' as ROQ, " &
                                    " case when datediff(d,convert(date,WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE),convert(date, getdate())) > 3 then 'Y' else 'N' end as IS_OVERDUE, " &
                                    " insp.gri_insp_qty, WMS_ITEM.ITM_UOM, '' AS ILBS_UOM2 " &
                                "FROM WMS_ITEM_LOC_BAL " &
                                "INNER JOIN WMS_STORER " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_STORER.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE=WMS_STORER.STORER_CODE " &
                                "LEFT OUTER JOIN WMS_ITEM " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_ITEM.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE=WMS_ITEM.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE=WMS_ITEM.ITM_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY=WMS_ITEM.PACK_KEY " &
                                "LEFT OUTER JOIN V_ALT_VEND_ITEM t " &
                                    "ON WMS_ITEM.IMP_CODE = t.IMP_CODE " &
                                    "AND WMS_ITEM.STORER_CODE = t.STORER_CODE " &
                                    "AND WMS_ITEM.ITM_CODE = t.ITM_CODE " &
                                    "AND WMS_ITEM.PACK_KEY = t.PACK_KEY " &
                                "LEFT OUTER JOIN WMS_WAREHOUSE " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WAREHOUSE.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WAREHOUSE.WH_CODE " &
                                "LEFT OUTER JOIN WMS_WH_FL " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WH_FL.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_FL.WH_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_FLOOR=WMS_WH_FL.FL_NUM " &
                                "LEFT OUTER JOIN WMS_WH_AREA " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WH_AREA.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_AREA.WH_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_FLOOR=WMS_WH_AREA.FL_NUM " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_AREA=WMS_WH_AREA.AR_CODE " &
                                "LEFT OUTER JOIN WMS_WH_RACK " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WH_RACK.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_RACK.WH_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_FLOOR=WMS_WH_RACK.FL_NUM " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_AREA=WMS_WH_RACK.AR_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_RACK=WMS_WH_RACK.RK_CODE " &
                                "LEFT OUTER JOIN ( " &
                                        "select imp_code, storer_code, ITM_CODE, PACK_KEY, ISNULL(ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, " &
                                        "ISNULL(ILOC_BATCH_NO, '') AS ILOC_BATCH_NO, sum(ILOC_BAL_QTY) as TOTAL_BAL_QTY " &
                                        "from WMS_ITEM_LOC_BAL " &
                                        "where ISNULL(ILOC_BAL_QTY, 0) != 0 " &
                                        "group by imp_code, storer_code, ITM_CODE, PACK_KEY, ISNULL(ILOC_PALLET_NO, '000'), ISNULL(ILOC_BATCH_NO, '')) TOTAL_BAL " &
                                    "on WMS_ITEM_LOC_BAL.IMP_CODE=TOTAL_BAL.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE=TOTAL_BAL.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE=TOTAL_BAL.ITM_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY=TOTAL_BAL.PACK_KEY " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000')=TOTAL_BAL.ILOC_PALLET_NO " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '')=TOTAL_BAL.ILOC_BATCH_NO " &
                                "LEFT OUTER JOIN ( " &
                                        "select hold.imp_code, " &
                                            "hold.storer_code, " &
                                            "hold.COD_ITM_CODE, " &
                                            "hold.COD_PACK_KEY, " &
                                            "hold.COD_PALLET_NO, " &
                                            "hold.COD_BATCH_NO, " &
                                            "sum(hold.hold_qty - ISNULL(pick.picked_qty, 0)) as hold_qty " &
                                        "from ( " &
                                            "select h.imp_code, " &
                                                "h.storer_code, " &
                                                "h.co_code, " &
                                                "h.COH_ITM_CODE as COD_ITM_CODE, " &
                                                "h.COH_PACK_KEY as COD_PACK_KEY, " &
                                                "ISNULL(h.COH_PALLET_NO, '000') as COD_PALLET_NO, " &
                                                "ISNULL(h.COH_BATCH_NO, '') as COD_BATCH_NO, " &
                                                "sum(ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0)) as hold_qty " &
                                            "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " &
                                            "where h.imp_code = d.imp_code " &
                                            "and h.storer_code = d.storer_code " &
                                            "and h.co_code = d.co_code " &
                                            "and h.cod_seq = d.cod_seq " &
                                            "and c.imp_code = d.imp_code " &
                                            "and c.storer_code = d.storer_code " &
                                            "and c.co_code = d.co_code " &
                                            "and h.coh_status <> 'RELEASE' " &
                                            "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " &
                                            "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " &
                                            "group by h.imp_code, h.storer_code, h.co_code, h.COH_ITM_CODE, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000'), ISNULL(h.COH_BATCH_NO, '')) hold " &
                                        "left outer join " &
                                            "( " &
                                            "select p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, sum(p.pld_item_qty) as picked_qty " &
                                            "from wms_do_picklist_d p, wms_delv_order d " &
                                            "where d.imp_code = p.imp_code " &
                                            "and d.storer_code = p.storer_code " &
                                            "and d.do_code = p.do_code " &
                                            "and d.do_status = 'PICKED' " &
                                            "group by p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, '') " &
                                            ") pick " &
                                        "on hold.imp_code = pick.imp_code " &
                                        "and hold.storer_code = pick.storer_code " &
                                        "and hold.co_code = pick.do_co_code " &
                                        "and hold.COD_ITM_CODE = pick.pld_item_no " &
                                        "and hold.COD_PACK_KEY = pick.pld_pack_key " &
                                        "and hold.COD_PALLET_NO = pick.pld_pallet_no " &
                                        "and hold.COD_BATCH_NO = pick.pld_batch_no " &
                                        "group by hold.imp_code, hold.storer_code, hold.COD_ITM_CODE, hold.COD_PACK_KEY, hold.COD_PALLET_NO, hold.COD_BATCH_NO) HOLD_STOCK " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE=HOLD_STOCK.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE=HOLD_STOCK.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE=HOLD_STOCK.COD_ITM_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY=HOLD_STOCK.COD_PACK_KEY " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000')=HOLD_STOCK.COD_PALLET_NO " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '')=HOLD_STOCK.COD_BATCH_NO " &
                                "LEFT OUTER JOIN ( " &
                                        "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " &
                                        "from wms_do_picklist_d p, wms_delv_order d " &
                                        "where d.imp_code = p.imp_code " &
                                        "and d.storer_code = p.storer_code " &
                                        "and d.do_code = p.do_code " &
                                        "and d.do_status = 'PICKED' " &
                                        "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE = PICK_ITEM.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = PICK_ITEM.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = PICK_ITEM.PLD_ITEM_NO " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = PICK_ITEM.PLD_PACK_KEY " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " &
                                    "AND WMS_ITEM_LOC_BAL.ILOC_LOC = PICK_ITEM.PLD_LOC " &
                                "LEFT OUTER JOIN ( " &
                                        "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, sum(p.pld_item_qty) as total_picked_qty " &
                                        "from wms_do_picklist_d p, wms_delv_order d " &
                                        "where d.imp_code = p.imp_code " &
                                        "and d.storer_code = p.storer_code " &
                                        "and d.do_code = p.do_code " &
                                        "and d.do_status = 'PICKED' " &
                                        "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, '')) TOTAL_PICK " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE = TOTAL_PICK.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = TOTAL_PICK.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = TOTAL_PICK.PLD_ITEM_NO " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = TOTAL_PICK.PLD_PACK_KEY " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = TOTAL_PICK.PLD_PALLET_NO " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = TOTAL_PICK.PLD_BATCH_NO " &
                                "LEFT OUTER JOIN ( " &
                                        "select imp_code, storer_code, cod_itm_code, cod_pack_key, cod_pallet_no, cod_batch_no, sum(cod_os_qty) as total_co_qty " &
                                        "from ( " &
                                            "select c2.imp_code, c2.storer_code, c2.co_code, c2.cod_itm_code, c2.cod_pack_key, ISNULL(c2.cod_pallet_no, '000') as cod_pallet_no, ISNULL(c2.cod_batch_no, '') as cod_batch_no, " &
                                            "ISNULL(sum(c2.cod_qty), 0) - ISNULL(max(c2.cod_post_qty), 0) as cod_os_qty " &
                                            "from WMS_CUST_ORDER_D c2, WMS_CUST_ORDER c1 " &
                                            "where c1.imp_code = c2.imp_code " &
                                            "and c1.storer_code = c2.storer_code " &
                                            "and c1.co_code = c2.co_code " &
                                            "and c1.co_status not in ('CANCELLED', 'CLOSED') " &
                                            "group by c2.imp_code, c2.storer_code, c2.co_code, c2.cod_itm_code, c2.cod_pack_key, ISNULL(c2.cod_pallet_no, '000'), ISNULL(c2.cod_batch_no, '') " &
                                            "having ISNULL(sum(c2.cod_qty), 0) - ISNULL(max(c2.cod_post_qty), 0) > 0) tbCO " &
                                        "group by imp_code, storer_code, cod_itm_code, cod_pack_key, cod_pallet_no, cod_batch_no) CO_ITEM " &
                                    "ON WMS_ITEM_LOC_BAL.IMP_CODE = CO_ITEM.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = CO_ITEM.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = CO_ITEM.COD_ITM_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = CO_ITEM.COD_PACK_KEY " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = CO_ITEM.COD_PALLET_NO " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = CO_ITEM.COD_BATCH_NO " &
                                " LEFT OUTER JOIN (" &
                                    " SELECT WMS_GOODSRCV_D.IMP_CODE, WMS_GOODSRCV_D.STORER_CODE, ISNULL(WMS_GOODSRCV_D.GRD_PALLET_NO,'000') AS GRD_PALLET_NO, ISNULL(WMS_GOODSRCV_D.GRD_BATCH_NO, '') AS GRD_BATCH_NO, " &
                                        " WMS_GOODSRCV_D.GRD_ITM_CODE, WMS_GOODSRCV_D.GRD_PACK_KEY, " &
                                        " SUM(WMS_GOODSRCV_D.GRD_RCV_QTY) AS TOTAL_RCV_QTY " &
                                    " FROM WMS_GOODSRCV " &
                                    " INNER JOIN WMS_GOODSRCV_D " &
                                    " ON WMS_GOODSRCV.IMP_CODE = WMS_GOODSRCV_D.IMP_CODE AND  " &
                                        " WMS_GOODSRCV.STORER_CODE = WMS_GOODSRCV_D.STORER_CODE AND " &
                                        " WMS_GOODSRCV.GR_CODE = WMS_GOODSRCV_D.GR_CODE " &
                                    " WHERE WMS_GOODSRCV.GR_STATUS NOT IN ('CANCELLED','POSTED') " &
                                    " GROUP BY WMS_GOODSRCV_D.IMP_CODE, WMS_GOODSRCV_D.STORER_CODE, WMS_GOODSRCV_D.GRD_PALLET_NO, ISNULL(WMS_GOODSRCV_D.GRD_BATCH_NO, ''), " &
                                            " WMS_GOODSRCV_D.GRD_ITM_CODE, WMS_GOODSRCV_D.GRD_PACK_KEY) GR_ITM  " &
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = GR_ITM.IMP_CODE AND " &
                                    " WMS_ITEM_LOC_BAL.STORER_CODE = GR_ITM.STORER_CODE AND " &
                                    " WMS_ITEM_LOC_BAL.ITM_CODE = GR_ITM.GRD_ITM_CODE AND " &
                                    " WMS_ITEM_LOC_BAL.PACK_KEY = GR_ITM.GRD_PACK_KEY AND " &
                                    " ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = GR_ITM.GRD_PALLET_NO AND " &
                                    " ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = GR_ITM.GRD_BATCH_NO " &
                                "LEFT OUTER JOIN ( " &
                                        "select p.imp_code, p.storer_code, p.gra_itm_code, p.gra_pack_key, ISNULL(p.gra_pallet_no, '000') as gra_pallet_no, isnull(p.gra_batch_no, '') as gra_batch_no, p.gra_loc, sum(i.gri_insp_qty) as gri_insp_qty " &
                                        "from wms_goodsrcv_pa p, wms_goodsrcv_insp i " &
                                        "where p.imp_code = i.imp_code " &
                                        "and p.storer_code = i.storer_code " &
                                        "and p.gr_code = i.gr_code " &
                                        "and p.gra_itm_code = i.gri_itm_code " &
                                        "and p.gra_pack_key = i.gri_pack_key " &
                                        "and p.gra_pallet_no = i.gri_pallet_no " &
                                        "and p.gra_batch_no = i.gri_batch_no " &
                                        "and i.gri_status <> 'DONE' " &
                                        "group by p.imp_code, p.storer_code, p.gra_itm_code, p.gra_pack_key, ISNULL(p.gra_pallet_no, '000'), isnull(p.gra_batch_no, ''), p.gra_loc) insp " &
                                    "on WMS_ITEM_LOC_BAL.IMP_CODE = insp.IMP_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.STORER_CODE = insp.STORER_CODE " &
                                    "AND WMS_ITEM_LOC_BAL.ITM_CODE = insp.gra_itm_code " &
                                    "AND WMS_ITEM_LOC_BAL.PACK_KEY = insp.gra_pack_key " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = insp.gra_pallet_no " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = insp.gra_batch_no " &
                                    "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_LOC, '') = insp.gra_loc " &
                                "WHERE 2=2 AND WMS_ITEM_LOC_BAL.ILOC_BAL_QTY != 0 and ISNULL(wms_item.ITM_SERIAL_NO_YN,'N') <> 'Y' "

                        unionSQL = " SELECT WMS_STORER.STO_SHORTNAME, WMS_ITEM.ITM_SKU_NO,WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, " &
                                       " WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC,  " &
                                       " WMS_ITEM_LOC_BAL.ILOC_WH, WMS_ITEM_LOC_BAL.ILOC_FLOOR, WMS_ITEM_LOC_BAL.ILOC_AREA, WMS_ITEM_LOC_BAL.ILOC_RACK,  " &
                                       " WMS_ITEM_LOC_BAL.ILOC_BIN, 1 as ILOC_BAL_QTY, ISNULL(WMS_ITEM.ITM_PCS_PER_UOM, 1) AS ITM_PCS_PER_UOM,  " &
                                       " ISNULL(WMS_ITEM.ITM_PCS_PER_UOM, 1) * WMS_ITEM_LOC_BAL.ILOC_BAL_QTY AS TOTAL_NO, WMS_ITEM_LOC_BAL.ILOC_BAL_CBM,  " &
                                       " WMS_ITEM_LOC_BAL.ILOC_BAL_KG, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, CONVERT(varchar, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, 103)  " &
                                       " AS ILOC_EXPIRY_DATE, CONVERT(varchar, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE, 103) AS ILOC_MANU_DATE, WMS_ITEM_LOC_BAL.VND_CODE,  " &
                                       " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_TEMP_FR, WMS_ITEM.ITM_TEMP_TO, ISNULL(WMS_WAREHOUSE.WH_NAME, WMS_ITEM_LOC_BAL.ILOC_WH)  " &
                                       " AS WH_NAME, ISNULL(WMS_WH_FL.FL_NAME, WMS_ITEM_LOC_BAL.ILOC_FLOOR) AS FL_NAME, ISNULL(WMS_WH_AREA.AR_NAME,  " &
                                       " WMS_ITEM_LOC_BAL.ILOC_AREA) AS AR_NAME, ISNULL(WMS_WH_RACK.RK_NAME, WMS_ITEM_LOC_BAL.ILOC_RACK) AS RK_NAME,  " &
                                       " TOTAL_BAL.TOTAL_BAL_QTY, NULL as hold_qty, PICK_ITEM.picked_qty, PICK_ITEM.picked_qty as total_picked_qty, CO_ITEM.total_co_qty,  " &
                                       " ISNULL(TOTAL_BAL.TOTAL_BAL_QTY, 0) - 0 AS STOCK_AVAIL_QTY, ISNULL(TOTAL_BAL.TOTAL_BAL_QTY, 0)  " &
                                       " - 0 - ISNULL(TOTAL_PICK.total_picked_qty, 0) AS AVAIL_QTY, t.AITM_QTY_PER_CTN,  " &
                                       " CASE WHEN t.aitm_qty_per_ctn > 0 THEN ceiling(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY / t.aitm_qty_per_ctn) ELSE 0 END AS NO_OF_CARTON,  " &
                                       " WMS_ITEM_LOC_BAL.ILOC_BAL_QTY * t.CARTON_CBM AS TOTAL_carton_cbm, WMS_ITEM_LOC_BAL.ILOC_BAL_QTY * t.AITM_VOL AS TOTAL_carton_KG, " &
                                       " WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID, " &
                                       " WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_EMB, " &
                                       " CONVERT(varchar, WMS_ITEM.ITM_PD_RCV_DATE, 103) as ITM_PD_RCV_DATE, WMS_ITEM.ITM_PD_CONTRACT_NO, WMS_ITEM.ITM_PD_ARR_NOTICE_NO, WMS_ITEM.ITM_PD_COND_OF_SPARES, WMS_ITEM.ITM_PD_ST_1_YEAR, WMS_ITEM.ITM_PD_ST_OVER_1_YEAR, " &
                                       " WMS_ITEM.ITM_DRAWING_NO, WMS_ITEM.ITM_SERIAL_NO_YN, WMS_ITEM.ITM_STACKABLE_YN, WMS_ITEM.ITM_INSP_YN, WMS_ITEM.ITM_SCRAP_YN,  " &
                                       " WMS_ITEM.ITM_NONSTOCK_YN, WMS_ITEM.ITM_DG_YN, WMS_ITEM.ITM_REQ_STORE_HUM_YN, WMS_ITEM.ITM_REQ_STORE_AIRCON_YN,  " &
                                       " WMS_ITEM.ITM_CHE_CLASS, WMS_ITEM.ITM_REQ_MSDS_YN, WMS_ITEM.ITM_WEIGHT_TYPE, WMS_ITEM.ITM_ORO_YN, WMS_ITEM.ITEM_PRICE_CLASS,  " &
                                       " WMS_ITEM_LOC_BAL.SYS_LUB, WMS_ITEM_LOC_BAL.SYS_LUD, WMS_ITEM_LOC_BAL.SYS_LUD, WMS_ITEM.ITM_ROP_APL, WMS_ITEM.ITM_ROP_LAM, WMS_ITEM.ITM_ROP_CABLE,WMS_ITEM.ITM_ROP_NP, " &
                                       " replace(replace(replace(replace(replace(replace(replace(replace(replace(WMS_ITEM.ITM_CHE_CLASS,'1F','Flammable'),'2E','Explosive'),'3O', 'Oxidizing'),'4H','Harmful'),'5T','Toxic'),'6C','Corrosive'),'7I','Irritant'),'8C','Carcinogen'),'NA','NA') as che_class, " &
                                       " gr_itm.total_rcv_qty,WMS_ITEM_LOC_BAL.ILOC_PO_NO, NUll as BORD_QTY, WMS_ITEM.ITM_CRITICAL_YN, WMS_ITEM.ITM_RESTRICTED_ITEM,'' as ROQ, " &
                                       " case when datediff(d,convert(date,WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE),convert(date, getdate())) > 3 then 'Y' else 'N' end as IS_OVERDUE, " &
                                       " insp.gri_insp_qty, WMS_ITEM.ITM_UOM, WMS_ITEM_LOC_BAL_S.ILBS_UOM2 " &
                                   " FROM WMS_ITEM_LOC_BAL INNER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ  " &
                                   " INNER JOIN WMS_STORER ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_STORER.IMP_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.STORER_CODE = WMS_STORER.STORER_CODE LEFT OUTER JOIN " &
                                   " WMS_ITEM ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY LEFT OUTER JOIN " &
                                   " V_ALT_VEND_ITEM AS t ON WMS_ITEM.IMP_CODE = t.IMP_CODE AND WMS_ITEM.STORER_CODE = t.STORER_CODE AND  " &
                                   " WMS_ITEM.ITM_CODE = t.ITM_CODE AND WMS_ITEM.PACK_KEY = t.PACK_KEY LEFT OUTER JOIN " &
                                   " WMS_WAREHOUSE ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE LEFT OUTER JOIN " &
                                   " WMS_WH_FL ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_FL.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_FL.WH_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_FL.FL_NUM LEFT OUTER JOIN " &
                                   " WMS_WH_AREA ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE LEFT OUTER JOIN " &
                                   " WMS_WH_RACK ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_RACK.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_RACK.WH_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_RACK.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_RACK.AR_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ILOC_RACK = WMS_WH_RACK.RK_CODE LEFT OUTER JOIN " &
                                   " (SELECT IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ISNULL(ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, ISNULL(ILOC_BATCH_NO, '')  " &
                                   " AS ILOC_BATCH_NO, SUM(ILOC_BAL_QTY) AS TOTAL_BAL_QTY " &
                                   " FROM WMS_ITEM_LOC_BAL " &
                                   " WHERE (ISNULL(ILOC_BAL_QTY, 0) != 0) " &
                                   " GROUP BY IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ISNULL(ILOC_PALLET_NO, '000'), ISNULL(ILOC_BATCH_NO, '')) AS TOTAL_BAL ON  " &
                                   " WMS_ITEM_LOC_BAL.IMP_CODE = TOTAL_BAL.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = TOTAL_BAL.STORER_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ITM_CODE = TOTAL_BAL.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = TOTAL_BAL.PACK_KEY AND  " &
                                   " ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = TOTAL_BAL.ILOC_PALLET_NO AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '')  " &
                                   " = TOTAL_BAL.ILOC_BATCH_NO LEFT OUTER JOIN " &
                                   " (SELECT        p.IMP_CODE, p.STORER_CODE, p.PLD_ITEM_NO, p.PLD_PACK_KEY, ISNULL(p.PLD_PALLET_NO, '000') AS pld_pallet_no, ISNULL(p.PLD_BATCH_NO, '') AS pld_batch_no, p.PLD_LOC, SUM(p.PLD_ITEM_QTY) AS picked_qty, p.PLD_SERIAL_NO  " &
                                   " FROM WMS_DO_PICKLIST_D AS p INNER JOIN " &
                                   " WMS_DELV_ORDER AS d ON p.IMP_CODE = d.IMP_CODE AND p.STORER_CODE = d.STORER_CODE AND p.DO_CODE = d.DO_CODE " &
                                   " WHERE (d.DO_STATUS = 'PICKED') " &
                                   " GROUP BY p.IMP_CODE, p.STORER_CODE, p.PLD_ITEM_NO, p.PLD_PACK_KEY, ISNULL(p.PLD_PALLET_NO, '000'), ISNULL(p.PLD_BATCH_NO, ''), p.PLD_LOC,PLD_SERIAL_NO)  " &
                                   " AS PICK_ITEM ON WMS_ITEM_LOC_BAL.IMP_CODE = PICK_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = PICK_ITEM.STORER_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ITM_CODE = PICK_ITEM.PLD_ITEM_NO AND WMS_ITEM_LOC_BAL.PACK_KEY = PICK_ITEM.PLD_PACK_KEY AND  " &
                                   " ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = PICK_ITEM.pld_pallet_no AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '')  " &
                                   " = PICK_ITEM.pld_batch_no AND WMS_ITEM_LOC_BAL.ILOC_LOC = PICK_ITEM.PLD_LOC AND WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO = PICK_ITEM.PLD_SERIAL_NO  LEFT OUTER JOIN " &
                                   " (SELECT p.IMP_CODE, p.STORER_CODE, p.PLD_ITEM_NO, p.PLD_PACK_KEY, ISNULL(p.PLD_PALLET_NO, '000') AS pld_pallet_no, ISNULL(p.PLD_BATCH_NO, '') AS pld_batch_no, SUM(p.PLD_ITEM_QTY) AS total_picked_qty " &
                                   " FROM WMS_DO_PICKLIST_D AS p INNER JOIN " &
                                   " WMS_DELV_ORDER AS d ON p.IMP_CODE = d.IMP_CODE AND p.STORER_CODE = d.STORER_CODE AND p.DO_CODE = d.DO_CODE " &
                                   " WHERE        (d.DO_STATUS = 'PICKED') " &
                                   " GROUP BY p.IMP_CODE, p.STORER_CODE, p.PLD_ITEM_NO, p.PLD_PACK_KEY, ISNULL(p.PLD_PALLET_NO, '000'), ISNULL(p.PLD_BATCH_NO, ''))  " &
                                   " AS TOTAL_PICK ON WMS_ITEM_LOC_BAL.IMP_CODE = TOTAL_PICK.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = TOTAL_PICK.STORER_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ITM_CODE = TOTAL_PICK.PLD_ITEM_NO AND WMS_ITEM_LOC_BAL.PACK_KEY = TOTAL_PICK.PLD_PACK_KEY AND  " &
                                   " ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = TOTAL_PICK.pld_pallet_no AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '')  " &
                                   " = TOTAL_PICK.pld_batch_no LEFT OUTER JOIN " &
                                   " (SELECT IMP_CODE, STORER_CODE, COD_ITM_CODE, COD_PACK_KEY, cod_pallet_no, cod_batch_no, SUM(cod_os_qty) AS total_co_qty " &
                                   " FROM (SELECT c2.IMP_CODE, c2.STORER_CODE, c2.CO_CODE, c2.COD_ITM_CODE, c2.COD_PACK_KEY, ISNULL(c2.COD_PALLET_NO, '000')  " &
                                   "  AS cod_pallet_no, ISNULL(c2.COD_BATCH_NO, '') AS cod_batch_no, ISNULL(SUM(c2.COD_QTY), 0) - ISNULL(MAX(c2.COD_POST_QTY), 0) AS cod_os_qty " &
                                   " FROM WMS_CUST_ORDER_D AS c2 INNER JOIN " &
                                   " WMS_CUST_ORDER AS c1 ON c2.IMP_CODE = c1.IMP_CODE AND c2.STORER_CODE = c1.STORER_CODE AND  " &
                                   " c2.CO_CODE = c1.CO_CODE " &
                                   " WHERE (c1.CO_STATUS NOT IN ('CANCELLED', 'CLOSED')) " &
                                   " GROUP BY c2.IMP_CODE, c2.STORER_CODE, c2.CO_CODE, c2.COD_ITM_CODE, c2.COD_PACK_KEY, ISNULL(c2.COD_PALLET_NO, '000'),  " &
                                   " ISNULL(c2.COD_BATCH_NO, '') " &
                                   " HAVING         (ISNULL(SUM(c2.COD_QTY), 0) - ISNULL(MAX(c2.COD_POST_QTY), 0) > 0)) AS tbCO " &
                                   " GROUP BY IMP_CODE, STORER_CODE, COD_ITM_CODE, COD_PACK_KEY, cod_pallet_no, cod_batch_no) AS CO_ITEM ON  " &
                                   " WMS_ITEM_LOC_BAL.IMP_CODE = CO_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = CO_ITEM.STORER_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ITM_CODE = CO_ITEM.COD_ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = CO_ITEM.COD_PACK_KEY AND  " &
                                   " ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = CO_ITEM.cod_pallet_no AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = CO_ITEM.cod_batch_no " &
                                   " LEFT OUTER JOIN (SELECT WMS_GOODSRCV_D.IMP_CODE, WMS_GOODSRCV_D.STORER_CODE, ISNULL(WMS_GOODSRCV_D.GRD_PALLET_NO,'000') AS GRD_PALLET_NO, ISNULL(WMS_GOODSRCV_D.GRD_BATCH_NO, '') AS GRD_BATCH_NO, WMS_GOODSRCV_D.GRD_ITM_CODE, WMS_GOODSRCV_D.GRD_PACK_KEY,  " &
                                   " SUM(WMS_GOODSRCV_D.GRD_RCV_QTY) AS TOTAL_RCV_QTY " &
                                   " FROM WMS_GOODSRCV INNER JOIN " &
                                   " WMS_GOODSRCV_D ON WMS_GOODSRCV.IMP_CODE = WMS_GOODSRCV_D.IMP_CODE AND  " &
                                   " WMS_GOODSRCV.STORER_CODE = WMS_GOODSRCV_D.STORER_CODE AND WMS_GOODSRCV.GR_CODE = WMS_GOODSRCV_D.GR_CODE " &
                                   " WHERE WMS_GOODSRCV.GR_STATUS NOT IN ('CANCELLED','POSTED') " &
                                   " GROUP BY WMS_GOODSRCV_D.IMP_CODE, WMS_GOODSRCV_D.STORER_CODE,WMS_GOODSRCV_D.GRD_PALLET_NO, ISNULL(WMS_GOODSRCV_D.GRD_BATCH_NO, ''), WMS_GOODSRCV_D.GRD_ITM_CODE, WMS_GOODSRCV_D.GRD_PACK_KEY) GR_ITM  " &
                                   " ON WMS_ITEM_LOC_BAL.IMP_CODE = GR_ITM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = GR_ITM.STORER_CODE AND  " &
                                   " WMS_ITEM_LOC_BAL.ITM_CODE = GR_ITM.GRD_ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = GR_ITM.GRD_PACK_KEY AND  " &
                                   " ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = GR_ITM.GRD_PALLET_NO AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = GR_ITM.GRD_BATCH_NO " &
                                   "LEFT OUTER JOIN ( " &
                                            "select p.imp_code, p.storer_code, p.gra_itm_code, p.gra_pack_key, ISNULL(p.gra_pallet_no, '000') as gra_pallet_no, isnull(p.gra_batch_no, '') as gra_batch_no, p.gra_loc, sum(i.gri_insp_qty) as gri_insp_qty " &
                                            "from wms_goodsrcv_pa p, wms_goodsrcv_insp i " &
                                            "where p.imp_code = i.imp_code " &
                                            "and p.storer_code = i.storer_code " &
                                            "and p.gr_code = i.gr_code " &
                                            "and p.gra_itm_code = i.gri_itm_code " &
                                            "and p.gra_pack_key = i.gri_pack_key " &
                                            "and p.gra_pallet_no = i.gri_pallet_no " &
                                            "and p.gra_batch_no = i.gri_batch_no " &
                                            "and i.gri_status <> 'DONE' " &
                                            "group by p.imp_code, p.storer_code, p.gra_itm_code, p.gra_pack_key, ISNULL(p.gra_pallet_no, '000'), isnull(p.gra_batch_no, ''), p.gra_loc) insp " &
                                        "on WMS_ITEM_LOC_BAL.IMP_CODE = insp.IMP_CODE " &
                                        "AND WMS_ITEM_LOC_BAL.STORER_CODE = insp.STORER_CODE " &
                                        "AND WMS_ITEM_LOC_BAL.ITM_CODE = insp.gra_itm_code " &
                                        "AND WMS_ITEM_LOC_BAL.PACK_KEY = insp.gra_pack_key " &
                                        "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000') = insp.gra_pallet_no " &
                                        "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = insp.gra_batch_no " &
                                        "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_LOC, '') = insp.gra_loc " &
                                   " WHERE (2 = 2) AND (WMS_ITEM_LOC_BAL.ILOC_BAL_QTY != 0) AND ((WMS_ITEM.ITM_TYPE='CABLE' AND WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0) OR (ISNULL(WMS_ITEM.ITM_TYPE,'N') <> 'CABLE' AND WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0)) AND (WMS_ITEM.ITM_SERIAL_NO_YN  = 'Y') "

                        '"select h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000') as COD_PALLET_NO, d.COD_BATCH_NO, sum(ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0)) as hold_qty " & _
                        '"from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
                        '"where h.imp_code=d.imp_code " & _
                        '"and h.storer_code=d.storer_code " & _
                        '"and h.co_code=d.co_code " & _
                        '"and h.cod_seq=d.cod_seq " & _
                        '"and c.imp_code=d.imp_code " & _
                        '"and c.storer_code=d.storer_code " & _
                        '"and c.co_code=d.co_code " & _
                        '"and h.coh_status <> 'RELEASE' " & _
                        '"and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
                        '"and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " & _
                        '"group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO) HOLD_STOCK " & _

                        'value = " SELECT WMS_STORER.STO_SHORTNAME,  WMS_ITEM.ITM_SKU_NO," & _
                        '        " WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE,  " & _
                        '        " WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC,  " & _
                        '        " WMS_ITEM_LOC_BAL.ILOC_WH, WMS_ITEM_LOC_BAL.ILOC_FLOOR, WMS_ITEM_LOC_BAL.ILOC_AREA,  " & _
                        '        " WMS_ITEM_LOC_BAL.ILOC_RACK, WMS_ITEM_LOC_BAL.ILOC_BIN, WMS_ITEM_LOC_BAL.ILOC_BAL_QTY,  " & _
                        '        " ISNULL(WMS_ITEM.ITM_PCS_PER_UOM ,1) as ITM_PCS_PER_UOM, (ISNULL(WMS_ITEM.ITM_PCS_PER_UOM ,1) * WMS_ITEM_LOC_BAL.ILOC_BAL_QTY ) as TOTAL_NO, " & _
                        '        " WMS_ITEM_LOC_BAL.ILOC_BAL_CBM, WMS_ITEM_LOC_BAL.ILOC_BAL_KG,  " & _
                        '        " WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,  " & _
                        '        " TO_CHAR(ILOC_EXPIRY_DATE, 'DD/MM/YYYY') AS ILOC_EXPIRY_DATE,  " & _
                        '        " TO_CHAR(ILOC_MANU_DATE, 'DD/MM/YYYY') AS ILOC_MANU_DATE,  " & _
                        '        " WMS_ITEM_LOC_BAL.VND_CODE, " & _
                        '        " WMS_item.itm_name, WMS_ITEM.ITM_TEMP_FR, WMS_ITEM.ITM_TEMP_TO, " & _
                        '        " ISNULL(WMS_WAREHOUSE.WH_NAME, WMS_ITEM_LOC_BAL.ILOC_WH) as WH_NAME, " & _
                        '        " ISNULL(WMS_WH_FL.FL_NAME, WMS_ITEM_LOC_BAL.ILOC_FLOOR) as FL_NAME, " & _
                        '        " ISNULL(WMS_WH_AREA.AR_NAME, WMS_ITEM_LOC_BAL.ILOC_AREA) as AR_NAME, " & _
                        '        " ISNULL(WMS_WH_RACK.RK_NAME, WMS_ITEM_LOC_BAL.ILOC_RACK) as RK_NAME, " & _
                        '        " TOTAL_BAL.TOTAL_BAL_QTY, HOLD_STOCK.HOLD_QTY, " & _
                        '        " ISNULL(TOTAL_BAL.TOTAL_BAL_QTY, 0) - ISNULL(HOLD_STOCK.HOLD_QTY, 0) as AVAIL_QTY, " & _
                        '        " t.aitm_qty_per_ctn, " & _
                        '        " case when t.aitm_qty_per_ctn > 0 then ceil(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY / t.aitm_qty_per_ctn) else 0 end as NO_OF_CARTON, " & _
                        '        " case when t.aitm_qty_per_ctn > 0 then ceil(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY / t.aitm_qty_per_ctn) * t.carton_cbm else 0 end as TOTAL_carton_cbm, " & _
                        '        " case when t.aitm_qty_per_ctn > 0 then ceil(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY / t.aitm_qty_per_ctn) * t.AITM_VOL else 0 end as TOTAL_carton_KG " & _
                        '        " FROM WMS_ITEM_LOC_BAL INNER JOIN WMS_STORER ON " & _
                        '        " WMS_ITEM_LOC_BAL.IMP_CODE=WMS_STORER.IMP_CODE AND " & _
                        '        " WMS_ITEM_LOC_BAL.STORER_CODE=WMS_STORER.STORER_CODE " & _
                        '        " LEFT OUTER JOIN WMS_ITEM ON " & _
                        '        " WMS_ITEM_LOC_BAL.IMP_CODE=WMS_ITEM.IMP_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.STORER_CODE=WMS_ITEM.STORER_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ITM_CODE=WMS_ITEM.ITM_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.PACK_KEY=WMS_ITEM.PACK_KEY " & _
                        '        " LEFT OUTER JOIN V_ALT_VEND_ITEM t ON " & _
                        '        " WMS_ITEM.IMP_CODE = t.IMP_CODE " & _
                        '        " AND WMS_ITEM.STORER_CODE = t.STORER_CODE " & _
                        '        " AND WMS_ITEM.ITM_CODE = t.ITM_CODE " & _
                        '        " AND WMS_ITEM.PACK_KEY = t.PACK_KEY " & _
                        '        " LEFT OUTER JOIN WMS_WAREHOUSE ON " & _
                        '        " WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WAREHOUSE.IMP_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WAREHOUSE.WH_CODE " & _
                        '        " LEFT OUTER JOIN WMS_WH_FL ON " & _
                        '        " WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WH_FL.IMP_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_FL.WH_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_FLOOR=WMS_WH_FL.FL_NUM " & _
                        '        " LEFT OUTER JOIN WMS_WH_AREA ON " & _
                        '        " WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WH_AREA.IMP_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_AREA.WH_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_FLOOR=WMS_WH_AREA.FL_NUM " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_AREA=WMS_WH_AREA.AR_CODE " & _
                        '        " LEFT OUTER JOIN WMS_WH_RACK ON " & _
                        '        " WMS_ITEM_LOC_BAL.IMP_CODE=WMS_WH_RACK.IMP_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_WH=WMS_WH_RACK.WH_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_FLOOR=WMS_WH_RACK.FL_NUM " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_AREA=WMS_WH_RACK.AR_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_RACK=WMS_WH_RACK.RK_CODE " & _
                        '        " LEFT OUTER JOIN ( " & _
                        '        " select imp_code, storer_code, ITM_CODE, PACK_KEY, ISNULL(ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, " & _
                        '        " ISNULL(ILOC_BATCH_NO, '') AS ILOC_BATCH_NO, sum(ILOC_BAL_QTY) as TOTAL_BAL_QTY " & _
                        '        " from WMS_ITEM_LOC_BAL " & _
                        '        " where ISNULL(ILOC_BAL_QTY, 0) > 0 " & _
                        '        " group by imp_code, storer_code, ITM_CODE, PACK_KEY, ISNULL(ILOC_PALLET_NO, '000'), ISNULL(ILOC_BATCH_NO, '')) TOTAL_BAL " & _
                        '        " on WMS_ITEM_LOC_BAL.IMP_CODE=TOTAL_BAL.IMP_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.STORER_CODE=TOTAL_BAL.STORER_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ITM_CODE=TOTAL_BAL.ITM_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.PACK_KEY=TOTAL_BAL.PACK_KEY " & _
                        '        " AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000')=TOTAL_BAL.ILOC_PALLET_NO " & _
                        '        " AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '')=TOTAL_BAL.ILOC_BATCH_NO " & _
                        '        " LEFT OUTER JOIN (  " & _
                        '        " select h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000') as COD_PALLET_NO, d.COD_BATCH_NO, sum(ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0)) as hold_qty " & _
                        '        " from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
                        '        " where h.imp_code=d.imp_code " & _
                        '        " and h.storer_code=d.storer_code " & _
                        '        " and h.co_code=d.co_code " & _
                        '        " and h.cod_seq=d.cod_seq " & _
                        '        " and c.imp_code=d.imp_code " & _
                        '        " and c.storer_code=d.storer_code " & _
                        '        " and c.co_code=d.co_code " & _
                        '        " and h.coh_status <> 'RELEASE' " & _
                        '        " and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
                        '        " and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " & _
                        '        " group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO) HOLD_STOCK " & _
                        '        " ON WMS_ITEM_LOC_BAL.IMP_CODE=HOLD_STOCK.IMP_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.STORER_CODE=HOLD_STOCK.STORER_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.ITM_CODE=HOLD_STOCK.COD_ITM_CODE " & _
                        '        " AND WMS_ITEM_LOC_BAL.PACK_KEY=HOLD_STOCK.COD_PACK_KEY " & _
                        '        " AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '000')=HOLD_STOCK.COD_PALLET_NO " & _
                        '        " AND WMS_ITEM_LOC_BAL.ILOC_BATCH_NO=HOLD_STOCK.COD_BATCH_NO " & _
                        '        " WHERE 2=2 AND WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0 "

                        If whereSQL <> "" Then
                            value &= " and " & whereSQL
                            unionSQL &= " and " & whereSQL
                        End If


                        value &= " UNION " & unionSQL & " " & orderbySQL
                    Case "LOOKUP_DOWP"
                        If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                        value = srchSQL & " WHERE 1=1 " & criteriaSQL
                        If whereSQL <> "" Then value &= " and " & whereSQL

                        Dim tempSQL As String = ""
                        Dim tempSTR As String = ""

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("selectedDO")) Then
                            tempSQL = " AND WMS_DO_PICKLIST_D.DO_CODE not in("

                            Dim doArr As String()
                            Dim doStr As String = HttpContext.Current.Session("selectedDO")

                            doArr = gU.listToArray(doStr)


                            For i = 0 To doArr.Length - 1
                                tempSTR &= "'" & doArr(i) & "',"
                            Next

                            tempSTR = Left(tempSTR, Len(tempSTR) - 1)

                            tempSQL &= tempSTR & ") "

                        End If

                        If tempSQL <> "" Then value &= tempSQL

                        value &= groupbySQL & " " & orderbySQL

                    Case "LOOKUP_STKCHK_BAL"

                        Dim selectSQL As String = ""
                        Dim tempStr As String = ""
                        Dim unionSQL As String = ""
                        Dim areaSQL As String = ""
                        Dim orgQTYSQL As String = ""
                        Dim orgQTYSQL2 As String = ""
                        Dim tempWHStr As String = ""
                        Dim tempWHSQLStr As String = ""
                        Dim binValue1 As String = ""
                        Dim binValue2 As String = ""
                        Dim binListValue As String = ""
                        Dim otherItemValue As String = ""

                        Dim CCSTempStr As String = ""
                        Dim CCDateTempStr As String = ""
                        Dim CCIsCableTempStr As String = ""
                        Dim CCWHTempStr As String = ""
                        Dim CCEXCLUDE_OTHERTempStr As String = ""
                        Dim CCSHOW_OTHER_AREATempStr As String = ""
                        Dim CCbinTempStr As String = ""

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("CK_TYPE")) Then
                            If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("CK_START_DATE")) AndAlso gU.isValidDate(HttpContext.Current.Session("CK_START_DATE")) AndAlso HttpContext.Current.Session("CK_TYPE") = "CC" Then
                                CCDateTempStr &= " AND (WMS_ITEM.ITM_CC_DATE is null or WMS_ITEM.ITM_CC_DATE < convert(date,'" & gU.dbEncode(HttpContext.Current.Session("CK_START_DATE")) & "',103)) "
                            Else
                                CCDateTempStr &= " and (WMS_ITEM.ITM_CC_DATE is NULL OR dateadd(YEAR,isnull(cast(col.colc_eng_value as int),1), cast(WMS_ITEM.ITM_CC_DATE AS DATE)) < cast(getdate() as date)) "
                            End If

                            orgQTYSQL &= " Case when WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' then 1 else WMS_ITEM_LOC_BAL.ILOC_BAL_QTY end as ILOC_BAL_QTY,"
                            orgQTYSQL2 &= " WMS_ITEM_LOC_BAL_S.ILBS_QTY2, "
                        End If

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_AREA_CODE")) Then
                            areaSQL = " AND WMS_ITEM_LOC_BAL.ILOC_LOC like '" & gU.dbEncode(HttpContext.Current.Session("SEARCH_SESSION_PAGE_AREA_CODE")) & "%' "
                        End If

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_FLOOR_CODE")) Then
                            areaSQL = " AND WMS_ITEM_LOC_BAL.ILOC_LOC like '" & gU.dbEncode(HttpContext.Current.Session("SEARCH_SESSION_PAGE_FLOOR_CODE")) & "%' "
                        End If


                        If Not HttpContext.Current.Session("CK_WH") Is Nothing AndAlso HttpContext.Current.Session("CK_WH") <> "" Then
                            Dim tempSelect As String = " Select distinct wh_code from wms_warehouse where WH_MAIN_WH='" & gU.dbEncode(HttpContext.Current.Session("CK_WH")) & "'"

                            If Not HttpContext.Current.Session("LookupWH") Is Nothing AndAlso HttpContext.Current.Session("LookupWH") <> "" Then
                                tempSelect &= " AND WH_CODE='" & gU.dbEncode(HttpContext.Current.Session("LookupWH")) & "'"
                            End If

                            Dim tempDT As System.Data.DataTable = gDB.getDataTable(tempSelect)

                            If tempDT.Rows.Count > 0 Then
                                For i = 0 To tempDT.Rows.Count - 1
                                    tempWHStr &= "'" & tempDT.Rows(i).Item("wh_code").ToString.Trim & "',"
                                Next

                                If tempWHStr <> "" Then
                                    tempWHStr = Left(tempWHStr, Len(tempWHStr) - 1)
                                    CCWHTempStr = " AND WMS_ITEM_LOC_BAL.ILOC_WH in(" & tempWHStr & ") "
                                End If
                            End If

                        End If

                        Dim TopStr As String = ""
                        Dim tempStr1 As String = ""

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("selectedSEQ")) Then
                            Dim keyArr As String()
                            keyArr = gU.listToArray(HttpContext.Current.Session("selectedSEQ"))

                            Dim tempKeySQL As String = ""

                            For i = 0 To keyArr.Length - 1
                                tempKeySQL &= "'" & keyArr(i) & "',"
                            Next

                            If tempKeySQL <> "" Then
                                tempKeySQL = Left(tempKeySQL, Len(tempKeySQL) - 1)
                                tempStr1 &= " AND WMS_ITEM_LOC_BAL.STORER_CODE + '|*|' + WMS_ITEM_LOC_BAL.ITM_CODE + '|*|' + WMS_ITEM_LOC_BAL.PACK_KEY + '|*|' + isnull(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') + '|*|' +  isnull(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,'') + '|*|' + isnull(WMS_ITEM_LOC_BAL.ILOC_LOC,'') " & _
                                       " not in(" & tempKeySQL & ") "
                            End If
                        End If

                        tempStr = CCSTempStr & CCDateTempStr & CCIsCableTempStr & CCWHTempStr & CCEXCLUDE_OTHERTempStr & CCSHOW_OTHER_AREATempStr & areaSQL & CCbinTempStr & tempStr1
                        If whereSQL <> "" Then tempStr &= " and " & whereSQL

                        Dim updateSQL As String = ""
                        Dim tmpKey As String = ""


                        value = " SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE,WMS_ITEM_LOC_BAL.ILOC_BAL_QTY,WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, " & orgQTYSQL2 & " WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID," &
                                " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO," & orgQTYSQL & " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, " &
                                " WMS_ITEM.ITM_CC_DATE, WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_GP_CODE, " &
                                " WMS_ITEM.ITM_CODE + '#_#' + WMS_ITEM.PACK_KEY AS VALUE1, " &
                                " isnull(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') + '#_#' + isnull(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') AS VALUE2, Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'')  as VALUE, WMS_ITEM_LOC_BAL.ILOC_SEQ " &
                                " FROM WMS_ITEM_LOC_BAL INNER JOIN " &
                                " WMS_ITEM ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " &
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " &
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " &
                                " LEFT OUTER JOIN WMS_COL_CODE col on WMS_ITEM.ITEM_PRICE_CLASS = col.COLC_CODE and col.COLC_TABCOL='PRICE_CLASS_VALUE' " &
                                " WHERE WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' AND WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0 " & tempStr &
                                " UNION " &
                                " SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE,WMS_ITEM_LOC_BAL.ILOC_BAL_QTY,WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, " & orgQTYSQL2 & " WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID," &
                                " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO," & orgQTYSQL & " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, " &
                                " WMS_ITEM.ITM_CC_DATE, WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_GP_CODE, " &
                                " WMS_ITEM.ITM_CODE + '#_#' + WMS_ITEM.PACK_KEY AS VALUE1, " &
                                " isnull(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') + '#_#' + isnull(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') AS VALUE2, Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'')  as VALUE, WMS_ITEM_LOC_BAL.ILOC_SEQ " &
                                " FROM WMS_ITEM_LOC_BAL INNER JOIN " &
                                " WMS_ITEM ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " &
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " &
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " &
                                " LEFT OUTER JOIN WMS_COL_CODE col on WMS_ITEM.ITEM_PRICE_CLASS = col.COLC_CODE and col.COLC_TABCOL='PRICE_CLASS_VALUE' " &
                                " WHERE (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN is null) AND WMS_ITEM_LOC_BAL.ILOC_BAL_QTY> 0 " & tempStr
                        'And WMS_ITEM.ITM_CC='Y' 

                        binValue2 = " SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, " & orgQTYSQL2 & " WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID," & _
                                " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO," & orgQTYSQL & " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, " & _
                                " WMS_ITEM.ITM_CC_DATE, WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_GP_CODE, " & _
                                " WMS_ITEM.ITM_CODE + '#_#' + WMS_ITEM.PACK_KEY AS VALUE1, " & _
                                " isnull(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') + '#_#' + isnull(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') AS VALUE2, Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'')  as VALUE, WMS_ITEM_LOC_BAL.ILOC_SEQ " & _
                                " FROM WMS_ITEM_LOC_BAL INNER JOIN " & _
                                " WMS_ITEM ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                " LEFT OUTER JOIN WMS_COL_CODE col on WMS_ITEM.ITEM_PRICE_CLASS = col.COLC_CODE and col.COLC_TABCOL='PRICE_CLASS_VALUE' " & _
                                " WHERE WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' AND WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0 AND WMS_ITEM.ITM_CC='Y' " & tempStr & _
                                " UNION " & _
                                " SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, " & orgQTYSQL2 & " WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID," & _
                                " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO," & orgQTYSQL & " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, " & _
                                " WMS_ITEM.ITM_CC_DATE, WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_GP_CODE, " & _
                                " WMS_ITEM.ITM_CODE + '#_#' + WMS_ITEM.PACK_KEY AS VALUE1, " & _
                                " isnull(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') + '#_#' + isnull(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') AS VALUE2, Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'')  as VALUE, WMS_ITEM_LOC_BAL.ILOC_SEQ " & _
                                " FROM WMS_ITEM_LOC_BAL INNER JOIN " & _
                                " WMS_ITEM ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                " LEFT OUTER JOIN WMS_COL_CODE col on WMS_ITEM.ITEM_PRICE_CLASS = col.COLC_CODE and col.COLC_TABCOL='PRICE_CLASS_VALUE' " & _
                                " WHERE (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN is null) AND WMS_ITEM_LOC_BAL.ILOC_BAL_QTY> 0 AND WMS_ITEM.ITM_CC='Y' " & tempStr

                        value &= orderbySQL

                    Case "LOOKUP_ILOC_BAL"
                        Dim tempStr As String = ""

                        If (Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("IS_Cable")) AndAlso HttpContext.Current.Session("IS_Cable") = "Y") OrElse HttpContext.Current.Session("SEARCH_SESSION_PAGE_ISCABLE") = "Y" Then
                            tempStr &= " AND WMS_ITEM.itm_type='CABLE' "
                        End If

                        value = " SELECT WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE,WMS_ITEM.ITM_UOM, " &
                                " WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC, CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END AS ILOC_BAL_QTY,  " &
                                " WMS_ITEM_LOC_BAL.ILOC_SEQ, WMS_ITEM_LOC_BAL.ILOC_WH, WMS_ITEM_LOC_BAL.ILOC_FLOOR, WMS_ITEM_LOC_BAL.ILOC_AREA, " &
                                " WMS_ITEM_LOC_BAL.ILOC_RACK, WMS_ITEM_LOC_BAL.ILOC_BIN, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, " &
                                " Convert(varchar,WMS_ITEM_LOC_BAL.ILOC_MANU_DATE,103) as ILOC_MANU_DATE, Convert(varchar,WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE,103) as ILOC_EXPIRY_DATE," &
                                " WMS_ITEM.ITM_CODE + '#_#' + WMS_ITEM.PACK_KEY AS VALUE1, ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, '') " &
                                " + '#_#' + ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') AS VALUE2, Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'') AS value, WMS_ITEM.ITM_SKU_NO, " &
                                " WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL,WMS_ITEM_LOC_BAL_S.ILBS_SL " &
                                " FROM WMS_ITEM_LOC_BAL INNER JOIN " &
                                " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND " &
                                " WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY Left outer JOIN " &
                                " WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ "
                        If (HttpContext.Current.Session("PAGE_SESSION_MENU_CODE") = "OP_SADJ") Then
                            value &= " Where 1=1  " & tempStr
                        Else
                            value &= " Where 1=1 and CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN WMS_ITEM_LOC_BAL_S.ILBS_QTY2 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END > 0 " & tempStr
                        End If


                        If whereSQL <> "" Then value &= " and " & whereSQL

                        value &= orderbySQL

                    Case "LOOKUP_IM"
                        Dim tempStr As String = ""
                        If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                        value &= srchSQL & " WHERE 1=1 " & criteriaSQL

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SR_IS_Cable")) AndAlso HttpContext.Current.Session("SR_IS_Cable") = "Y" Then
                            tempStr &= " AND WMS_ITEM.itm_type='CABLE' "
                        End If

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_IS_CABLE")) Then
                            tempStr &= " AND WMS_ITEM.itm_type='CABLE' "
                        End If

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SR_IS_Scrap")) AndAlso HttpContext.Current.Session("SR_IS_Scrap") = "Y" Then
                            tempStr &= " AND WMS_ITEM.ITM_SCRAP_YN='Y' "
                        End If


                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SR_IS_NS")) AndAlso HttpContext.Current.Session("SR_IS_NS") = "Y" Then
                            tempStr &= " AND WMS_ITEM.ITM_NONSTOCK_YN='Y' "
                        Else
                            If Not String.IsNullOrEmpty(HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_ITM_NONSTOCK_YN")) Then
                                If HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_ITM_NONSTOCK_YN") = "N" Then
                                    tempStr &= " AND ISNULL(WMS_ITEM.ITM_NONSTOCK_YN,'N') <> 'Y' "

                                ElseIf HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_ITM_NONSTOCK_YN") = "Y" Then
                                    tempStr &= " AND WMS_ITEM.ITM_NONSTOCK_YN='Y' "
                                End If
                            End If
                        End If

                        If whereSQL <> "" Then value &= " and " & whereSQL
                        value &= tempStr & orderbySQL
                    Case "MAST_IM", "MAST_IM_PD"
                        Dim tempStr As String = ""
                        If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                        value &= srchSQL & " WHERE 1=1 " & criteriaSQL

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_DESC1")) Or Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_DESC2")) Or Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_DESC3")) Then
                            tempStr = " AND ( "
                            If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_DESC1")) Then
                                tempStr &= "WMS_ITEM.ITM_DESC LIKE '%" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_DESC1") & "%' "
                            End If
                            If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_DESC2")) Then
                                If tempStr <> " AND ( " Then
                                    tempStr &= " AND "
                                End If
                                tempStr &= "WMS_ITEM.ITM_DESC LIKE '%" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_DESC2") & "%' "
                            End If
                            If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_DESC3")) Then
                                If tempStr <> " AND ( " Then
                                    tempStr &= " AND "
                                End If
                                tempStr &= "WMS_ITEM.ITM_DESC LIKE '%" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_DESC3") & "%' "
                            End If
                            tempStr &= " ) "
                        End If
                        If fun_code = "MAST_IM" Then
                            tempStr = tempStr & " AND WMS_ITEM.STORER_CODE <> 'PD' "
                        End If
                        If whereSQL <> "" Then value &= " and " & whereSQL
                        value &= tempStr & orderbySQL

                    Case "LOOKUP_IM_STCHK"
                        Dim tempStr As String = ""
                        If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                        value &= srchSQL & " WHERE 1=1 " & criteriaSQL

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_IS_CABLE")) OrElse HttpContext.Current.Session("CK_IS_Cable") = "Y" Then
                            tempStr &= " AND WMS_ITEM.itm_type='CABLE' "
                        End If

                        If Not String.IsNullOrEmpty(HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_ITM_NONSTOCK_YN")) Then
                            If HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_ITM_NONSTOCK_YN") = "N" Then
                                tempStr &= " AND ISNULL(WMS_ITEM.ITM_NONSTOCK_YN,'N') <> 'Y' "

                            ElseIf HttpContext.Current.Session("SEARCH_SESSION_PAGE_WMS_ITEM_ITM_NONSTOCK_YN") = "Y" Then
                                tempStr &= " AND WMS_ITEM.ITM_NONSTOCK_YN='Y' "
                            End If
                        End If

                        If Not String.IsNullOrEmpty(HttpContext.Current.Session("SELECTED_ITMBAL")) Then
                            Dim tempArr As String()
                            Dim itmStr As String = ""
                            tempArr = gU.listToArray(HttpContext.Current.Session("SELECTED_ITMBAL"))

                            For i = 0 To tempArr.Length - 1
                                itmStr &= "'" & tempArr(i) & "',"
                            Next

                            If itmStr <> "" Then
                                itmStr = Left(itmStr, itmStr.Length - 1)
                            End If

                            tempStr &= " AND WMS_ITEM.ITM_CODE + '||' +  WMS_ITEM.PACK_KEY IN (" & itmStr & ") "

                        End If


                        If whereSQL <> "" Then value &= " and " & whereSQL
                        value &= tempStr & orderbySQL

                    Case "LOOKUP_SR"
                        Dim tempStr As String = ""
                        If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                        value &= srchSQL & " WHERE 1=1 " & criteriaSQL

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_LOCATION_CODE")) Then
                            tempStr &= " AND RTD_LOC LIKE '%" & gU.dbEncode(HttpContext.Current.Session("SEARCH_SESSION_PAGE_LOCATION_CODE")) & "%' "
                        End If

                        Dim tempWHStr As String = ""

                        If Not HttpContext.Current.Session("SR_MWH") Is Nothing AndAlso HttpContext.Current.Session("SR_MWH") <> "" Then
                            Dim tempSelect As String = " Select distinct wh_code from wms_warehouse where WH_MAIN_WH='" & gU.dbEncode(HttpContext.Current.Session("SR_MWH")) & "'"
                            Dim tempDT As System.Data.DataTable = gDB.getDataTable(tempSelect)

                            If tempDT.Rows.Count > 0 Then
                                For i = 0 To tempDT.Rows.Count - 1
                                    tempWHStr &= "'" & tempDT.Rows(i).Item("wh_code").ToString.Trim & "',"
                                Next

                                If tempWHStr <> "" Then
                                    tempWHStr = Left(tempWHStr, Len(tempWHStr) - 1)
                                    tempStr &= " AND WMS_WH_BIN.WH_CODE in(" & tempWHStr & ") "
                                End If
                            End If
                        End If

                        If tempStr <> "" Then value &= tempStr

                        If whereSQL <> "" Then value &= " and " & whereSQL

                        value &= tempStr & groupbySQL & orderbySQL

                    Case "LOOKUP_SRC_WIT"
                        Dim tempStr As String = ""

                        If Not HttpContext.Current.Session("SR_MWH") Is Nothing AndAlso HttpContext.Current.Session("SR_MWH") <> "" Then
                            tempStr = " AND DOD_WH_CODE='" & gU.dbEncode(HttpContext.Current.Session("SR_MWH")) & "' "
                        End If
                        If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                        value &= srchSQL & " WHERE 1=1 " & tempStr & criteriaSQL

                        If whereSQL <> "" Then value &= " and " & whereSQL

                        value &= groupbySQL & orderbySQL
                    Case "RPT_DOSUM"
                        Dim tempStr As String = ""
                        If whereSQL <> "" Then whereSQL = " and " & whereSQL

                        value = " SELECT DISTINCT WMS_DO_PICKLIST_D.PLD_SEQ, WMS_DELV_ORDER.DO_CODE, WMS_STORER.STO_SHORTNAME, WMS_DELV_ORDER.STORER_CODE, WMS_DELV_ORDER.DO_CO_CODE, " & _
                                " WMS_DELV_ORDER.DO_STATUS, CONVERT(Varchar, WMS_DELV_ORDER.DO_DATE, 103) AS DO_DATE,  " & _
                                " WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_NAME, WMS_DELV_ORDER_D.DOD_PACK_KEY, WMS_DELV_ORDER_D.DOD_QTY, " & _
                                " WMS_DO_PICKLIST_D.PLD_ITEM_QTY, WMS_DO_PICKLIST_D.PLD_WH, WMS_DO_PICKLIST_D.PLD_FLOOR, WMS_DO_PICKLIST_D.PLD_AREA,  " & _
                                " WMS_DO_PICKLIST_D.PLD_RACK, WMS_DO_PICKLIST_D.PLD_BIN, WMS_DO_PICKLIST_D.PLD_LOC, WMS_DELV_ORDER.SYS_LUD, WMS_DELV_ORDER.SYS_LUB,  " & _
                                " WMS_DELV_ORDER.SYS_CD as DO_SYS_CD, CO.SYS_CD AS CO_SYS_CD, WMS_DO_PICKLIST_D.PLD_PICKED_BY, WMS_DELV_ORDER.DO_TROLLEY_ID,  " & _
                                " WMS_DELV_ORDER.DO_DRUM_ID, WMS_DELV_ORDER.DO_POSTED_DATE, WMS_DELV_ORDER.DO_EDI_SIR_NO, WMS_DELV_ORDER.DO_EDI_WIT_NO, WMS_DO_PICKLIST_D.PLD_SERIAL_NO, WMS_DELV_ORDER.DO_POSTED_BY,  " & _
                                " WMS_ITEM.ITM_WEIGHT_TYPE, WMS_ITEM.ITM_CHE_CLASS, WMS_ITEM.ITM_REQ_STORE_HUM_YN, WMS_ITEM.ITM_REQ_STORE_AIRCON_YN, WMS_ITEM.ITM_TYPE,  " & _
                                " WMS_ITEM.ITM_DG_YN, WMS_ITEM.ITM_CRITICAL_YN, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_NONSTOCK_YN, WMS_ITEM_WH.IW_PREF_LOC1, IW1.IW_PREF_LOC1 as IW_PREF_LOC, " & _
                                " WMS_DO_PICKLIST_D.PLD_PICKED_DATE, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_RESTRICTED_ITEM,  " & _
                                " WMS_DO_PICKLIST_D.PLD_FOI_QTY, CO.COD_QTY, WMS_DO_PICKLIST_D.PLD_PL_LIST_NO, WMS_DO_PICKLIST_D.PLD_WAVE_PICK_NO, " & _
                                " Convert(varchar,WMS_DO_PICKLIST_D.PLD_EXPIRY_DATE, 103) as PLD_EXPIRY_DATE, convert(varchar,WMS_DO_PICKLIST_D.PLD_MANU_DATE,103) as PLD_MANU_DATE, WMS_ITEM.ITM_UOM, " & _
                                " WMS_DO_PICKLIST_D.PLD_PALLET_NO, WMS_DELV_ORDER_D.DOD_CARTON_NO, WMS_DELV_ORDER_D.DOD_PACK_NO, WMS_DELV_ORDER_D.DOD_PACK_TYPE, WMS_DELV_ORDER_D.DOD_UOM, " & _
                                " WMS_DELV_ORDER_D.DOD_PCS_UOM, WMS_DELV_ORDER_D.DOD_TOT_WGT, WMS_DELV_ORDER_D.DOD_TOT_CBM,WMS_DO_PICKLIST_D.PLD_BATCH_NO, WMS_DELV_ORDER_D.DOD_WH_CODE, " & _
                                " WMS_DELV_ORDER_D.DOD_REM, WMS_DELV_ORDER_D.DOD_VEND_CODE,CO.CO_REQ_BY,WMS_DELV_ORDER.DO_ISSUED_BY,  " & _
                                " WMS_DELV_ORDER.SYS_CB, NULL as ACC_NO, WMS_DELV_ORDER.DO_SHIP_TO " & _
                                " FROM WMS_DO_PICKLIST_D LEFT OUTER JOIN " & _
                                " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND  " & _
                                " WMS_ITEM.PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY AND WMS_ITEM.ITM_CODE = WMS_DO_PICKLIST_D.PLD_ITEM_NO LEFT OUTER JOIN " & _
                                " WMS_ITEM_WH ON WMS_DO_PICKLIST_D.IMP_CODE = WMS_ITEM_WH.IMP_CODE AND  " & _
                                " WMS_DO_PICKLIST_D.STORER_CODE = WMS_ITEM_WH.STORER_CODE AND WMS_DO_PICKLIST_D.PLD_ITEM_NO = WMS_ITEM_WH.ITM_CODE AND  " & _
                                " WMS_DO_PICKLIST_D.PLD_PACK_KEY = WMS_ITEM_WH.PACK_KEY AND WMS_DO_PICKLIST_D.PLD_WH = WMS_ITEM_WH.WH_CODE INNER JOIN " & _
                                " WMS_DELV_ORDER ON WMS_DO_PICKLIST_D.IMP_CODE = WMS_DELV_ORDER.IMP_CODE AND  " & _
                                " WMS_DO_PICKLIST_D.STORER_CODE = WMS_DELV_ORDER.STORER_CODE AND WMS_DO_PICKLIST_D.DO_CODE = WMS_DELV_ORDER.DO_CODE INNER JOIN " & _
                                " WMS_STORER ON WMS_DELV_ORDER.IMP_CODE = WMS_STORER.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_STORER.STORER_CODE LEFT OUTER JOIN " & _
                                "  (select WMS_CUST_ORDER.IMP_CODE, WMS_CUST_ORDER.STORER_CODE,WMS_CUST_ORDER.SYS_CD,  WMS_CUST_ORDER.co_code, WMS_CUST_ORDER_D.COD_ITM_CODE, WMS_CUST_ORDER_D.COD_PACK_KEY, SUM(WMS_CUST_ORDER_D.COD_QTY) as COD_QTY,WMS_CUST_ORDER.CO_REQ_BY,WMS_CUST_ORDER.CO_STATUS, sum(WMS_CUST_ORDER_D.COD_POST_QTY) as COD_POST_QTY  from wms_cust_order INNER JOIN " & _
                                " WMS_CUST_ORDER_D ON WMS_CUST_ORDER.IMP_CODE = WMS_CUST_ORDER_D.IMP_CODE AND  " & _
                                " WMS_CUST_ORDER.STORER_CODE = WMS_CUST_ORDER_D.STORER_CODE AND WMS_CUST_ORDER.CO_CODE = WMS_CUST_ORDER_D.CO_CODE group by WMS_CUST_ORDER.IMP_CODE, WMS_CUST_ORDER.STORER_CODE,WMS_CUST_ORDER.SYS_CD,  WMS_CUST_ORDER.co_code, WMS_CUST_ORDER_D.COD_ITM_CODE, WMS_CUST_ORDER_D.COD_PACK_KEY,WMS_CUST_ORDER.CO_REQ_BY,WMS_CUST_ORDER.CO_STATUS) co " & _
                                " ON WMS_DELV_ORDER.IMP_CODE = co.IMP_CODE AND  " & _
                                " WMS_DELV_ORDER.STORER_CODE = co.STORER_CODE AND WMS_DELV_ORDER.DO_CO_CODE = co.CO_CODE AND  CO.COD_ITM_CODE = WMS_DO_PICKLIST_D.PLD_ITEM_NO AND  " & _
                                " CO.COD_PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY " & _
                                " INNER JOIN WMS_DELV_ORDER_D ON WMS_DELV_ORDER_D.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND " & _
                                " WMS_DELV_ORDER_D.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND WMS_DELV_ORDER_D.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE AND  " & _
                                " WMS_DELV_ORDER_D.DOD_PALLET_NO = WMS_DO_PICKLIST_D.PLD_PALLET_NO AND  " & _
                                " WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY AND  " & _
                                " WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_DO_PICKLIST_D.PLD_ITEM_NO AND  " & _
                                " isnull(WMS_DELV_ORDER_D.DOD_BATCH_NO, '') = isnull(WMS_DO_PICKLIST_D.PLD_BATCH_NO, '') AND  " & _
                                " isnull(WMS_DELV_ORDER_D.DOD_SERIAL, '') = isnull(WMS_DO_PICKLIST_D.PLD_SERIAL_NO, '') " & _
                                " LEFT OUTER JOIN " & _
                                " WMS_ITEM_WH IW1 ON WMS_ITEM.IMP_CODE = IW1.IMP_CODE AND " & _
                                " WMS_ITEM.STORER_CODE = IW1.STORER_CODE AND WMS_ITEM.ITM_CODE = IW1.ITM_CODE AND " & _
                                " WMS_ITEM.PACK_KEY = IW1.PACK_KEY AND WMS_ITEM.ITM_PREF_WH = IW1.WH_CODE " & _
                                " WHERE (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' or  WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) " & whereSQL & _
                                " UNION ALL " & _
                                " SELECT DISTINCT WMS_DO_PICKLIST_D.PLD_SEQ, WMS_DELV_ORDER.DO_CODE, WMS_STORER.STO_SHORTNAME, WMS_DELV_ORDER.STORER_CODE, WMS_DELV_ORDER.DO_CO_CODE, " & _
                                " WMS_DELV_ORDER.DO_STATUS, CONVERT(Varchar, WMS_DELV_ORDER.DO_DATE, 103) AS DO_DATE,  " & _
                                " WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_NAME, WMS_DELV_ORDER_D.DOD_PACK_KEY, WMS_DELV_ORDER_D.DOD_QTY2,  " & _
                                " WMS_DO_PICKLIST_D.PLD_ITEM_QTY * WMS_DO_PICKLIST_D.PLD_QTY2, WMS_DO_PICKLIST_D.PLD_WH, WMS_DO_PICKLIST_D.PLD_FLOOR, WMS_DO_PICKLIST_D.PLD_AREA,  " & _
                                " WMS_DO_PICKLIST_D.PLD_RACK, WMS_DO_PICKLIST_D.PLD_BIN, WMS_DO_PICKLIST_D.PLD_LOC, WMS_DELV_ORDER.SYS_LUD, WMS_DELV_ORDER.SYS_LUB,  " & _
                                " WMS_DELV_ORDER.SYS_CD as DO_SYS_CD, CO.SYS_CD AS CO_SYS_CD, WMS_DO_PICKLIST_D.PLD_PICKED_BY, WMS_DELV_ORDER.DO_TROLLEY_ID,  " & _
                                " WMS_DELV_ORDER.DO_DRUM_ID, WMS_DELV_ORDER.DO_POSTED_DATE, WMS_DELV_ORDER.DO_EDI_SIR_NO, WMS_DELV_ORDER.DO_EDI_WIT_NO, WMS_DO_PICKLIST_D.PLD_SERIAL_NO, WMS_DELV_ORDER.DO_POSTED_BY,  " & _
                                " WMS_ITEM.ITM_WEIGHT_TYPE, WMS_ITEM.ITM_CHE_CLASS, WMS_ITEM.ITM_REQ_STORE_HUM_YN, WMS_ITEM.ITM_REQ_STORE_AIRCON_YN, WMS_ITEM.ITM_TYPE,  " & _
                                " WMS_ITEM.ITM_DG_YN, WMS_ITEM.ITM_CRITICAL_YN, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_NONSTOCK_YN, WMS_ITEM_WH.IW_PREF_LOC1, IW1.IW_PREF_LOC1 as IW_PREF_LOC, " & _
                                " WMS_DO_PICKLIST_D.PLD_PICKED_DATE, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_RESTRICTED_ITEM,  " & _
                                " WMS_DO_PICKLIST_D.PLD_FOI_QTY * WMS_DO_PICKLIST_D.PLD_QTY2, CO.COD_QTY, WMS_DO_PICKLIST_D.PLD_PL_LIST_NO, WMS_DO_PICKLIST_D.PLD_WAVE_PICK_NO, " & _
                                " Convert(varchar,WMS_DO_PICKLIST_D.PLD_EXPIRY_DATE, 103) as PLD_EXPIRY_DATE, convert(varchar,WMS_DO_PICKLIST_D.PLD_MANU_DATE,103) as PLD_MANU_DATE, WMS_ITEM.ITM_UOM, " & _
                                " WMS_DO_PICKLIST_D.PLD_PALLET_NO, WMS_DELV_ORDER_D.DOD_CARTON_NO, WMS_DELV_ORDER_D.DOD_PACK_NO, WMS_DELV_ORDER_D.DOD_PACK_TYPE, WMS_ITEM.ITM_UOM2, " & _
                                " WMS_DELV_ORDER_D.DOD_PCS_UOM, WMS_DELV_ORDER_D.DOD_TOT_WGT, WMS_DELV_ORDER_D.DOD_TOT_CBM,WMS_DO_PICKLIST_D.PLD_BATCH_NO, WMS_DELV_ORDER_D.DOD_WH_CODE, " & _
                                " WMS_DELV_ORDER_D.DOD_REM, WMS_DELV_ORDER_D.DOD_VEND_CODE,CO.CO_REQ_BY,WMS_DELV_ORDER.DO_ISSUED_BY,  " & _
                                " WMS_DELV_ORDER.SYS_CB, NULL as ACC_NO, WMS_DELV_ORDER.DO_SHIP_TO " & _
                                " FROM WMS_DO_PICKLIST_D LEFT OUTER JOIN " & _
                                " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND  " & _
                                " WMS_ITEM.PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY AND WMS_ITEM.ITM_CODE = WMS_DO_PICKLIST_D.PLD_ITEM_NO LEFT OUTER JOIN " & _
                                " WMS_ITEM_WH ON WMS_DO_PICKLIST_D.IMP_CODE = WMS_ITEM_WH.IMP_CODE AND  " & _
                                " WMS_DO_PICKLIST_D.STORER_CODE = WMS_ITEM_WH.STORER_CODE AND WMS_DO_PICKLIST_D.PLD_ITEM_NO = WMS_ITEM_WH.ITM_CODE AND  " & _
                                " WMS_DO_PICKLIST_D.PLD_PACK_KEY = WMS_ITEM_WH.PACK_KEY AND WMS_DO_PICKLIST_D.PLD_WH = WMS_ITEM_WH.WH_CODE INNER JOIN " & _
                                " WMS_DELV_ORDER ON WMS_DO_PICKLIST_D.IMP_CODE = WMS_DELV_ORDER.IMP_CODE AND  " & _
                                " WMS_DO_PICKLIST_D.STORER_CODE = WMS_DELV_ORDER.STORER_CODE AND WMS_DO_PICKLIST_D.DO_CODE = WMS_DELV_ORDER.DO_CODE INNER JOIN " & _
                                " WMS_STORER ON WMS_DELV_ORDER.IMP_CODE = WMS_STORER.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_STORER.STORER_CODE LEFT OUTER JOIN " & _
                                "  (select WMS_CUST_ORDER.IMP_CODE, WMS_CUST_ORDER.STORER_CODE,WMS_CUST_ORDER.SYS_CD,  WMS_CUST_ORDER.co_code, WMS_CUST_ORDER_D.COD_ITM_CODE, WMS_CUST_ORDER_D.COD_PACK_KEY, SUM(WMS_CUST_ORDER_D.COD_QTY2) as COD_QTY,WMS_CUST_ORDER.CO_REQ_BY,WMS_CUST_ORDER.CO_STATUS, sum(WMS_CUST_ORDER_D.COD_POST_QTY * WMS_CUST_ORDER_D.COD_QTY2) as COD_POST_QTY  from wms_cust_order INNER JOIN " & _
                                " WMS_CUST_ORDER_D ON WMS_CUST_ORDER.IMP_CODE = WMS_CUST_ORDER_D.IMP_CODE AND  " & _
                                " WMS_CUST_ORDER.STORER_CODE = WMS_CUST_ORDER_D.STORER_CODE AND WMS_CUST_ORDER.CO_CODE = WMS_CUST_ORDER_D.CO_CODE group by WMS_CUST_ORDER.IMP_CODE, WMS_CUST_ORDER.STORER_CODE,WMS_CUST_ORDER.SYS_CD,  WMS_CUST_ORDER.co_code, WMS_CUST_ORDER_D.COD_ITM_CODE, WMS_CUST_ORDER_D.COD_PACK_KEY,WMS_CUST_ORDER.CO_REQ_BY,WMS_CUST_ORDER.CO_STATUS) co " & _
                                " ON WMS_DELV_ORDER.IMP_CODE = co.IMP_CODE AND  " & _
                                " WMS_DELV_ORDER.STORER_CODE = co.STORER_CODE AND WMS_DELV_ORDER.DO_CO_CODE = co.CO_CODE AND  CO.COD_ITM_CODE = WMS_DO_PICKLIST_D.PLD_ITEM_NO AND  " & _
                                " CO.COD_PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY " & _
                                " INNER JOIN WMS_DELV_ORDER_D ON WMS_DELV_ORDER_D.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND " & _
                                " WMS_DELV_ORDER_D.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND WMS_DELV_ORDER_D.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE AND  " & _
                                " WMS_DELV_ORDER_D.DOD_PALLET_NO = WMS_DO_PICKLIST_D.PLD_PALLET_NO AND  " & _
                                " WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_DO_PICKLIST_D.PLD_PACK_KEY AND  " & _
                                " WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_DO_PICKLIST_D.PLD_ITEM_NO AND  " & _
                                " isnull(WMS_DELV_ORDER_D.DOD_BATCH_NO, '') = isnull(WMS_DO_PICKLIST_D.PLD_BATCH_NO, '') " & _
                                " LEFT OUTER JOIN " & _
                                " WMS_ITEM_WH IW1 ON WMS_ITEM.IMP_CODE = IW1.IMP_CODE AND " & _
                                " WMS_ITEM.STORER_CODE = IW1.STORER_CODE AND WMS_ITEM.ITM_CODE = IW1.ITM_CODE AND " & _
                                " WMS_ITEM.PACK_KEY = IW1.PACK_KEY AND WMS_ITEM.ITM_PREF_WH = IW1.WH_CODE " & _
                                " WHERE WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' " & whereSQL

                        'value &= srchSQL & " WHERE 1=1 AND " & criteriaSQL

                        value &= tempStr & orderbySQL

                    Case "RPT_SRSUM"
                        'Dim tempStr As String = ""
                        'Dim totalSQL As String = ""

                        If whereSQL <> "" Then whereSQL = " and " & whereSQL

                        value = " SELECT WMS_WH_BIN.BN_CSMS_CODE, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_STOCK_RETURN.RT_REF_NO2, WMS_STOCK_RETURN.RT_TYPE, WMS_STOCK_RETURN.RT_DATE, CONVERT(varchar,WMS_STOCK_RETURN.RT_DATE,103) as  RT_DATE_D, " & _
                                " WMS_STOCK_RETURN.RT_STATUS, WMS_ITEM.ITM_SKU_NO, WMS_STOCK_RETURN_D.RTD_DRUM_ID, WMS_STOCK_RETURN_D.RTD_SERIAL_NO, WMS_STOCK_RETURN_D.RTD_BATCH_NO,  " & _
                                " WMS_STOCK_RETURN_D.RTD_RCV_QTY, WMS_ITEM.ITM_UOM, WMS_STOCK_RETURN.RT_C8_YN, WMS_STOCK_RETURN_D.RTD_KG, WMS_STOCK_RETURN_D.RTD_LOC, WMS_STOCK_RETURN.RT_RCV_BY,  " & _
                                " WMS_STOCK_RETURN.RT_BY, WMS_STOCK_RETURN.RT_REF_DOC_NO, WMS_STOCK_RETURN_D.RT_CODE, WMS_STOCK_RETURN.RT_WH,RTD_ITM_NAME, WMS_STOCK_RETURN.RT_REM, WMS_STOCK_RETURN.RT_CUS_CODE, '' as SIR_NO " & _
                                " FROM WMS_STOCK_RETURN INNER JOIN " & _
                                " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE LEFT OUTER JOIN " & _
                                " WMS_ITEM ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM.PACK_KEY LEFT OUTER JOIN " & _
                                " WMS_WH_BIN ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_STOCK_RETURN_D.RTD_LOC = WMS_WH_BIN.LOC_KEY " & _
                                " WHERE (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) " & whereSQL & _
                                " UNION  " & _
                                " SELECT WMS_WH_BIN.BN_CSMS_CODE, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_STOCK_RETURN.RT_REF_NO2, WMS_STOCK_RETURN.RT_TYPE, WMS_STOCK_RETURN.RT_DATE, CONVERT(varchar,WMS_STOCK_RETURN.RT_DATE,103) as  RT_DATE_D,  " & _
                                " WMS_STOCK_RETURN.RT_STATUS, WMS_ITEM.ITM_SKU_NO, WMS_STOCK_RETURN_D.RTD_DRUM_ID, WMS_STOCK_RETURN_D.RTD_SERIAL_NO, WMS_STOCK_RETURN_D.RTD_BATCH_NO,  " & _
                                " WMS_STOCK_RETURN_D.RTD_QTY2, WMS_STOCK_RETURN_D.RTD_UOM2, WMS_STOCK_RETURN.RT_C8_YN, WMS_STOCK_RETURN_D.RTD_KG, WMS_STOCK_RETURN_D.RTD_LOC, WMS_STOCK_RETURN.RT_RCV_BY,  " & _
                                " WMS_STOCK_RETURN.RT_BY, WMS_STOCK_RETURN.RT_REF_DOC_NO, WMS_STOCK_RETURN_D.RT_CODE, WMS_STOCK_RETURN.RT_WH, RTD_ITM_NAME, WMS_STOCK_RETURN.RT_REM, WMS_STOCK_RETURN.RT_CUS_CODE, '' as SIR_NO " & _
                                " FROM WMS_STOCK_RETURN INNER JOIN " & _
                                " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE LEFT OUTER JOIN " & _
                                " WMS_ITEM ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM.PACK_KEY LEFT OUTER JOIN " & _
                                " WMS_WH_BIN ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_STOCK_RETURN_D.RTD_LOC = WMS_WH_BIN.LOC_KEY " & _
                                " WHERE WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' " & whereSQL

                        'value &= srchSQL & " WHERE 1=1 AND " & criteriaSQL
                        'totalSQL = " SELECT 1 as union_order, NULL, NULL, NULL, NULL, NULL, " & _
                        '           " NULL, NULL, NULL, NULL, NULL, NULL, NULL,'Total KG',  " & _
                        '           " NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL,NULL, NULL, NULL,null, null, null,null, null," & _
                        '           " null, null,null, null,null, null, SUM(WMS_STOCK_RETURN_D.RTD_KG), null, null,null, null, " & _
                        '           " null, null, null, null, null, null, null,null " & _
                        '           " FROM WMS_STOCK_RETURN INNER JOIN " & _
                        '           " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND  " & _
                        '           " WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND  " & _
                        '           " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE INNER JOIN " & _
                        '           " WMS_ITEM ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                        '           " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " & _
                        '           " V_ALT_VEND_ITEM ON WMS_ITEM.IMP_CODE = V_ALT_VEND_ITEM.IMP_CODE AND WMS_ITEM.STORER_CODE = V_ALT_VEND_ITEM.STORER_CODE AND  " & _
                        '           " WMS_ITEM.ITM_CODE = V_ALT_VEND_ITEM.ITM_CODE AND WMS_ITEM.PACK_KEY = V_ALT_VEND_ITEM.PACK_KEY INNER JOIN " & _
                        '           " WMS_STORER ON WMS_STOCK_RETURN.IMP_CODE = WMS_STORER.IMP_CODE AND  " & _
                        '           " WMS_STOCK_RETURN.STORER_CODE = WMS_STORER.STORER_CODE left outer JOIN " & _
                        '           " WMS_STOCK_ISSUE ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_ISSUE.IMP_CODE AND " & _
                        '           " WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_ISSUE.STORER_CODE AND WMS_STOCK_RETURN.RT_REF_DOC_NO = WMS_STOCK_ISSUE.IS_CODE " & _
                        '           " WHERE 1=1"

                        'If whereSQL <> "" Then totalSQL &= " and " & whereSQL

                        'value &= " UNION " & totalSQL

                        'If orderbySQL <> "" Then orderbySQL = " ORDER BY union_order, " & Replace(orderbySQL.ToUpper, "ORDER BY", "")

                        If orderbySQL = "" Then orderbySQL = " ORDER BY RT_DATE DESC, ITM_SKU_NO "

                        value &= orderbySQL
                    Case "RPT_DGR"

                        If whereSQL <> "" Then whereSQL = " and " & whereSQL

                        value = " SELECT WMS_GOODSRCV.STORER_CODE, WMS_STORER.STO_SHORTNAME, WMS_GOODSRCV.GR_CODE, WMS_GOODSRCV.PRJ_CODE, " & _
                                "  WMS_GOODSRCV.GR_STATUS, WMS_GOODSRCV.GR_TYPE, Convert(varchar,WMS_GOODSRCV.GR_DATE, 103) AS GR_DATE, " & _
                                "  WMS_GOODSRCV.GR_RCV_BY, WMS_GOODSRCV.PO_CODE, WMS_GOODSRCV.GR_DOC_TYPE, WMS_GOODSRCV.GR_DOC_NO, " & _
                                "  Convert(varchar,WMS_GOODSRCV.GR_DATE, 103) AS PO_DATE, WMS_GOODSRCV.VND_CODE, WMS_GOODSRCV.VND_NAME, " & _
                                "  WMS_GOODSRCV.GR_VND_DNREF, WMS_GOODSRCV.GR_REM, WMS_GOODSRCV.GR_TRACK_NO, WMS_GOODSRCV.SYS_LUD, " & _
                                "  WMS_GOODSRCV_D.GRD_LOC, WMS_GOODSRCV_D.GRD_WH, WMS_GOODSRCV_D.GRD_ITM_CODE, WMS_GOODSRCV_D.GRD_PACK_KEY, WMS_GOODSRCV_D.GRD_BATCH_NO, " & _
                                "  WMS_GOODSRCV_D.GRD_ITM_NAME, WMS_GOODSRCV_D.GRD_RCV_QTY, " & _
                                "  Convert(varchar,WMS_GOODSRCV.GR_DATE, 103) as GRD_EXPIRY_DATE_D, Convert(varchar,WMS_GOODSRCV.GR_DATE, 103) as GRD_MANU_DATE_D, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_CAP_REV, WMS_ITEM.ITEM_PRICE_CLASS, WMS_ITEM.ITM_CAP_VALUE, IW1.IW_PREF_LOC1 as IW_PREF_LOC, " & _
                                "  WMS_GOODSRCV_D_S.GRS_SERIAL_NO, NULL as GRS_DRUM_ID, WMS_GOODSRCV.GR_POSTED_DATE, WMS_GOODSRCV_D.GRD_UOM, WMS_GOODSRCV.SYS_CD, WMS_GOODSRCV_D.GRD_INSP_QTY, WMS_GOODSRCV_D.GRD_INSP_REQ, WMS_GOODSRCV_D.GRD_WH, GR_EDI_PO_NO " & _
                                " FROM WMS_GOODSRCV " & _
                                " INNER JOIN WMS_GOODSRCV_D " & _
                                " ON WMS_GOODSRCV.GR_CODE      = WMS_GOODSRCV_D.GR_CODE " & _
                                " AND WMS_GOODSRCV.IMP_CODE    = WMS_GOODSRCV_D.IMP_CODE " & _
                                " AND WMS_GOODSRCV.STORER_CODE = WMS_GOODSRCV_D.STORER_CODE " & _
                                " INNER JOIN WMS_STORER " & _
                                " ON WMS_GOODSRCV.IMP_CODE    =WMS_STORER.IMP_CODE " & _
                                " AND WMS_GOODSRCV.STORER_CODE =WMS_STORER.STORER_CODE " & _
                                " LEFT OUTER JOIN WMS_ITEM " & _
                                " ON WMS_GOODSRCV_D.IMP_CODE    = WMS_ITEM.IMP_CODE " & _
                                " AND WMS_GOODSRCV_D.STORER_CODE = WMS_ITEM.STORER_CODE " & _
                                " AND WMS_GOODSRCV_D.GRD_ITM_CODE = WMS_ITEM.ITM_CODE " & _
                                " AND WMS_GOODSRCV_D.GRD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " LEFT OUTER JOIN  " & _
                                " WMS_ITEM_WH IW1 ON WMS_ITEM.IMP_CODE = IW1.IMP_CODE AND   " & _
                                " WMS_ITEM.STORER_CODE = IW1.STORER_CODE AND WMS_ITEM.ITM_CODE = IW1.ITM_CODE AND   " & _
                                " WMS_ITEM.PACK_KEY = IW1.PACK_KEY AND WMS_ITEM.ITM_PREF_WH = IW1.WH_CODE " & _
                                " LEFT OUTER JOIN WMS_GOODSRCV_D_S ON WMS_GOODSRCV_D.IMP_CODE = WMS_GOODSRCV_D_S.IMP_CODE AND WMS_GOODSRCV_D.STORER_CODE = WMS_GOODSRCV_D_S.STORER_CODE AND  " & _
                                " WMS_GOODSRCV_D.GR_CODE = WMS_GOODSRCV_D_S.GR_CODE AND WMS_GOODSRCV_D.GRD_ITM_CODE = WMS_GOODSRCV_D_S.GRS_ITM_CODE AND  " & _
                                " WMS_GOODSRCV_D.GRD_PACK_KEY = WMS_GOODSRCV_D_S.GRS_PACK_KEY AND ISNULL(WMS_GOODSRCV_D.GRD_PALLET_NO,'000') = ISNULL(WMS_GOODSRCV_D_S.GRS_PALLET_NO,'000') AND  " & _
                                " ISNULL(WMS_GOODSRCV_D.GRD_BATCH_NO,'') = ISNULL(WMS_GOODSRCV_D_S.GRS_BATCH_NO,'') " & _
                                " WHERE (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' or WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) " & whereSQL & _
                                " UNION ALL " & _
                                " SELECT WMS_GOODSRCV.STORER_CODE, WMS_STORER.STO_SHORTNAME, WMS_GOODSRCV.GR_CODE, WMS_GOODSRCV.PRJ_CODE, " & _
                                "  WMS_GOODSRCV.GR_STATUS, WMS_GOODSRCV.GR_TYPE, Convert(varchar,WMS_GOODSRCV.GR_DATE, 103) AS GR_DATE, " & _
                                "  WMS_GOODSRCV.GR_RCV_BY, WMS_GOODSRCV.PO_CODE, WMS_GOODSRCV.GR_DOC_TYPE, WMS_GOODSRCV.GR_DOC_NO, " & _
                                "  Convert(varchar,WMS_GOODSRCV.GR_DATE, 103) AS PO_DATE, WMS_GOODSRCV.VND_CODE, WMS_GOODSRCV.VND_NAME, " & _
                                "  WMS_GOODSRCV.GR_VND_DNREF, WMS_GOODSRCV.GR_REM, WMS_GOODSRCV.GR_TRACK_NO, WMS_GOODSRCV.SYS_LUD, " & _
                                "  WMS_GOODSRCV_D.GRD_LOC, WMS_GOODSRCV_D.GRD_WH, WMS_GOODSRCV_D.GRD_ITM_CODE, WMS_GOODSRCV_D.GRD_PACK_KEY, WMS_GOODSRCV_D.GRD_BATCH_NO, " & _
                                "  WMS_GOODSRCV_D.GRD_ITM_NAME, WMS_GOODSRCV_D.GRD_QTY2, " & _
                                "  Convert(varchar,WMS_GOODSRCV.GR_DATE, 103) as GRD_EXPIRY_DATE_D, Convert(varchar,WMS_GOODSRCV.GR_DATE, 103) as GRD_MANU_DATE_D, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_CAP_REV, WMS_ITEM.ITEM_PRICE_CLASS, WMS_ITEM.ITM_CAP_VALUE, IW1.IW_PREF_LOC1 as IW_PREF_LOC, " & _
                                "  WMS_GOODSRCV_D_S.GRS_SERIAL_NO, NULL as GRS_DRUM_ID, WMS_GOODSRCV.GR_POSTED_DATE, ISNULL(WMS_ITEM.ITM_UOM2, WMS_GOODSRCV_D.GRD_UOM), WMS_GOODSRCV.SYS_CD, WMS_GOODSRCV_D.GRD_INSP_QTY * WMS_GOODSRCV_D.GRD_QTY2, WMS_GOODSRCV_D.GRD_INSP_REQ, WMS_GOODSRCV_D.GRD_WH, GR_EDI_PO_NO " & _
                                " FROM WMS_GOODSRCV " & _
                                " INNER JOIN WMS_GOODSRCV_D " & _
                                " ON WMS_GOODSRCV.GR_CODE      = WMS_GOODSRCV_D.GR_CODE " & _
                                " AND WMS_GOODSRCV.IMP_CODE    = WMS_GOODSRCV_D.IMP_CODE " & _
                                " AND WMS_GOODSRCV.STORER_CODE = WMS_GOODSRCV_D.STORER_CODE " & _
                                " INNER JOIN WMS_STORER " & _
                                " ON WMS_GOODSRCV.IMP_CODE    =WMS_STORER.IMP_CODE " & _
                                " AND WMS_GOODSRCV.STORER_CODE =WMS_STORER.STORER_CODE " & _
                                " LEFT OUTER JOIN WMS_ITEM " & _
                                " ON WMS_GOODSRCV_D.IMP_CODE    = WMS_ITEM.IMP_CODE " & _
                                " AND WMS_GOODSRCV_D.STORER_CODE = WMS_ITEM.STORER_CODE " & _
                                " AND WMS_GOODSRCV_D.GRD_ITM_CODE = WMS_ITEM.ITM_CODE " & _
                                " AND WMS_GOODSRCV_D.GRD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " LEFT OUTER JOIN  " & _
                                " WMS_ITEM_WH IW1 ON WMS_ITEM.IMP_CODE = IW1.IMP_CODE AND   " & _
                                " WMS_ITEM.STORER_CODE = IW1.STORER_CODE AND WMS_ITEM.ITM_CODE = IW1.ITM_CODE AND   " & _
                                " WMS_ITEM.PACK_KEY = IW1.PACK_KEY AND WMS_ITEM.ITM_PREF_WH = IW1.WH_CODE " & _
                                " LEFT OUTER JOIN WMS_GOODSRCV_D_S ON WMS_GOODSRCV_D.IMP_CODE = WMS_GOODSRCV_D_S.IMP_CODE AND WMS_GOODSRCV_D.STORER_CODE = WMS_GOODSRCV_D_S.STORER_CODE AND  " & _
                                " WMS_GOODSRCV_D.GR_CODE = WMS_GOODSRCV_D_S.GR_CODE AND WMS_GOODSRCV_D.GRD_ITM_CODE = WMS_GOODSRCV_D_S.GRS_ITM_CODE AND  " & _
                                " WMS_GOODSRCV_D.GRD_PACK_KEY = WMS_GOODSRCV_D_S.GRS_PACK_KEY AND ISNULL(WMS_GOODSRCV_D.GRD_PALLET_NO,'000') = ISNULL(WMS_GOODSRCV_D_S.GRS_PALLET_NO,'000') AND  " & _
                                " ISNULL(WMS_GOODSRCV_D.GRD_BATCH_NO,'') = ISNULL(WMS_GOODSRCV_D_S.GRS_BATCH_NO,'') " & _
                                " WHERE WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' " & whereSQL

                        value &= orderbySQL

                    Case "IB_RO"

                        If HttpContext.Current.Request("sl") <> "" AndAlso HttpContext.Current.Request("sl") = True Then

                            Dim tempStr As String = ""
                            If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                            value &= srchSQL & " WHERE 1=1 " & criteriaSQL

                            If Not HttpContext.Current.Request("cmWH") Is Nothing AndAlso HttpContext.Current.Request("cmWH") <> "" Then
                                tempStr &= " AND WMS_REPLENISH.RO_WH_CODE='" & HttpContext.Current.Request("cmWH") & "' "
                            End If

                            If Not HttpContext.Current.Session("mPDFlag") Is Nothing AndAlso HttpContext.Current.Session("mPDFlag") = "PD" Then
                                tempStr &= " AND WMS_REPLENISH.STORER_CODE='PD' "
                            Else
                                tempStr &= " AND WMS_REPLENISH.STORER_CODE not in ('PD') "
                            End If

                            tempStr &= " AND WMS_REPLENISH.RO_STATUS not in ('CLOSED','CANCELLED') "

                            If whereSQL <> "" Then value &= " AND " & whereSQL
                            value &= tempStr & groupbySQL & orderbySQL
                        End If



                    Case "IB_SR"
                        If HttpContext.Current.Request("sl") <> "" AndAlso HttpContext.Current.Request("sl") = True Then
                            Dim tempStr As String = ""
                            If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                            value &= srchSQL & " WHERE 1=1 " & criteriaSQL

                            If Not HttpContext.Current.Request("cmWH") Is Nothing AndAlso HttpContext.Current.Request("cmWH") <> "" Then
                                tempStr &= " AND WMS_STOCK_RETURN.RT_WH='" & HttpContext.Current.Request("cmWH") & "' "
                            End If

                            If Not HttpContext.Current.Session("mPDFlag") Is Nothing AndAlso HttpContext.Current.Session("mPDFlag") = "PD" Then
                                tempStr &= " AND WMS_STOCK_RETURN.STORER_CODE='PD' "
                            Else
                                tempStr &= " AND WMS_STOCK_RETURN.STORER_CODE not in ('PD') "
                            End If

                            tempStr &= " AND WMS_STOCK_RETURN.RT_STATUS not in ('CLOSED','POSTED','CANCELLED') "

                            If whereSQL <> "" Then value &= " AND " & whereSQL
                            value &= tempStr & groupbySQL & orderbySQL
                        End If



                    Case "OP_WO"
                        If HttpContext.Current.Request("sl") <> "" AndAlso HttpContext.Current.Request("sl") = True Then
                            Dim tempStr As String = ""
                            If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                            value &= srchSQL & " WHERE 1=1 " & criteriaSQL

                            tempStr &= " AND EXISTS (SELECT 1 FROM WMS_STOCK_TRANSFER " & _
       "WHERE M.IMP_CODE = WMS_STOCK_TRANSFER.IMP_CODE AND M.STORER_CODE = WMS_STOCK_TRANSFER.STORER_CODE " & _
       "AND M.WO_TR_CODE = WMS_STOCK_TRANSFER.TR_CODE " & _
       "AND TR_STATUS NOT in ('POSTED','CANCELLED')) AND WO_STATUS NOT IN ('CANCELLED') "

                            If whereSQL <> "" Then value &= " AND " & whereSQL
                            value &= tempStr & " " & groupbySQL & " " & orderbySQL
                        End If


                    Case "OB_CO"
                        If HttpContext.Current.Request("sl") <> "" AndAlso HttpContext.Current.Request("sl") = True Then
                            Dim tempStr As String = ""
                            If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                            value &= srchSQL & " WHERE 1=1 " & criteriaSQL

                            If Not HttpContext.Current.Request("cmWH") Is Nothing AndAlso HttpContext.Current.Request("cmWH") <> "" Then
                                tempStr &= " AND WMS_CUST_ORDER_D.COD_WH_CODE='" & HttpContext.Current.Request("cmWH") & "' "
                            End If

                            If Not HttpContext.Current.Session("mPDFlag") Is Nothing AndAlso HttpContext.Current.Session("mPDFlag") = "PD" Then
                                tempStr &= " AND WMS_CUST_ORDER_D.STORER_CODE='PD' "
                            Else
                                tempStr &= " AND WMS_CUST_ORDER_D.STORER_CODE not in ('PD') "
                            End If

                            tempStr &= " AND WMS_CUST_ORDER.CO_STATUS not in ('CLOSED','CANCELLED') "

                            If whereSQL <> "" Then value &= " AND " & whereSQL
                            value &= tempStr & " " & groupbySQL & " " & orderbySQL
                        End If


                    Case "OB_DO"
                        Dim tempStr As String = ""

                        If HttpContext.Current.Request("sl") <> "" AndAlso HttpContext.Current.Request("sl") = True Then

                            If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL
                            value &= srchSQL & " WHERE 1=1 " & criteriaSQL

                            If Not HttpContext.Current.Request("cmWH") Is Nothing AndAlso HttpContext.Current.Request("cmWH") <> "" Then
                                tempStr &= " AND WMS_DELV_ORDER_D.DOD_WH_CODE='" & HttpContext.Current.Request("cmWH") & "' "
                            End If

                            If Not HttpContext.Current.Session("mPDFlag") Is Nothing AndAlso HttpContext.Current.Session("mPDFlag") = "PD" Then
                                tempStr &= " AND WMS_DELV_ORDER_D.STORER_CODE='PD' "
                            Else
                                tempStr &= " AND WMS_DELV_ORDER_D.STORER_CODE not in ('PD') "
                            End If

                            tempStr &= " AND WMS_DELV_ORDER.DO_STATUS not in ('POSTED','CANCELLED') "

                            If whereSQL <> "" Then value &= " AND " & whereSQL
                            'value &= tempStr & " " & groupbySQL & " " & orderbySQL
                        End If

                        value = " SELECT WMS_DELV_ORDER.DO_CODE,WMS_DELV_ORDER.ROUTE_ID," &
                                " WMS_DELV_ORDER.STORER_CODE," &
                                " WMS_DELV_ORDER.DO_CUS_REF_NO, WMS_DELV_ORDER.DO_INV_NO," &
                                " WMS_DELV_ORDER.DO_STATUS," &
                                " WMS_DELV_ORDER.DO_PROJECT_NO,WMS_DELV_ORDER.DO_COUNTRY_DEL," &
                                " WMS_DELV_ORDER.DO_TRACK_NO,WMS_DELV_ORDER.DO_EDI_SIR_NO,WMS_DELV_ORDER.DO_EDI_WIT_NO,WMS_DELV_ORDER.DO_CO_CODE," &
                                " convert(varchar,WMS_DELV_ORDER.DO_DATE, 103) as DO_DATE," &
                                " WMS_DELV_ORDER.SYS_CD," &
                                " case when DO_TO_STORER_REM is null then NULL else WMS_DELV_ORDER.SYS_LUD end as SYS_LUD," &
                                " convert(varchar,WMS_DELV_ORDER.SYS_CD, 103) + ' ' + convert(varchar,WMS_DELV_ORDER.SYS_CD, 14) as SYS_CD," &
                                " case when DO_STORER_REM is null then NULL else WMS_DELV_ORDER.DO_STORER_LUD end as DO_STORER_LUD ," &
                                " WMS_DELV_ORDER.DO_DELI_STATUS," &
                                " WMS_DELV_ORDER.DO_DELI_RMKS," &
                                " WMS_DELV_ORDER.DO_CONF_DELDATE," &
                                " WMS_STORER.STO_SHORTNAME," &
                                " ISNULL(WMS_DELV_ORDER.DO_DELI_FLAG, 'N') as DO_DELI_FLAG," &
                                " WMS_DELV_ORDER.SYS_CB, WMS_DELV_ORDER.DO_VEHICLE_NO, " &
                                " DO_TO_STORER_REM, DO_STORER_REM, DO_FTRACK_NO, '' as SUB_FTRACK_NO, SUM(DOD_QTY) AS SUM_QTY, DOV_WEIGHT as SUM_KG " &
                                " FROM WMS_DELV_ORDER" &
                                " INNER JOIN WMS_STORER ON WMS_DELV_ORDER.IMP_CODE = WMS_STORER.IMP_CODE" &
                                " AND WMS_DELV_ORDER.STORER_CODE = WMS_STORER.STORER_CODE" &
                                " LEFT OUTER JOIN WMS_DELV_ORDER_D " &
                                " ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE" &
                                " AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE" &
                                " AND WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE" &
                                " LEFT OUTER JOIN WMS_ITEM ON WMS_DELV_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE" &
                                " AND WMS_DELV_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE" &
                                " AND WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY" &
                                " AND WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE" &
                                " LEFT OUTER JOIN WMS_CUST_ORDER_D" &
                                " ON WMS_CUST_ORDER_D.CO_CODE = WMS_DELV_ORDER.DO_CO_CODE" &
                                " AND WMS_CUST_ORDER_D.IMP_CODE = WMS_DELV_ORDER.IMP_CODE" &
                                " AND WMS_CUST_ORDER_D.STORER_CODE = WMS_DELV_ORDER.STORER_CODE" &
                                " AND WMS_CUST_ORDER_D.COD_ITM_CODE = WMS_DELV_ORDER_D.DOD_ITM_CODE" &
                                " AND WMS_CUST_ORDER_D.COD_PACK_KEY = WMS_DELV_ORDER_D.DOD_PACK_KEY" &
                                " AND ISNULL(WMS_CUST_ORDER_D.COD_BATCH_NO,'') = ISNULL(WMS_DELV_ORDER_D.DOD_BATCH_NO,'')" &
                                " AND ISNULL(wms_cust_order_d.cod_pallet_no,'') = ISNULL(wms_delv_order_d.dod_pallet_no,'') " &
                                " LEFT OUTER JOIN (SELECT WMS_DELV_ORDER_VIDEO.DO_CODE,WMS_DELV_ORDER_VIDEO.IMP_CODE,WMS_DELV_ORDER_VIDEO.STORER_CODE,sum(DOV_WEIGHT) as DOV_WEIGHT FROM WMS_DELV_ORDER_VIDEO " &
                                " GROUP BY WMS_DELV_ORDER_VIDEO.DO_CODE,WMS_DELV_ORDER_VIDEO.IMP_CODE,WMS_DELV_ORDER_VIDEO.STORER_CODE) A " &
                                " ON A.DO_CODE = WMS_DELV_ORDER.DO_CODE " &
                                " AND A.IMP_CODE = WMS_DELV_ORDER.IMP_CODE " &
                                " AND A.STORER_CODE = WMS_DELV_ORDER.STORER_CODE " &
                                " WHERE NOT EXISTS (SELECT 1 FROM WMS_DO_BOX_TRACK " &
                                " WHERE WMS_DO_BOX_TRACK.IMP_CODE = WMS_DELV_ORDER.IMP_CODE " &
                                " AND WMS_DO_BOX_TRACK.STORER_CODE = WMS_DELV_ORDER.STORER_CODE " &
                                " AND WMS_DO_BOX_TRACK.DO_CODE = WMS_DELV_ORDER.DO_CODE) "
                        If whereSQL <> "" Then value &= " AND " & whereSQL & tempStr
                        If HttpContext.Current.Session("usr_type") = "T" Then
                            value &= " AND WMS_DELV_ORDER.STORER_CODE = '" & HttpContext.Current.Session("usr_pref_storer") & "' "
                        End If
                        value = value & " GROUP BY WMS_DELV_ORDER.DO_CODE, WMS_DELV_ORDER.ROUTE_ID," &
                                " WMS_DELV_ORDER.STORER_CODE," &
                                " WMS_DELV_ORDER.DO_CUS_REF_NO, WMS_DELV_ORDER.DO_INV_NO," &
                                " WMS_DELV_ORDER.DO_STATUS," &
                                " WMS_DELV_ORDER.DO_PROJECT_NO,WMS_DELV_ORDER.DO_COUNTRY_DEL," &
                                " WMS_DELV_ORDER.DO_TRACK_NO,WMS_DELV_ORDER.DO_EDI_SIR_NO,WMS_DELV_ORDER.DO_EDI_WIT_NO,WMS_DELV_ORDER.DO_CO_CODE," &
                                " convert(varchar,WMS_DELV_ORDER.DO_DATE, 103)," &
                                " WMS_DELV_ORDER.SYS_CD," &
                                " case when DO_TO_STORER_REM is null then NULL else WMS_DELV_ORDER.SYS_LUD end," &
                                " convert(varchar,WMS_DELV_ORDER.SYS_CD, 103) + ' ' + convert(varchar,WMS_DELV_ORDER.SYS_CD, 14)," &
                                " case when DO_STORER_REM is null then NULL else WMS_DELV_ORDER.DO_STORER_LUD end," &
                                " WMS_DELV_ORDER.DO_DELI_STATUS," &
                                " WMS_DELV_ORDER.DO_DELI_RMKS," &
                                " WMS_DELV_ORDER.DO_CONF_DELDATE," &
                                " WMS_STORER.STO_SHORTNAME," &
                                " ISNULL(WMS_DELV_ORDER.DO_DELI_FLAG, 'N')," &
                                " WMS_DELV_ORDER.SYS_CB, WMS_DELV_ORDER.DO_VEHICLE_NO, " &
                                " DO_TO_STORER_REM, DO_STORER_REM, DO_FTRACK_NO,DOV_WEIGHT "
                        value = value & " UNION " &
                                " SELECT DISTINCT WMS_DELV_ORDER.DO_CODE, WMS_DELV_ORDER.ROUTE_ID," &
                                " WMS_DELV_ORDER.STORER_CODE," &
                                " WMS_DELV_ORDER.DO_CUS_REF_NO, WMS_DELV_ORDER.DO_INV_NO," &
                                " WMS_DELV_ORDER.DO_STATUS," &
                                " WMS_DELV_ORDER.DO_PROJECT_NO,WMS_DELV_ORDER.DO_COUNTRY_DEL," &
                                " WMS_DELV_ORDER.DO_TRACK_NO,WMS_DELV_ORDER.DO_EDI_SIR_NO,WMS_DELV_ORDER.DO_EDI_WIT_NO,WMS_DELV_ORDER.DO_CO_CODE," &
                                " convert(varchar,WMS_DELV_ORDER.DO_DATE, 103) as DO_DATE," &
                                " WMS_DELV_ORDER.SYS_CD," &
                                " case when DO_TO_STORER_REM is null then NULL else WMS_DELV_ORDER.SYS_LUD end as SYS_LUD," &
                                " convert(varchar,WMS_DELV_ORDER.SYS_CD, 103) + ' ' + convert(varchar,WMS_DELV_ORDER.SYS_CD, 14) as SYS_CD," &
                                " case when DO_STORER_REM is null then NULL else WMS_DELV_ORDER.DO_STORER_LUD end as DO_STORER_LUD ," &
                                " DOB_STATUS," &
                                " DOB_RMKS," &
                                " DOB_DATE," &
                                " WMS_STORER.STO_SHORTNAME," &
                                " ISNULL(WMS_DELV_ORDER.DO_DELI_FLAG, 'N') as DO_DELI_FLAG," &
                                " WMS_DELV_ORDER.SYS_CB, WMS_DELV_ORDER.DO_VEHICLE_NO, " &
                                " DO_TO_STORER_REM, DO_STORER_REM, DOB_FTRACK_NO, DOB_SUB_FTRACK_NO, SUM(DOD_QTY) AS SUM_QTY, DOV_WEIGHT as SUM_KG " &
                                " FROM WMS_DELV_ORDER" &
                                " INNER JOIN WMS_STORER ON WMS_DELV_ORDER.IMP_CODE = WMS_STORER.IMP_CODE" &
                                " AND WMS_DELV_ORDER.STORER_CODE = WMS_STORER.STORER_CODE" &
                                " INNER JOIN WMS_DO_BOX_TRACK " &
                                " ON WMS_DELV_ORDER.IMP_CODE = WMS_DO_BOX_TRACK.IMP_CODE" &
                                " AND WMS_DELV_ORDER.STORER_CODE = WMS_DO_BOX_TRACK.STORER_CODE" &
                                " AND WMS_DELV_ORDER.DO_CODE = WMS_DO_BOX_TRACK.DO_CODE" &
                                " LEFT OUTER JOIN WMS_DELV_ORDER_D " &
                                " ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE" &
                                " AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE" &
                                " AND WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE" &
                                " LEFT OUTER JOIN WMS_ITEM ON WMS_DELV_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE" &
                                " AND WMS_DELV_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE" &
                                " AND WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY" &
                                " AND WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE" &
                                " LEFT OUTER JOIN WMS_CUST_ORDER_D" &
                                " ON WMS_CUST_ORDER_D.CO_CODE = WMS_DELV_ORDER.DO_CO_CODE" &
                                " AND WMS_CUST_ORDER_D.IMP_CODE = WMS_DELV_ORDER.IMP_CODE" &
                                " AND WMS_CUST_ORDER_D.STORER_CODE = WMS_DELV_ORDER.STORER_CODE" &
                                " AND WMS_CUST_ORDER_D.COD_ITM_CODE = WMS_DELV_ORDER_D.DOD_ITM_CODE" &
                                " AND WMS_CUST_ORDER_D.COD_PACK_KEY = WMS_DELV_ORDER_D.DOD_PACK_KEY" &
                                " AND ISNULL(WMS_CUST_ORDER_D.COD_BATCH_NO,'') = ISNULL(WMS_DELV_ORDER_D.DOD_BATCH_NO,'')" &
                                " AND ISNULL(wms_cust_order_d.cod_pallet_no,'') = ISNULL(wms_delv_order_d.dod_pallet_no,'') " &
                                " LEFT OUTER JOIN (SELECT WMS_DELV_ORDER_VIDEO.DO_CODE,WMS_DELV_ORDER_VIDEO.IMP_CODE,WMS_DELV_ORDER_VIDEO.STORER_CODE,sum(DOV_WEIGHT) as DOV_WEIGHT FROM WMS_DELV_ORDER_VIDEO " &
                                " GROUP BY WMS_DELV_ORDER_VIDEO.DO_CODE,WMS_DELV_ORDER_VIDEO.IMP_CODE,WMS_DELV_ORDER_VIDEO.STORER_CODE) A " &
                                " ON A.DO_CODE = WMS_DELV_ORDER.DO_CODE " &
                                " AND A.IMP_CODE = WMS_DELV_ORDER.IMP_CODE " &
                                " AND A.STORER_CODE = WMS_DELV_ORDER.STORER_CODE " &
                                "WHERE 1 = 1 "
                        If whereSQL <> "" Then value &= " AND " & whereSQL & tempStr
                        If HttpContext.Current.Session("usr_type") = "T" Then
                            value &= " AND WMS_DELV_ORDER.STORER_CODE = '" & HttpContext.Current.Session("usr_pref_storer") & "' "
                        End If
                        value = value & " GROUP BY WMS_DELV_ORDER.DO_CODE, WMS_DELV_ORDER.ROUTE_ID," &
                                " WMS_DELV_ORDER.STORER_CODE," &
                                " WMS_DELV_ORDER.DO_CUS_REF_NO, WMS_DELV_ORDER.DO_INV_NO," &
                                " WMS_DELV_ORDER.DO_STATUS," &
                                " WMS_DELV_ORDER.DO_PROJECT_NO,WMS_DELV_ORDER.DO_COUNTRY_DEL," &
                                " WMS_DELV_ORDER.DO_TRACK_NO,WMS_DELV_ORDER.DO_EDI_SIR_NO,WMS_DELV_ORDER.DO_EDI_WIT_NO,WMS_DELV_ORDER.DO_CO_CODE," &
                                " convert(varchar,WMS_DELV_ORDER.DO_DATE, 103)," &
                                " WMS_DELV_ORDER.SYS_CD," &
                                " case when DO_TO_STORER_REM is null then NULL else WMS_DELV_ORDER.SYS_LUD end," &
                                " convert(varchar,WMS_DELV_ORDER.SYS_CD, 103) + ' ' + convert(varchar,WMS_DELV_ORDER.SYS_CD, 14)," &
                                " case when DO_STORER_REM is null then NULL else WMS_DELV_ORDER.DO_STORER_LUD end," &
                                " DOB_STATUS," &
                                " DOB_RMKS," &
                                " DOB_DATE," &
                                " WMS_STORER.STO_SHORTNAME," &
                                " ISNULL(WMS_DELV_ORDER.DO_DELI_FLAG, 'N')," &
                                " WMS_DELV_ORDER.SYS_CB, WMS_DELV_ORDER.DO_VEHICLE_NO, " &
                                " DO_TO_STORER_REM, DO_STORER_REM, DOB_FTRACK_NO, DOB_SUB_FTRACK_NO,DOV_WEIGHT "

                        value &= " ORDER BY 1 DESC"

                    Case "RPT_CC"
                        Dim tempStr As String = ""

                        value = " SELECT WMS_CC_LOC_SEQ.CCLS_DISPLAY_SEQ, WMS_STOCK_CHECK.CK_CODE, ISNULL(WMS_COL_CODE.COLC_ENG_VALUE, WMS_STOCK_CHECK_D.CKD_STATUS) as CK_STATUS, convert(varchar, WMS_STOCK_CHECK.CK_DATE,103) as CK_DATE, WMS_STOCK_CHECK.CK_WH, " & _
                                " convert(varchar, WMS_STOCK_CHECK.CK_ACTUAL_DATE, 103) as CK_ACTUAL_DATE, convert(varchar,WMS_STOCK_CHECK.CK_START_DATE,103) as CK_START_DATE, WMS_STOCK_CHECK.CK_PERIOD," & _
                                " WMS_STOCK_CHECK_D.CKD_LOC, WMS_STOCK_CHECK_D.CKD_ORG_QTY, WMS_STOCK_CHECK_D.CKD_PACK_KEY, " & _
                                " WMS_STOCK_CHECK_D.CKD_SERIAL_NO, WMS_STOCK_CHECK_D.CKD_CC_DATE, Case when WMS_ITEM.ITM_TYPE = 'CABLE' then WMS_STOCK_CHECK_D.CKD_UOM2 else ITM_UOM end as ITM_UOM, WMS_STOCK_CHECK_D.CKD_UOM2, WMS_STOCK_CHECK_D.CKD_ORG_QTY2, WMS_STOCK_CHECK_D.CKD_REV_QTY2, " & _
                                " WMS_STOCK_CHECK_D.CKD_VAR_QTY2, WMS_STOCK_CHECK_D.CKD_ACTUAL_QTY2, WMS_STOCK_CHECK_D.CKD_BOOK_QTY2, " & _
                                " WMS_STOCK_CHECK_D.CKD_RECHECK_BY, WMS_STOCK_CHECK_D.CKD_BY, WMS_STOCK_CHECK_D.CKD_CC1_DATE, WMS_STOCK_CHECK_D.CKD_WITNESS, WMS_ITEM.ITM_SKU_NO,WMS_ITEM.ITM_NAME, " & _
                                " Case when WMS_ITEM.ITM_TYPE = 'CABLE' then WMS_STOCK_CHECK_D.CKD_BOOK_QTY2 else WMS_STOCK_CHECK_D.CKD_BOOK_QTY end as CKD_BOOK_QTY, " & _
                                " Case when WMS_ITEM.ITM_TYPE = 'CABLE' then WMS_STOCK_CHECK_D.CKD_REV_QTY2 else WMS_STOCK_CHECK_D.CKD_REV_QTY end as CKD_REV_QTY, " & _
                                " Case when WMS_ITEM.ITM_TYPE = 'CABLE' then WMS_STOCK_CHECK_D.CKD_ACTUAL_QTY2 else WMS_STOCK_CHECK_D.CKD_ACTUAL_QTY end as CKD_ACTUAL_QTY, " & _
                                " Case when WMS_ITEM.ITM_TYPE = 'CABLE' then WMS_STOCK_CHECK_D.CKD_REV_QTY2 else WMS_STOCK_CHECK_D.CKD_REV_QTY end - Case when WMS_ITEM.ITM_TYPE = 'CABLE' then WMS_STOCK_CHECK_D.CKD_ORG_QTY2 else WMS_STOCK_CHECK_D.CKD_ORG_QTY end as CHK_VAR, " & _
                                " Case when WMS_ITEM.ITM_TYPE = 'CABLE' then WMS_STOCK_CHECK_D.CKD_VAR_QTY2 else WMS_STOCK_CHECK_D.CKD_VAR_QTY end as CKD_VAR_QTY, " & _
                                " WMS_WH_BIN.BN_CSMS_CODE, WMS_WH_BIN.WH_CODE, " & _
                                " YEAR(WMS_STOCK_CHECK.CK_DATE) as TARGET_START_YEAR, WMS_STOCK_CHECK_D.CKD_IN_PDA, WMS_STOCK_CHECK_D.CKD_RECHK_IN_PDA " & _
                                " FROM WMS_STOCK_CHECK " & _
                                " INNER JOIN WMS_STOCK_CHECK_D ON WMS_STOCK_CHECK.IMP_CODE = WMS_STOCK_CHECK_D.IMP_CODE AND WMS_STOCK_CHECK.STORER_CODE = WMS_STOCK_CHECK_D.STORER_CODE AND " & _
                                " WMS_STOCK_CHECK.CK_CODE = WMS_STOCK_CHECK_D.CK_CODE " & _
                                " INNER JOIN WMS_ITEM ON WMS_STOCK_CHECK_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_CHECK_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                                " WMS_STOCK_CHECK_D.CKD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_CHECK_D.CKD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " LEFT OUTER JOIN WMS_CC_LOC_SEQ ON ISNULL(WMS_STOCK_CHECK.CK_PERIOD,'ALL') = WMS_CC_LOC_SEQ.CCLS_PERIOD AND WMS_STOCK_CHECK_D.CKD_LOC = WMS_CC_LOC_SEQ.CCLS_LOC " & _
                                " AND ISNULL(year(WMS_STOCK_CHECK.CK_ACTUAL_DATE), year(getdate())) = WMS_CC_LOC_SEQ.CCLS_YEAR " & _
                                " LEFT OUTER JOIN WMS_COL_CODE ON WMS_STOCK_CHECK_D.CKD_STATUS = WMS_COL_CODE.COLC_CODE AND WMS_COL_CODE.COLC_TABCOL = 'WMS_STOCK_CHECK_D.CKD_STATUS' " & _
                                " LEFT OUTER JOIN WMS_WH_BIN ON WMS_STOCK_CHECK_D.CKD_LOC = WMS_WH_BIN.LOC_KEY " & _
                                " WHERE 1=1 "

                        If whereSQL <> "" Then value &= " AND " & whereSQL

                        value &= orderbySQL

                    Case "OP_MTL"

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_LAST_MTL")) Then
                            Dim addSQL As String = ""

                            If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_STORER_CODE")) Then
                                addSQL &= " AND WMS_MTL.STORER_CODE='" & gU.dbEncode(HttpContext.Current.Session("SEARCH_SESSION_PAGE_STORER_CODE")) & "' "
                            End If

                            value = " SELECT WMS_STOCK_TRANS_REQ.TQ_EDI_RFT_NO, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_MTL.IMP_CODE, WMS_MTL.STORER_CODE, WMS_MTL.MTL_NO, " & _
                                    " WMS_MTL.MTL_DATE, WMS_MTL.MTL_REF_NO, WMS_MTL.MTL_STATUS, WMS_MTL.MTL_BY, WMS_MTL.MTL_COLLECTED_DATE, WMS_MTL.MTL_WH, WMS_MTL.MTL_TO_WH, " & _
                                    " WMS_MTL.MTL_REMARKS, WMS_MTL.MTL_TOTAL_ITEMS, WMS_MTL.MTL_TOTAL_PCS, WMS_MTL_DTL.MTLD_SEQ, WMS_MTL_DTL.MTLD_LIST_ID, WMS_MTL_DTL.MTLD_PALLET_NO, " & _
                                    " WMS_MTL_DTL.MTLD_DATE_TIME, WMS_MTL_DTL.MTLD_DOC_TYPE, WMS_MTL_DTL.MTLD_DOC_NO, WMS_MTL_DTL.MTLD_DOC_SEQ, WMS_MTL_DTL.MTLD_ITM_CODE, " & _
                                    " WMS_MTL_DTL.MTLD_PACK_KEY, WMS_MTL_DTL.MTLD_BATCH_NO, WMS_MTL_DTL.MTLD_W, WMS_MTL_DTL.MTLD_L, WMS_MTL_DTL.MTLD_H, WMS_MTL_DTL.MTLD_CBM, WMS_MTL_DTL.MTLD_KG, " & _
                                    " WMS_MTL_DTL.MTLD_ORG_QTY1, WMS_MTL_DTL.MTLD_ORG_UOM1, WMS_MTL_DTL.MTLD_ORG_QTY2, WMS_MTL_DTL.MTLD_ORG_UOM2, WMS_MTL_DTL.MTLD_ORG_QTY3, " & _
                                    " WMS_MTL_DTL.MTLD_ORG_UOM3, WMS_MTL_DTL.MTLD_QTY1, WMS_MTL_DTL.MTLD_UOM1, WMS_MTL_DTL.MTLD_QTY2, WMS_MTL_DTL.MTLD_UOM2, WMS_MTL_DTL.MTLD_QTY3, " & _
                                    " WMS_MTL_DTL.MTLD_UOM3, WMS_MTL_DTL.MTLD_DESC, WMS_MTL_DTL.MTLD_STATUS, WMS_MTL_DTL.MTLD_REMARKS, WMS_REPLENISH.RO_EDI_PO_NO, " & _
                                    " WMS_MTL.SYS_CB, WMS_MTL.SYS_CD " & _
                                    " FROM WMS_MTL LEFT OUTER JOIN WMS_MTL_DTL ON WMS_MTL.IMP_CODE = WMS_MTL_DTL.IMP_CODE AND WMS_MTL.STORER_CODE = WMS_MTL_DTL.STORER_CODE AND WMS_MTL.MTL_NO = WMS_MTL_DTL.MTL_NO " & _
                                    " LEFT OUTER JOIN WMS_REPLENISH ON WMS_MTL_DTL.IMP_CODE = WMS_REPLENISH.IMP_CODE AND WMS_MTL_DTL.STORER_CODE = WMS_REPLENISH.STORER_CODE AND " & _
                                    " WMS_MTL_DTL.MTLD_DOC_NO = WMS_REPLENISH.RO_CODE AND WMS_MTL_DTL.MTLD_DOC_TYPE = 'RO' " & _
                                    " LEFT OUTER JOIN WMS_STOCK_TRANS_REQ ON WMS_MTL_DTL.MTLD_DOC_NO = WMS_STOCK_TRANS_REQ.TQ_CODE AND WMS_MTL_DTL.IMP_CODE = WMS_STOCK_TRANS_REQ.IMP_CODE AND " & _
                                    " WMS_MTL_DTL.STORER_CODE = WMS_STOCK_TRANS_REQ.STORER_CODE AND WMS_MTL_DTL.MTLD_DOC_TYPE = 'STQ' " & _
                                    " LEFT OUTER JOIN WMS_ITEM ON WMS_MTL_DTL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_MTL_DTL.STORER_CODE = WMS_ITEM.STORER_CODE AND WMS_MTL_DTL.MTLD_ITM_CODE = WMS_ITEM.ITM_CODE AND " & _
                                    " WMS_MTL_DTL.MTLD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                    " INNER JOIN (Select max(mtl_no) as mtl_no from wms_mtl where MTL_STATUS not in ('CLOSED', 'CANCELLED')) t " & _
                                    " ON WMS_MTL.MTL_NO = t.MTL_NO "
                            If addSQL <> "" Then value &= " WHERE 1=1 " & addSQL
                        End If

                    Case "LOOKUP_MTLRO"

                        Dim tempStr As String = ""
                        If criteriaSQL <> "" Then criteriaSQL = " and " & criteriaSQL

                        value &= srchSQL & " WHERE 1=1 "

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("MTL_MWH")) Then
                            If InStr(value, "MASTERSQL") > 0 Then
                                tempStr = " AND WMS_REPLENISH.RO_WH_CODE <> '" & gU.dbEncode(HttpContext.Current.Session("MTL_MWH")) & "' "

                            ElseIf InStr(value, "DETAILSQL") > 0 Then
                                tempStr = " AND WMS_REPLENISH_D.ROD_WH_CODE <> '" & gU.dbEncode(HttpContext.Current.Session("MTL_MWH")) & "' "
                            End If
                        End If

                        If whereSQL <> "" Then value &= " AND " & whereSQL
                        value &= tempStr & criteriaSQL & " " & groupbySQL & " " & orderbySQL

                    Case "RPT_CSMS_DSCP_NEW"

                        Dim tempStr As String = ""

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_CSBA_LOC")) Then

                            tempStr = HttpContext.Current.Session("SEARCH_SESSION_PAGE_CSBA_LOC")
                            tempStr = " AND WMS_WAREHOUSE.WH_MAIN_WH IN ('" & tempStr.Replace(", ", "', '") & "') "

                        End If

                        If whereSQL <> "" Then whereSQL = " AND " & whereSQL

                        value = " SELECT WMS_CSMS_BAL.IMP_CODE, WMS_CSMS_BAL.STORER_CODE, WMS_CSMS_BAL.ITM_CODE, WMS_CSMS_BAL.PACK_KEY, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, SUM(WMS_CSMS_BAL.CSBA_CSMS_BAL_QTY) AS CSMS_BAL_QTY, SUM(WMS_CSMS_BAL.CSBA_WMS_BAL_QTY) " & _
                                " AS CSBA_WMS_BAL_QTY, ABS(SUM(WMS_CSMS_BAL.CSBA_CSMS_BAL_QTY) - SUM(WMS_CSMS_BAL.CSBA_WMS_BAL_QTY)) AS OLD_disc_qty, SUM(ITM_BAL.ILOC_BAL_QTY) as ILOC_BAL_QTY, SUM(PICKED_DO.PLD_ITEM_QTY) as PICKED_QTY,  " & _
                                " SUM(ITM_BAL_INSP.ILOC_BAL_QTY) as INSP_QTY, SUM(STK_RTD.RTD_RCV_QTY) as RTD_QTY, " & _
                                " ISNULL(SUM(ITM_BAL.ILOC_BAL_QTY),0) + ISNULL(SUM(PICKED_DO.PLD_ITEM_QTY),0) as UNRES_PICK, (ISNULL(SUM(ITM_BAL.ILOC_BAL_QTY),0) + ISNULL(SUM(PICKED_DO.PLD_ITEM_QTY),0)) - ISNULL(SUM(WMS_CSMS_BAL.CSBA_CSMS_BAL_QTY),0) as DISC_QTY, " & _
                                " '' as BLK_QTY " & _
                                " FROM WMS_CSMS_BAL LEFT OUTER JOIN " & _
                                " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_CSMS_BAL.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_CSMS_BAL.STORER_CODE AND  " & _
                                " WMS_ITEM.ITM_CODE = WMS_CSMS_BAL.ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_CSMS_BAL.PACK_KEY " & _
                                " LEFT OUTER JOIN (SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY,  " & _
                                " sum(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY) as ILOC_BAL_QTY" & _
                                " FROM WMS_ITEM_LOC_BAL INNER JOIN " & _
                                " WMS_WH_AREA ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE AND WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM INNER JOIN " & _
                                " WMS_WAREHOUSE ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE " & _
                                " WHERE ISNULL(WMS_WH_AREA.AR_INSP_AREA, 'N') <> 'Y' " & tempStr & " " & _
                                " group by WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY) ITM_BAL ON " & _
                                " ITM_BAL.IMP_CODE = WMS_CSMS_BAL.IMP_CODE AND ITM_BAL.STORER_CODE = WMS_CSMS_BAL.STORER_CODE AND  " & _
                                " ITM_BAL.ITM_CODE = WMS_CSMS_BAL.ITM_CODE AND ITM_BAL.PACK_KEY = WMS_CSMS_BAL.PACK_KEY   " & _
                                " LEFT OUTER JOIN  (SELECT WMS_DO_PICKLIST_D.PLD_ITEM_QTY, WMS_DO_PICKLIST_D.IMP_CODE, WMS_DO_PICKLIST_D.STORER_CODE, WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_DO_PICKLIST_D.PLD_PACK_KEY " & _
                                " FROM WMS_DELV_ORDER INNER JOIN " & _
                                " WMS_DO_PICKLIST_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND  " & _
                                " WMS_DELV_ORDER.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " & _
                                " LEFT OUTER JOIN WMS_WH_BIN ON WMS_WH_BIN.LOC_KEY = WMS_DO_PICKLIST_D.PLD_LOC AND WMS_WH_BIN.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_WAREHOUSE.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_WAREHOUSE.WH_CODE = WMS_WH_BIN.WH_CODE " & _
                                " WHERE WMS_DELV_ORDER.DO_STATUS = 'PICKED'" & tempStr & ") PICKED_DO ON  " & _
                                " PICKED_DO.IMP_CODE = WMS_CSMS_BAL.IMP_CODE AND PICKED_DO.STORER_CODE = WMS_CSMS_BAL.STORER_CODE AND  " & _
                                " PICKED_DO.PLD_ITEM_NO = WMS_CSMS_BAL.ITM_CODE AND PICKED_DO.PLD_PACK_KEY = WMS_CSMS_BAL.PACK_KEY   " & _
                                " LEFT OUTER JOIN (SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY,  " & _
                                " sum(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY) as ILOC_BAL_QTY " & _
                                " FROM WMS_ITEM_LOC_BAL INNER JOIN " & _
                                " WMS_WH_AREA ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE AND WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM INNER JOIN " & _
                                " WMS_WAREHOUSE ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE " & _
                                " WHERE  ISNULL(WMS_WH_AREA.AR_INSP_AREA, 'N') = 'Y' " & tempStr & " " & _
                                " group by WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY) ITM_BAL_INSP ON " & _
                                " ITM_BAL_INSP.IMP_CODE = WMS_CSMS_BAL.IMP_CODE AND ITM_BAL_INSP.STORER_CODE = WMS_CSMS_BAL.STORER_CODE AND  " & _
                                " ITM_BAL_INSP.ITM_CODE = WMS_CSMS_BAL.ITM_CODE AND ITM_BAL_INSP.PACK_KEY = WMS_CSMS_BAL.PACK_KEY " & _
                                " LEFT OUTER JOIN (SELECT WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY,  " & _
                                " sum(WMS_STOCK_RETURN_D.RTD_RCV_QTY) as RTD_RCV_QTY " & _
                                " FROM WMS_STOCK_RETURN INNER JOIN " & _
                                " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE " & _
                                " LEFT OUTER JOIN WMS_WH_BIN ON WMS_WH_BIN.LOC_KEY = WMS_STOCK_RETURN_D.RTD_LOC AND WMS_WH_BIN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_WAREHOUSE.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_WAREHOUSE.WH_CODE = WMS_WH_BIN.WH_CODE " & _
                                " WHERE (WMS_STOCK_RETURN.RT_STATUS in ('APPROVED', 'PENDING') OR (WMS_STOCK_RETURN.RT_STATUS = 'POSTED' AND isnull(WMS_STOCK_RETURN.RT_CUS_CODE,'') = '')) " & tempStr & " " & _
                                " group by WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY) STK_RTD " & _
                                " ON STK_RTD.IMP_CODE = WMS_CSMS_BAL.IMP_CODE AND STK_RTD.STORER_CODE = WMS_CSMS_BAL.STORER_CODE AND  " & _
                                " STK_RTD.RTD_ITM_CODE = WMS_CSMS_BAL.ITM_CODE AND STK_RTD.RTD_PACK_KEY = WMS_CSMS_BAL.PACK_KEY   " & _
                                " WHERE (WMS_CSMS_BAL.BATCH_NO IN " & _
                                " (SELECT CAST(MAX(IMP_BATCH_ID) AS varchar) AS BATCH_NO " & _
                                " FROM WMS_INTF_LOG " & _
                                " WHERE (ITF_IMP_TYPE = 'IM') AND (ITF_STATUS = 'SUCCESS'))) " & _
                                " and isnull(WMS_ITEM.ITM_SERIAL_NO_YN,'N') <> 'Y' " & whereSQL & _
                                " group by WMS_CSMS_BAL.IMP_CODE, WMS_CSMS_BAL.STORER_CODE, WMS_CSMS_BAL.ITM_CODE, WMS_CSMS_BAL.PACK_KEY, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC " & _
                                " UNION " & _
                                "   SELECT WMS_CSMS_BAL.IMP_CODE, WMS_CSMS_BAL.STORER_CODE, WMS_CSMS_BAL.ITM_CODE, WMS_CSMS_BAL.PACK_KEY,WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, SUM(WMS_CSMS_BAL.CSBA_CSMS_BAL_QTY) AS CSMS_BAL_QTY, SUM(WMS_CSMS_BAL.CSBA_WMS_BAL_QTY)  " & _
                                " AS CSBA_WMS_BAL_QTY, ABS(SUM(WMS_CSMS_BAL.CSBA_CSMS_BAL_QTY) - SUM(WMS_CSMS_BAL.CSBA_WMS_BAL_QTY)) AS OLD_disc_qty, SUM(ITM_BAL.ILBS_QTY2) as ILOC_BAL_QTY, SUM(PICKED_DO.PLD_QTY2) as PICKED_QTY,  " & _
                                " SUM(ITM_BAL_INSP.ILBS_QTY2) as INSP_QTY, SUM(STK_RTD.RTD_QTY2) as RTD_QTY, " & _
                                " ISNULL(SUM(ITM_BAL.ILBS_QTY2),0) + isnull(SUM(PICKED_DO.PLD_QTY2),0) as UNRES_PICK, (isnull(SUM(ITM_BAL.ILBS_QTY2),0) + isnull(SUM(PICKED_DO.PLD_QTY2),0)) - isnull(SUM(WMS_CSMS_BAL.CSBA_CSMS_BAL_QTY),0) as DISC_QTY, " & _
                                " '' as BLK_QTY " & _
                                " FROM WMS_CSMS_BAL LEFT OUTER JOIN " & _
                                " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_CSMS_BAL.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_CSMS_BAL.STORER_CODE AND  " & _
                                " WMS_ITEM.ITM_CODE = WMS_CSMS_BAL.ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_CSMS_BAL.PACK_KEY " & _
                                " LEFT OUTER JOIN (SELECT WMS_ITEM_LOC_BAL_S.IMP_CODE, WMS_ITEM_LOC_BAL_S.STORER_CODE, WMS_ITEM_LOC_BAL_S.ITM_CODE, WMS_ITEM_LOC_BAL_S.PACK_KEY, sum(WMS_ITEM_LOC_BAL_S.ILBS_QTY2) as ILBS_QTY2" & _
                                " FROM WMS_ITEM_LOC_BAL_S INNER JOIN " & _
                                " WMS_ITEM_LOC_BAL ON WMS_ITEM_LOC_BAL_S.ILOC_SEQ = WMS_ITEM_LOC_BAL.ILOC_SEQ INNER JOIN " & _
                                " WMS_WH_AREA ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE INNER JOIN " & _
                                " WMS_WAREHOUSE ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE " & _
                                " WHERE WMS_WH_AREA.AR_INSP_AREA <> 'Y' group by WMS_ITEM_LOC_BAL_S.IMP_CODE, WMS_ITEM_LOC_BAL_S.STORER_CODE, WMS_ITEM_LOC_BAL_S.ITM_CODE, WMS_ITEM_LOC_BAL_S.PACK_KEY) ITM_BAL ON " & _
                                " ITM_BAL.IMP_CODE = WMS_CSMS_BAL.IMP_CODE AND ITM_BAL.STORER_CODE = WMS_CSMS_BAL.STORER_CODE AND  " & _
                                " ITM_BAL.ITM_CODE = WMS_CSMS_BAL.ITM_CODE AND ITM_BAL.PACK_KEY = WMS_CSMS_BAL.PACK_KEY   " & _
                                " LEFT OUTER JOIN  (SELECT WMS_DO_PICKLIST_D.PLD_QTY2, WMS_DO_PICKLIST_D.IMP_CODE, WMS_DO_PICKLIST_D.STORER_CODE, WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_DO_PICKLIST_D.PLD_PACK_KEY " & _
                                " FROM WMS_DELV_ORDER INNER JOIN " & _
                                " WMS_DO_PICKLIST_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND  " & _
                                " WMS_DELV_ORDER.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " & _
                                " LEFT OUTER JOIN WMS_WH_BIN ON WMS_WH_BIN.LOC_KEY = WMS_DO_PICKLIST_D.PLD_LOC AND WMS_WH_BIN.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_WAREHOUSE.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_WAREHOUSE.WH_CODE = WMS_WH_BIN.WH_CODE " & _
                                " WHERE WMS_DELV_ORDER.DO_STATUS = 'PICKED') PICKED_DO ON  " & _
                                " PICKED_DO.IMP_CODE = WMS_CSMS_BAL.IMP_CODE AND PICKED_DO.STORER_CODE = WMS_CSMS_BAL.STORER_CODE AND  " & _
                                " PICKED_DO.PLD_ITEM_NO = WMS_CSMS_BAL.ITM_CODE AND PICKED_DO.PLD_PACK_KEY = WMS_CSMS_BAL.PACK_KEY   " & _
                                " LEFT OUTER JOIN (SELECT WMS_ITEM_LOC_BAL_S.IMP_CODE, WMS_ITEM_LOC_BAL_S.STORER_CODE, WMS_ITEM_LOC_BAL_S.ITM_CODE, WMS_ITEM_LOC_BAL_S.PACK_KEY, sum(WMS_ITEM_LOC_BAL_S.ILBS_QTY2) as ILBS_QTY2 " & _
                                " FROM WMS_ITEM_LOC_BAL_S INNER JOIN " & _
                                " WMS_ITEM_LOC_BAL ON WMS_ITEM_LOC_BAL_S.ILOC_SEQ = WMS_ITEM_LOC_BAL.ILOC_SEQ INNER JOIN " & _
                                " WMS_WH_AREA ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE INNER JOIN " & _
                                " WMS_WAREHOUSE ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE " & _
                                " WHERE WMS_WH_AREA.AR_INSP_AREA = 'Y' group by WMS_ITEM_LOC_BAL_S.IMP_CODE, WMS_ITEM_LOC_BAL_S.STORER_CODE, WMS_ITEM_LOC_BAL_S.ITM_CODE, WMS_ITEM_LOC_BAL_S.PACK_KEY) ITM_BAL_INSP ON " & _
                                " ITM_BAL_INSP.IMP_CODE = WMS_CSMS_BAL.IMP_CODE AND ITM_BAL_INSP.STORER_CODE = WMS_CSMS_BAL.STORER_CODE AND  " & _
                                " ITM_BAL_INSP.ITM_CODE = WMS_CSMS_BAL.ITM_CODE AND ITM_BAL_INSP.PACK_KEY = WMS_CSMS_BAL.PACK_KEY " & _
                                " LEFT OUTER JOIN (SELECT WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY,  " & _
                                " sum(WMS_STOCK_RETURN_D.RTD_QTY2) as RTD_QTY2 " & _
                                " FROM WMS_STOCK_RETURN INNER JOIN " & _
                                " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE " & _
                                " LEFT OUTER JOIN WMS_WH_BIN ON WMS_WH_BIN.LOC_KEY = WMS_STOCK_RETURN_D.RTD_LOC AND WMS_WH_BIN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE ON WMS_WAREHOUSE.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_WAREHOUSE.WH_CODE = WMS_WH_BIN.WH_CODE " & _
                                " WHERE (WMS_STOCK_RETURN.RT_STATUS in ('APPROVED', 'PENDING') OR (WMS_STOCK_RETURN.RT_STATUS = 'POSTED' AND WMS_STOCK_RETURN.RT_CUS_CODE is null)) " & _
                                " group by WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY) STK_RTD " & _
                                " ON STK_RTD.IMP_CODE = WMS_CSMS_BAL.IMP_CODE AND STK_RTD.STORER_CODE = WMS_CSMS_BAL.STORER_CODE AND  " & _
                                " STK_RTD.RTD_ITM_CODE = WMS_CSMS_BAL.ITM_CODE AND STK_RTD.RTD_PACK_KEY = WMS_CSMS_BAL.PACK_KEY   " & _
                                " WHERE (WMS_CSMS_BAL.BATCH_NO IN " & _
                                " (SELECT CAST(MAX(IMP_BATCH_ID) AS varchar) AS BATCH_NO " & _
                                " FROM WMS_INTF_LOG " & _
                                " WHERE (ITF_IMP_TYPE = 'IM') AND (ITF_STATUS = 'SUCCESS'))) " & _
                                " and isnull(WMS_ITEM.ITM_SERIAL_NO_YN,'N') = 'Y' " & whereSQL & _
                                " group by WMS_CSMS_BAL.IMP_CODE, WMS_CSMS_BAL.STORER_CODE, WMS_CSMS_BAL.ITM_CODE, WMS_CSMS_BAL.PACK_KEY,WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO, WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC "

                        If orderbySQL = "" Then orderbySQL = " ORDER BY ITM_SKU_NO "

                        value &= orderbySQL

                    Case "HKE_STK_BAL"

                        Dim tempLocBalStr As String = ""
                        Dim tempROStr As String = ""
                        Dim tempSRStr As String = ""
                        Dim tempSTFStr As String = ""

                        If HttpContext.Current.Session("usr_type") = "T" Then

                        End If



                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_RT_TYPE")) Then

                            tempSRStr = HttpContext.Current.Session("SEARCH_SESSION_PAGE_RT_TYPE")
                            tempSRStr = " AND WMS_STOCK_RETURN.RT_TYPE IN ('" & tempSRStr.Replace(", ", "', '") & "') "

                            If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_STORER_CODE")) Then
                                tempSRStr = " AND WMS_STOCK_RETURN.STORER_CODE = '" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_STORER_CODE") & "' "
                            End If

                            tempLocBalStr = " AND EXISTS (Select 1 FROM WMS_STOCK_RETURN INNER JOIN " & _
                                            " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND " & _
                                            " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE INNER JOIN " & _
                                            " WMS_ITEM_LOC_BAL ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND " & _
                                            " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY AND " & _
                                            " WMS_STOCK_RETURN_D.RTD_PALLET_NO = WMS_ITEM_LOC_BAL.ILOC_PALLET_NO AND WMS_STOCK_RETURN_D.RTD_BATCH_NO = WMS_ITEM_LOC_BAL.ILOC_BATCH_NO " & _
                                            " WHERE WMS_STOCK_RETURN.RT_STATUS in ('APPROVED', 'PENDING') " & tempSRStr & ")"

                            tempROStr = " AND EXISTS(SELECT 1 FROM WMS_STOCK_RETURN INNER JOIN " & _
                                        " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND " & _
                                        " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE INNER JOIN " & _
                                        " WMS_REPLENISH_D ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_REPLENISH_D.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_REPLENISH_D.STORER_CODE AND " & _
                                        " WMS_STOCK_RETURN_D.RTD_PALLET_NO = WMS_REPLENISH_D.ROD_PALLET_NO AND WMS_STOCK_RETURN_D.RTD_BATCH_NO = WMS_REPLENISH_D.ROD_BATCH_NO AND " & _
                                        " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_REPLENISH_D.ROD_ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_REPLENISH_D.ROD_PACK_KEY " & _
                                        " WHERE WMS_STOCK_RETURN.RT_STATUS IN ('APPROVED', 'PENDING') " & tempSRStr & ") "

                            tempSTFStr = " AND EXISTS(SELECT 1 FROM WMS_STOCK_RETURN INNER JOIN " & _
                                         " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND " & _
                                         " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE INNER JOIN " & _
                                         " WMS_STOCK_TRANSFER_D ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_STOCK_TRANSFER_D.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_STOCK_TRANSFER_D.STORER_CODE AND " & _
                                         " WMS_STOCK_RETURN_D.RTD_PALLET_NO = WMS_STOCK_TRANSFER_D.TRD_PALLET_NO_FR AND WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_STOCK_TRANSFER_D.ITM_CODE AND " & _
                                         " WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_STOCK_TRANSFER_D.PACK_KEY AND WMS_STOCK_RETURN_D.RTD_BATCH_NO = WMS_STOCK_TRANSFER_D.TRD_BATCH_NO_FR " & _
                                         " WHERE WMS_STOCK_RETURN.RT_STATUS IN ('APPROVED', 'PENDING') " & tempSRStr & ") "
                        End If

                        If whereSQL <> "" Then whereSQL = " AND " & whereSQL

                        'If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_STORER_CODE")) Then
                        'whereSQL = " AND WMS_ITEM.STORER_CODE = '" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_STORER_CODE") & "' "
                        'End If

                        value = " Select WH_MAIN_WH, TOTAL_BAL.IMP_CODE, TOTAL_BAL.STORER_CODE, ITM_CODE, PACK_KEY, ITM_NAME, ITM_DESC, ITM_GP_CODE, ITM_DIV_CODE, ITM_UOM, ITM_SKU_NO, " & _
                                " SUM(ISNULL(ILOC_BAL_QTY,0)) as ILOC_BAL_QTY, SUM(ISNULL(UNRES_QTY,0)) - SUM(ISNULL(PICKED_QTY,0)) as UNRES_QTY, SUM(ISNULL(INSP_QTY,0)) as INSP_QTY, SUM(ISNULL(PICKED_QTY,0)) as PICKED_QTY, " & _
                                " SUM(ISNULL(BLK_QTY,0)) as BLK_QTY, SUM(ISNULL(REST_QTY,0)) as REST_QTY, SUM(ISNULL(PEND_RT_QTY,0)) as PEND_RT_QTY, SUM(ISNULL(RSV_QTY,0)) AS RSV_QTY, SUM(ISNULL(INTRANS_QTY,0)) AS INTRANS_QTY, SUM(ISNULL(ON_ORDER_QTY,0)) AS ON_ORDER_QTY, SUM(ISNULL(BACK_ORDER_QTY,0)) AS BACK_ORDER_QTY,  " & _
                                " SUM(ISNULL(UNRES_QTY,0)) + SUM(ISNULL(INSP_QTY,0)) + SUM(ISNULL(BLK_QTY,0)) + SUM(ISNULL(REST_QTY,0)) + SUM(ISNULL(PEND_RT_QTY,0)) as TOTAL_ON_HAND, WMS_STORER.STO_NAME, " & _
                                " CASE WHEN ITM_TYPE = 'CABLE' THEN 'Cable' ELSE 'General' END AS ITM_TYPE, Convert(varchar, ILOC_EXPIRY_DATE ,103) as ILOC_EXPIRY_DATE , Convert(varchar, ILOC_MANU_DATE,103) as ILOC_MANU_DATE FROM " & _
                                " (" & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " SUM(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY) as ILOC_BAL_QTY, NULL as UNRES_QTY, NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " FROM WMS_ITEM_LOC_BAL " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE " & _
                                " ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE " & _
                                " AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE ILOC_LOC NOT IN ('HK010RJ00R1000') AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) AND (WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0) " & whereSQL & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " UNION ALL " & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, SUM(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY) as UNRES_QTY, NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE  " & _
                                " FROM WMS_ITEM_LOC_BAL " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE " & _
                                " ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE " & _
                                " AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " INNER JOIN WMS_WH_AREA " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE " & _
                                " WHERE ILOC_LOC NOT IN ('HK010RJ00R1000') AND WMS_WH_AREA.AR_INSP_AREA <> 'Y' " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) AND (WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0) " & whereSQL & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " UNION ALL " & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY, NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " SUM(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY) as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE  " & _
                                " FROM WMS_ITEM_LOC_BAL " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE " & _
                                " ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE " & _
                                " AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " INNER JOIN WMS_WH_AREA " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE " & _
                                " WHERE ILOC_LOC = ('HK010RJ00R1000') " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) AND (WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0) " & whereSQL & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " UNION ALL " & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY, SUM(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY) as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " FROM WMS_ITEM_LOC_BAL " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE " & _
                                " ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE " & _
                                " AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " INNER JOIN WMS_WH_AREA " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE " & _
                                " WHERE ILOC_LOC NOT IN ('HK010RJ00R1000') AND WMS_WH_AREA.AR_INSP_AREA = 'Y' " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) AND (WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0) " & whereSQL & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " UNION ALL " & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_DO_PICKLIST_D.IMP_CODE, WMS_DO_PICKLIST_D.STORER_CODE, WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_DO_PICKLIST_D.PLD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_DO_PICKLIST_D.PLD_PALLET_NO, WMS_DO_PICKLIST_D.PLD_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY, NULL as INSP_QTY, SUM(WMS_DO_PICKLIST_D.PLD_ITEM_QTY) as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, NULL, NULL " & _
                                " FROM WMS_DELV_ORDER INNER JOIN WMS_DO_PICKLIST_D " & _
                                " ON WMS_DELV_ORDER.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND  " & _
                                " WMS_DELV_ORDER.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE " & _
                                " ON WMS_DO_PICKLIST_D.PLD_WH = WMS_WAREHOUSE.WH_CODE " & _
                                " AND WMS_DO_PICKLIST_D.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_DO_PICKLIST_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_DO_PICKLIST_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_DO_PICKLIST_D.PLD_ITEM_NO = WMS_ITEM.ITM_CODE AND WMS_DO_PICKLIST_D.PLD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE WMS_DELV_ORDER.DO_STATUS = 'PICKED' " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) AND (WMS_DO_PICKLIST_D.PLD_ITEM_QTY > 0) " & whereSQL & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_DO_PICKLIST_D.IMP_CODE, WMS_DO_PICKLIST_D.STORER_CODE, WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_DO_PICKLIST_D.PLD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_DO_PICKLIST_D.PLD_PALLET_NO, WMS_DO_PICKLIST_D.PLD_BATCH_NO, WMS_ITEM.ITM_TYPE " & _
                                " UNION ALL " & _
                                " SELECT WMS_CUST_ORDER_D.COD_WH_CODE, WMS_CUST_ORDER_D.IMP_CODE, WMS_CUST_ORDER_D.STORER_CODE, WMS_CUST_ORDER_D.COD_ITM_CODE, WMS_CUST_ORDER_D.COD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_CUST_ORDER_D.COD_PALLET_NO, WMS_CUST_ORDER_D.COD_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY, NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, SUM(WMS_CUST_ORDER_D.COD_QTY - ISNULL(WMS_CUST_ORDER_D.COD_POST_QTY,0)) as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, NULL, NULL " & _
                                " FROM WMS_CUST_ORDER INNER JOIN WMS_CUST_ORDER_D " & _
                                " ON WMS_CUST_ORDER.IMP_CODE = WMS_CUST_ORDER_D.IMP_CODE AND WMS_CUST_ORDER.STORER_CODE = WMS_CUST_ORDER_D.STORER_CODE AND  " & _
                                " WMS_CUST_ORDER.CO_CODE = WMS_CUST_ORDER_D.CO_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_CUST_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_CUST_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_CUST_ORDER_D.COD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_CUST_ORDER_D.COD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE WMS_CUST_ORDER.CO_STATUS NOT IN ('CANCELLED','CLOSED') " & _
                                " AND 1 = 2 AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) AND (WMS_CUST_ORDER_D.COD_QTY - ISNULL(WMS_CUST_ORDER_D.COD_POST_QTY,0) > 0) " & Replace(whereSQL, "WH_MAIN_WH", "COD_WH_CODE") & _
                                " GROUP BY WMS_CUST_ORDER_D.COD_WH_CODE, WMS_CUST_ORDER_D.IMP_CODE, WMS_CUST_ORDER_D.STORER_CODE, WMS_CUST_ORDER_D.COD_ITM_CODE, WMS_CUST_ORDER_D.COD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_CUST_ORDER_D.COD_PALLET_NO, WMS_CUST_ORDER_D.COD_BATCH_NO, WMS_ITEM.ITM_TYPE " & _
                                " UNION ALL " & _
                                " SELECT  WMS_STOCK_RETURN.RT_WH, WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_STOCK_RETURN_D.RTD_PALLET_NO, WMS_STOCK_RETURN_D.RTD_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY,NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, SUM(WMS_STOCK_RETURN_D.RTD_RCV_QTY) as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, NULL, NULL " & _
                                " FROM WMS_STOCK_RETURN INNER JOIN " & _
                                " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE INNER JOIN " & _
                                " WMS_ITEM ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE WMS_STOCK_RETURN.RT_STATUS in ('APPROVED', 'PENDING') " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) " & Replace(whereSQL, "WH_MAIN_WH", "RT_WH") & " " & tempSRStr & _
                                " GROUP BY WMS_STOCK_RETURN.RT_WH, WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_STOCK_RETURN_D.RTD_PALLET_NO, WMS_STOCK_RETURN_D.RTD_BATCH_NO,WMS_ITEM.ITM_TYPE, WMS_ITEM.ITM_TYPE " & _
                                " UNION ALL " & _
                                " SELECT WMS_STOCK_TRANSFER.TR_WH_FR, WMS_STOCK_TRANSFER_D.IMP_CODE, WMS_STOCK_TRANSFER_D.STORER_CODE,WMS_STOCK_TRANSFER_D.ITM_CODE, WMS_STOCK_TRANSFER_D.PACK_KEY, WMS_ITEM.ITM_NAME, " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO, TRD_PALLET_NO_FR, TRD_BATCH_NO_FR, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY,NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, SUM(TRD_QTY) as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, NULL, NULL " & _
                                " FROM WMS_STOCK_TRANSFER INNER JOIN " & _
                                " WMS_STOCK_TRANSFER_D ON WMS_STOCK_TRANSFER.IMP_CODE = WMS_STOCK_TRANSFER_D.IMP_CODE AND WMS_STOCK_TRANSFER.STORER_CODE = WMS_STOCK_TRANSFER_D.STORER_CODE AND " & _
                                " WMS_STOCK_TRANSFER.TR_CODE = WMS_STOCK_TRANSFER_D.TR_CODE INNER JOIN " & _
                                " WMS_ITEM ON WMS_STOCK_TRANSFER_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_TRANSFER_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                                " WMS_STOCK_TRANSFER_D.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_TRANSFER_D.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE Substring(WMS_STOCK_TRANSFER_D.TR_CODE,1,2) ='IT' " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) AND WMS_STOCK_TRANSFER.TR_STATUS = 'ISSUED' " & Replace(whereSQL, "WH_MAIN_WH", "TR_WH_FR") & " " & tempSTFStr & _
                                " Group by WMS_STOCK_TRANSFER.TR_WH_FR, WMS_STOCK_TRANSFER_D.IMP_CODE, WMS_STOCK_TRANSFER_D.STORER_CODE,WMS_STOCK_TRANSFER_D.ITM_CODE, WMS_STOCK_TRANSFER_D.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO, TRD_PALLET_NO_FR, TRD_BATCH_NO_FR, WMS_ITEM.ITM_TYPE " & _
                                " UNION ALL " & _
                                " SELECT WMS_REPLENISH_D.rod_wh_code, WMS_REPLENISH_D.IMP_CODE, WMS_REPLENISH_D.STORER_CODE,WMS_REPLENISH_D.ROD_ITM_CODE, WMS_REPLENISH_D.ROD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO, WMS_REPLENISH_D.ROD_PALLET_NO, WMS_REPLENISH_D.ROD_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY,NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, SUM(ISNULL(WMS_REPLENISH_D.ROD_QTY,0) - ISNULL(WMS_REPLENISH_D.ROD_POST_QTY,0)) as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, NULL, NULL " & _
                                " FROM WMS_REPLENISH INNER JOIN " & _
                                " WMS_REPLENISH_D ON WMS_REPLENISH.IMP_CODE = WMS_REPLENISH_D.IMP_CODE AND WMS_REPLENISH.STORER_CODE = WMS_REPLENISH_D.STORER_CODE AND  " & _
                                " WMS_REPLENISH.RO_CODE = WMS_REPLENISH_D.RO_CODE INNER JOIN " & _
                                " WMS_ITEM ON WMS_REPLENISH_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_REPLENISH_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_REPLENISH_D.ROD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_REPLENISH_D.ROD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE RO_STATUS NOT IN ('CANCELLED','CLOSED') AND RO_ONBEHALF_RO_NO is NULL " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) AND ISNULL(WMS_REPLENISH_D.ROD_QTY,0) - ISNULL(WMS_REPLENISH_D.ROD_POST_QTY,0) > 0 " & Replace(whereSQL, "WH_MAIN_WH", "rod_wh_code") & " " & tempROStr & _
                                " Group by WMS_REPLENISH_D.rod_wh_code, WMS_REPLENISH_D.IMP_CODE, WMS_REPLENISH_D.STORER_CODE,WMS_REPLENISH_D.ROD_ITM_CODE, WMS_REPLENISH_D.ROD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO, WMS_REPLENISH_D.ROD_PALLET_NO, WMS_REPLENISH_D.ROD_BATCH_NO,WMS_ITEM.ITM_TYPE " & _
                                " UNION ALL " & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " SUM(WMS_ITEM_LOC_BAL_S.ILBS_QTY2) as ILOC_BAL_QTY, NULL as UNRES_QTY, NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE  " & _
                                " FROM WMS_ITEM_LOC_BAL  " & _
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S " & _
                                " ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE " & _
                                " ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE" & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE ILOC_LOC NOT IN ('HK010RJ00R1000') AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'Y') AND (WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0) " & whereSQL & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " UNION ALL " & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, SUM(WMS_ITEM_LOC_BAL_S.ILBS_QTY2) as UNRES_QTY, NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY , WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " FROM WMS_ITEM_LOC_BAL  " & _
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE " & _
                                " ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE" & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " INNER JOIN WMS_WH_AREA ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE " & _
                                " WHERE ILOC_LOC NOT IN ('HK010RJ00R1000') AND WMS_WH_AREA.AR_INSP_AREA <> 'Y' " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'Y') AND (WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0) AND (WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0) " & whereSQL & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " UNION ALL " & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY, NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " SUM(WMS_ITEM_LOC_BAL_S.ILBS_QTY2) as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY , WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " FROM WMS_ITEM_LOC_BAL  " & _
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE " & _
                                " ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE" & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " INNER JOIN WMS_WH_AREA ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE " & _
                                " WHERE ILOC_LOC IN ('HK010RJ00R1000') " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'Y') AND (WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0) AND (WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0) " & whereSQL & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " UNION ALL " & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY, SUM(WMS_ITEM_LOC_BAL_S.ILBS_QTY2) as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE  " & _
                                " FROM WMS_ITEM_LOC_BAL  " & _
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE " & _
                                " ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE" & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " INNER JOIN WMS_WH_AREA ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE " & _
                                " WHERE ILOC_LOC NOT IN ('HK010RJ00R1000') AND WMS_WH_AREA.AR_INSP_AREA = 'Y' " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'Y') AND (WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0) AND (WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0) " & whereSQL & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM.ITM_TYPE, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE " & _
                                " UNION ALL " & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_DO_PICKLIST_D.IMP_CODE, WMS_DO_PICKLIST_D.STORER_CODE, WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_DO_PICKLIST_D.PLD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_DO_PICKLIST_D.PLD_PALLET_NO, WMS_DO_PICKLIST_D.PLD_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY, NULL as INSP_QTY, SUM(WMS_DO_PICKLIST_D.PLD_QTY2) as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, NULL, NULL " & _
                                " FROM WMS_DELV_ORDER INNER JOIN WMS_DO_PICKLIST_D " & _
                                " ON WMS_DELV_ORDER.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND  " & _
                                " WMS_DELV_ORDER.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " & _
                                " LEFT OUTER JOIN WMS_WAREHOUSE " & _
                                " ON WMS_DO_PICKLIST_D.PLD_WH = WMS_WAREHOUSE.WH_CODE " & _
                                " AND WMS_DO_PICKLIST_D.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_DO_PICKLIST_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_DO_PICKLIST_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_DO_PICKLIST_D.PLD_ITEM_NO = WMS_ITEM.ITM_CODE AND WMS_DO_PICKLIST_D.PLD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE WMS_DELV_ORDER.DO_STATUS = 'PICKED' " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'Y') AND (WMS_DO_PICKLIST_D.PLD_QTY2 > 0) " & whereSQL & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_DO_PICKLIST_D.IMP_CODE, WMS_DO_PICKLIST_D.STORER_CODE, WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_DO_PICKLIST_D.PLD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_DO_PICKLIST_D.PLD_PALLET_NO, WMS_DO_PICKLIST_D.PLD_BATCH_NO, WMS_ITEM.ITM_TYPE " & _
                                " UNION ALL " & _
                                " SELECT WMS_CUST_ORDER_D.COD_WH_CODE, WMS_CUST_ORDER_D.IMP_CODE, WMS_CUST_ORDER_D.STORER_CODE, WMS_CUST_ORDER_D.COD_ITM_CODE, WMS_CUST_ORDER_D.COD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_CUST_ORDER_D.COD_PALLET_NO, WMS_CUST_ORDER_D.COD_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY, NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, SUM((WMS_CUST_ORDER_D.COD_QTY - ISNULL(WMS_CUST_ORDER_D.COD_POST_QTY,0))*WMS_CUST_ORDER_D.COD_QTY2) as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, NULL, NULL " & _
                                " FROM WMS_CUST_ORDER INNER JOIN WMS_CUST_ORDER_D " & _
                                " ON WMS_CUST_ORDER.IMP_CODE = WMS_CUST_ORDER_D.IMP_CODE AND WMS_CUST_ORDER.STORER_CODE = WMS_CUST_ORDER_D.STORER_CODE AND  " & _
                                " WMS_CUST_ORDER.CO_CODE = WMS_CUST_ORDER_D.CO_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_CUST_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_CUST_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_CUST_ORDER_D.COD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_CUST_ORDER_D.COD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE WMS_CUST_ORDER.CO_STATUS NOT IN ('CANCELLED','CLOSED') " & _
                                " AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'Y') AND (WMS_CUST_ORDER_D.COD_QTY - ISNULL(WMS_CUST_ORDER_D.COD_POST_QTY,0) > 0) " & Replace(whereSQL, "WH_MAIN_WH", "COD_WH_CODE") & _
                                " GROUP BY WMS_CUST_ORDER_D.COD_WH_CODE, WMS_CUST_ORDER_D.IMP_CODE, WMS_CUST_ORDER_D.STORER_CODE, WMS_CUST_ORDER_D.COD_ITM_CODE, WMS_CUST_ORDER_D.COD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_CUST_ORDER_D.COD_PALLET_NO, WMS_CUST_ORDER_D.COD_BATCH_NO, WMS_ITEM.ITM_TYPE " & _
                                " UNION ALL " & _
                                " SELECT  WMS_STOCK_RETURN.RT_WH, WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_STOCK_RETURN_D.RTD_PALLET_NO, WMS_STOCK_RETURN_D.RTD_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY,NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, SUM(WMS_STOCK_RETURN_D.RTD_QTY2) as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, NULL, NULL " & _
                                " FROM WMS_STOCK_RETURN " & _
                                " INNER JOIN WMS_STOCK_RETURN_D " & _
                                " ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE WMS_STOCK_RETURN.RT_STATUS in ('APPROVED', 'PENDING') AND WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' " & Replace(whereSQL, "WH_MAIN_WH", "RT_WH") & " " & tempSRStr & _
                                " GROUP BY WMS_STOCK_RETURN.RT_WH, WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_STOCK_RETURN_D.RTD_PALLET_NO, WMS_STOCK_RETURN_D.RTD_BATCH_NO,WMS_ITEM.ITM_TYPE " & _
                                " UNION ALL " & _
                                " SELECT WMS_STOCK_TRANSFER.TR_WH_FR, WMS_STOCK_TRANSFER_D.IMP_CODE, WMS_STOCK_TRANSFER_D.STORER_CODE,WMS_STOCK_TRANSFER_D.ITM_CODE, WMS_STOCK_TRANSFER_D.PACK_KEY, WMS_ITEM.ITM_NAME, " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO, TRD_PALLET_NO_FR, TRD_BATCH_NO_FR, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY,NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, SUM(TRD_QTY2) as INTRANS_QTY, NULL as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, NULL, NULL " & _
                                " FROM WMS_STOCK_TRANSFER INNER JOIN WMS_STOCK_TRANSFER_D " & _
                                " ON WMS_STOCK_TRANSFER.IMP_CODE = WMS_STOCK_TRANSFER_D.IMP_CODE AND WMS_STOCK_TRANSFER.STORER_CODE = WMS_STOCK_TRANSFER_D.STORER_CODE AND " & _
                                " WMS_STOCK_TRANSFER.TR_CODE = WMS_STOCK_TRANSFER_D.TR_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_STOCK_TRANSFER_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_TRANSFER_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                                " WMS_STOCK_TRANSFER_D.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_TRANSFER_D.PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE Substring(WMS_STOCK_TRANSFER_D.TR_CODE,1,2) ='IT' AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'Y') AND WMS_STOCK_TRANSFER.TR_STATUS = 'ISSUED' " & Replace(whereSQL, "WH_MAIN_WH", "TR_WH_FR") & " " & tempSTFStr & _
                                " Group by WMS_STOCK_TRANSFER.TR_WH_FR, WMS_STOCK_TRANSFER_D.IMP_CODE, WMS_STOCK_TRANSFER_D.STORER_CODE,WMS_STOCK_TRANSFER_D.ITM_CODE, WMS_STOCK_TRANSFER_D.PACK_KEY, WMS_ITEM.ITM_NAME, " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO, TRD_PALLET_NO_FR, TRD_BATCH_NO_FR, WMS_ITEM.ITM_TYPE " & _
                                " UNION ALL " & _
                                " SELECT WMS_REPLENISH_D.rod_wh_code, WMS_REPLENISH_D.IMP_CODE, WMS_REPLENISH_D.STORER_CODE,WMS_REPLENISH_D.ROD_ITM_CODE, WMS_REPLENISH_D.ROD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO, WMS_REPLENISH_D.ROD_PALLET_NO, WMS_REPLENISH_D.ROD_BATCH_NO, WMS_ITEM.ITM_TYPE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY,NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, NULL as RSV_QTY, NULL as INTRANS_QTY, SUM((ISNULL(WMS_REPLENISH_D.ROD_QTY,0) - ISNULL(WMS_REPLENISH_D.ROD_POST_QTY,0)) * ISNULL(WMS_REPLENISH_D.ROD_QTY2,0)) as ON_ORDER_QTY, NULL as BACK_ORDER_QTY, NULL, NULL " & _
                                " FROM WMS_REPLENISH INNER JOIN WMS_REPLENISH_D " & _
                                " ON WMS_REPLENISH.IMP_CODE = WMS_REPLENISH_D.IMP_CODE AND WMS_REPLENISH.STORER_CODE = WMS_REPLENISH_D.STORER_CODE AND  " & _
                                " WMS_REPLENISH.RO_CODE = WMS_REPLENISH_D.RO_CODE " & _
                                " INNER JOIN WMS_ITEM " & _
                                " ON WMS_REPLENISH_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_REPLENISH_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_REPLENISH_D.ROD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_REPLENISH_D.ROD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " WHERE WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' AND RO_STATUS NOT IN ('CANCELLED','CLOSED') AND RO_ONBEHALF_RO_NO is NULL " & _
                                " AND ((ISNULL(WMS_REPLENISH_D.ROD_QTY,0) - ISNULL(WMS_REPLENISH_D.ROD_POST_QTY,0)) * ISNULL(WMS_REPLENISH_D.ROD_QTY2,0)) > 0 " & Replace(whereSQL, "WH_MAIN_WH", "rod_wh_code") & " " & tempROStr & _
                                " Group by WMS_REPLENISH_D.rod_wh_code, WMS_REPLENISH_D.IMP_CODE, WMS_REPLENISH_D.STORER_CODE,WMS_REPLENISH_D.ROD_ITM_CODE, WMS_REPLENISH_D.ROD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO, WMS_REPLENISH_D.ROD_PALLET_NO, WMS_REPLENISH_D.ROD_BATCH_NO,WMS_ITEM.ITM_TYPE) TOTAL_BAL " & _
                                " LEFT OUTER JOIN WMS_STORER ON " & _
                                " TOTAL_BAL.IMP_CODE = WMS_STORER.IMP_CODE And TOTAL_BAL.STORER_CODE = WMS_STORER.STORER_CODE " & _
                                " WHERE 1=1 " & Replace(whereSQL, "WMS_ITEM.STORER_CODE", "TOTAL_BAL.STORER_CODE") & _
                                " GROUP BY WH_MAIN_WH, TOTAL_BAL.IMP_CODE, TOTAL_BAL.STORER_CODE, ITM_CODE, PACK_KEY, ITM_NAME, ITM_DESC, ITM_GP_CODE, ITM_DIV_CODE, ITM_UOM, ITM_SKU_NO, WMS_STORER.STO_NAME, " & _
                                " CASE WHEN ITM_TYPE = 'CABLE' THEN 'Cable' ELSE 'General' END, Convert(varchar, ILOC_EXPIRY_DATE ,103), Convert(varchar, ILOC_MANU_DATE,103)"

                        If orderbySQL = "" Then orderbySQL = " ORDER BY ITM_SKU_NO "
                        value &= orderbySQL

                    Case "HKE_INV_RPT"

                        Dim tempLocBalStr As String = ""
                        Dim tempSRStr As String = ""

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_RT_TYPE")) Then

                            tempSRStr = HttpContext.Current.Session("SEARCH_SESSION_PAGE_RT_TYPE")
                            tempSRStr = " AND WMS_STOCK_RETURN.RT_TYPE IN ('" & tempSRStr.Replace(", ", "', '") & "') "

                            tempLocBalStr = " AND EXISTS (Select 1 FROM WMS_STOCK_RETURN INNER JOIN " & _
                                            " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND " & _
                                            " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE INNER JOIN " & _
                                            " WMS_ITEM_LOC_BAL ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND " & _
                                            " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY AND " & _
                                            " WMS_STOCK_RETURN_D.RTD_PALLET_NO = WMS_ITEM_LOC_BAL.ILOC_PALLET_NO AND WMS_STOCK_RETURN_D.RTD_BATCH_NO = WMS_ITEM_LOC_BAL.ILOC_BATCH_NO " & _
                                            " WHERE WMS_STOCK_RETURN.RT_STATUS in ('APPROVED', 'PENDING') " & tempSRStr & ")"

                        End If

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_FR_MANU_DATE")) Then
                            tempLocBalStr &= " AND WMS_ITEM_LOC_BAL.ILOC_MANU_DATE >= CONVERT(DATE,'" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_FR_MANU_DATE") & "'," & gU.getConfig("DDFORMATNO") & ") "
                            tempSRStr &= " AND WMS_STOCK_RETURN_D.RTD_MANU_DATE >= CONVERT(DATE,'" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_FR_MANU_DATE") & "'," & gU.getConfig("DDFORMATNO") & ") "
                        End If

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_TO_MANU_DATE")) Then
                            tempLocBalStr &= " AND WMS_ITEM_LOC_BAL.ILOC_MANU_DATE < CONVERT(DATE,'" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_TO_MANU_DATE") & "'," & gU.getConfig("DDFORMATNO") & ") + 1 "
                            tempSRStr &= " AND WMS_STOCK_RETURN_D.RTD_MANU_DATE < CONVERT(DATE,'" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_TO_MANU_DATE") & "'," & gU.getConfig("DDFORMATNO") & ") + 1 "
                        End If

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_FR_EXP_DATE")) Then
                            tempLocBalStr &= " AND WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE >= CONVERT(DATE,'" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_FR_EXP_DATE") & "'," & gU.getConfig("DDFORMATNO") & ") "
                            tempSRStr &= " AND WMS_STOCK_RETURN_D.RTD_EXPIRY_DATE >= CONVERT(DATE,'" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_FR_EXP_DATE") & "'," & gU.getConfig("DDFORMATNO") & ") "
                        End If

                        If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("SEARCH_SESSION_PAGE_TO_EXP_DATE")) Then
                            tempLocBalStr &= " AND WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE < CONVERT(DATE,'" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_TO_EXP_DATE") & "'," & gU.getConfig("DDFORMATNO") & ") + 1 "
                            tempSRStr &= " AND WMS_STOCK_RETURN_D.RTD_EXPIRY_DATE < CONVERT(DATE,'" & HttpContext.Current.Session("SEARCH_SESSION_PAGE_TO_EXP_DATE") & "'," & gU.getConfig("DDFORMATNO") & ") + 1 "
                        End If


                        If whereSQL <> "" Then whereSQL = " AND " & whereSQL

                        value = " Select WH_MAIN_WH , TOTAL_RECORD.IMP_CODE, TOTAL_RECORD.STORER_CODE, ITM_CODE, PACK_KEY, ITM_NAME, " & _
                                " ITM_DESC, ITM_GP_CODE, ITM_DIV_CODE, ITM_UOM, ITM_SKU_NO, ILOC_PALLET_NO, ILOC_BATCH_NO, Convert(varchar, ILOC_EXPIRY_DATE ,103) as ILOC_EXPIRY_DATE , Convert(varchar, ILOC_MANU_DATE,103) as ILOC_MANU_DATE, " & _
                                " SUM(ISNULL(ILOC_BAL_QTY,0)) as ILOC_BAL_QTY, SUM(ISNULL(UNRES_QTY,0)) - SUM(ISNULL(PICKED_QTY,0)) as UNRES_QTY, SUM(ISNULL(INSP_QTY,0)) as INSP_QTY, SUM(ISNULL(PICKED_QTY,0)) as PICKED_QTY, " & _
                                " SUM(ISNULL(BLK_QTY,0)) as BLK_QTY, SUM(ISNULL(REST_QTY,0)) as REST_QTY, SUM(ISNULL(PEND_RT_QTY,0)) as PEND_RT_QTY, " & _
                                " SUM(ISNULL(UNRES_QTY,0)) + SUM(ISNULL(INSP_QTY,0)) + SUM(ISNULL(BLK_QTY,0)) + SUM(ISNULL(REST_QTY,0)) + SUM(ISNULL(PEND_RT_QTY,0)) as TOTAL_BAL, " & _
                                " ITM_TYPE,WH_CODE,ILOC_LOC as BN_CSMS_CODE,ILBS_SERIAL_NO,ILBS_DRUM_ID,ILOC_LOC,ILOC_FLOOR, ILOC_AREA, ILOC_RACK, ILOC_BIN,WMS_STORER.STO_NAME FROM " & _
                                " (SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE, " & _
                                " SUM(WMS_ITEM_LOC_BAL.ILOC_BAL_QTY) as ILOC_BAL_QTY, SUM(BAL_QTY.ILOC_BAL_QTY) as UNRES_QTY, SUM(INSP_QTY.ILOC_BAL_QTY) as INSP_QTY, SUM(PICKED_BAL.PLD_ITEM_QTY) as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY,NULL as PEND_RT_QTY, " & _
                                " WMS_ITEM.ITM_TYPE, WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE, NULL as ILBS_SERIAL_NO, NULL as ILBS_DRUM_ID, WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM_LOC_BAL.ILOC_FLOOR, WMS_ITEM_LOC_BAL.ILOC_AREA, WMS_ITEM_LOC_BAL.ILOC_RACK, WMS_ITEM_LOC_BAL.ILOC_BIN " & _
                                " FROM WMS_ITEM_LOC_BAL INNER JOIN " & _
                                " WMS_ITEM ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY LEFT OUTER JOIN " & _
                                " WMS_WAREHOUSE ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                                " LEFT OUTER JOIN WMS_WH_BIN ON " & _
                                " WMS_ITEM_LOC_BAL.ILOC_LOC  = WMS_WH_BIN.LOC_KEY AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_BIN.IMP_CODE " & _
                                " LEFT OUTER JOIN  " & _
                                " (SELECT WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, " & _
                                " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM_LOC_BAL.ILOC_BAL_QTY " & _
                                " FROM WMS_ITEM_LOC_BAL INNER JOIN " & _
                                " WMS_WH_AREA ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE " & _
                                " WHERE WMS_WH_AREA.AR_INSP_AREA <> 'Y' AND ILOC_BAL_QTY > 0 ) BAL_QTY   " & _
                                " ON BAL_QTY.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND BAL_QTY.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE " & _
                                " AND BAL_QTY.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND BAL_QTY.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " & _
                                " AND BAL_QTY.ILOC_LOC = WMS_ITEM_LOC_BAL.ILOC_LOC AND ISNULL(BAL_QTY.ILOC_PALLET_NO,'000') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') " & _
                                " AND ISNULL(BAL_QTY.ILOC_BATCH_NO,'') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,'') " & _
                                " LEFT OUTER JOIN  " & _
                                " (SELECT WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, " & _
                                " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM_LOC_BAL.ILOC_BAL_QTY " & _
                                " FROM WMS_ITEM_LOC_BAL INNER JOIN " & _
                                " WMS_WH_AREA ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE " & _
                                " WHERE WMS_WH_AREA.AR_INSP_AREA = 'Y' AND ILOC_BAL_QTY > 0 ) INSP_QTY   " & _
                                " ON INSP_QTY.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND INSP_QTY.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE " & _
                                " AND INSP_QTY.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND INSP_QTY.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " & _
                                " AND INSP_QTY.ILOC_LOC = WMS_ITEM_LOC_BAL.ILOC_LOC AND ISNULL(INSP_QTY.ILOC_PALLET_NO,'000') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') " & _
                                " AND ISNULL(INSP_QTY.ILOC_BATCH_NO,'') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,'') " & _
                                " LEFT OUTER JOIN  " & _
                                " (SELECT WMS_DO_PICKLIST_D.IMP_CODE, WMS_DO_PICKLIST_D.STORER_CODE, WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_DO_PICKLIST_D.PLD_PACK_KEY, WMS_DO_PICKLIST_D.PLD_PALLET_NO,WMS_DO_PICKLIST_D.PLD_LOC,   " & _
                                " SUM(WMS_DO_PICKLIST_D.PLD_ITEM_QTY) as PLD_ITEM_QTY, WMS_DO_PICKLIST_D.PLD_BATCH_NO " & _
                                " FROM WMS_DELV_ORDER INNER JOIN " & _
                                " WMS_DO_PICKLIST_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND  " & _
                                " WMS_DELV_ORDER.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " & _
                                " WHERE WMS_DELV_ORDER.DO_STATUS = 'PICKED' GROUP BY WMS_DO_PICKLIST_D.IMP_CODE, WMS_DO_PICKLIST_D.STORER_CODE, WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_DO_PICKLIST_D.PLD_PACK_KEY, WMS_DO_PICKLIST_D.PLD_PALLET_NO,WMS_DO_PICKLIST_D.PLD_LOC,   " & _
                                " WMS_DO_PICKLIST_D.PLD_BATCH_NO) PICKED_BAL " & _
                                " ON PICKED_BAL.PLD_LOC = WMS_ITEM_LOC_BAL.ILOC_LOC AND PICKED_BAL.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND  " & _
                                " PICKED_BAL.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND PICKED_BAL.PLD_ITEM_NO = WMS_ITEM_LOC_BAL.ITM_CODE AND  " & _
                                " PICKED_BAL.PLD_PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY AND ISNULL(PICKED_BAL.PLD_PALLET_NO,'000') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') AND  " & _
                                " ISNULL(PICKED_BAL.PLD_BATCH_NO,'') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,'') " & _
                                " WHERE (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) AND (WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0) " & tempLocBalStr & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE, WMS_ITEM.ITM_TYPE, WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE, WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM_LOC_BAL.ILOC_FLOOR, WMS_ITEM_LOC_BAL.ILOC_AREA, WMS_ITEM_LOC_BAL.ILOC_RACK, WMS_ITEM_LOC_BAL.ILOC_BIN " & _
                                " UNION ALL " & _
                                " SELECT  WMS_STOCK_RETURN.RT_WH, WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_STOCK_RETURN_D.RTD_PALLET_NO, WMS_STOCK_RETURN_D.RTD_BATCH_NO, NULL as ILOC_EXPIRY_DATE, NULL as ILOC_MANU_DATE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY,NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, SUM(WMS_STOCK_RETURN_D.RTD_RCV_QTY) as PEND_RT_QTY, " & _
                                " WMS_ITEM.ITM_TYPE, WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE,NULL as ILBS_SERIAL_NO, NULL as ILBS_DRUM_ID, WMS_STOCK_RETURN_D.RTD_LOC, WMS_WH_BIN.FL_NUM, WMS_WH_BIN.AR_CODE, WMS_WH_BIN.RK_CODE, WMS_WH_BIN.BN_CODE " & _
                                " FROM WMS_STOCK_RETURN INNER JOIN " & _
                                " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE INNER JOIN " & _
                                " WMS_ITEM ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " LEFT OUTER JOIN WMS_WH_BIN ON " & _
                                " WMS_STOCK_RETURN_D.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_STOCK_RETURN_D.RTD_LOC = WMS_WH_BIN.LOC_KEY " & _
                                " WHERE WMS_STOCK_RETURN.RT_STATUS in ('APPROVED', 'PENDING') AND (WMS_ITEM.ITM_SERIAL_NO_YN = 'N' OR WMS_ITEM.ITM_SERIAL_NO_YN IS NULL) " & tempSRStr & _
                                " GROUP BY WMS_STOCK_RETURN.RT_WH, WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM, WMS_ITEM.ITM_SKU_NO,WMS_STOCK_RETURN_D.RTD_PALLET_NO, WMS_STOCK_RETURN_D.RTD_BATCH_NO,WMS_ITEM.ITM_TYPE, WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE, WMS_STOCK_RETURN_D.RTD_LOC, WMS_WH_BIN.FL_NUM, WMS_WH_BIN.AR_CODE, WMS_WH_BIN.RK_CODE, WMS_WH_BIN.BN_CODE " & _
                                " UNION ALL " & _
                                " SELECT WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE, " & _
                                " SUM(WMS_ITEM_LOC_BAL_S.ILBS_QTY2) as ILOC_BAL_QTY, SUM(BAL_QTY.ILBS_QTY2) as UNRES_QTY, SUM(INSP_QTY.ILBS_QTY2) as INSP_QTY, SUM(PICKED_BAL.PLD_QTY2) as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, NULL as PEND_RT_QTY, " & _
                                " WMS_ITEM.ITM_TYPE,WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE, WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID, WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM_LOC_BAL.ILOC_FLOOR, WMS_ITEM_LOC_BAL.ILOC_AREA, WMS_ITEM_LOC_BAL.ILOC_RACK, WMS_ITEM_LOC_BAL.ILOC_BIN " & _
                                " FROM WMS_ITEM_LOC_BAL  " & _
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                " INNER JOIN " & _
                                " WMS_ITEM ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_ITEM_LOC_BAL.STORER_CODE = WMS_ITEM.STORER_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_ITEM_LOC_BAL.PACK_KEY = WMS_ITEM.PACK_KEY LEFT OUTER JOIN " & _
                                " WMS_WAREHOUSE ON WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WAREHOUSE.WH_CODE AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WAREHOUSE.IMP_CODE " & _
                                " LEFT OUTER JOIN WMS_WH_BIN ON " & _
                                " WMS_ITEM_LOC_BAL.ILOC_LOC  = WMS_WH_BIN.LOC_KEY AND WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_BIN.IMP_CODE " & _
                                " LEFT OUTER JOIN  " & _
                                " (SELECT WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, " & _
                                " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID " & _
                                " FROM WMS_ITEM_LOC_BAL  " & _
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                " INNER JOIN WMS_WH_AREA ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE " & _
                                " WHERE WMS_WH_AREA.AR_INSP_AREA <> 'Y' AND ILOC_BAL_QTY > 0 AND ILBS_QTY2 > 0 ) BAL_QTY   " & _
                                " ON BAL_QTY.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND BAL_QTY.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE " & _
                                " AND BAL_QTY.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND BAL_QTY.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " & _
                                " AND BAL_QTY.ILOC_LOC = WMS_ITEM_LOC_BAL.ILOC_LOC AND ISNULL(BAL_QTY.ILOC_PALLET_NO,'000') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') " & _
                                " AND ISNULL(BAL_QTY.ILOC_BATCH_NO,'') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,'') AND ISNULL(WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO,'') = isNULL(BAL_QTY.ILBS_SERIAL_NO,'') " & _
                                " LEFT OUTER JOIN  " & _
                                " (SELECT WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, " & _
                                " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID " & _
                                " FROM WMS_ITEM_LOC_BAL  " & _
                                " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                " INNER JOIN WMS_WH_AREA ON WMS_ITEM_LOC_BAL.IMP_CODE = WMS_WH_AREA.IMP_CODE AND WMS_ITEM_LOC_BAL.ILOC_WH = WMS_WH_AREA.WH_CODE AND  " & _
                                " WMS_ITEM_LOC_BAL.ILOC_FLOOR = WMS_WH_AREA.FL_NUM AND WMS_ITEM_LOC_BAL.ILOC_AREA = WMS_WH_AREA.AR_CODE " & _
                                " WHERE WMS_WH_AREA.AR_INSP_AREA = 'Y' AND ILOC_BAL_QTY > 0 AND ILBS_QTY2 > 0 ) INSP_QTY   " & _
                                " ON INSP_QTY.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND INSP_QTY.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE " & _
                                " AND INSP_QTY.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND INSP_QTY.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " & _
                                " AND INSP_QTY.ILOC_LOC = WMS_ITEM_LOC_BAL.ILOC_LOC AND ISNULL(INSP_QTY.ILOC_PALLET_NO,'000') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') " & _
                                " AND ISNULL(INSP_QTY.ILOC_BATCH_NO,'') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,'') AND isnull(WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO,'') = isnull(INSP_QTY.ILBS_SERIAL_NO,'') " & _
                                " LEFT OUTER JOIN  " & _
                                " (SELECT WMS_DO_PICKLIST_D.IMP_CODE, WMS_DO_PICKLIST_D.STORER_CODE, WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_DO_PICKLIST_D.PLD_PACK_KEY, WMS_DO_PICKLIST_D.PLD_PALLET_NO,WMS_DO_PICKLIST_D.PLD_LOC,   " & _
                                " WMS_DO_PICKLIST_D.PLD_QTY2, WMS_DO_PICKLIST_D.PLD_BATCH_NO, WMS_DO_PICKLIST_D.PLD_SERIAL_NO " & _
                                " FROM WMS_DELV_ORDER INNER JOIN " & _
                                " WMS_DO_PICKLIST_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DO_PICKLIST_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DO_PICKLIST_D.STORER_CODE AND  " & _
                                " WMS_DELV_ORDER.DO_CODE = WMS_DO_PICKLIST_D.DO_CODE " & _
                                " WHERE WMS_DELV_ORDER.DO_STATUS = 'PICKED') PICKED_BAL " & _
                                " ON PICKED_BAL.PLD_LOC = WMS_ITEM_LOC_BAL.ILOC_LOC AND PICKED_BAL.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND  " & _
                                " PICKED_BAL.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND PICKED_BAL.PLD_ITEM_NO = WMS_ITEM_LOC_BAL.ITM_CODE AND  " & _
                                " PICKED_BAL.PLD_PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY AND ISNULL(PICKED_BAL.PLD_PALLET_NO,'000') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') AND  " & _
                                " ISNULL(PICKED_BAL.PLD_BATCH_NO,'') = ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,'')  AND ISNULL(WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO,'') = isNULL(PICKED_BAL.PLD_SERIAL_NO,'') " & _
                                " WHERE (WMS_ITEM.ITM_SERIAL_NO_YN = 'Y') AND (WMS_ITEM_LOC_BAL.ILOC_BAL_QTY > 0) AND (WMS_ITEM_LOC_BAL_S.ILBS_QTY2 > 0) " & tempLocBalStr & _
                                " GROUP BY WMS_WAREHOUSE.WH_MAIN_WH, WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE, WMS_ITEM.ITM_TYPE, " & _
                                " WMS_WH_BIN.WH_CODE,WMS_WH_BIN.BN_CSMS_CODE, WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID, WMS_ITEM_LOC_BAL.ILOC_LOC, WMS_ITEM_LOC_BAL.ILOC_FLOOR, WMS_ITEM_LOC_BAL.ILOC_AREA, WMS_ITEM_LOC_BAL.ILOC_RACK, WMS_ITEM_LOC_BAL.ILOC_BIN " & _
                                " UNION ALL " & _
                                " SELECT  WMS_STOCK_RETURN.RT_WH, WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY, WMS_ITEM.ITM_NAME,  " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_STOCK_RETURN_D.RTD_PALLET_NO, WMS_STOCK_RETURN_D.RTD_BATCH_NO, NULL as ILOC_EXPIRY_DATE, NULL as ILOC_MANU_DATE, " & _
                                " NULL as ILOC_BAL_QTY, NULL as UNRES_QTY,NULL as INSP_QTY, NULL as PICKED_QTY, " & _
                                " NULL as BLK_QTY, NULL as REST_QTY, SUM(WMS_STOCK_RETURN_D.RTD_QTY2) as PEND_RT_QTY, " & _
                                " WMS_ITEM.ITM_TYPE,WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE, WMS_STOCK_RETURN_D.RTD_SERIAL_NO, WMS_STOCK_RETURN_D.RTD_DRUM_ID, WMS_STOCK_RETURN_D.RTD_LOC, WMS_WH_BIN.FL_NUM, WMS_WH_BIN.AR_CODE, WMS_WH_BIN.RK_CODE, WMS_WH_BIN.BN_CODE " & _
                                " FROM WMS_STOCK_RETURN INNER JOIN " & _
                                " WMS_STOCK_RETURN_D ON WMS_STOCK_RETURN.IMP_CODE = WMS_STOCK_RETURN_D.IMP_CODE AND WMS_STOCK_RETURN.STORER_CODE = WMS_STOCK_RETURN_D.STORER_CODE AND " & _
                                " WMS_STOCK_RETURN.RT_CODE = WMS_STOCK_RETURN_D.RT_CODE INNER JOIN " & _
                                " WMS_ITEM ON WMS_STOCK_RETURN_D.IMP_CODE = WMS_ITEM.IMP_CODE AND WMS_STOCK_RETURN_D.STORER_CODE = WMS_ITEM.STORER_CODE AND " & _
                                " WMS_STOCK_RETURN_D.RTD_ITM_CODE = WMS_ITEM.ITM_CODE AND WMS_STOCK_RETURN_D.RTD_PACK_KEY = WMS_ITEM.PACK_KEY " & _
                                " LEFT OUTER JOIN WMS_WH_BIN ON " & _
                                " WMS_STOCK_RETURN_D.IMP_CODE = WMS_WH_BIN.IMP_CODE AND WMS_STOCK_RETURN_D.RTD_LOC = WMS_WH_BIN.LOC_KEY " & _
                                " WHERE WMS_STOCK_RETURN.RT_STATUS in ('APPROVED', 'PENDING') AND WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' " & tempSRStr & _
                                " GROUP BY WMS_STOCK_RETURN.RT_WH, WMS_STOCK_RETURN_D.IMP_CODE, WMS_STOCK_RETURN_D.STORER_CODE, WMS_STOCK_RETURN_D.RTD_ITM_CODE, WMS_STOCK_RETURN_D.RTD_PACK_KEY, WMS_ITEM.ITM_NAME, " & _
                                " WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_GP_CODE, WMS_ITEM.ITM_DIV_CODE, WMS_ITEM.ITM_UOM2, WMS_ITEM.ITM_SKU_NO,WMS_STOCK_RETURN_D.RTD_PALLET_NO, WMS_STOCK_RETURN_D.RTD_BATCH_NO,WMS_ITEM.ITM_TYPE, WMS_WH_BIN.WH_CODE, WMS_WH_BIN.BN_CSMS_CODE, WMS_STOCK_RETURN_D.RTD_SERIAL_NO, WMS_STOCK_RETURN_D.RTD_DRUM_ID, WMS_STOCK_RETURN_D.RTD_LOC, WMS_WH_BIN.FL_NUM, WMS_WH_BIN.AR_CODE, WMS_WH_BIN.RK_CODE, WMS_WH_BIN.BN_CODE) " & _
                                " TOTAL_RECORD " & _
                                " LEFT OUTER JOIN WMS_STORER ON " & _
                                " TOTAL_RECORD.IMP_CODE = WMS_STORER.IMP_CODE And TOTAL_RECORD.STORER_CODE = WMS_STORER.STORER_CODE " & _
                                " WHERE 1 = 1 " & whereSQL & _
                                " GROUP BY WH_MAIN_WH, TOTAL_RECORD.IMP_CODE, TOTAL_RECORD.STORER_CODE, ITM_CODE, PACK_KEY, ITM_NAME, ITM_DESC, ITM_GP_CODE, ITM_DIV_CODE, ITM_UOM, ITM_SKU_NO, ILOC_PALLET_NO, ILOC_BATCH_NO, Convert(varchar, ILOC_EXPIRY_DATE ,103), Convert(varchar, ILOC_MANU_DATE,103), ITM_TYPE,WH_CODE,BN_CSMS_CODE,ILBS_SERIAL_NO,ILBS_DRUM_ID,ILOC_LOC,ILOC_FLOOR, ILOC_AREA, ILOC_RACK, ILOC_BIN,WMS_STORER.STO_NAME "

                        If orderbySQL = "" Then orderbySQL = " ORDER BY ITM_SKU_NO, WH_MAIN_WH, WH_CODE,BN_CSMS_CODE "
                        value &= orderbySQL



                End Select
            Case 1
                REM======Maintanience=====

        End Select

        If value = "" Then
            value = srchSQL

            If srchSQL.Contains("WHERE") Then
                value += " AND "
            Else
                value += " WHERE 1 = 1 "
            End If

            If whereSQL <> "" Then
                value += "AND " & whereSQL
            End If

            If criteriaSQL <> "" Then
                criteriaSQL = " AND " & criteriaSQL
            End If

            value += criteriaSQL & " " & groupbySQL & " " & orderbySQL
        End If


        Return value.Trim

    End Function

    Public Sub Main_afterSave_Fun(ByVal srch_id As String, Optional ByRef cnn As SqlConnection = Nothing, Optional ByRef transaction As SqlTransaction = Nothing)
        Select Case srch_id

        End Select
    End Sub
End Module
