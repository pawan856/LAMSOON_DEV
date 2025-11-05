Imports Microsoft.VisualBasic
Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Public Class PickListTable
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private gU As New GeneralUtils

    Dim pickListDt, pickSerialDt As DataTable 'pick list datatable which going to be return
    Dim pickList As List(Of DtItems) 'pick list contains selected loc and item, used for sorting, will be converted into pickListDt at last
    'Dim dtList As List(Of DtItems)
    Dim dtDict As Dictionary(Of String, DtItems)    'contains all rows of datatable
    'Dim confirmedDict As Dictionary(Of String, DtItems) 'contains confirmed pick list
    Dim itmDict As Dictionary(Of String, Double)    'list of all items
    'Dim pickedItmDict As Dictionary(Of String, Double)
    'Dim lackItmDict As Dictionary(Of String, Double)    'list of items do not hv a full contain wh
    'Dim locDict As Dictionary(Of String, LocBal)    'list of all locations
    Private itm_combo_limit As Integer = 1    'Limit the max combo count for each item, prevent the combo getting too big and eat all memory

    Private pickOrder As String
    'BY_BATCH_DATE: WMS_DATE_CODE.DC_CONV_DATE
    'BY_IN_DATE:    WMS_IN_TX.SYS_CD
    'BY_EXP_DATE:   WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE
    'BY_MANU_DATE:  WMS_ITEM_LOC_BAL.ILOC_MANU_DATE
    'NIL

    Private pickByBatchNo As Boolean = True

    Dim cnn As SqlConnection
    Dim transaction As SqlTransaction
    Dim lSTORER_CODE, lIMP_CODE, lCO_CODE, lDO_CODE As String

    Private itmWhList As String = ""

    Structure DtItems
        Dim dt_key, itm_key, itm_code, pack_key, pallet_no, loc, wh, floor, area, rack, bin, batch, sku_no, expiry_date, manu_date, bn_csms_code As String
        Dim qty, bal_qty, org_do_qty, org_bal_qty, org_stock_qty, total_bal_qty, hold_qty As Double
        Dim total_foi_qty, total_itm_qty As Double  'For validation use
        Dim dod_disp_seq As Integer
    End Structure

    Structure DtSerialItem
        Dim hasQty2 As Boolean
        Dim itm_key, itm_code, pack_key, pallet_no, batch_no As String
        Dim bal_qty, min_qty2, total_qty2 As Double
        Dim qty2List As List(Of Double)
        Dim dod_disp_seq As Integer
    End Structure


    'Structure LocBal
    '    Dim loc, wh, floor, area, rack, bin, rank As String
    '    'Dim full_itm_key_list, lack_itm_key_list As List(Of String)
    '    Dim isPicked As Boolean
    '    'Dim pickedDict As Dictionary(Of String, Double)
    'End Structure

    Structure Combo
        Dim itmList As List(Of DtItems)
        Dim cmBalQty As Double
        Dim isValid As Boolean
    End Structure

    Structure Rank
        Dim wh, floor, area, rack, bin As Integer
    End Structure

    Public Sub New(ByVal co_code As String, ByVal do_code As String, Optional ByVal storer_code As String = "", Optional ByVal imp_code As String = "", Optional pickOrderType As String = "")
        lSTORER_CODE = storer_code
        lIMP_CODE = imp_code
        lCO_CODE = co_code
        lDO_CODE = do_code

        If pickOrderType <> "" Then
            pickOrder = pickOrderType
        Else
            pickOrder = "BY_IN_DATE"
        End If

        If pickOrderType = "NIL" Then
            itm_combo_limit = 5
        End If
    End Sub

    Public Function generatePickList(ByRef dtl_dt As DataTable, Optional ByRef pConn As SqlConnection = Nothing, Optional ByRef pTransaction As SqlTransaction = Nothing) As DataTable
        Dim selDt As DataTable
        Dim siDict As Dictionary(Of String, DtSerialItem)

        If pConn IsNot Nothing Then
            cnn = pConn
        End If

        If pTransaction IsNot Nothing Then
            transaction = pTransaction
        End If

        Try
            If dtl_dt.Rows.Count = 0 Then
                Return Nothing
            Else
                itmWhList = getItemWh(dtl_dt)

                siDict = getSerialDict(dtl_dt)

                genSerialPickList(siDict)


                selDt = getItemBal(dtl_dt)
                setBalanceTable(selDt)

                Return getPickList()
            End If
        'Catch ex As Exception
        '    Throw ex
        Finally
            If pConn IsNot Nothing Then
                cnn.Close()
                cnn = Nothing
            End If
        End Try
    End Function

    Private Function getItemWh(ByRef dtl_dt As DataTable) As String
        Dim i As Integer
        Dim tmpList, whList As String
        Dim tmpDt As DataTable
        Dim selectSql As String

        tmpList = ""
        whList = ""

        For i = 0 To dtl_dt.Rows.Count - 1
            If Not gU.inList(tmpList, dtl_dt.Rows(i).Item("DOD_WH_CODE").ToString.Trim) Then
                tmpList = gU.appendToList(tmpList, dtl_dt.Rows(i).Item("DOD_WH_CODE").ToString.Trim)
            End If
        Next

        If tmpList <> "" Then
            selectSql = "select distinct WH_CODE " & _
                        "from wms_warehouse " & _
                        "where WH_MAIN_WH in (" & gU.dbConvList(tmpList) & ") "

            tmpDt = gDB.getDataTable(selectSql, cnn, transaction)

            For i = 0 To tmpDt.Rows.Count - 1
                whList = gU.appendToList(whList, tmpDt.Rows(i).Item("WH_CODE").ToString.Trim)
            Next
        End If

        Return whList
    End Function

    Public Function getSerialDict(ByRef dtl_dt As DataTable) As Dictionary(Of String, DtSerialItem)
        Dim itmKey, itmKeyNoBatch As String
        Dim siDict As Dictionary(Of String, DtSerialItem)
        Dim tmpSerialItem As DtSerialItem
        Dim tmpQtyList As List(Of Double)

        siDict = New Dictionary(Of String, DtSerialItem)

        For i = 0 To dtl_dt.Rows.Count - 1
            If DB.decodeDBNull(dtl_dt.Rows(i).Item("DOD_QTY"), 0) > 0 And dtl_dt.Rows(i).Item("mFlag").ToString <> "D" Then
                itmKey = dtl_dt.Rows(i).Item("dod_itm_code").ToString & "#_#" & _
                        dtl_dt.Rows(i).Item("dod_pack_key").ToString & "#_#" & _
                        gU.decodeNullOrEmpty(dtl_dt.Rows(i).Item("dod_pallet_no").ToString, "000")

                itmKeyNoBatch = ""

                '###
                If pickByBatchNo Then
                    If dtl_dt.Rows(i).Item("dod_batch_no").ToString.Trim <> "" Then
                        itmKeyNoBatch = itmKey
                        itmKey = itmKey & "#_#" & dtl_dt.Rows(i).Item("dod_batch_no").ToString.Trim
                    End If
                End If


                If dtl_dt.Rows(i).Item("ITM_SERIAL_NO_YN").ToString.Trim = "Y" Then
                    If Not siDict.ContainsKey(itmKey) Then

                        tmpSerialItem = New DtSerialItem

                        tmpSerialItem.itm_key = itmKey
                        tmpSerialItem.itm_code = dtl_dt.Rows(i).Item("dod_itm_code").ToString.Trim
                        tmpSerialItem.pack_key = dtl_dt.Rows(i).Item("dod_pack_key").ToString.Trim
                        tmpSerialItem.pallet_no = gU.decodeNullOrEmpty(dtl_dt.Rows(i).Item("dod_pallet_no").ToString.Trim, "000")
                        tmpSerialItem.batch_no = dtl_dt.Rows(i).Item("dod_batch_no").ToString.Trim
                        tmpSerialItem.dod_disp_seq = dtl_dt.Rows(i).Item("dod_disp_seq").ToString.Trim

                        'If dtl_dt.Rows(i).Item("DOD_UOM2").ToString.Trim <> "" Then
                        If dtl_dt.Rows(i).Item("ITM_TYPE").ToString.Trim = "CABLE" Then

                            If dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim = "" Then
                                Throw New Exception("Qty 2 cannot be empty for item: " & itmKey)
                            End If

                            tmpSerialItem.hasQty2 = True

                            tmpSerialItem.min_qty2 = CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim)

                            tmpSerialItem.total_qty2 = CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim)

                            tmpQtyList = New List(Of Double)

                            tmpQtyList.Add(CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim))

                            tmpSerialItem.qty2List = tmpQtyList

                            tmpSerialItem.bal_qty = 1
                        Else
                            tmpSerialItem.hasQty2 = False

                            tmpSerialItem.bal_qty = CDbl(dtl_dt.Rows(i).Item("DOD_QTY").ToString.Trim)
                        End If

                        siDict.Add(itmKey, tmpSerialItem)

                    Else
                        tmpSerialItem = siDict.Item(itmKey)

                        If tmpSerialItem.hasQty2 Then
                            If dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim = "" Then
                                Throw New Exception("Qty 2 cannot be empty for item: " & itmKey)
                            End If

                            If CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim) < tmpSerialItem.min_qty2 Then
                                tmpSerialItem.min_qty2 = CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim)
                            End If

                            tmpSerialItem.total_qty2 += CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim)

                            tmpQtyList = tmpSerialItem.qty2List

                            tmpQtyList.Add(CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim))
                        End If
                    End If
                End If
            End If
        Next

        Return siDict

    End Function


    Private Sub genSerialPickList(siDict As Dictionary(Of String, DtSerialItem))
        Dim keys As Dictionary(Of String, DtSerialItem).KeyCollection
        Dim i, j, k As Integer
        Dim selectSql As String
        Dim tmpQtyList As List(Of Double)
        Dim tmpBalDt As DataTable
        Dim pickedSerialKeyList As String
        Dim tmpSerialItem As DtSerialItem

        pickSerialDt = New DataTable

        pickSerialDt.Columns.Add("PLD_ITEM_NO", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("ITM_SKU_NO", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_PACK_KEY", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_PALLET_NO", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_DO_QTY", Type.GetType("System.Double"))
        pickSerialDt.Columns.Add("PLD_QTY2", Type.GetType("System.Double"))
        pickSerialDt.Columns.Add("PLD_SERIAL_NO", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("ITM_SERIAL_NO_YN", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_FOI_QTY", Type.GetType("System.Double"))
        pickSerialDt.Columns.Add("STOCK_QTY", Type.GetType("System.Double"))
        pickSerialDt.Columns.Add("ILOC_BAL_QTY", Type.GetType("System.Double"))
        pickSerialDt.Columns.Add("ILOC_EXPIRY_DATE", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("ILOC_MANU_DATE", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_WH", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_LOC", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_FLOOR", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_AREA", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_RACK", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_BIN", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("PLD_BATCH_NO", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("HOLD_QTY", Type.GetType("System.String"))
        pickSerialDt.Columns.Add("TOTAL_BAL", Type.GetType("System.String"))

        pickSerialDt.Columns.Add("DOD_DISP_SEQ", Type.GetType("System.Int32"))
        pickSerialDt.Columns.Add("BN_CSMS_CODE", Type.GetType("System.String"))

        keys = siDict.Keys

        pickedSerialKeyList = ""

        For i = 0 To keys.Count - 1

            tmpSerialItem = siDict(keys(i))

            selectSql = "select l.ITM_CODE, l.PACK_KEY, l.ILOC_PALLET_NO, 0.0 as COD_QTY, 0.0 as TOTAL_HOLD_QTY, 0.0 as TOTAL_BAL_QTY, 0.0 as AVAIL_BAL," & _
                            "l.ILOC_BAL_QTY, l.ILOC_LOC, " & _
                            "l.ILOC_WH, l.ILOC_FLOOR, l.ILOC_AREA, l.ILOC_RACK, l.ILOC_BIN, l.ILOC_BATCH_NO, " & _
                            "l.ILOC_EXPIRY_DATE, l.ILOC_MANU_DATE, i.ITM_SKU_NO, s.ILBS_SERIAL_NO, pick_item.pld_serial_no, s.ILBS_SEQ, " & _
                            "ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(PICK_ITEM.picked_qty, 0) as STOCK_QTY, " & _
                            "ISNULL(s.ILBS_QTY2, 0) - ISNULL(PICK_ITEM.picked_qty2, 0) as STOCK_QTY2, " & _
                            "Convert(varchar, l.ILOC_EXPIRY_DATE, " & gU.getConfig("DDFORMATNo") & ") as EXPIRY_DATE_STR, " & _
                            "Convert(varchar, l.ILOC_MANU_DATE, " & gU.getConfig("DDFORMATNo") & ") as MANU_DATE_STR, " & _
                            "b.BN_CSMS_CODE " & _
                        "from WMS_ITEM_LOC_BAL l " & _
                        "inner join WMS_ITEM_LOC_BAL_S s " & _
                        "on l.ILOC_SEQ = s.ILOC_SEQ " & _
                        "inner join WMS_WH_AREA a " & _
                        "on l.IMP_CODE = a.IMP_CODE " & _
                            "and l.ILOC_WH = a.WH_CODE " & _
                            "and l.ILOC_FLOOR = a.FL_NUM " & _
                            "and l.ILOC_AREA = a.AR_CODE " & _
                            "and ISNULL(a.AR_DAMAGE_YN, '') <> 'Y' " & _
                        "inner join WMS_WH_BIN b " & _
                        "on l.ILOC_LOC = b.LOC_KEY " & _
                        "inner join WMS_ITEM i " & _
                        "on l.IMP_CODE = i.IMP_CODE " & _
                            "and l.STORER_CODE = i.STORER_CODE " & _
                            "and l.ITM_CODE = i.ITM_CODE " & _
                            "and l.PACK_KEY = i.PACK_KEY " & _
                        "left outer join WMS_DATE_CODE DC " & _
                        "on l.ILOC_BATCH_NO = DC.DC_DATE_CODE " & _
                            "and l.STORER_CODE = DC.STORER_CODE " & _
                            "and l.IMP_CODE = DC.IMP_CODE " & _
                        "LEFT OUTER JOIN ( " & _
                            "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, p.pld_serial_no, sum(p.pld_item_qty) as picked_qty, sum(p.pld_qty2) as picked_qty2 " & _
                            "from wms_do_picklist_d p, wms_delv_order d " & _
                            "where d.imp_code = p.imp_code " & _
                            "and d.storer_code = p.storer_code " & _
                            "and d.do_code = p.do_code " & _
                            "and d.do_status = 'PICKED' " & _
                            "and d.do_code <> '" & gU.dbEncode(lDO_CODE) & "' " & _
                            "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc, p.pld_serial_no) PICK_ITEM " & _
                        "ON l.IMP_CODE = PICK_ITEM.IMP_CODE " & _
                            "AND l.STORER_CODE = PICK_ITEM.STORER_CODE " & _
                            "AND l.ITM_CODE = PICK_ITEM.PLD_ITEM_NO " & _
                            "AND l.PACK_KEY = PICK_ITEM.PLD_PACK_KEY " & _
                            "AND ISNULL(l.ILOC_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " & _
                            "AND ISNULL(l.ILOC_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " & _
                            "AND l.ILOC_LOC = PICK_ITEM.PLD_LOC " & _
                            "AND s.ILBS_SERIAL_NO = PICK_ITEM.PLD_SERIAL_NO "


            If itmWhList <> "" Then
                selectSql = selectSql & _
                            "left outer join wms_item_wh iw " & _
                            "on iw.IMP_CODE = i.IMP_CODE " & _
                            "and iw.STORER_CODE = i.STORER_CODE " & _
                            "and iw.ITM_CODE = i.ITM_CODE " & _
                            "and iw.PACK_KEY = i.PACK_KEY " & _
                            "and iw.WH_CODE in (" & gU.dbConvList(itmWhList) & ") "

            End If

            selectSql = selectSql & _
                        "where l.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
                        "and l.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' " & _
                        "and l.ITM_CODE = '" & gU.dbEncode(tmpSerialItem.itm_code) & "' " & _
                        "and l.PACK_KEY = '" & gU.dbEncode(tmpSerialItem.pack_key) & "' " & _
                        "and isnull(l.ILOC_PALLET_NO, '000') = '" & gU.dbEncode(tmpSerialItem.pallet_no) & "' " & _
                        "and isnull(l.ILOC_BATCH_NO, '') = '" & gU.dbEncode(tmpSerialItem.batch_no) & "' "

            '"and l.ILOC_LOC in (i.itm_pref_loc, i.itm_pref_loc2, case when isnull(i.itm_pref_loc, '') = '' and isnull(i.itm_pref_loc2, '') = '' then l.ILOC_LOC else null end) "

            If pickedSerialKeyList <> "" Then
                selectSql = selectSql & _
                            "and s.ILBS_SEQ not in (" & pickedSerialKeyList & ") "
            End If
            

            If tmpSerialItem.hasQty2 Then
                selectSql = selectSql & _
                            "and ISNULL(s.ILBS_QTY2, 0) - ISNULL(PICK_ITEM.picked_qty2, 0) >= " & gU.dbEncode(tmpSerialItem.min_qty2) & " "
            Else
                'selectSql = selectSql & _
                '            "and ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(PICK_ITEM.picked_qty, 0) > 0 "

                selectSql = selectSql & _
                            "and pick_item.pld_serial_no is null " & _
                            "and ISNULL(s.ILBS_QTY2, 0) > 0 "

                'pick_item.pld_serial_no
            End If

            If itmWhList <> "" Then
                selectSql = selectSql & _
                            "and l.ILOC_WH in (" & gU.dbConvList(itmWhList) & ") " & _
                            "and isnull(iw.IW_PICK_LOC, l.ILOC_LOC) = l.ILOC_LOC "

                '"and (not exists (" & _
                '    "select 1 from wms_item_wh iw " & _
                '    "where iw.IMP_CODE = i.IMP_CODE " & _
                '    "and iw.STORER_CODE = i.STORER_CODE " & _
                '    "and iw.ITM_CODE = i.ITM_CODE " & _
                '    "and iw.PACK_KEY = i.PACK_KEY " & _
                '    "and iw.WH_CODE in (" & gU.dbConvList(itmWhList) & ")) " & _
                '"or exists (" & _
                '    "select 1 from wms_item_wh iw " & _
                '    "where iw.IMP_CODE = i.IMP_CODE " & _
                '    "and iw.STORER_CODE = i.STORER_CODE " & _
                '    "and iw.ITM_CODE = i.ITM_CODE " & _
                '    "and iw.PACK_KEY = i.PACK_KEY " & _
                '    "and iw.WH_CODE in (" & gU.dbConvList(itmWhList) & ") " & _
                '    "and isnull(iw.IW_PICK_LOC, l.ILOC_LOC) = l.ILOC_LOC))) "
            End If

            selectSql = selectSql & _
                        "order by ISNULL(s.ILBS_QTY2, 0) - ISNULL(PICK_ITEM.picked_qty2, 0) "

            tmpBalDt = gDB.getDataTable(selectSql, cnn, transaction)

            tmpQtyList = tmpSerialItem.qty2List

            If tmpSerialItem.hasQty2 Then
                'First find out all just fit cable
                For j = tmpQtyList.Count - 1 To 0 Step -1

                    For k = tmpBalDt.Rows.Count - 1 To 0 Step -1
                        If tmpQtyList(j) = CDbl(tmpBalDt.Rows(k).Item("STOCK_QTY2").ToString.Trim) Then
                            'Not gU.inList(pickedSerialKeyList, tmpBalDt.Rows(k).Item("ILBS_SEQ").ToString.Trim)

                            pickSerialDt.Rows.Add()

                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_ITEM_NO") = tmpBalDt.Rows(k).Item("ITM_CODE").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ITM_SKU_NO") = tmpBalDt.Rows(k).Item("ITM_SKU_NO").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PACK_KEY") = tmpBalDt.Rows(k).Item("PACK_KEY").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PALLET_NO") = tmpBalDt.Rows(k).Item("ILOC_PALLET_NO").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_QTY2") = tmpQtyList(j)
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_SERIAL_NO") = tmpBalDt.Rows(k).Item("ILBS_SERIAL_NO").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FOI_QTY") = 1
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_DO_QTY") = 1
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("STOCK_QTY") = 1
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_BAL_QTY") = 1
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_EXPIRY_DATE") = tmpBalDt.Rows(k).Item("EXPIRY_DATE_STR").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_MANU_DATE") = tmpBalDt.Rows(k).Item("MANU_DATE_STR").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_WH") = tmpBalDt.Rows(k).Item("ILOC_WH").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_LOC") = tmpBalDt.Rows(k).Item("ILOC_LOC").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FLOOR") = tmpBalDt.Rows(k).Item("ILOC_FLOOR").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_AREA") = tmpBalDt.Rows(k).Item("ILOC_AREA").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_RACK") = tmpBalDt.Rows(k).Item("ILOC_RACK").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BIN") = tmpBalDt.Rows(k).Item("ILOC_BIN").ToString.Trim
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BATCH_NO") = tmpBalDt.Rows(k).Item("ILOC_BATCH_NO").ToString.Trim

                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("BN_CSMS_CODE") = tmpBalDt.Rows(k).Item("BN_CSMS_CODE").ToString.Trim

                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("DOD_DISP_SEQ") = tmpSerialItem.dod_disp_seq

                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("HOLD_QTY") = 0
                            pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("TOTAL_BAL") = 1

                            pickedSerialKeyList = gU.appendToList(pickedSerialKeyList, tmpBalDt.Rows(k).Item("ILBS_SEQ").ToString.Trim)

                            tmpBalDt.Rows(k).Delete()

                            tmpBalDt.AcceptChanges()

                            tmpSerialItem.total_qty2 = tmpSerialItem.total_qty2 - tmpQtyList(j)

                            tmpQtyList.RemoveAt(j)

                            Exit For
                        End If
                    Next
                Next

                'If qty2 = total qty2
                If tmpSerialItem.total_qty2 > 0 AndAlso tmpQtyList.Count > 0 Then
                    For k = tmpBalDt.Rows.Count - 1 To 0 Step -1
                        If tmpSerialItem.total_qty2 = CDbl(tmpBalDt.Rows(k).Item("STOCK_QTY2").ToString.Trim) Then

                            For j = tmpQtyList.Count - 1 To 0 Step -1
                                pickSerialDt.Rows.Add()

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_ITEM_NO") = tmpBalDt.Rows(k).Item("ITM_CODE").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ITM_SKU_NO") = tmpBalDt.Rows(k).Item("ITM_SKU_NO").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PACK_KEY") = tmpBalDt.Rows(k).Item("PACK_KEY").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PALLET_NO") = tmpBalDt.Rows(k).Item("ILOC_PALLET_NO").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_QTY2") = tmpQtyList(j)
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_SERIAL_NO") = tmpBalDt.Rows(k).Item("ILBS_SERIAL_NO").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FOI_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_DO_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("STOCK_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_BAL_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_EXPIRY_DATE") = tmpBalDt.Rows(k).Item("EXPIRY_DATE_STR").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_MANU_DATE") = tmpBalDt.Rows(k).Item("MANU_DATE_STR").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_WH") = tmpBalDt.Rows(k).Item("ILOC_WH").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_LOC") = tmpBalDt.Rows(k).Item("ILOC_LOC").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FLOOR") = tmpBalDt.Rows(k).Item("ILOC_FLOOR").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_AREA") = tmpBalDt.Rows(k).Item("ILOC_AREA").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_RACK") = tmpBalDt.Rows(k).Item("ILOC_RACK").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BIN") = tmpBalDt.Rows(k).Item("ILOC_BIN").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BATCH_NO") = tmpBalDt.Rows(k).Item("ILOC_BATCH_NO").ToString.Trim

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("BN_CSMS_CODE") = tmpBalDt.Rows(k).Item("BN_CSMS_CODE").ToString.Trim

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("DOD_DISP_SEQ") = tmpSerialItem.dod_disp_seq

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("HOLD_QTY") = 0
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("TOTAL_BAL") = 1

                                tmpQtyList.RemoveAt(j)
                            Next

                            pickedSerialKeyList = gU.appendToList(pickedSerialKeyList, tmpBalDt.Rows(k).Item("ILBS_SEQ").ToString.Trim)

                            tmpBalDt.Rows(k).Delete()

                            tmpBalDt.AcceptChanges()

                            tmpSerialItem.total_qty2 = 0

                            Exit For
                        End If
                    Next
                End If


                'If qty2 > total qty2
                If tmpSerialItem.total_qty2 > 0 AndAlso tmpQtyList.Count > 0 Then
                    For k = tmpBalDt.Rows.Count - 1 To 0 Step -1
                        If tmpSerialItem.total_qty2 <= CDbl(tmpBalDt.Rows(k).Item("STOCK_QTY2").ToString.Trim) Then

                            For j = tmpQtyList.Count - 1 To 0 Step -1
                                pickSerialDt.Rows.Add()

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_ITEM_NO") = tmpBalDt.Rows(k).Item("ITM_CODE").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ITM_SKU_NO") = tmpBalDt.Rows(k).Item("ITM_SKU_NO").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PACK_KEY") = tmpBalDt.Rows(k).Item("PACK_KEY").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PALLET_NO") = tmpBalDt.Rows(k).Item("ILOC_PALLET_NO").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_QTY2") = tmpQtyList(j)
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_SERIAL_NO") = tmpBalDt.Rows(k).Item("ILBS_SERIAL_NO").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FOI_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_DO_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("STOCK_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_BAL_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_EXPIRY_DATE") = tmpBalDt.Rows(k).Item("EXPIRY_DATE_STR").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_MANU_DATE") = tmpBalDt.Rows(k).Item("MANU_DATE_STR").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_WH") = tmpBalDt.Rows(k).Item("ILOC_WH").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_LOC") = tmpBalDt.Rows(k).Item("ILOC_LOC").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FLOOR") = tmpBalDt.Rows(k).Item("ILOC_FLOOR").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_AREA") = tmpBalDt.Rows(k).Item("ILOC_AREA").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_RACK") = tmpBalDt.Rows(k).Item("ILOC_RACK").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BIN") = tmpBalDt.Rows(k).Item("ILOC_BIN").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BATCH_NO") = tmpBalDt.Rows(k).Item("ILOC_BATCH_NO").ToString.Trim

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("BN_CSMS_CODE") = tmpBalDt.Rows(k).Item("BN_CSMS_CODE").ToString.Trim

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("DOD_DISP_SEQ") = tmpSerialItem.dod_disp_seq

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("HOLD_QTY") = 0
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("TOTAL_BAL") = 1

                                tmpQtyList.RemoveAt(j)
                            Next

                            pickedSerialKeyList = gU.appendToList(pickedSerialKeyList, tmpBalDt.Rows(k).Item("ILBS_SEQ").ToString.Trim)

                            'tmpBalDt.Rows(k).Delete()

                            tmpBalDt.Rows(k).Item("STOCK_QTY2") = CDbl(tmpBalDt.Rows(k).Item("STOCK_QTY2").ToString.Trim) - tmpSerialItem.total_qty2

                            tmpBalDt.AcceptChanges()

                            tmpSerialItem.total_qty2 = 0

                            Exit For
                        End If
                    Next
                End If


                'Pick rest which stock qty2 > req qty2
                If tmpSerialItem.total_qty2 > 0 Then
                    For j = tmpQtyList.Count - 1 To 0 Step -1
                        For k = tmpBalDt.Rows.Count - 1 To 0 Step -1
                            If tmpQtyList(j) <= CDbl(tmpBalDt.Rows(k).Item("STOCK_QTY2").ToString.Trim) Then

                                pickSerialDt.Rows.Add()

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_ITEM_NO") = tmpBalDt.Rows(k).Item("ITM_CODE").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ITM_SKU_NO") = tmpBalDt.Rows(k).Item("ITM_SKU_NO").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PACK_KEY") = tmpBalDt.Rows(k).Item("PACK_KEY").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PALLET_NO") = tmpBalDt.Rows(k).Item("ILOC_PALLET_NO").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_QTY2") = tmpQtyList(j)
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_SERIAL_NO") = tmpBalDt.Rows(k).Item("ILBS_SERIAL_NO").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FOI_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_DO_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("STOCK_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_BAL_QTY") = 1
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_EXPIRY_DATE") = tmpBalDt.Rows(k).Item("EXPIRY_DATE_STR").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_MANU_DATE") = tmpBalDt.Rows(k).Item("MANU_DATE_STR").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_WH") = tmpBalDt.Rows(k).Item("ILOC_WH").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_LOC") = tmpBalDt.Rows(k).Item("ILOC_LOC").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FLOOR") = tmpBalDt.Rows(k).Item("ILOC_FLOOR").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_AREA") = tmpBalDt.Rows(k).Item("ILOC_AREA").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_RACK") = tmpBalDt.Rows(k).Item("ILOC_RACK").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BIN") = tmpBalDt.Rows(k).Item("ILOC_BIN").ToString.Trim
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BATCH_NO") = tmpBalDt.Rows(k).Item("ILOC_BATCH_NO").ToString.Trim

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("BN_CSMS_CODE") = tmpBalDt.Rows(k).Item("BN_CSMS_CODE").ToString.Trim

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("DOD_DISP_SEQ") = tmpSerialItem.dod_disp_seq

                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("HOLD_QTY") = 0
                                pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("TOTAL_BAL") = 1

                                pickedSerialKeyList = gU.appendToList(pickedSerialKeyList, tmpBalDt.Rows(k).Item("ILBS_SEQ").ToString.Trim)

                                'tmpBalDt.Rows(k).Delete()

                                tmpBalDt.Rows(k).Item("STOCK_QTY2") = CDbl(tmpBalDt.Rows(k).Item("STOCK_QTY2").ToString.Trim) - tmpQtyList(j)

                                tmpBalDt.AcceptChanges()

                                tmpSerialItem.total_qty2 = tmpSerialItem.total_qty2 - tmpQtyList(j)

                                tmpQtyList.RemoveAt(j)

                                Exit For
                            End If
                        Next
                    Next
                End If
                
            Else

                Dim genRowCnt As Integer

                If CInt(tmpSerialItem.bal_qty) > tmpBalDt.Rows.Count Then
                    genRowCnt = tmpBalDt.Rows.Count
                Else
                    genRowCnt = CInt(tmpSerialItem.bal_qty)
                End If

                For j = 0 To genRowCnt - 1
                    pickSerialDt.Rows.Add()

                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_ITEM_NO") = tmpBalDt.Rows(j).Item("ITM_CODE").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ITM_SKU_NO") = tmpBalDt.Rows(j).Item("ITM_SKU_NO").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PACK_KEY") = tmpBalDt.Rows(j).Item("PACK_KEY").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PALLET_NO") = tmpBalDt.Rows(j).Item("ILOC_PALLET_NO").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_QTY2") = 1
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_SERIAL_NO") = tmpBalDt.Rows(j).Item("ILBS_SERIAL_NO").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FOI_QTY") = 1
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_DO_QTY") = 1
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("STOCK_QTY") = 1
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_BAL_QTY") = 1
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_EXPIRY_DATE") = tmpBalDt.Rows(j).Item("EXPIRY_DATE_STR").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_MANU_DATE") = tmpBalDt.Rows(j).Item("MANU_DATE_STR").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_WH") = tmpBalDt.Rows(j).Item("ILOC_WH").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_LOC") = tmpBalDt.Rows(j).Item("ILOC_LOC").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FLOOR") = tmpBalDt.Rows(j).Item("ILOC_FLOOR").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_AREA") = tmpBalDt.Rows(j).Item("ILOC_AREA").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_RACK") = tmpBalDt.Rows(j).Item("ILOC_RACK").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BIN") = tmpBalDt.Rows(j).Item("ILOC_BIN").ToString.Trim
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BATCH_NO") = tmpBalDt.Rows(j).Item("ILOC_BATCH_NO").ToString.Trim

                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("BN_CSMS_CODE") = tmpBalDt.Rows(j).Item("BN_CSMS_CODE").ToString.Trim

                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("DOD_DISP_SEQ") = tmpSerialItem.dod_disp_seq

                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("HOLD_QTY") = 0
                    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("TOTAL_BAL") = 1

                    pickedSerialKeyList = gU.appendToList(pickedSerialKeyList, tmpBalDt.Rows(j).Item("ILBS_SEQ").ToString.Trim)
                Next

                'If tmpBalDt.Rows.Count > 0 Then
                '    pickSerialDt.Rows.Add()

                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_ITEM_NO") = tmpBalDt.Rows(0).Item("ITM_CODE").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ITM_SKU_NO") = tmpBalDt.Rows(0).Item("ITM_SKU_NO").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PACK_KEY") = tmpBalDt.Rows(0).Item("PACK_KEY").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_PALLET_NO") = tmpBalDt.Rows(0).Item("ILOC_PALLET_NO").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_QTY2") = 0
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_SERIAL_NO") = tmpBalDt.Rows(0).Item("ILBS_SERIAL_NO").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FOI_QTY") = 1
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_DO_QTY") = 1
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("STOCK_QTY") = 1
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_BAL_QTY") = 1
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_EXPIRY_DATE") = tmpBalDt.Rows(0).Item("EXPIRY_DATE_STR").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("ILOC_MANU_DATE") = tmpBalDt.Rows(0).Item("MANU_DATE_STR").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_WH") = tmpBalDt.Rows(0).Item("ILOC_WH").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_LOC") = tmpBalDt.Rows(0).Item("ILOC_LOC").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_FLOOR") = tmpBalDt.Rows(0).Item("ILOC_FLOOR").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_AREA") = tmpBalDt.Rows(0).Item("ILOC_AREA").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_RACK") = tmpBalDt.Rows(0).Item("ILOC_RACK").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BIN") = tmpBalDt.Rows(0).Item("ILOC_BIN").ToString.Trim
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("PLD_BATCH_NO") = tmpBalDt.Rows(0).Item("ILOC_BATCH_NO").ToString.Trim

                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("BN_CSMS_CODE") = tmpBalDt.Rows(0).Item("BN_CSMS_CODE").ToString.Trim

                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("DOD_DISP_SEQ") = tmpSerialItem.dod_disp_seq

                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("HOLD_QTY") = 0
                '    pickSerialDt.Rows(pickSerialDt.Rows.Count - 1).Item("TOTAL_BAL") = 1

                '    pickedSerialKeyList = gU.appendToList(pickedSerialKeyList, tmpBalDt.Rows(0).Item("ILBS_SEQ").ToString.Trim)
                'End If
            End If

        Next

    End Sub



    Public Function getItemBal(ByRef dtl_dt As DataTable) As DataTable
        Dim selectSql, selectSql1, selectSql2, selectSql3, selSqlTmp1, selSqlTmp2, selHldSql As String
        Dim itm_key_list1, itm_key_list2, itm_key_list3, itm_key_list2_tmp, itm_key_list3_tmp As String
        Dim itmDict, itmDictNoBatch As Dictionary(Of String, Double)
        Dim reqDict As Dictionary(Of String, Double)
        Dim dspSeqDict, dspSeqDictNoBatch As Dictionary(Of String, Integer)
        Dim itmKey, itmKeyNoBatch As String
        Dim LOC_dt, hldDt As DataTable
        Dim siDict As Dictionary(Of String, DtSerialItem)
        Dim tmpItmKeys As Dictionary(Of String, Double).KeyCollection
        Dim itmDictOS As Dictionary(Of String, Double)
        Dim tmpOSKey As String
        Dim i, j As Long
        Dim tmpAvailQty As Double
        'Dim tmpSerialItem As DtSerialItem
        'Dim tmpQtyList As List(Of Double)

        selectSql = ""
        selectSql1 = ""
        selectSql2 = ""
        selectSql3 = ""
        itm_key_list1 = ""
        itm_key_list2 = ""
        itm_key_list3 = ""
        itm_key_list3_tmp = ""
        itmDict = New Dictionary(Of String, Double)
        reqDict = New Dictionary(Of String, Double)
        siDict = New Dictionary(Of String, DtSerialItem)
        dspSeqDict = New Dictionary(Of String, Integer)
        dspSeqDictNoBatch = New Dictionary(Of String, Integer)

        For i = 0 To dtl_dt.Rows.Count - 1
            If DB.decodeDBNull(dtl_dt.Rows(i).Item("DOD_QTY"), 0) > 0 And dtl_dt.Rows(i).Item("mFlag").ToString <> "D" Then
                itmKey = dtl_dt.Rows(i).Item("dod_itm_code").ToString & "#_#" & _
                        dtl_dt.Rows(i).Item("dod_pack_key").ToString & "#_#" & _
                        gU.decodeNullOrEmpty(dtl_dt.Rows(i).Item("dod_pallet_no").ToString, "000")

                itmKeyNoBatch = ""

                '###
                If pickByBatchNo Then
                    If dtl_dt.Rows(i).Item("dod_batch_no").ToString.Trim <> "" Then
                        itmKeyNoBatch = itmKey
                        itmKey = itmKey & "#_#" & dtl_dt.Rows(i).Item("dod_batch_no").ToString.Trim
                    End If
                End If


                If dtl_dt.Rows(i).Item("ITM_SERIAL_NO_YN").ToString.Trim = "Y" Then
                    'If Not siDict.ContainsKey(itmKey) Then

                    '    tmpSerialItem = New DtSerialItem

                    '    tmpSerialItem.itm_key = itmKey
                    '    tmpSerialItem.itm_code = dtl_dt.Rows(i).Item("dod_itm_code").ToString.Trim
                    '    tmpSerialItem.pack_key = dtl_dt.Rows(i).Item("dod_pack_key").ToString.Trim
                    '    tmpSerialItem.pallet_no = gU.decodeNullOrEmpty(dtl_dt.Rows(i).Item("dod_pallet_no").ToString.Trim, "000")
                    '    tmpSerialItem.batch_no = dtl_dt.Rows(i).Item("dod_batch_no").ToString.Trim

                    '    If dtl_dt.Rows(i).Item("DOD_UOM2").ToString.Trim <> "" Then

                    '        If dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim = "" Then
                    '            Throw New Exception("Qty 2 cannot be empty for item: " & itmKey)
                    '        End If

                    '        tmpSerialItem.hasQty2 = True

                    '        tmpSerialItem.min_qty2 = CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim)

                    '        tmpSerialItem.total_qty2 = CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim)

                    '        tmpQtyList = New List(Of Double)

                    '        tmpQtyList.Add(CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim))

                    '        tmpSerialItem.qty2List = tmpQtyList

                    '    Else
                    '        tmpSerialItem.hasQty2 = False
                    '    End If

                    '    siDict.Add(itmKey, tmpSerialItem)

                    'Else
                    '    tmpSerialItem = siDict.Item(itmKey)

                    '    If tmpSerialItem.hasQty2 Then
                    '        If dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim = "" Then
                    '            Throw New Exception("Qty 2 cannot be empty for item: " & itmKey)
                    '        End If

                    '        If CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim) < tmpSerialItem.min_qty2 Then
                    '            tmpSerialItem.min_qty2 = CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim)
                    '        End If

                    '        tmpSerialItem.total_qty2 += CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim)

                    '        tmpQtyList = tmpSerialItem.qty2List

                    '        tmpQtyList.Add(CDbl(dtl_dt.Rows(i).Item("DOD_QTY2").ToString.Trim))
                    '    End If
                    'End If
                Else

                    If Not reqDict.ContainsKey(itmKey) Then
                        reqDict.Add(itmKey, 0)
                    End If

                    If Not itmDict.ContainsKey(itmKey) Then
                        itmDict.Add(itmKey, dtl_dt.Rows(i).Item("DOD_QTY"))
                        dspSeqDict.Add(itmKey, dtl_dt.Rows(i).Item("DOD_DISP_SEQ"))
                    Else
                        itmDict.Item(itmKey) = itmDict.Item(itmKey) + dtl_dt.Rows(i).Item("DOD_QTY")
                    End If

                    If itmKeyNoBatch <> "" AndAlso Not itmDict.ContainsKey(itmKeyNoBatch) AndAlso Not dspSeqDictNoBatch.ContainsKey(itmKeyNoBatch) Then
                        dspSeqDictNoBatch.Add(itmKeyNoBatch, dtl_dt.Rows(i).Item("DOD_DISP_SEQ"))
                    End If

                    If pickByBatchNo AndAlso dtl_dt.Rows(i).Item("dod_batch_no").ToString.Trim <> "" Then
                        'If itm_key_list1 <> "" Then
                        '    itm_key_list1 = itm_key_list1 & ", "
                        'End If
                        'itm_key_list1 = itm_key_list1 & "'" & gU.dbEncode(itmKey) & "'"

                        itm_key_list1 = gU.appendToList(itm_key_list1, gU.dbEncode(itmKey))

                        'list 3 is used to store the item key without batch for the item with batch
                        itm_key_list3_tmp = gU.appendToList(itm_key_list3_tmp, gU.dbEncode(itmKeyNoBatch))
                    Else
                        'If itm_key_list2 <> "" Then
                        '    itm_key_list2 = itm_key_list2 & ", "
                        'End If
                        'itm_key_list2 = itm_key_list2 & "'" & gU.dbEncode(itmKey) & "'"
                        itm_key_list2 = gU.appendToList(itm_key_list2, gU.dbEncode(itmKey))
                    End If

                    'If selectSql <> "" Then
                    '    selectSql = selectSql & ", "
                    'End If

                    'selectSql = selectSql & "'" & gU.dbEncode(itmKey) & "'"
                End If
            End If
        Next

        itm_key_list2_tmp = itm_key_list2

        'If same item exists in both with batch and empty batch, remove from list 2 to prevend select the same item (with batch) more than one time
        itm_key_list2 = gU.listSubtract(itm_key_list2_tmp, itm_key_list3_tmp)

        'Use list 3 to select item with no batch by adding batch no is null criteria
        itm_key_list3 = gU.listUnion(itm_key_list2_tmp, itm_key_list3_tmp)

        itm_key_list1 = gU.dbConvList(itm_key_list1)

        itm_key_list2 = gU.dbConvList(itm_key_list2)

        itm_key_list3 = gU.dbConvList(itm_key_list3)


        'Note: if there is no holding qty, TOTAL_BAL_QTY and AVAIL_BAL is ZERO
        selSqlTmp1 = "select l.ITM_CODE, l.PACK_KEY, l.ILOC_PALLET_NO, 0.0 as COD_QTY, 0.0 as TOTAL_HOLD_QTY, 0.0 as TOTAL_BAL_QTY, 0.0 as AVAIL_BAL," & _
                            "l.ILOC_BAL_QTY, l.ILOC_LOC, " & _
                            "l.ILOC_WH, l.ILOC_FLOOR, l.ILOC_AREA, l.ILOC_RACK, l.ILOC_BIN, l.ILOC_BATCH_NO, " & _
                            "l.ILOC_EXPIRY_DATE, l.ILOC_MANU_DATE, i.ITM_SKU_NO, " & _
                            "ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(PICK_ITEM.picked_qty, 0) as STOCK_QTY, " & _
                            "Convert(varchar, l.ILOC_EXPIRY_DATE, " & gU.getConfig("DDFORMATNo") & ") as EXPIRY_DATE_STR, " & _
                            "Convert(varchar, l.ILOC_MANU_DATE, " & gU.getConfig("DDFORMATNo") & ") as MANU_DATE_STR, " & _
                            "null as DOD_DISP_SEQ, b.BN_CSMS_CODE, "


        If pickOrder = "BY_IN_DATE" Then
            selSqlTmp1 = selSqlTmp1 & _
                        "tx.max_stock_in_date, "
        End If

        selSqlTmp2 = "from WMS_ITEM_LOC_BAL l " & _
                        "inner join WMS_WH_AREA a " & _
                        "on l.IMP_CODE = a.IMP_CODE " & _
                            "and l.ILOC_WH = a.WH_CODE " & _
                            "and l.ILOC_FLOOR = a.FL_NUM " & _
                            "and l.ILOC_AREA = a.AR_CODE " & _
                            "and ISNULL(a.AR_DAMAGE_YN, '') <> 'Y' " & _
                        "inner join WMS_WH_BIN b " & _
                        "on l.ILOC_LOC = b.LOC_KEY " & _
                        "inner join WMS_ITEM i " & _
                        "on l.IMP_CODE = i.IMP_CODE " & _
                            "and l.STORER_CODE = i.STORER_CODE " & _
                            "and l.ITM_CODE = i.ITM_CODE " & _
                            "and l.PACK_KEY = i.PACK_KEY " & _
                        "left outer join WMS_DATE_CODE DC " & _
                        "on l.ILOC_BATCH_NO = DC.DC_DATE_CODE " & _
                            "and l.STORER_CODE = DC.STORER_CODE " & _
                            "and l.IMP_CODE = DC.IMP_CODE " & _
                        "LEFT OUTER JOIN ( " & _
                            "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " & _
                            "from wms_do_picklist_d p, wms_delv_order d " & _
                            "where d.imp_code = p.imp_code " & _
                            "and d.storer_code = p.storer_code " & _
                            "and d.do_code = p.do_code " & _
                            "and d.do_status = 'PICKED' " & _
                            "and d.do_code <> '" & gU.dbEncode(lDO_CODE) & "' " & _
                            "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " & _
                        "ON l.IMP_CODE = PICK_ITEM.IMP_CODE " & _
                        "AND l.STORER_CODE = PICK_ITEM.STORER_CODE " & _
                        "AND l.ITM_CODE = PICK_ITEM.PLD_ITEM_NO " & _
                        "AND l.PACK_KEY = PICK_ITEM.PLD_PACK_KEY " & _
                        "AND ISNULL(l.ILOC_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " & _
                        "AND ISNULL(l.ILOC_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " & _
                        "AND l.ILOC_LOC = PICK_ITEM.PLD_LOC "

        If pickOrder = "BY_IN_DATE" Then
            selSqlTmp2 = selSqlTmp2 & _
                        "left outer join (" & _
                            "select imp_code, storer_code, itm_code, pack_key, io_pallet_no, io_batch_no, io_loc, max(sys_cd) as max_stock_in_date " & _
                            "from wms_in_tx " & _
                            "where io_type = 'IN' " & _
                            "and io_doc = 'GR' " & _
                            "and io_bal_before = 0 " & _
                            "group by imp_code, storer_code, itm_code, pack_key, io_pallet_no, io_batch_no, io_loc) tx " & _
                        "on l.IMP_CODE = tx.IMP_CODE " & _
                        "and l.storer_code = tx.storer_code " & _
                        "and l.itm_code = tx.itm_code " & _
                        "and l.pack_key = tx.pack_key " & _
                        "and ISNULL(l.iloc_pallet_no, '') = ISNULL(tx.io_pallet_no, '') " & _
                        "and ISNULL(l.iloc_batch_no, '') = ISNULL(tx.io_batch_no, '') " & _
                        "and l.iloc_loc = tx.io_loc "
        End If


        If itmWhList <> "" Then
            selSqlTmp2 = selSqlTmp2 & _
                        "left outer join wms_item_wh iw " & _
                        "on iw.IMP_CODE = i.IMP_CODE " & _
                        "and iw.STORER_CODE = i.STORER_CODE " & _
                        "and iw.ITM_CODE = i.ITM_CODE " & _
                        "and iw.PACK_KEY = i.PACK_KEY " & _
                        "and iw.WH_CODE in (" & gU.dbConvList(itmWhList) & ") " 

        End If

        selSqlTmp2 = selSqlTmp2 & _
                    "where l.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
                    "and l.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' " & _
                    "and ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(PICK_ITEM.picked_qty, 0) > 0 "

        '"and l.ILOC_LOC in (i.itm_pref_loc, i.itm_pref_loc2, case when isnull(i.itm_pref_loc, '') = '' and isnull(i.itm_pref_loc2, '') = '' then l.ILOC_LOC else null end) "

        If itmWhList <> "" Then
            selSqlTmp2 = selSqlTmp2 & _
                        "and l.ILOC_WH in (" & gU.dbConvList(itmWhList) & ") " & _
                        "and isnull(iw.IW_PICK_LOC, l.ILOC_LOC) = l.ILOC_LOC "

            '"and (not exists (" & _
            '        "select 1 from wms_item_wh iw " & _
            '        "where iw.IMP_CODE = i.IMP_CODE " & _
            '        "and iw.STORER_CODE = i.STORER_CODE " & _
            '        "and iw.ITM_CODE = i.ITM_CODE " & _
            '        "and iw.PACK_KEY = i.PACK_KEY " & _
            '        "and iw.WH_CODE in (" & gU.dbConvList(itmWhList) & ")) " & _
            '    "or exists (" & _
            '        "select 1 from wms_item_wh iw " & _
            '        "where iw.IMP_CODE = i.IMP_CODE " & _
            '        "and iw.STORER_CODE = i.STORER_CODE " & _
            '        "and iw.ITM_CODE = i.ITM_CODE " & _
            '        "and iw.PACK_KEY = i.PACK_KEY " & _
            '        "and iw.WH_CODE in (" & gU.dbConvList(itmWhList) & ") " & _
            '        "and isnull(iw.IW_PICK_LOC, l.ILOC_LOC) = l.ILOC_LOC))) "
        End If


        If itm_key_list1 = "" AndAlso itm_key_list2 = "" Then
            selectSql = selSqlTmp1 & " 'Y' as BATCH_FLAG " & selSqlTmp2 & "and 1=0 "
        Else
            If itm_key_list1 <> "" Then
                selectSql1 = selSqlTmp1 & " 'Y' as BATCH_FLAG " & selSqlTmp2 & _
                        "and l.ITM_CODE + '#_#' + l.PACK_KEY + '#_#' + isnull(l.ILOC_PALLET_NO, '') + '#_#' + isnull(l.ILOC_BATCH_NO, '') in (" & itm_key_list1 & ") "
            End If

            If itm_key_list2 <> "" Then
                selectSql2 = selSqlTmp1 & " 'N' as BATCH_FLAG " & selSqlTmp2 & _
                        "and l.ITM_CODE + '#_#' + l.PACK_KEY + '#_#' + isnull(l.ILOC_PALLET_NO, '') in (" & itm_key_list2 & ") "
            End If

            If itm_key_list3 <> "" Then
                selectSql3 = selSqlTmp1 & " 'N' as BATCH_FLAG " & selSqlTmp2 & _
                        "and l.ITM_CODE + '#_#' + l.PACK_KEY + '#_#' + isnull(l.ILOC_PALLET_NO, '') in (" & itm_key_list3 & ") " & _
                        "and isnull(l.ILOC_BATCH_NO, '') = '' "
            End If


            If selectSql1 <> "" Then
                selectSql = selectSql1
            End If

            If selectSql2 <> "" Then
                If selectSql <> "" Then
                    selectSql = selectSql & " union " & selectSql2
                Else
                    selectSql = selectSql2
                End If
            End If

            If selectSql3 <> "" Then
                If selectSql <> "" Then
                    selectSql = selectSql & " union " & selectSql3
                Else
                    selectSql = selectSql3
                End If
            End If
        End If



        'BY_BATCH_DATE: WMS_DATE_CODE.DC_CONV_DATE
        'BY_IN_DATE:    WMS_IN_TX.SYS_CD
        'BY_EXP_DATE:   WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE
        'BY_MANU_DATE:  WMS_ITEM_LOC_BAL.ILOC_MANU_DATE
        'NIL

        If pickOrder = "BY_BATCH_DATE" Then
            selectSql = selectSql & _
                    "order by ITM_CODE, PACK_KEY, DC_CONV_DATE, ILOC_BATCH_NO, STOCK_QTY desc "
        ElseIf pickOrder = "BY_IN_DATE" Then
            selectSql = selectSql & _
                    "order by ITM_CODE, PACK_KEY, ILOC_BATCH_NO, max_stock_in_date,STOCK_QTY desc "
        ElseIf pickOrder = "BY_EXP_DATE" Then
            selectSql = selectSql & _
                    "order by ITM_CODE, PACK_KEY, ILOC_EXPIRY_DATE, ILOC_BATCH_NO, STOCK_QTY desc "
        ElseIf pickOrder = "BY_MANU_DATE" Then
            selectSql = selectSql & _
                    "order by ITM_CODE, PACK_KEY, ILOC_MANU_DATE, ILOC_BATCH_NO, STOCK_QTY desc "
        End If



        '"order by l.ITM_CODE, DC.DC_CONV_DATE, l.PACK_KEY, l.ILOC_BATCH_NO, l.ILOC_BAL_QTY desc "

        LOC_dt = gDB.getDataTable(selectSql, cnn, transaction)


        

        tmpItmKeys = itmDict.Keys

        'Remove dspSeqDictNoBatch item if it is already in itmDict
        For i = 0 To tmpItmKeys.Count - 1
            If dspSeqDictNoBatch.ContainsKey(tmpItmKeys(i)) Then
                dspSeqDictNoBatch.Remove(tmpItmKeys(i))
            End If
        Next

        itmDictNoBatch = New Dictionary(Of String, Double)(itmDict)

        '###SAMUEL LOGIC
        For i = 0 To LOC_dt.Rows.Count - 1
            If pickByBatchNo AndAlso itmDict.ContainsKey(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString) Then
                LOC_dt.Rows(i).Item("COD_QTY") = itmDict.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString)

                If dspSeqDict.ContainsKey(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString) Then
                    LOC_dt.Rows(i).Item("DOD_DISP_SEQ") = dspSeqDict.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString)
                End If

                'itmDictNoBatch.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString)

            ElseIf itmDict.ContainsKey(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString) Then
                LOC_dt.Rows(i).Item("COD_QTY") = itmDict.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString)

                If dspSeqDict.ContainsKey(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString) Then
                    LOC_dt.Rows(i).Item("DOD_DISP_SEQ") = dspSeqDict.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString)
                End If

            ElseIf dspSeqDictNoBatch.ContainsKey(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString) Then
                'LOC_dt.Rows(i).Item("COD_QTY") = 0
                LOC_dt.Rows(i).Item("DOD_DISP_SEQ") = dspSeqDictNoBatch.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString)
            End If
        Next

        'Select items which are holding by CO, and their total balance in stock
        updateHoldBal(LOC_dt, True)



        'Two purpose:
        '1. Find out oustanding qty of item with batch no.
        '2. Update display seq for item without batch no. (while same item with batch no. exists)
        If pickByBatchNo Then
            For i = 0 To LOC_dt.Rows.Count - 1
                If itmDictNoBatch.ContainsKey(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString) Then
                    'Remove items with no batch no, because there is no outstanding requested qty for batch item for them (we only find out remaining qty for item with batch)
                    itmDictNoBatch.Remove(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString)

                ElseIf itmDictNoBatch.ContainsKey(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString) Then

                    If CDbl(LOC_dt.Rows(i).Item("TOTAL_HOLD_QTY")) > 0 Then
                        tmpAvailQty = gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("AVAIL_BAL"), 0)
                    Else
                        tmpAvailQty = gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("STOCK_QTY"), 0)
                    End If

                    If itmDictNoBatch.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString) > tmpAvailQty Then
                        itmDictNoBatch.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString) -= tmpAvailQty
                    Else
                        itmDictNoBatch.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString) = 0
                    End If
                End If
            Next

            itmDictOS = New Dictionary(Of String, Double)

            tmpItmKeys = itmDictNoBatch.Keys

            For i = tmpItmKeys.Count - 1 To 0 Step -1
                If itmDictNoBatch.Item(tmpItmKeys(i)) <= 0 Then
                    itmDictNoBatch.Remove(tmpItmKeys(i))
                End If
            Next

            For i = 0 To tmpItmKeys.Count - 1
                If itmDictNoBatch.Item(tmpItmKeys(i)) > 0 Then
                    tmpOSKey = Left(tmpItmKeys(i), InStrRev(tmpItmKeys(i), "#_#") - 1)

                    If Not itmDictOS.ContainsKey(tmpOSKey) Then
                        itmDictOS.Add(tmpOSKey, itmDictNoBatch.Item(tmpItmKeys(i)))
                    Else
                        itmDictOS.Item(tmpOSKey) = itmDictOS.Item(tmpOSKey) + itmDictNoBatch.Item(tmpItmKeys(i))
                    End If
                End If
            Next

            If itmDictOS.Count > 0 Then
                For i = 0 To LOC_dt.Rows.Count - 1
                    If LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim = "" Then
                        If itmDictOS.ContainsKey(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString) Then
                            LOC_dt.Rows(i).Item("COD_QTY") = gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("COD_QTY"), 0) + itmDictOS.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString)
                        End If

                        'If LOC_dt.Rows(i).Item("DOD_DISP_SEQ").ToString.Trim = "" Then
                        '    LOC_dt.Rows(i).Item("DOD_DISP_SEQ") = dspSeqDictNoBatch.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString)
                        'End If
                    End If
                Next

                'Update available qty again
                updateHoldBal(LOC_dt, True)
            End If
        End If
        






        '### REPLACE BY FUNCTION updateHoldBal *******************************************************************************
        ''Select items which are holding by CO, and their total balance in stock
        'selHldSql = "select hd.COD_ITM_CODE, hd.COD_PACK_KEY, hd.COD_PALLET_NO, hd.COD_BATCH_NO, hd.HOLD_QTY, ISNULL(l.BAL_QTY, 0.0) as BAL_QTY " & _
        '            "from ( "

        'If itm_key_list1 <> "" Then
        '    selHldSql = selHldSql & _
        '                    "select h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000') as COD_PALLET_NO, d.COD_BATCH_NO, sum(h.coh_in_stock_qty) as hold_qty " & _
        '                    "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
        '                    "where h.imp_code = d.imp_code " & _
        '                    "and h.storer_code = d.storer_code " & _
        '                    "and h.co_code = d.co_code " & _
        '                    "and h.cod_seq = d.cod_seq " & _
        '                    "and c.imp_code = d.imp_code " & _
        '                    "and c.storer_code = d.storer_code " & _
        '                    "and c.co_code = d.co_code " & _
        '                    "and h.coh_status <> 'RELEASE' " & _
        '                    "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
        '                    "and ISNULL(h.coh_in_stock_qty, 0) > 0 " & _
        '                    "and d.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' " & _
        '                    "and d.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
        '                    "and d.COD_ITM_CODE || '#_#' || d.COD_PACK_KEY || '#_#' || ISNULL(d.COD_PALLET_NO, '000') || '#_#' || d.COD_BATCH_NO in (" & itm_key_list1 & ") " & _
        '                    "group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO "
        'End If

        'If itm_key_list2 <> "" Then
        '    If itm_key_list1 <> "" Then
        '        selHldSql = selHldSql & "union "
        '    End If

        '    selHldSql = selHldSql & _
        '                    "select h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000') as COD_PALLET_NO, d.COD_BATCH_NO, sum(h.coh_in_stock_qty) as hold_qty " & _
        '                    "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
        '                    "where h.imp_code = d.imp_code " & _
        '                    "and h.storer_code = d.storer_code " & _
        '                    "and h.co_code = d.co_code " & _
        '                    "and h.cod_seq = d.cod_seq " & _
        '                    "and c.imp_code = d.imp_code " & _
        '                    "and c.storer_code = d.storer_code " & _
        '                    "and c.co_code = d.co_code " & _
        '                    "and h.coh_status <> 'RELEASE' " & _
        '                    "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
        '                    "and ISNULL(h.coh_in_stock_qty, 0) > 0 " & _
        '                    "and d.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' " & _
        '                    "and d.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
        '                    "and d.COD_ITM_CODE || '#_#' || d.COD_PACK_KEY || '#_#' || ISNULL(d.COD_PALLET_NO, '000') in (" & itm_key_list2 & ") " & _
        '                    "group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO "
        'End If

        'selHldSql = selHldSql & _
        '                ") hd,  " & _
        '                "( " & _
        '                    "select imp_code, storer_code, ITM_CODE, PACK_KEY, ILOC_PALLET_NO, " & _
        '                    "   ISNULL(ILOC_BATCH_NO, '') AS ILOC_BATCH_NO, sum(ILOC_BAL_QTY) as BAL_QTY " & _
        '                    "from WMS_ITEM_LOC_BAL " & _
        '                    "where ISNULL(ILOC_BAL_QTY, 0) > 0 " & _
        '                    "group by imp_code, storer_code, ITM_CODE, PACK_KEY, ILOC_PALLET_NO, ISNULL(ILOC_BATCH_NO, '')) l " & _
        '            "where hd.imp_code = l.imp_code (+) " & _
        '            "and hd.storer_code = l.storer_code (+) " & _
        '            "and hd.COD_ITM_CODE = l.ITM_CODE (+) " & _
        '            "and hd.COD_PACK_KEY = l.PACK_KEY (+) " & _
        '            "and hd.COD_PALLET_NO = l.ILOC_PALLET_NO (+) " & _
        '            "and hd.COD_BATCH_NO = l.ILOC_BATCH_NO (+) "

        'hldDt = gDB.getDataTable(selHldSql, cnn, transaction)

        'Dim tmpHldQty As Double

        'For j = 0 To hldDt.Rows.Count - 1
        '    tmpHldQty = CDbl(hldDt.Rows(j).Item("HOLD_QTY"))

        '    For i = 0 To LOC_dt.Rows.Count - 1
        '        If LOC_dt.Rows(i).Item("ITM_CODE").ToString.Trim = hldDt.Rows(j).Item("COD_ITM_CODE").ToString.Trim AndAlso _
        '            LOC_dt.Rows(i).Item("PACK_KEY").ToString.Trim = hldDt.Rows(j).Item("COD_PACK_KEY").ToString.Trim AndAlso _
        '            LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim = hldDt.Rows(j).Item("COD_PALLET_NO").ToString.Trim AndAlso _
        '            LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim = hldDt.Rows(j).Item("COD_BATCH_NO").ToString.Trim Then

        '            'TOTAL_HOLD_QTY,TOTAL_BAL_QTY,AVAIL_BAL
        '            LOC_dt.Rows(i).Item("TOTAL_HOLD_QTY") = CDbl(hldDt.Rows(j).Item("HOLD_QTY"))
        '            LOC_dt.Rows(i).Item("TOTAL_BAL_QTY") = CDbl(hldDt.Rows(j).Item("BAL_QTY"))

        '            If tmpHldQty > 0 Then
        '                If CDbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY")) > tmpHldQty Then
        '                    LOC_dt.Rows(i).Item("AVAIL_BAL") = CDbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY")) - tmpHldQty
        '                    tmpHldQty = 0
        '                Else
        '                    LOC_dt.Rows(i).Item("AVAIL_BAL") = 0
        '                    tmpHldQty = tmpHldQty - CDbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY"))
        '                End If
        '            Else
        '                LOC_dt.Rows(i).Item("AVAIL_BAL") = CDbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY"))
        '            End If
        '        End If
        '    Next
        'Next
        '************************************************************************************************


        Return LOC_dt
    End Function

    Public Sub updateTotalQtyForGV(ByRef gv As GridView)
        Dim itmKey As String
        Dim tmpItm As DtItems
        Dim itmQtyDict As New Dictionary(Of String, DtItems)
        Dim i As Long

        If gv.Rows.Count > 0 Then
            For i = 0 To gv.Rows.Count - 1
                'If CType(gv.Rows(i).FindControl("pld_batch_no"), Label).Text.Trim <> "" Then
                itmKey = CType(gv.Rows(i).FindControl("pld_item_no"), Label).Text.Trim & "#_#" &
                        CType(gv.Rows(i).FindControl("pld_pack_key"), Label).Text.Trim & "#_#" &
                        gU.decodeNullOrEmpty(CType(gv.Rows(i).FindControl("pld_pallet_no"), TextBox).Text.Trim, "000") & "#_#" &
                        CType(gv.Rows(i).FindControl("pld_batch_no"), TextBox).Text.Trim

                If itmQtyDict.ContainsKey(itmKey) Then
                    tmpItm = itmQtyDict.Item(itmKey)
                    tmpItm.total_foi_qty = tmpItm.total_foi_qty + CDbl(gU.decodeNullOrEmpty(CType(gv.Rows(i).FindControl("pld_foi_qty"), TextBox).Text.Trim, "0"))
                    tmpItm.total_itm_qty = tmpItm.total_itm_qty + CDbl(gU.decodeNullOrEmpty(CType(gv.Rows(i).FindControl("pld_item_qty"), TextBox).Text.Trim, "0"))
                Else
                    tmpItm = New DtItems
                    tmpItm.total_foi_qty = CDbl(gU.decodeNullOrEmpty(CType(gv.Rows(i).FindControl("pld_foi_qty"), TextBox).Text.Trim, "0"))
                    tmpItm.total_itm_qty = CDbl(gU.decodeNullOrEmpty(CType(gv.Rows(i).FindControl("pld_item_qty"), TextBox).Text.Trim, "0"))

                    itmQtyDict.Add(itmKey, tmpItm)
                End If
                'End If
            Next

            For i = 0 To gv.Rows.Count - 1
                itmKey = CType(gv.Rows(i).FindControl("pld_item_no"), Label).Text.Trim & "#_#" &
                            CType(gv.Rows(i).FindControl("pld_pack_key"), Label).Text.Trim & "#_#" &
                            gU.decodeNullOrEmpty(CType(gv.Rows(i).FindControl("pld_pallet_no"), TextBox).Text.Trim, "000") & "#_#" &
                            CType(gv.Rows(i).FindControl("pld_batch_no"), TextBox).Text.Trim

                tmpItm = itmQtyDict.Item(itmKey)

                CType(gv.Rows(i).FindControl("total_foi_qty"), HiddenField).Value = tmpItm.total_foi_qty
                CType(gv.Rows(i).FindControl("total_item_qty"), HiddenField).Value = tmpItm.total_itm_qty
            Next
        End If
    End Sub

    Public Function updateTotalQty(ByRef PL_dt As DataTable) As DataTable
        Dim itmKey As String
        Dim tmpItm As DtItems
        Dim itmQtyDict As New Dictionary(Of String, DtItems)
        Dim i As Long

        For i = 0 To PL_dt.Rows.Count - 1
            'If PL_dt.Rows(i).Item("PLD_BATCH_NO").ToString.Trim <> "" Then
            itmKey = PL_dt.Rows(i).Item("PLD_ITEM_NO").ToString & "#_#" & _
                    PL_dt.Rows(i).Item("PLD_PACK_KEY").ToString & "#_#" & _
                    gU.decodeNullOrEmpty(PL_dt.Rows(i).Item("PLD_PALLET_NO").ToString, "000") & "#_#" & _
                    PL_dt.Rows(i).Item("PLD_BATCH_NO").ToString.Trim

            If itmQtyDict.ContainsKey(itmKey) Then
                tmpItm = itmQtyDict.Item(itmKey)
                tmpItm.total_foi_qty = tmpItm.total_foi_qty + DB.decodeDBNull(PL_dt.Rows(i).Item("PLD_FOI_QTY"), 0)
                tmpItm.total_itm_qty = tmpItm.total_itm_qty + DB.decodeDBNull(PL_dt.Rows(i).Item("PLD_ITEM_QTY"), 0)
            Else
                tmpItm = New DtItems
                tmpItm.total_foi_qty = DB.decodeDBNull(PL_dt.Rows(i).Item("PLD_FOI_QTY"), 0)
                tmpItm.total_itm_qty = DB.decodeDBNull(PL_dt.Rows(i).Item("PLD_ITEM_QTY"), 0)

                itmQtyDict.Add(itmKey, tmpItm)
            End If
            'End If
        Next

        For i = 0 To PL_dt.Rows.Count - 1
            itmKey = PL_dt.Rows(i).Item("PLD_ITEM_NO").ToString & "#_#" & _
                        PL_dt.Rows(i).Item("PLD_PACK_KEY").ToString & "#_#" & _
                        gU.decodeNullOrEmpty(PL_dt.Rows(i).Item("PLD_PALLET_NO").ToString, "000") & "#_#" & _
                        PL_dt.Rows(i).Item("PLD_BATCH_NO").ToString.Trim

            tmpItm = itmQtyDict.Item(itmKey)

            PL_dt.Rows(i).Item("TOTAL_FOI_QTY") = tmpItm.total_foi_qty
            PL_dt.Rows(i).Item("TOTAL_ITEM_QTY") = tmpItm.total_itm_qty
        Next

        Return PL_dt
    End Function

    Public Function updateAvailBal(ByRef PL_dt As DataTable, Optional ByVal isLocTable As Boolean = False) As DataTable
        Dim selectSql As String
        Dim i As Long
        Dim tmpDt As DataTable

        For i = 0 To PL_dt.Rows.Count - 1
            If PL_dt.Rows(i).Item("mFlag").ToString.Trim <> "D" AndAlso PL_dt.Rows(i).Item("PLD_LOC").ToString.Trim <> "" Then

                If PL_dt.Rows(i).Item("ITM_SERIAL_NO_YN").ToString.Trim = "Y" Then

                    selectSql = "SELECT max(l.ILOC_BAL_QTY) as ILOC_BAL_QTY, " & _
                                    "count(*) as stock_qty " & _
                                "from WMS_ITEM_LOC_BAL l " & _
                                "inner join WMS_ITEM_LOC_BAL_S s " & _
                                "on l.ILOC_SEQ = s.ILOC_SEQ " & _
                                "left outer join ( " & _
                                    "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, p.pld_serial_no, sum(p.PLD_QTY2) as picked_qty2 " & _
                                    "from wms_do_picklist_d p, wms_delv_order d " & _
                                    "where d.imp_code = p.imp_code " & _
                                    "and d.storer_code = p.storer_code " & _
                                    "and d.do_code = p.do_code " & _
                                    "and d.do_status = 'PICKED' " & _
                                    "and d.do_code <> '" & gU.dbEncode(lDO_CODE) & "' " & _
                                    "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc, p.pld_serial_no) PICK_ITEM_S " & _
                                "on l.IMP_CODE = PICK_ITEM_S.IMP_CODE " & _
                                    "AND l.STORER_CODE = PICK_ITEM_S.STORER_CODE " & _
                                    "AND l.ITM_CODE = PICK_ITEM_S.PLD_ITEM_NO " & _
                                    "AND l.PACK_KEY = PICK_ITEM_S.PLD_PACK_KEY  " & _
                                    "AND ISNULL(l.ILOC_PALLET_NO, '000') = PICK_ITEM_S.PLD_PALLET_NO " & _
                                    "AND ISNULL(l.ILOC_BATCH_NO, '') = PICK_ITEM_S.PLD_BATCH_NO " & _
                                    "AND l.ILOC_LOC = PICK_ITEM_S.PLD_LOC " & _
                                    "AND s.ILBS_SERIAL_NO = PICK_ITEM_S.PLD_SERIAL_NO " & _
                                "where l.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' " & _
                                "and l.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
                                "and l.ITM_CODE = '" & gU.dbEncode(PL_dt.Rows(i).Item("PLD_ITEM_NO").ToString.Trim) & "' " & _
                                "and l.PACK_KEY = '" & gU.dbEncode(PL_dt.Rows(i).Item("PLD_PACK_KEY").ToString.Trim) & "' " & _
                                "and ISNULL(l.ILOC_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNull(PL_dt.Rows(i).Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' " & _
                                "and ISNULL(l.ILOC_BATCH_NO, '') = '" & gU.dbEncode(PL_dt.Rows(i).Item("PLD_BATCH_NO").ToString.Trim) & "' " & _
                                "and l.ILOC_LOC = '" & gU.dbEncode(PL_dt.Rows(i).Item("PLD_LOC").ToString.Trim) & "' " & _
                                "and ISNULL(s.ILBS_QTY2, 0) - ISNULL(PICK_ITEM_S.picked_qty2, 0) >= " & gU.dbEncode(gU.decodeNullOrEmpty(PL_dt.Rows(i).Item("PLD_QTY2").ToString.Trim, "0")) & " "

                Else
                    selectSql = "SELECT l.ILOC_BAL_QTY, " & _
                                    "ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(pick_item.picked_qty, 0) as stock_qty " & _
                                "from WMS_ITEM_LOC_BAL l " & _
                                "LEFT OUTER JOIN ( " & _
                                    "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " & _
                                    "from wms_do_picklist_d p, wms_delv_order d " & _
                                    "where d.imp_code = p.imp_code " & _
                                    "and d.storer_code = p.storer_code " & _
                                    "and d.do_code = p.do_code " & _
                                    "and d.do_status = 'PICKED' " & _
                                    "and d.do_code <> '" & gU.dbEncode(lDO_CODE) & "' " & _
                                    "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " & _
                                "ON l.IMP_CODE = PICK_ITEM.IMP_CODE " & _
                                    "AND l.STORER_CODE = PICK_ITEM.STORER_CODE " & _
                                    "AND l.ITM_CODE = PICK_ITEM.PLD_ITEM_NO " & _
                                    "AND l.PACK_KEY = PICK_ITEM.PLD_PACK_KEY " & _
                                    "AND ISNULL(l.ILOC_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " & _
                                    "AND ISNULL(l.ILOC_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " & _
                                    "AND l.ILOC_LOC = PICK_ITEM.PLD_LOC " & _
                                "where l.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' " & _
                                "and l.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
                                "and l.ITM_CODE = '" & gU.dbEncode(PL_dt.Rows(i).Item("PLD_ITEM_NO").ToString.Trim) & "' " & _
                                "and l.PACK_KEY = '" & gU.dbEncode(PL_dt.Rows(i).Item("PLD_PACK_KEY").ToString.Trim) & "' " & _
                                "and ISNULL(l.ILOC_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNull(PL_dt.Rows(i).Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' " & _
                                "and ISNULL(l.ILOC_BATCH_NO, '') = '" & gU.dbEncode(PL_dt.Rows(i).Item("PLD_BATCH_NO").ToString.Trim) & "' " & _
                                "and l.ILOC_LOC = '" & gU.dbEncode(PL_dt.Rows(i).Item("PLD_LOC").ToString.Trim) & "' "

                End If

                tmpDt = gDB.getDataTable(selectSql, cnn, transaction)

                If tmpDt.Rows.Count > 0 Then
                    PL_dt.Rows(i).Item("ILOC_BAL_QTY") = tmpDt.Rows(0).Item("ILOC_BAL_QTY")
                    PL_dt.Rows(i).Item("STOCK_QTY") = tmpDt.Rows(0).Item("STOCK_QTY")
                Else
                    PL_dt.Rows(i).Item("ILOC_BAL_QTY") = 0
                    PL_dt.Rows(i).Item("STOCK_QTY") = 0
                End If
            End If
        Next

        PL_dt.AcceptChanges()

        Return PL_dt
    End Function

    Public Function updateHoldBal(ByRef PL_dt As DataTable, Optional ByVal isLocTable As Boolean = False) As DataTable
        Dim selHldSql As String
        Dim itmKey As String
        Dim hldDt As DataTable
        Dim itmKeyList As String
        Dim i, j As Long

        itmKeyList = ""
        'itm_key_list1 = ""
        'itm_key_list2 = ""

        For i = 0 To PL_dt.Rows.Count - 1

            If isLocTable Then
                'If PL_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim <> "" Then
                itmKey = PL_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & _
                        PL_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & _
                        gU.decodeNullOrEmpty(PL_dt.Rows(i).Item("ILOC_PALLET_NO").ToString, "000") & "#_#" & _
                        PL_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim

                itmKeyList = gU.appendToList(itmKeyList, itmKey)
                'End If
            Else
                'If PL_dt.Rows(i).Item("PLD_BATCH_NO").ToString.Trim <> "" Then
                itmKey = PL_dt.Rows(i).Item("PLD_ITEM_NO").ToString & "#_#" & _
                        PL_dt.Rows(i).Item("PLD_PACK_KEY").ToString & "#_#" & _
                        gU.decodeNullOrEmpty(PL_dt.Rows(i).Item("PLD_PALLET_NO").ToString, "000") & "#_#" & _
                        PL_dt.Rows(i).Item("PLD_BATCH_NO").ToString.Trim

                itmKeyList = gU.appendToList(itmKeyList, itmKey)
                'End If
            End If
        Next

        If itmKeyList <> "" Then
            itmKeyList = gU.dbConvList(itmKeyList)
        End If

        If itmKeyList <> "" Then
            'Select items which are holding by CO, and their total balance in stock
            selHldSql = "select hd.COD_ITM_CODE, hd.COD_PACK_KEY, hd.COD_PALLET_NO, hd.COD_BATCH_NO, hd.HOLD_QTY, ISNULL(l.BAL_QTY, 0.0) as BAL_QTY " & _
                        "from ( "

            selHldSql = selHldSql & _
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
                                            "h.COH_ITM_CODE as COD_ITM_CODE, " & _
                                            "h.COH_PACK_KEY as COD_PACK_KEY, " & _
                                            "ISNULL(h.COH_PALLET_NO, '000') as COD_PALLET_NO, " & _
                                            "h.COH_BATCH_NO as COD_BATCH_NO, " & _
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
                                        "and h.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' " & _
                                        "and h.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
                                        "and c.co_code <> '" & gU.dbEncode(lCO_CODE) & "' " & _
                                        "and h.COH_ITM_CODE + '#_#' + h.COH_PACK_KEY + '#_#' + ISNULL(h.COH_PALLET_NO, '000') + '#_#' + isnull(h.COH_BATCH_NO, '') in (" & itmKeyList & ") " & _
                                        "group by h.imp_code, h.storer_code, h.co_code, h.COH_ITM_CODE, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000'), h.COH_BATCH_NO) hold " & _
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
                                    "having sum(hold.hold_qty - ISNULL(pick.picked_qty, 0)) > 0 " & _
                            ") hd left outer join  " & _
                            "( " & _
                                "select imp_code, storer_code, ITM_CODE, PACK_KEY, ISNULL(ILOC_PALLET_NO, '000') as ILOC_PALLET_NO, " & _
                                    "ISNULL(ILOC_BATCH_NO, '') AS ILOC_BATCH_NO, sum(ILOC_BAL_QTY) as BAL_QTY " & _
                                "from WMS_ITEM_LOC_BAL " & _
                                "where ISNULL(ILOC_BAL_QTY, 0) > 0 " & _
                                "group by imp_code, storer_code, ITM_CODE, PACK_KEY, ISNULL(ILOC_PALLET_NO, '000'), ISNULL(ILOC_BATCH_NO, '') " & _
                            ") l " & _
                        "on hd.imp_code = l.imp_code  " & _
                        "and hd.storer_code = l.storer_code  " & _
                        "and hd.COD_ITM_CODE = l.ITM_CODE  " & _
                        "and hd.COD_PACK_KEY = l.PACK_KEY  " & _
                        "and hd.COD_PALLET_NO = l.ILOC_PALLET_NO  " & _
                        "and hd.COD_BATCH_NO = l.ILOC_BATCH_NO  "

            '"and d.COD_ITM_CODE + '#_#' + d.COD_PACK_KEY + '#_#' + ISNULL(d.COD_PALLET_NO, '000') + '#_#' + d.COD_BATCH_NO in (" & itmKeyList & ") " & _
            '"group by h.imp_code, h.storer_code, h.co_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO) hold " & _

            '            "select h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000') as COD_PALLET_NO, d.COD_BATCH_NO, sum(ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0)) as hold_qty " & _
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
            '"and d.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' " & _
            '"and d.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
            '"and c.co_code <> '" & gU.dbEncode(lCO_CODE) & "' " & _
            '"and d.COD_ITM_CODE || '#_#' || d.COD_PACK_KEY || '#_#' || ISNULL(d.COD_PALLET_NO, '000') || '#_#' || d.COD_BATCH_NO in (" & itmKeyList & ") " & _
            '"group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO " & _



            '"(" & _
            '                   "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, sum(p.pld_item_qty) as picked_qty " & _
            '                   "from wms_do_picklist_d p, wms_delv_order d " & _
            '                   "where d.imp_code = p.imp_code " & _
            '                   "and d.storer_code = p.storer_code " & _
            '                   "and d.do_code = p.do_code " & _
            '                   "and d.do_status = 'PICKED' " & _
            '                   "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, '') " & _
            '               ") pck " & _

            '"and hd.imp_code = pck.imp_code (+) " & _
            '"and hd.storer_code = pck.storer_code (+) " & _
            '"and hd.COD_ITM_CODE = pck.pld_item_no (+) " & _
            '"and hd.COD_PACK_KEY = pck.pld_pack_key (+) " & _
            '"and hd.COD_PALLET_NO = pck.pld_pallet_no (+) " & _
            '"and hd.COD_BATCH_NO = pck.pld_batch_no (+) "

            hldDt = gDB.getDataTable(selHldSql, cnn, transaction)

            Dim tmpHldQty As Double

            For j = 0 To hldDt.Rows.Count - 1
                tmpHldQty = CDbl(hldDt.Rows(j).Item("HOLD_QTY"))

                For i = 0 To PL_dt.Rows.Count - 1

                    If isLocTable Then
                        If PL_dt.Rows(i).Item("ITM_CODE").ToString.Trim = hldDt.Rows(j).Item("COD_ITM_CODE").ToString.Trim AndAlso _
                        PL_dt.Rows(i).Item("PACK_KEY").ToString.Trim = hldDt.Rows(j).Item("COD_PACK_KEY").ToString.Trim AndAlso _
                        PL_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim = hldDt.Rows(j).Item("COD_PALLET_NO").ToString.Trim AndAlso _
                        PL_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim = hldDt.Rows(j).Item("COD_BATCH_NO").ToString.Trim Then

                            'TOTAL_HOLD_QTY,TOTAL_BAL_QTY,AVAIL_BAL
                            PL_dt.Rows(i).Item("TOTAL_HOLD_QTY") = CDbl(hldDt.Rows(j).Item("HOLD_QTY"))
                            PL_dt.Rows(i).Item("TOTAL_BAL_QTY") = CDbl(hldDt.Rows(j).Item("BAL_QTY"))

                            'Change ILOC_BAL_QTY to STOCK_QTY
                            If tmpHldQty > 0 Then
                                If gU.decodeEmptyCdbl(PL_dt.Rows(i).Item("STOCK_QTY").ToString.Trim, 0) > tmpHldQty Then
                                    PL_dt.Rows(i).Item("AVAIL_BAL") = CDbl(PL_dt.Rows(i).Item("STOCK_QTY")) - tmpHldQty
                                    tmpHldQty = 0
                                Else
                                    PL_dt.Rows(i).Item("AVAIL_BAL") = 0
                                    tmpHldQty = tmpHldQty - gU.decodeEmptyCdbl(PL_dt.Rows(i).Item("STOCK_QTY").ToString.Trim, 0)
                                End If
                            Else
                                PL_dt.Rows(i).Item("AVAIL_BAL") = gU.decodeEmptyCdbl(PL_dt.Rows(i).Item("STOCK_QTY").ToString.Trim, 0)
                            End If
                        End If
                    Else
                        If PL_dt.Rows(i).Item("PLD_ITEM_NO").ToString.Trim = hldDt.Rows(j).Item("COD_ITM_CODE").ToString.Trim AndAlso _
                        PL_dt.Rows(i).Item("PLD_PACK_KEY").ToString.Trim = hldDt.Rows(j).Item("COD_PACK_KEY").ToString.Trim AndAlso _
                        PL_dt.Rows(i).Item("PLD_PALLET_NO").ToString.Trim = hldDt.Rows(j).Item("COD_PALLET_NO").ToString.Trim AndAlso _
                        PL_dt.Rows(i).Item("PLD_BATCH_NO").ToString.Trim = hldDt.Rows(j).Item("COD_BATCH_NO").ToString.Trim Then

                            'TOTAL_HOLD_QTY,TOTAL_BAL_QTY,AVAIL_BAL
                            PL_dt.Rows(i).Item("HOLD_QTY") = CDbl(hldDt.Rows(j).Item("HOLD_QTY"))
                            PL_dt.Rows(i).Item("TOTAL_BAL") = CDbl(hldDt.Rows(j).Item("BAL_QTY"))

                            If tmpHldQty > 0 Then
                                If gU.decodeEmptyCdbl(PL_dt.Rows(i).Item("STOCK_QTY").ToString.Trim, 0) > tmpHldQty Then
                                    PL_dt.Rows(i).Item("AVAIL_QTY") = gU.decodeEmptyCdbl(PL_dt.Rows(i).Item("STOCK_QTY").ToString.Trim, 0) - tmpHldQty
                                    tmpHldQty = 0
                                Else
                                    PL_dt.Rows(i).Item("AVAIL_QTY") = 0
                                    tmpHldQty = tmpHldQty - gU.decodeEmptyCdbl(PL_dt.Rows(i).Item("STOCK_QTY").ToString.Trim, 0)
                                End If
                            Else
                                PL_dt.Rows(i).Item("AVAIL_QTY") = gU.decodeEmptyCdbl(PL_dt.Rows(i).Item("STOCK_QTY").ToString.Trim, 0)
                            End If
                        End If
                    End If
                Next
            Next
        End If

        Return PL_dt
    End Function

    Public Sub setBalanceTable(ByRef selDt As DataTable)
        Dim tmpItm As DtItems

        pickListDt = New DataTable
        pickListDt.Columns.Add("PLD_ITEM_NO", Type.GetType("System.String"))
        pickListDt.Columns.Add("ITM_SKU_NO", Type.GetType("System.String"))
        pickListDt.Columns.Add("PLD_PACK_KEY", Type.GetType("System.String"))
        pickListDt.Columns.Add("PLD_PALLET_NO", Type.GetType("System.String"))
        pickListDt.Columns.Add("PLD_DO_QTY", Type.GetType("System.Double"))
        'pickListDt.Columns.Add("PLD_ITEM_QTY", Type.GetType("System.Double"))

        pickListDt.Columns.Add("PLD_QTY2", Type.GetType("System.Double"))
        pickListDt.Columns.Add("PLD_SERIAL_NO", Type.GetType("System.String"))
        pickListDt.Columns.Add("ITM_SERIAL_NO_YN", Type.GetType("System.String"))

        pickListDt.Columns.Add("PLD_FOI_QTY", Type.GetType("System.Double"))
        pickListDt.Columns.Add("STOCK_QTY", Type.GetType("System.Double"))
        pickListDt.Columns.Add("ILOC_BAL_QTY", Type.GetType("System.Double"))
        pickListDt.Columns.Add("ILOC_EXPIRY_DATE", Type.GetType("System.String"))
        pickListDt.Columns.Add("ILOC_MANU_DATE", Type.GetType("System.String"))
        pickListDt.Columns.Add("PLD_WH", Type.GetType("System.String"))
        pickListDt.Columns.Add("PLD_LOC", Type.GetType("System.String"))
        pickListDt.Columns.Add("PLD_FLOOR", Type.GetType("System.String"))
        pickListDt.Columns.Add("PLD_AREA", Type.GetType("System.String"))
        pickListDt.Columns.Add("PLD_RACK", Type.GetType("System.String"))
        pickListDt.Columns.Add("PLD_BIN", Type.GetType("System.String"))
        pickListDt.Columns.Add("PLD_BATCH_NO", Type.GetType("System.String"))

        pickListDt.Columns.Add("DOD_DISP_SEQ", Type.GetType("System.Int32"))
        pickListDt.Columns.Add("BN_CSMS_CODE", Type.GetType("System.String"))

        pickListDt.Columns.Add("HOLD_QTY", Type.GetType("System.String"))
        pickListDt.Columns.Add("TOTAL_BAL", Type.GetType("System.String"))

        itmDict = New Dictionary(Of String, Double)
        'pickedItmDict = New Dictionary(Of String, Double)
        'lackItmDict = New Dictionary(Of String, Double)
        'locDict = New Dictionary(Of String, LocBal)
        'confirmedDict = New Dictionary(Of String, DtItems)
        'dtList = New List(Of DtItems)
        dtDict = New Dictionary(Of String, DtItems)

        'selDt is ordered by ITM_CODE, PACK_KEY, ILOC_BAL_QTY desc
        For i As Integer = 0 To selDt.Rows.Count - 1
            'itm_code, pack_key, loc_key, loc, wh, floor, area, rack, bin
            'qty, bal_qty

            'If AVAIL_BAL = 0, all items are hold, so no need select stock
            If CDbl(selDt.Rows(i).Item("TOTAL_HOLD_QTY")) > 0 AndAlso CDbl(selDt.Rows(i).Item("AVAIL_BAL")) = 0 Then
                Continue For
            End If

            'If there is no request qty
            If CDbl(selDt.Rows(i).Item("COD_QTY")) = 0 Then
                Continue For
            End If

            tmpItm = New DtItems
            tmpItm.dt_key = selDt.Rows(i).Item("ITM_CODE").ToString.Trim & "#_#" & _
                                selDt.Rows(i).Item("PACK_KEY").ToString.Trim & "#_#" & _
                                selDt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim & "#_#" & _
                                selDt.Rows(i).Item("ILOC_LOC").ToString.Trim & "#_#" & _
                                selDt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim

            If selDt.Rows(i).Item("BATCH_FLAG").ToString.Trim = "Y" Then
                tmpItm.itm_key = selDt.Rows(i).Item("ITM_CODE").ToString.Trim & "#_#" & _
                            selDt.Rows(i).Item("PACK_KEY").ToString.Trim & "#_#" & _
                            selDt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim & "#_#" & _
                            selDt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
            Else
                tmpItm.itm_key = selDt.Rows(i).Item("ITM_CODE").ToString.Trim & "#_#" & _
                            selDt.Rows(i).Item("PACK_KEY").ToString.Trim & "#_#" & _
                            selDt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
            End If

            tmpItm.itm_code = selDt.Rows(i).Item("ITM_CODE").ToString.Trim
            tmpItm.sku_no = selDt.Rows(i).Item("ITM_SKU_NO").ToString.Trim
            tmpItm.pack_key = selDt.Rows(i).Item("PACK_KEY").ToString.Trim
            tmpItm.pallet_no = selDt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
            tmpItm.loc = selDt.Rows(i).Item("ILOC_LOC").ToString.Trim
            tmpItm.expiry_date = selDt.Rows(i).Item("EXPIRY_DATE_STR").ToString.Trim
            tmpItm.manu_date = selDt.Rows(i).Item("MANU_DATE_STR").ToString.Trim
            tmpItm.wh = selDt.Rows(i).Item("ILOC_WH").ToString.Trim
            tmpItm.floor = selDt.Rows(i).Item("ILOC_FLOOR").ToString.Trim
            tmpItm.area = selDt.Rows(i).Item("ILOC_AREA").ToString.Trim
            tmpItm.rack = selDt.Rows(i).Item("ILOC_RACK").ToString.Trim
            tmpItm.bin = selDt.Rows(i).Item("ILOC_BIN").ToString.Trim

            tmpItm.dod_disp_seq = selDt.Rows(i).Item("DOD_DISP_SEQ").ToString.Trim
            tmpItm.bn_csms_code = selDt.Rows(i).Item("BN_CSMS_CODE").ToString.Trim

            '??? -----------------------
            ''tmpItm.qty = CDbl(selDt.Rows(i).Item("COD_QTY"))
            'tmpItm.qty = 0
            'tmpItm.bal_qty = CDbl(selDt.Rows(i).Item("ILOC_BAL_QTY"))
            'tmpItm.org_qty = CDbl(selDt.Rows(i).Item("COD_QTY"))
            '??? -----------------------

            tmpItm.qty = CDbl(selDt.Rows(i).Item("COD_QTY"))

            tmpItm.org_do_qty = CDbl(selDt.Rows(i).Item("COD_QTY"))
            tmpItm.org_bal_qty = CDbl(selDt.Rows(i).Item("ILOC_BAL_QTY"))

            tmpItm.org_stock_qty = CDbl(selDt.Rows(i).Item("STOCK_QTY"))

            'TOTAL_HOLD_QTY,TOTAL_BAL_QTY,AVAIL_BAL
            tmpItm.total_bal_qty = CDbl(selDt.Rows(i).Item("TOTAL_BAL_QTY"))
            tmpItm.hold_qty = CDbl(selDt.Rows(i).Item("TOTAL_HOLD_QTY"))


            'If hold stock exist, use AVAIL_BAL
            If CDbl(selDt.Rows(i).Item("TOTAL_HOLD_QTY")) > 0 Then
                tmpItm.bal_qty = CDbl(selDt.Rows(i).Item("AVAIL_BAL"))
            Else
                'tmpItm.bal_qty = CDbl(selDt.Rows(i).Item("ILOC_BAL_QTY"))
                tmpItm.bal_qty = CDbl(selDt.Rows(i).Item("STOCK_QTY"))
            End If

            tmpItm.batch = selDt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
            If Not dtDict.ContainsKey(tmpItm.dt_key) Then
                dtDict.Add(tmpItm.dt_key, tmpItm)
            End If
            If Not itmDict.ContainsKey(tmpItm.itm_key) Then
                itmDict.Add(tmpItm.itm_key, tmpItm.qty)
            End If
        Next
    End Sub

    Public Function getPickList() As DataTable
        Dim i As Integer

        'Add serial item first
        For i = 0 To pickSerialDt.Rows.Count - 1
            pickListDt.Rows.Add()
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_ITEM_NO") = pickSerialDt.Rows(i).Item("PLD_ITEM_NO")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("ITM_SKU_NO") = pickSerialDt.Rows(i).Item("ITM_SKU_NO")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_PACK_KEY") = pickSerialDt.Rows(i).Item("PLD_PACK_KEY")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_PALLET_NO") = pickSerialDt.Rows(i).Item("PLD_PALLET_NO")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_QTY2") = pickSerialDt.Rows(i).Item("PLD_QTY2")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_SERIAL_NO") = pickSerialDt.Rows(i).Item("PLD_SERIAL_NO")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("ITM_SERIAL_NO_YN") = "Y"
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_FOI_QTY") = pickSerialDt.Rows(i).Item("PLD_FOI_QTY")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_DO_QTY") = pickSerialDt.Rows(i).Item("PLD_DO_QTY")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("STOCK_QTY") = pickSerialDt.Rows(i).Item("STOCK_QTY")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("ILOC_BAL_QTY") = pickSerialDt.Rows(i).Item("ILOC_BAL_QTY")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("ILOC_EXPIRY_DATE") = pickSerialDt.Rows(i).Item("ILOC_EXPIRY_DATE")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("ILOC_MANU_DATE") = pickSerialDt.Rows(i).Item("ILOC_MANU_DATE")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_WH") = pickSerialDt.Rows(i).Item("PLD_WH")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_LOC") = pickSerialDt.Rows(i).Item("PLD_LOC")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_FLOOR") = pickSerialDt.Rows(i).Item("PLD_FLOOR")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_AREA") = pickSerialDt.Rows(i).Item("PLD_AREA")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_RACK") = pickSerialDt.Rows(i).Item("PLD_RACK")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_BIN") = pickSerialDt.Rows(i).Item("PLD_BIN")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_BATCH_NO") = pickSerialDt.Rows(i).Item("PLD_BATCH_NO")

            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("DOD_DISP_SEQ") = pickSerialDt.Rows(i).Item("DOD_DISP_SEQ")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("BN_CSMS_CODE") = pickSerialDt.Rows(i).Item("BN_CSMS_CODE")

            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("HOLD_QTY") = pickSerialDt.Rows(i).Item("HOLD_QTY")
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("TOTAL_BAL") = pickSerialDt.Rows(i).Item("TOTAL_BAL")
        Next
        '-------------------------------------------------------------------------------------------------------------------------------


        pickList = New List(Of DtItems)

        'Find out all loc that contains unique item combination first
        pickUniqueItem()

        'Find out the best item combo and pick
        pickBestItems()

        pickList.Sort(AddressOf CompareDtByLoc)

        For i = 0 To pickList.Count - 1
            pickListDt.Rows.Add()
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_ITEM_NO") = pickList(i).itm_code
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("ITM_SKU_NO") = pickList(i).sku_no
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_PACK_KEY") = pickList(i).pack_key
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_PALLET_NO") = pickList(i).pallet_no
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_QTY2") = 0
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_SERIAL_NO") = ""
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("ITM_SERIAL_NO_YN") = "N"
            'pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_ITEM_QTY") = pickList(i).qty
            'pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_ITEM_QTY") = pickList(i).bal_qty
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_FOI_QTY") = pickList(i).bal_qty
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_DO_QTY") = pickList(i).org_do_qty
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("STOCK_QTY") = pickList(i).org_stock_qty
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("ILOC_BAL_QTY") = pickList(i).org_bal_qty
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("ILOC_EXPIRY_DATE") = pickList(i).expiry_date
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("ILOC_MANU_DATE") = pickList(i).manu_date
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_WH") = pickList(i).wh
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_LOC") = pickList(i).loc
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_FLOOR") = pickList(i).floor
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_AREA") = pickList(i).area
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_RACK") = pickList(i).rack
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_BIN") = pickList(i).bin
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("PLD_BATCH_NO") = pickList(i).batch

            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("DOD_DISP_SEQ") = pickList(i).dod_disp_seq
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("BN_CSMS_CODE") = pickList(i).bn_csms_code

            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("HOLD_QTY") = pickList(i).hold_qty
            pickListDt.Rows(pickListDt.Rows.Count - 1).Item("TOTAL_BAL") = pickList(i).total_bal_qty
        Next

        'Dim dv As DataView

        'dv = pickListDt.DefaultView

        'pickListDt.DefaultView.Sort = "DOD_DISP_SEQ, ITM_SKU_NO, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_BATCH_NO, cint(PLD_SPLIT_FR), cint(PLD_SEQ)"

        'dt.AsEnumerable()
        '              .OrderBy(r => int.Parse(r.Field<String>("RollNo")))
        '.CopyToDataTable()



        'Dim query = From q In pickListDt.AsEnumerable() _
        '                     Order By _
        '                     q("DOD_DISP_SEQ"), _
        '                     q("ITM_SKU_NO").ToString, _
        '                     q("PLD_ITEM_NO").ToString, _
        '                     q("PLD_PACK_KEY"), _
        '                     q("PLD_PALLET_NO").ToString, _
        '                     q("PLD_BATCH_NO").ToString, _
        '                     gU.decodeEmptyCInt(q("PLD_SPLIT_FR").ToString, 0), _
        '                     gU.decodeEmptyCInt(q("PLD_SEQ").ToString, 0)

        'Dim query = From q In pickListDt.AsEnumerable() _
        '                     Order By gU.decodeEmptyCInt(q("pld_split_fr_int").ToString, 0), _
        '                     gU.decodeEmptyCInt(q("pld_Seq").ToString, 0), _
        '                     q("PLD_BATCH_NO").ToString, _
        '                     q("PLD_WH").ToString, _
        '                     gU.decodeEmptyCInt(q("PLD_FLOOR").ToString, 0), _
        '                     q("PLD_AREA").ToString, _
        '                     q("PLD_RACK").ToString, _
        '                     q("PLD_BIN").ToString, _
        '                     gU.decodeEmptyCdbl(q("PLD_FOI_QTY").ToString, 0), _
        '                     gU.decodeEmptyCdbl(q("PLD_ITEM_QTY").ToString, 0) Descending

        'pickListDt = query.CopyToDataTable()





        'pickListDt = dv.ToTable

        getPickList = pickListDt

    End Function

    'Used to sort pickList
    Private Shared Function CompareDtByLoc(ByVal x As DtItems, ByVal y As DtItems) As Integer
        If IsNothing(x) Then
            If IsNothing(y) Then
                ' If x is Nothing and y is Nothing, they're equal. 
                Return 0
            Else
                ' If x is Nothing and y is not Nothing, y is greater. 
                Return -1
            End If
        Else
            ' If x is not Nothing...
            If IsNothing(y) Then
                ' ...and y is Nothing, x is greater.
                Return 1
            Else
                ' ...and y is not Nothing, compare the loc
                Dim retval As Integer = x.wh.CompareTo(y.wh)

                If retval <> 0 Then
                    Return retval
                Else
                    retval = x.floor.CompareTo(y.floor)
                    If retval <> 0 Then
                        Return retval
                    Else
                        retval = x.area.CompareTo(y.area)
                        If retval <> 0 Then
                            Return retval
                        Else
                            retval = x.rack.CompareTo(y.rack)
                            If retval <> 0 Then
                                Return retval
                            Else
                                retval = x.bin.CompareTo(y.bin)
                                Return retval
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Function

    Private Sub pickBestItems()
        Dim itmComboDict As Dictionary(Of String, List(Of Combo))
        Dim allComboList As List(Of List(Of Combo))
        Dim bestComboList As List(Of DtItems)
        Dim tmpDtItem As DtItems
        Dim i As Integer

        If itmDict.Count > 0 Then
            'Get combos for each item
            itmComboDict = getFullItemCombo()
            'Get all combos within all items
            allComboList = getAllCombo(itmComboDict)
            'Get the best combo among all valid combo
            bestComboList = getBestCombo(allComboList)

            For i = 0 To bestComboList.Count - 1
                tmpDtItem = bestComboList(i)

                If itmDict.ContainsKey(tmpDtItem.itm_key) Then
                    If itmDict(tmpDtItem.itm_key) <= tmpDtItem.bal_qty Then
                        If itmDict(tmpDtItem.itm_key) < tmpDtItem.bal_qty Then
                            tmpDtItem.bal_qty = itmDict(tmpDtItem.itm_key)
                        End If
                        itmDict.Remove(tmpDtItem.itm_key)
                    Else
                        itmDict(tmpDtItem.itm_key) = itmDict(tmpDtItem.itm_key) - tmpDtItem.bal_qty
                    End If
                End If

                pickList.Add(tmpDtItem)
                dtDict.Remove(tmpDtItem.dt_key)
            Next
        End If
    End Sub

    Private Function getFullItemCombo() As Dictionary(Of String, List(Of Combo))
        Dim i As Integer
        Dim prevDtItem, tmpDtItem As DtItems
        Dim keys As Dictionary(Of String, DtItems).KeyCollection
        Dim itmList As List(Of DtItems)
        Dim itmCombo As List(Of Combo)
        Dim itmComboDict As Dictionary(Of String, List(Of Combo))

        keys = dtDict.Keys

        tmpDtItem = Nothing
        prevDtItem = Nothing
        itmList = New List(Of DtItems)

        If keys.Count > 0 Then
            itmComboDict = New Dictionary(Of String, List(Of Combo))


            For i = 0 To keys.Count - 1
                tmpDtItem = dtDict.Item(keys(i))

                If i <> 0 Then
                    If tmpDtItem.itm_key <> prevDtItem.itm_key Then
                        itmCombo = getItemCombo(itmList)
                        itmComboDict.Add(prevDtItem.itm_key, itmCombo)
                        itmList = New List(Of DtItems)
                    End If
                End If
                itmList.Add(tmpDtItem)
                prevDtItem = tmpDtItem
            Next
            itmCombo = getItemCombo(itmList)
            itmComboDict.Add(prevDtItem.itm_key, itmCombo)

            getFullItemCombo = itmComboDict
        Else
            getFullItemCombo = Nothing
        End If
    End Function

    Private Function getAllCombo(ByRef itmComboDict As Dictionary(Of String, List(Of Combo))) As List(Of List(Of Combo))
        Dim i, j, k As Integer
        Dim keys As Dictionary(Of String, List(Of Combo)).KeyCollection
        Dim itemComboList, tmpComboList As List(Of Combo)
        Dim allComboList As List(Of List(Of Combo))
        Dim copyList As List(Of Combo)

        keys = itmComboDict.Keys
        If keys.Count > 0 Then
            allComboList = New List(Of List(Of Combo))

            'Loop to generate all combination of all items
            For i = 0 To keys.Count - 1
                itemComboList = itmComboDict.Item(keys(i))

                If i = 0 Then
                    'Make initial list from first items
                    For j = 0 To itemComboList.Count - 1
                        tmpComboList = New List(Of Combo)
                        tmpComboList.Add(itemComboList(j))
                        allComboList.Add(tmpComboList)
                    Next
                Else
                    'Generate different combo to previous list
                    For k = 0 To allComboList.Count - 1
                        'Copy the existing combo list
                        copyList = New List(Of Combo)(allComboList(k))
                        'For each new item combo, generate initial x new item combo lists
                        For j = 0 To itemComboList.Count - 1
                            If j = 0 Then
                                'Use the existing list for the first item
                                allComboList(k).Add(itemComboList(j))
                            ElseIf j = itemComboList.Count - 1 Then
                                'Use the copy item for the last item
                                copyList.Add(itemComboList(j))
                                allComboList.Add(copyList)
                            Else
                                'Copy from copy item to generate new list
                                tmpComboList = New List(Of Combo)(copyList)
                                tmpComboList.Add(itemComboList(j))
                                allComboList.Add(tmpComboList)
                            End If
                        Next
                    Next
                End If
            Next
            getAllCombo = allComboList
        Else
            getAllCombo = Nothing
        End If
    End Function

    Private Function getBestCombo(ByRef allComboList As List(Of List(Of Combo))) As List(Of DtItems)
        Dim i, j, k As Integer
        Dim comboRank, bestRank As Rank
        Dim whDict, floorDict, areaDict, rackDict, binDict As Dictionary(Of String, String)
        Dim tmpWhDict, tmpFloorDict, tmpAreaDict, tmpRackDict, tmpBinDict As Dictionary(Of String, String)
        Dim tmpDtItem As DtItems
        Dim isWorse, firstCombo As Boolean
        Dim bestRankKey As Integer

        If allComboList.Count > 0 Then
            'Now we have all valid combination of item, next is to find out the best one
            whDict = New Dictionary(Of String, String)
            floorDict = New Dictionary(Of String, String)
            areaDict = New Dictionary(Of String, String)
            rackDict = New Dictionary(Of String, String)
            binDict = New Dictionary(Of String, String)

            'Scan for picked items first
            For i = 0 To pickList.Count - 1
                If whDict.ContainsKey(pickList(i).wh) Then
                    If floorDict.ContainsKey(pickList(i).wh & "#_#" & pickList(i).floor) Then
                        If areaDict.ContainsKey(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area) Then
                            If rackDict.ContainsKey(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack) Then
                                If Not binDict.ContainsKey(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack & "#_#" & pickList(i).bin) Then
                                    binDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack & "#_#" & pickList(i).bin, "")
                                End If
                            Else
                                rackDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack, "")
                                binDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack & "#_#" & pickList(i).bin, "")
                            End If

                        Else
                            areaDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area, "")
                            rackDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack, "")
                            binDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack & "#_#" & pickList(i).bin, "")
                        End If
                    Else
                        floorDict.Add(pickList(i).wh & "#_#" & pickList(i).floor, "")
                        areaDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area, "")
                        rackDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack, "")
                        binDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack & "#_#" & pickList(i).bin, "")
                    End If
                Else
                    whDict.Add(pickList(i).wh, "")
                    floorDict.Add(pickList(i).wh & "#_#" & pickList(i).floor, "")
                    areaDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area, "")
                    rackDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack, "")
                    binDict.Add(pickList(i).wh & "#_#" & pickList(i).floor & "#_#" & pickList(i).area & "#_#" & pickList(i).rack & "#_#" & pickList(i).bin, "")
                End If
            Next

            bestRank = Nothing
            bestRankKey = 0
            firstCombo = True

            'Loop to find out the best item combo
            For i = 0 To allComboList.Count - 1
                'Initial a new rank for each combo
                comboRank = New Rank
                comboRank.wh = 0
                comboRank.floor = 0
                comboRank.area = 0
                comboRank.bin = 0

                'Make a copy of picked location dicts
                tmpWhDict = New Dictionary(Of String, String)(whDict)
                tmpFloorDict = New Dictionary(Of String, String)(floorDict)
                tmpAreaDict = New Dictionary(Of String, String)(areaDict)
                tmpRackDict = New Dictionary(Of String, String)(rackDict)
                tmpBinDict = New Dictionary(Of String, String)(binDict)

                isWorse = False

                For j = 0 To allComboList(i).Count - 1

                    For k = 0 To allComboList(i)(j).itmList.Count - 1
                        tmpDtItem = allComboList(i)(j).itmList(k)

                        'Check for picked items to evaluate rank
                        If tmpWhDict.ContainsKey(tmpDtItem.wh) Then
                            If tmpFloorDict.ContainsKey(tmpDtItem.wh & "#_#" & tmpDtItem.floor) Then
                                If tmpAreaDict.ContainsKey(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area) Then
                                    If tmpRackDict.ContainsKey(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.rack) Then
                                        If Not tmpBinDict.ContainsKey(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.rack & "#_#" & tmpDtItem.bin) Then
                                            comboRank.bin += comboRank.bin + 1
                                            tmpBinDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.rack & "#_#" & tmpDtItem.bin, "")
                                        End If
                                    Else
                                        comboRank.rack += comboRank.rack + 1
                                        tmpRackDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.rack, "")
                                        tmpBinDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.bin & "#_#" & tmpDtItem.rack, "")
                                    End If
                                Else
                                    comboRank.area += comboRank.area + 1
                                    tmpAreaDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area, "")
                                    tmpRackDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.rack, "")
                                    tmpBinDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.bin & "#_#" & tmpDtItem.rack, "")
                                End If
                            Else
                                comboRank.floor += comboRank.floor + 1
                                tmpFloorDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor, "")
                                tmpAreaDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area, "")
                                tmpRackDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.rack, "")
                                tmpBinDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.bin & "#_#" & tmpDtItem.rack, "")
                            End If
                        Else
                            comboRank.wh = comboRank.wh + 1
                            tmpWhDict.Add(tmpDtItem.wh, "")
                            tmpFloorDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor, "")
                            tmpAreaDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area, "")
                            tmpRackDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.rack, "")
                            tmpBinDict.Add(tmpDtItem.wh & "#_#" & tmpDtItem.floor & "#_#" & tmpDtItem.area & "#_#" & tmpDtItem.bin & "#_#" & tmpDtItem.rack, "")
                        End If

                        If Not firstCombo Then
                            If Not isBetterRank(bestRank, comboRank) Then
                                isWorse = True
                                Exit For
                            End If
                        End If
                    Next

                    If isWorse Then
                        Exit For
                    End If
                Next

                If firstCombo Then
                    bestRank = comboRank
                    bestRankKey = i
                    firstCombo = False
                ElseIf Not isWorse Then
                    If isBetterRank(bestRank, comboRank) Then
                        bestRank = comboRank
                        bestRankKey = i
                    End If
                End If
            Next

            getBestCombo = New List(Of DtItems)

            For j = 0 To allComboList(bestRankKey).Count - 1
                For k = 0 To allComboList(bestRankKey)(j).itmList.Count - 1
                    getBestCombo.Add(allComboList(bestRankKey)(j).itmList(k))
                Next
            Next
        Else
            getBestCombo = Nothing
        End If
    End Function

    Private Function isBetterRank(ByRef bestRank As Rank, ByRef comboRank As Rank) As Boolean
        isBetterRank = False

        If comboRank.wh < bestRank.wh Then
            isBetterRank = True
        ElseIf comboRank.wh = bestRank.wh Then
            If comboRank.floor < bestRank.floor Then
                isBetterRank = True
            ElseIf comboRank.floor = bestRank.floor Then
                If comboRank.area < bestRank.area Then
                    isBetterRank = True
                ElseIf comboRank.area = bestRank.area Then
                    If comboRank.rack < bestRank.rack Then
                        isBetterRank = True
                    ElseIf comboRank.rack = bestRank.rack Then
                        If comboRank.bin < bestRank.bin Then
                            isBetterRank = True
                        End If
                    End If
                End If
            End If
        End If
    End Function

    Private Function getItemCombo(ByRef itmList As List(Of DtItems)) As List(Of Combo)
        Dim i, j As Integer
        Dim requiredQty As Double
        Dim itemComboList As List(Of Combo)
        Dim tmpItmList As List(Of DtItems)
        Dim copyList As List(Of DtItems)
        Dim tmpCombo As Combo
        Dim validComboCnt As Integer = 0

        itemComboList = New List(Of Combo)

        If itmList.Count > 0 Then
            requiredQty = itmDict(itmList(0).itm_key)

            'Loop to get all location combo for one item
            For i = 0 To itmList.Count - 1
                For j = 0 To itemComboList.Count - 1
                    'If the combo is already valid, no more item is needed
                    If Not itemComboList(j).isValid Then
                        tmpItmList = itemComboList(j).itmList
                        copyList = New List(Of DtItems)(tmpItmList)
                        copyList.Add(itmList(i))
                        tmpCombo = New Combo
                        tmpCombo.cmBalQty = itemComboList(j).cmBalQty + itmList(i).bal_qty
                        If tmpCombo.cmBalQty < requiredQty Then
                            tmpCombo.isValid = False
                        Else
                            tmpCombo.isValid = True
                            validComboCnt = validComboCnt + 1
                        End If
                        tmpCombo.itmList = copyList
                        itemComboList.Add(tmpCombo)

                        If validComboCnt >= itm_combo_limit Then
                            Exit For
                        End If
                    End If
                Next

                If validComboCnt >= itm_combo_limit Then
                    Exit For
                End If

                tmpItmList = New List(Of DtItems)
                tmpItmList.Add(itmList(i))
                tmpCombo = New Combo
                tmpCombo.cmBalQty = itmList(i).bal_qty
                If tmpCombo.cmBalQty < requiredQty Then
                    tmpCombo.isValid = False
                Else
                    tmpCombo.isValid = True
                    validComboCnt = validComboCnt + 1
                End If
                tmpCombo.itmList = tmpItmList
                itemComboList.Add(tmpCombo)

                If validComboCnt >= itm_combo_limit Then
                    Exit For
                End If
            Next

            For i = itemComboList.Count - 1 To 0 Step -1
                If Not itemComboList(i).isValid Then
                    itemComboList.RemoveAt(i)
                End If
            Next
        End If

        getItemCombo = itemComboList
    End Function


    Private Sub pickUniqueItem()
        'Dim locBalItem As LocBal
        Dim prevDtItem, tmpDtItem, uniqueItem, prevListItem As DtItems
        Dim i, j As Integer
        Dim keys As Dictionary(Of String, DtItems).KeyCollection
        Dim isUnique As Boolean
        Dim cmBalQty, pickedQty As Double
        Dim uniqueList As List(Of DtItems)
        Dim prevItemList As List(Of DtItems)
        Dim removeList As List(Of String)

        keys = dtDict.Keys

        tmpDtItem = Nothing
        prevDtItem = Nothing
        uniqueItem = Nothing
        isUnique = True
        cmBalQty = 0
        pickedQty = 0
        uniqueList = New List(Of DtItems)
        prevItemList = New List(Of DtItems)

        If keys.Count > 0 Then
            'Note: The dtDict is well order with item key and bal qty (desc)
            For i = 0 To keys.Count - 1
                tmpDtItem = dtDict.Item(keys(i))

                If i <> 0 Then
                    If tmpDtItem.itm_key = prevDtItem.itm_key Then
                        'Check if cumulative qty already > req qty, if not then more loc is need to fullfill the req qty, so still unique
                        If cmBalQty >= tmpDtItem.qty Then
                            'if replace last unique loc with current loc, req qty also can be fullfilled with cumulative qty, that means the combination is not unique
                            If cmBalQty - uniqueList(uniqueList.Count - 1).bal_qty + tmpDtItem.bal_qty >= tmpDtItem.qty Then
                                isUnique = False
                            End If
                        End If
                        If isUnique Then
                            uniqueList.Add(tmpDtItem)
                        End If
                    ElseIf tmpDtItem.itm_key <> prevDtItem.itm_key Then
                        If isUnique Then
                            pickedQty = 0
                            For j = 0 To uniqueList.Count - 1
                                uniqueItem = uniqueList(j)

                                'Mark picked flag in loc
                                'locBalItem = locDict.Item(uniqueItem.loc)
                                'locBalItem.isPicked = True

                                'add item to picklist
                                If pickedQty + uniqueItem.bal_qty > uniqueItem.qty Then
                                    uniqueItem.bal_qty = uniqueItem.qty - pickedQty
                                End If

                                pickList.Add(uniqueItem)

                                pickedQty = pickedQty + uniqueItem.bal_qty

                                'Unique list <> picklist, if pickedQty enough then no need pick the rest items
                                If pickedQty >= uniqueItem.qty Then
                                    Exit For
                                End If
                            Next
                            'pickedItmDict.Add(prevDtItem.itm_key, pickedQty)
                            itmDict.Remove(prevDtItem.itm_key)
                        Else
                            For j = 0 To prevItemList.Count - 1
                                prevListItem = prevItemList(j)

                                'If cumulative qty - item qty < req qty, then this item is unique (altough combination not unique)
                                If cmBalQty - prevItemList(j).bal_qty < prevDtItem.qty Then
                                    'Mark picked flag in loc
                                    'locBalItem = locDict.Item(prevItemList(j).loc)
                                    'locBalItem.isPicked = True

                                    'pickedItmDict.Add(prevDtItem.itm_key, prevItemList(j).bal_qty)

                                    If prevListItem.bal_qty > itmDict.Item(prevDtItem.itm_key) Then
                                        prevListItem.bal_qty = itmDict.Item(prevDtItem.itm_key)
                                        itmDict.Remove(prevDtItem.itm_key)
                                        'add item to picklist
                                        pickList.Add(prevListItem)
                                        Exit For
                                    Else
                                        itmDict.Item(prevDtItem.itm_key) = itmDict.Item(prevDtItem.itm_key) - prevListItem.bal_qty
                                        'add item to picklist
                                        pickList.Add(prevListItem)
                                    End If
                                    'itmDict.Item(prevDtItem.itm_key) = itmDict.Item(prevDtItem.itm_key) - prevListItem.bal_qty
                                End If
                            Next
                        End If
                        isUnique = True
                        cmBalQty = 0
                        uniqueList = New List(Of DtItems)
                        uniqueList.Add(tmpDtItem)
                        prevItemList = New List(Of DtItems)
                    End If
                Else
                    uniqueList.Add(tmpDtItem)
                End If
                cmBalQty = cmBalQty + tmpDtItem.bal_qty
                prevItemList.Add(tmpDtItem)
                prevDtItem = tmpDtItem
            Next

            'Run the same logic for "If tmpDtItem.itm_key <> prevDtItem.itm_key Then" at the end of loop -----------------------------------------------------
            If isUnique Then
                pickedQty = 0
                For j = 0 To uniqueList.Count - 1
                    uniqueItem = uniqueList(j)

                    'Mark picked flag in loc
                    'locBalItem = locDict.Item(uniqueItem.loc)
                    'locBalItem.isPicked = True

                    'add item to picklist
                    If pickedQty + uniqueItem.bal_qty > uniqueItem.qty Then
                        uniqueItem.bal_qty = uniqueItem.qty - pickedQty
                    End If

                    pickList.Add(uniqueItem)

                    pickedQty = pickedQty + uniqueItem.bal_qty

                    'Unique list <> picklist, if pickedQty enough then no need pick the rest items
                    If pickedQty >= uniqueItem.qty Then
                        Exit For
                    End If
                Next
                'pickedItmDict.Add(prevDtItem.itm_key, pickedQty)
                itmDict.Remove(prevDtItem.itm_key)
            Else
                For j = 0 To prevItemList.Count - 1
                    prevListItem = prevItemList(j)

                    'If cumulative qty - item qty < req qty, then this item is unique (altough combination not unique)
                    If cmBalQty - prevListItem.bal_qty < prevDtItem.qty Then
                        'Mark picked flag in loc
                        'locBalItem = locDict.Item(prevListItem.loc)
                        'locBalItem.isPicked = True
                        'add item to picklist

                        'pickedItmDict.Add(prevDtItem.itm_key, prevListItem.bal_qty)

                        If prevListItem.bal_qty > itmDict.Item(prevDtItem.itm_key) Then
                            prevListItem.bal_qty = itmDict.Item(prevDtItem.itm_key)
                            itmDict.Remove(prevDtItem.itm_key)
                            'add item to picklist
                            pickList.Add(prevListItem)
                            Exit For
                        Else
                            itmDict.Item(prevDtItem.itm_key) = itmDict.Item(prevDtItem.itm_key) - prevListItem.bal_qty
                            'add item to picklist
                            pickList.Add(prevListItem)
                        End If
                        'itmDict.Item(prevDtItem.itm_key) = itmDict.Item(prevDtItem.itm_key) - prevListItem.bal_qty
                    End If
                Next
            End If
            '-------------------------------------------------------------------------------------------------------------------------------------------------------------

            'Remove picked dt items
            For i = 0 To pickList.Count - 1
                dtDict.Remove(pickList(i).dt_key)
            Next

            'Find out all items which have been picked with other locations
            removeList = New List(Of String)
            keys = dtDict.Keys

            For i = 0 To keys.Count - 1
                If Not itmDict.ContainsKey(dtDict.Item(keys(i)).itm_key) Then
                    removeList.Add(dtDict.Item(keys(i)).dt_key)
                End If
            Next
            'Remove items
            For i = 0 To removeList.Count - 1
                dtDict.Remove(removeList(i))
            Next

        End If
    End Sub

    Public Property isPickByBatchNo() As Boolean
        Get
            Return pickByBatchNo
        End Get
        Set(value As Boolean)
            pickByBatchNo = value
        End Set
    End Property
End Class
