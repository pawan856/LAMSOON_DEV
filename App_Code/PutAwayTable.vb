Imports Microsoft.VisualBasic
Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Public Class PutAwayTable
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private gU As New GeneralUtils

    Private cnn As SqlConnection
    Private transaction As SqlTransaction

    Private paDelList As String
    Private grDt, paDt As DataTable
    Private paSeq As String

    Private inspLoc As Dictionary(Of String, itmLocQty)

    Structure itmStruct
        Dim itmName As String
        Dim qty As Double
        Dim itmSeq As String
        Dim batchNo As String
        Dim refNo As String
        Dim skuNo As String
        Dim expDate As String
        Dim manuDate As String
        Dim whCode As String
        Dim MainWhCode As String
        Dim reqInsp As String
    End Structure

    Structure itmLocQty
        Dim loc As String
        Dim wh As String
        Dim qty As Double
    End Structure

    Public Sub New(Optional ByRef pCnn As SqlConnection = Nothing, Optional ByRef pTrans As SqlTransaction = Nothing)
        paDt = Nothing
        paSeq = "0"
        paDelList = ""
        cnn = pCnn
        transaction = pTrans
    End Sub

    Public Sub New(ByRef paraGrDt As DataTable, ByRef paraPaDt As DataTable, ByVal paraPaSeq As String, ByVal paraPaDelList As String, Optional ByRef pCnn As SqlConnection = Nothing, Optional ByRef pTrans As SqlTransaction = Nothing)
        grDt = paraGrDt
        paDt = paraPaDt
        paSeq = paraPaSeq
        paDelList = paraPaDelList
        cnn = pCnn
        transaction = pTrans
    End Sub

    Private Function getItemLocQty(ByVal itemCode As String, ByVal packKey As String, ByVal itemQty As Long, ByVal warehouse As String, ByVal reqInspYN As String) As List(Of itmLocQty)
        Dim selectSql As String
        Dim tmpDt As DataTable
        Dim tmpLocList As List(Of itmLocQty)
        Dim tmpItemLocQty As itmLocQty
        Dim IW_PREF_LOC1, IW_PREF_LOC2, LOC1_UTIL_TYPE, LOC2_UTIL_TYPE, ITEM_CBM, ITEM_KG, ITEM_AREA As String
        Dim LOC1_CBM, LOC2_CBM, LOC1_AREA, LOC2_AREA, LOC1_CBM_BAL, LOC1_KG_BAL, LOC1_AREA_BAL, LOC2_CBM_BAL, LOC2_KG_BAL, LOC2_AREA_BAL, LOC1_WH, LOC2_WH As String
        Dim utlLimit As Double = 0.75
        Dim checkLoc2 As Boolean = False
        Dim useLoc1, useLoc2 As Boolean

        tmpLocList = New List(Of itmLocQty)


        If reqInspYN = "Y" Then
            If inspLoc.ContainsKey(warehouse) Then
                tmpItemLocQty = New itmLocQty
                tmpItemLocQty.loc = inspLoc.Item(warehouse).loc
                tmpItemLocQty.wh = inspLoc.Item(warehouse).wh
                tmpItemLocQty.qty = itemQty

                tmpLocList.Add(tmpItemLocQty)

                Return tmpLocList
            End If
        End If

        selectSql = "select iw.IW_PREF_LOC1, iw.IW_PREF_LOC2, " & _
                        "b1.BN_UTILIZATION_TYPE as LOC1_UTIL_TYPE, " & _
                        "b2.BN_UTILIZATION_TYPE as LOC2_UTIL_TYPE, " & _
                        "v.carton_cbm as item_cbm, v.aitm_vol as item_kg, v.aitm_length * v.aitm_width as item_area, " & _
                        "b1.BN_LENGTH * b1.BN_WIDTH * b1.BN_DEPTH / 1000000 as LOC1_CBM, " & _
                        "b1.BN_LENGTH * b1.BN_WIDTH * b1.BN_DEPTH / 1000000 as LOC2_CBM, " & _
                        "b1.BN_LENGTH * b1.BN_WIDTH / 10000 as LOC1_AREA, " & _
                        "b1.BN_LENGTH * b1.BN_WIDTH / 10000 as LOC2_AREA, " & _
                        "bal1.total_cbm as loc1_CBM_BAL, bal1.total_kg as loc1_kg_bal, bal1.total_area as loc1_area_bal, " & _
                        "bal2.total_cbm as loc2_CBM_BAL, bal2.total_kg as loc2_kg_bal, bal2.total_area as loc2_area_bal, " & _
                        "b1.wh_code as loc1_wh, b2.wh_code as loc2_wh " & _
                    "from WMS_ITEM_WH iw " & _
                    "inner join v_alt_vend_item v " & _
                    "on iw.ITM_CODE = v.ITM_CODE " & _
                    "and iw.PACK_KEY = v.PACK_KEY " & _
                    "left outer join V_LOCATION b1 " & _
                    "on iw.IW_PREF_LOC1 = b1.loc " & _
                    "left outer join V_LOCATION b2 " & _
                    "on iw.IW_PREF_LOC2 = b2.loc " & _
                    "left outer join ( " & _
                        "select b.ILOC_LOC, sum(i.carton_cbm * b.ILOC_BAL_QTY) as total_cbm, sum(i.aitm_vol * b.ILOC_BAL_QTY) as total_kg, sum(i.aitm_length * i.aitm_width * b.ILOC_BAL_QTY) as total_area " & _
                        "from WMS_ITEM_LOC_BAL b, v_alt_vend_item i " & _
                        "where b.imp_code = i.imp_code " & _
                        "and b.STORER_CODE = i.STORER_CODE " & _
                        "and b.ITM_CODE = i.ITM_CODE " & _
                        "and b.pack_key = i.PACK_KEY " & _
                        "group by b.ILOC_LOC) bal1 " & _
                    "on iw.IW_PREF_LOC1 = bal1.ILOC_LOC " & _
                    "left outer join ( " & _
                        "select b.ILOC_LOC, sum(i.carton_cbm * b.ILOC_BAL_QTY) as total_cbm, sum(i.aitm_vol * b.ILOC_BAL_QTY) as total_kg, sum(i.aitm_length * i.aitm_width * b.ILOC_BAL_QTY) as total_area " & _
                        "from WMS_ITEM_LOC_BAL b, v_alt_vend_item i " & _
                        "where b.imp_code = i.imp_code " & _
                        "and b.STORER_CODE = i.STORER_CODE " & _
                        "and b.ITM_CODE = i.ITM_CODE " & _
                        "and b.pack_key = i.PACK_KEY " & _
                        "group by b.ILOC_LOC) bal2 " & _
                    "on iw.IW_PREF_LOC2 = bal2.ILOC_LOC " & _
                    "where iw.ITM_CODE = '" & gU.dbEncode(itemCode) & "' " & _
                    "and iw.PACK_KEY = '" & gU.dbEncode(packKey) & "' " & _
                    "and iw.WH_CODE = '" & gU.dbEncode(warehouse) & "' "

        tmpDt = gDB.getDataTable(selectSql, cnn, transaction)

        If tmpDt.Rows.Count > 0 Then
            IW_PREF_LOC1 = tmpDt.Rows(0).Item("IW_PREF_LOC1").ToString.Trim
            IW_PREF_LOC2 = tmpDt.Rows(0).Item("IW_PREF_LOC2").ToString.Trim
            LOC1_UTIL_TYPE = tmpDt.Rows(0).Item("LOC1_UTIL_TYPE").ToString.Trim
            LOC2_UTIL_TYPE = tmpDt.Rows(0).Item("LOC2_UTIL_TYPE").ToString.Trim
            ITEM_CBM = tmpDt.Rows(0).Item("ITEM_CBM").ToString.Trim
            ITEM_KG = tmpDt.Rows(0).Item("ITEM_KG").ToString.Trim
            ITEM_AREA = tmpDt.Rows(0).Item("ITEM_AREA").ToString.Trim
            LOC1_CBM = tmpDt.Rows(0).Item("LOC1_CBM").ToString.Trim
            LOC2_CBM = tmpDt.Rows(0).Item("LOC2_CBM").ToString.Trim
            LOC1_AREA = tmpDt.Rows(0).Item("LOC1_AREA").ToString.Trim
            LOC2_AREA = tmpDt.Rows(0).Item("LOC2_AREA").ToString.Trim
            LOC1_CBM_BAL = tmpDt.Rows(0).Item("LOC1_CBM_BAL").ToString.Trim
            LOC1_KG_BAL = tmpDt.Rows(0).Item("LOC1_KG_BAL").ToString.Trim
            LOC1_AREA_BAL = tmpDt.Rows(0).Item("LOC1_AREA_BAL").ToString.Trim
            LOC2_CBM_BAL = tmpDt.Rows(0).Item("LOC2_CBM_BAL").ToString.Trim
            LOC2_KG_BAL = tmpDt.Rows(0).Item("LOC2_KG_BAL").ToString.Trim
            LOC2_AREA_BAL = tmpDt.Rows(0).Item("LOC2_AREA_BAL").ToString.Trim

            LOC1_WH = tmpDt.Rows(0).Item("LOC1_WH").ToString.Trim
            LOC2_WH = tmpDt.Rows(0).Item("LOC2_WH").ToString.Trim

            useLoc1 = False
            useLoc2 = False

            If IW_PREF_LOC1 <> "" AndAlso IW_PREF_LOC2 = "" Then
                useLoc1 = True
            ElseIf IW_PREF_LOC2 <> "" AndAlso IW_PREF_LOC1 = "" Then
                useLoc2 = True

            ElseIf IW_PREF_LOC2 = "" AndAlso IW_PREF_LOC1 = "" Then


            Else

                'If LOC1_UTIL_TYPE is empty, no need check ultilization, use prefer location 1
                If LOC1_UTIL_TYPE = "" Then
                    useLoc1 = True

                ElseIf LOC1_UTIL_TYPE = "CBM" Then

                    'If location 1 CBM is empty or item CBM is empt, no need check ultilization, use prefer location 1
                    If LOC1_CBM = "" OrElse LOC1_CBM = "0" OrElse ITEM_CBM = "" OrElse ITEM_CBM = "0" Then
                        useLoc1 = True

                        'If item CBM + balance CBM <= location 1 utiliation limit, use prefer location 1
                    ElseIf (LOC1_CBM - CDbl(ITEM_CBM) * itemQty + gU.decodeEmptyCdbl(LOC1_CBM_BAL, 0)) / LOC1_CBM <= utlLimit Then
                        useLoc1 = True
                    End If

                ElseIf LOC1_UTIL_TYPE = "AREA" Then

                    'If location 1 AREA is empty or item AREA is empt, no need check ultilization, use prefer location 1
                    If LOC1_AREA = "" OrElse LOC1_AREA = "0" OrElse ITEM_AREA = "" OrElse ITEM_AREA = "0" Then
                        useLoc1 = True

                        'If item AREA + balance AREA <= location 1 utiliation limit, use prefer location 1
                    ElseIf (LOC1_AREA - CDbl(ITEM_AREA) * itemQty + gU.decodeEmptyCdbl(LOC1_AREA_BAL, 0)) / LOC1_AREA <= utlLimit Then
                        useLoc1 = True
                    End If

                End If

                If Not useLoc1 Then

                    If LOC2_UTIL_TYPE = "CBM" Then
                        If LOC2_CBM = "" OrElse LOC2_CBM = "0" Then
                            useLoc2 = True

                            'If item CBM + balance CBM <= location 2 utiliation limit, use prefer location 2
                        ElseIf (LOC2_CBM - CDbl(ITEM_CBM) * itemQty + gU.decodeEmptyCdbl(LOC2_CBM_BAL, 0)) / LOC2_CBM <= utlLimit Then
                            useLoc2 = True
                        End If

                    ElseIf LOC1_UTIL_TYPE = "AREA" Then

                        'If location 1 AREA is empty or item AREA is empt, no need check ultilization, use prefer location 1
                        If LOC1_AREA = "" OrElse LOC1_AREA = "0" OrElse ITEM_AREA = "" OrElse ITEM_AREA = "0" Then
                            useLoc2 = True

                            'If item AREA + balance AREA <= location 1 utiliation limit, use prefer location 1
                        ElseIf (LOC1_AREA - CDbl(ITEM_AREA) * itemQty + gU.decodeEmptyCdbl(LOC1_AREA_BAL, 0)) / LOC1_AREA <= utlLimit Then
                            useLoc2 = True
                        End If
                    End If

                    'If both location utilization is full, use back location 1
                    'This part can implement more complex logic in the future
                    If Not useLoc2 Then
                        useLoc1 = True
                    End If

                End If
            End If


            If useLoc1 Then
                tmpItemLocQty = New itmLocQty
                tmpItemLocQty.loc = IW_PREF_LOC1
                tmpItemLocQty.wh = LOC1_WH
                tmpItemLocQty.qty = itemQty

                tmpLocList.Add(tmpItemLocQty)
            ElseIf useLoc2 Then
                tmpItemLocQty = New itmLocQty
                tmpItemLocQty.loc = IW_PREF_LOC2
                tmpItemLocQty.wh = LOC2_WH
                tmpItemLocQty.qty = itemQty

                tmpLocList.Add(tmpItemLocQty)
            End If

            Return tmpLocList
        Else
            Return tmpLocList
        End If

    End Function

    Private Sub setInspDict()
        Dim selectSql As String
        Dim i As Integer
        Dim tmpDt As DataTable
        Dim tmpItem As itmLocQty

        inspLoc = New Dictionary(Of String, itmLocQty)

        selectSql = "select w.wh_main_wh, b.WH_CODE, ISNULL(b.WH_CODE, '') " & _
                        "+ ISNULL(ssma_oracle.lpad_nvarchar(b.FL_NUM, 2, N'0'), '') " & _
                        "+ ISNULL(ssma_oracle.lpad_varchar(b.AR_CODE, 3, '0'), '') " & _
                        "+ ISNULL(ssma_oracle.lpad_varchar(b.RK_CODE, 4, '0'), '') " & _
                        "+ ISNULL(ssma_oracle.lpad_varchar(b.BN_CODE, 3, '0'), '') AS LOC " & _
                    "from wms_wh_area a, wms_wh_bin b, WMS_WAREHOUSE w " & _
                    "where a.imp_code = b.imp_code " & _
                    "and a.wh_code = b.wh_code " & _
                    "and a.fl_num = b.fl_num " & _
                    "and a.ar_code = b.ar_code " & _
                    "and b.imp_code = w.imp_code " & _
                    "and b.WH_CODE = w.WH_CODE " & _
                    "and a.ar_insp_area = 'Y' " & _
                    "order by b.wh_code "

        'selectSql = "select b.wh_code, ISNULL(b.WH_CODE, '') " & _

        tmpDt = gDB.getDataTable(selectSql, cnn, transaction)

        For i = 0 To tmpDt.Rows.Count - 1
            If Not inspLoc.ContainsKey(tmpDt.Rows(i).Item("wh_main_wh").ToString.Trim) Then
                tmpItem = New itmLocQty
                tmpItem.wh = tmpDt.Rows(i).Item("WH_CODE").ToString.Trim
                tmpItem.loc = tmpDt.Rows(i).Item("loc").ToString.Trim
                inspLoc.Add(tmpDt.Rows(i).Item("wh_main_wh").ToString.Trim, tmpItem)
            End If
        Next
    End Sub

    Public Sub genPutAway(ByVal regenFlag As Boolean)
        Dim rows_count As Integer
        Dim i, j As Integer
        Dim plSeq As Integer
        Dim itmDict, paDict As Dictionary(Of String, itmStruct)
        Dim itmKey As String
        Dim keys As Dictionary(Of String, itmStruct).KeyCollection
        Dim keyArray As String()
        Dim tmpStruct As itmStruct
        Dim paLocList As List(Of itmLocQty)
        Dim delPaList As List(Of String)

        If paDt Is Nothing Then
            Throw New Exception("Missing Puty Away datatable!")
        End If

        If grDt Is Nothing Then
            Throw New Exception("Missing Good Receive datatable!")
        End If

        setInspDict()

        If regenFlag Then
            'Delete all existing put away
            For i = paDt.Rows.Count - 1 To 0 Step -1
                If paDt.Rows(i).Item("mFlag") = "N" Then
                    paDt.Rows(i).Delete()
                Else
                    paDelList = gU.appendToList(paDelList, paDt.Rows(i).Item("gra_seq"))
                    paDt.Rows(i).Delete()
                End If
            Next
            paDt.AcceptChanges()

            'Set pa seq to zero
            paSeq = "0"

            plSeq = 0
        Else
            plSeq = paSeq
        End If

        If grDt.Rows.Count > 0 Then
            itmDict = New Dictionary(Of String, itmStruct)

            'Get items information
            For i = 0 To grDt.Rows.Count - 1
                If DB.decodeDBNull(grDt.Rows(i).Item("GRD_RCV_QTY"), 0) > 0 AndAlso grDt.Rows(i).Item("mFlag").ToString <> "D" AndAlso grDt.Rows(i).Item("GRD_ON_BEHALF").ToString.Trim <> "Y" Then
                    itmKey = grDt.Rows(i).Item("grd_itm_code").ToString & "#_#" &
                                grDt.Rows(i).Item("grd_pack_key").ToString & "#_#" &
                                grDt.Rows(i).Item("grd_pallet_no").ToString & "#_#" &
                                grDt.Rows(i).Item("GRD_BATCH_NO").ToString

                    If Not itmDict.ContainsKey(itmKey) Then
                        tmpStruct = New itmStruct
                        tmpStruct.itmName = gU.decodeNullOrEmpty(grDt.Rows(i).Item("GRD_ITM_NAME").ToString, "")
                        tmpStruct.itmSeq = gU.decodeNullOrEmpty(grDt.Rows(i).Item("GRD_SEQ").ToString, "")
                        tmpStruct.qty = gU.decodeEmptyCdbl(grDt.Rows(i).Item("GRD_RCV_QTY"), 0)
                        tmpStruct.batchNo = gU.decodeNullOrEmpty(grDt.Rows(i).Item("GRD_BATCH_NO").ToString, "")
                        tmpStruct.skuNo = gU.decodeNullOrEmpty(grDt.Rows(i).Item("ITM_SKU_NO").ToString, "")

                        tmpStruct.expDate = grDt.Rows(i).Item("GRD_EXPIRY_DATE").ToString.Trim
                        tmpStruct.manuDate = grDt.Rows(i).Item("GRD_MANU_DATE").ToString.Trim

                        'tmpStruct.whCode = grDt.Rows(i).Item("GRD_WH").ToString.Trim
                        tmpStruct.MainWhCode = grDt.Rows(i).Item("GRD_WH").ToString.Trim

                        tmpStruct.reqInsp = grDt.Rows(i).Item("GRD_INSP_REQ").ToString.Trim

                        itmDict.Add(itmKey, tmpStruct)
                    Else
                        tmpStruct = itmDict.Item(itmKey)
                        tmpStruct.qty = tmpStruct.qty + gU.decodeEmptyCdbl(grDt.Rows(i).Item("GRD_RCV_QTY"), 0)

                        itmDict(itmKey) = tmpStruct
                    End If
                End If
            Next

            paDict = New Dictionary(Of String, itmStruct)

            delPaList = New List(Of String)

            If Not regenFlag Then
                'Get current put away information
                For i = 0 To paDt.Rows.Count - 1

                    itmKey = paDt.Rows(i).Item("GRA_ITM_CODE").ToString & "#_#" &
                                paDt.Rows(i).Item("GRA_PACK_KEY").ToString & "#_#" &
                                paDt.Rows(i).Item("GRA_PALLET_NO").ToString & "#_#" &
                                paDt.Rows(i).Item("GRA_BATCH_NO").ToString

                    If Not paDict.ContainsKey(itmKey) Then
                        tmpStruct = New itmStruct
                        tmpStruct.itmName = gU.decodeNullOrEmpty(paDt.Rows(i).Item("DSP_ITM_NAME").ToString, "")
                        tmpStruct.itmSeq = gU.decodeNullOrEmpty(paDt.Rows(i).Item("GRA_SPLIT_FR").ToString, "")
                        If (paDt.Rows(i).Item("GRA_SUG_QTY") IsNot DBNull.Value) Then
                            tmpStruct.qty = gU.decodeEmptyCdbl(paDt.Rows(i).Item("GRA_SUG_QTY"), 0)
                        Else
                            tmpStruct.qty = 0
                        End If
                        tmpStruct.batchNo = gU.decodeNullOrEmpty(paDt.Rows(i).Item("GRA_BATCH_NO").ToString, "")
                            tmpStruct.skuNo = gU.decodeNullOrEmpty(paDt.Rows(i).Item("ITM_SKU_NO").ToString, "")

                            tmpStruct.expDate = paDt.Rows(i).Item("GRA_EXPIRY_DATE").ToString.Trim
                            tmpStruct.manuDate = paDt.Rows(i).Item("GRA_MANU_DATE").ToString.Trim

                            tmpStruct.whCode = paDt.Rows(i).Item("GRA_WH").ToString.Trim

                            tmpStruct.reqInsp = ""

                            paDict.Add(itmKey, tmpStruct)
                        Else
                        tmpStruct = paDict.Item(itmKey)
                        If (paDt.Rows(i).Item("GRA_SUG_QTY") IsNot DBNull.Value) Then
                            tmpStruct.qty = tmpStruct.qty + gU.decodeEmptyCdbl(paDt.Rows(i).Item("GRA_SUG_QTY"), 0)
                        End If

                        paDict(itmKey) = tmpStruct
                    End If
                Next

                keys = paDict.Keys

                For i = 0 To keys.Count - 1
                    If itmDict.ContainsKey(keys(i)) Then

                        tmpStruct = itmDict.Item(keys(i))

                        If tmpStruct.qty > paDict.Item(keys(i)).qty Then
                            tmpStruct.qty = tmpStruct.qty - paDict.Item(keys(i)).qty
                        Else
                            'Assume recv qty will not descrease
                            itmDict.Remove(keys(i))
                        End If
                    Else
                        delPaList.Add(keys(i))
                    End If
                Next
            End If


            keys = itmDict.Keys

            If keys.Count > 0 Then

                'Insert put away records
                For i = 0 To keys.Count - 1

                    keyArray = Split(keys(i), "#_#")

                    If Not regenFlag AndAlso paDict.ContainsKey(keys(i)) Then
                        'If item already exists in put away, increase the put away qty
                        For j = paDt.Rows.Count - 1 To 0 Step -1
                            itmKey = paDt.Rows(j).Item("GRA_ITM_CODE").ToString & "#_#" &
                                        paDt.Rows(j).Item("GRA_PACK_KEY").ToString & "#_#" &
                                        paDt.Rows(j).Item("GRA_PALLET_NO").ToString & "#_#" &
                                        paDt.Rows(j).Item("GRA_BATCH_NO").ToString

                            If keys(i) = itmKey Then
                                If (paDt.Rows(i).Item("GRA_SUG_QTY") IsNot DBNull.Value) Then
                                    paDt.Rows(j).Item("GRA_SUG_QTY") = paDt.Rows(j).Item("GRA_SUG_QTY") + itmDict.Item(keys(i)).qty
                                End If

                                Exit For
                            End If
                        Next
                    Else
                        'If it is a new item or re-gen put away, find the preferred location and add row
                        'paLocList = getItemLocQty(keyArray(0), keyArray(1), itmDict.Item(keys(i)).qty, itmDict.Item(keys(i)).whCode, itmDict.Item(keys(i)).reqInsp)
                        paLocList = getItemLocQty(keyArray(0), keyArray(1), itmDict.Item(keys(i)).qty, itmDict.Item(keys(i)).MainWhCode, itmDict.Item(keys(i)).reqInsp)

                        If paLocList.Count > 0 Then
                            For j = 0 To paLocList.Count - 1
                                paDt.Rows.Add()

                                plSeq = plSeq + 1

                                rows_count = paDt.Rows.Count

                                paDt.Rows(rows_count - 1).Item("GRA_SEQ") = CStr(plSeq)
                                paDt.Rows(rows_count - 1).Item("GRA_DISP_SEQ") = CStr(plSeq)
                                paDt.Rows(rows_count - 1).Item("GRA_PA_LIST_NO") = "1"
                                paDt.Rows(rows_count - 1).Item("GRA_ITM_CODE") = keyArray(0)
                                paDt.Rows(rows_count - 1).Item("DSP_ITM_NAME") = itmDict.Item(keys(i)).itmName
                                paDt.Rows(rows_count - 1).Item("GRA_PACK_KEY") = keyArray(1)
                                paDt.Rows(rows_count - 1).Item("GRA_PALLET_NO") = keyArray(2)
                                paDt.Rows(rows_count - 1).Item("GRA_REF_NO") = ""
                                paDt.Rows(rows_count - 1).Item("ITM_SKU_NO") = itmDict.Item(keys(i)).skuNo

                                paDt.Rows(rows_count - 1).Item("GRA_PA_QTY") = 0
                                paDt.Rows(rows_count - 1).Item("GRA_SUG_QTY") = paLocList(j).qty
                                paDt.Rows(rows_count - 1).Item("GRA_SPLIT_FR") = itmDict.Item(keys(i)).itmSeq
                                'paDt.Rows(rows_count - 1).Item("GRA_WH") = itmDict.Item(keys(i)).whCode
                                paDt.Rows(rows_count - 1).Item("GRA_WH") = paLocList(j).wh
                                paDt.Rows(rows_count - 1).Item("GRA_LOC") = paLocList(j).loc
                                paDt.Rows(rows_count - 1).Item("mFlag") = "N"

                                paDt.Rows(rows_count - 1).Item("GRA_BATCH_NO") = keyArray(3)

                                paDt.Rows(rows_count - 1).Item("GRA_EXPIRY_DATE") = itmDict.Item(keys(i)).expDate
                                paDt.Rows(rows_count - 1).Item("GRA_MANU_DATE") = itmDict.Item(keys(i)).manuDate

                                paDt.Rows(rows_count - 1).Item("GRA_REJ_QTY") = 0
                                paDt.Rows(rows_count - 1).Item("GRA_REJ_REASON") = ""
                            Next

                        Else
                            paDt.Rows.Add()

                            plSeq = plSeq + 1

                            rows_count = paDt.Rows.Count

                            paDt.Rows(rows_count - 1).Item("GRA_SEQ") = CStr(plSeq)
                            paDt.Rows(rows_count - 1).Item("GRA_DISP_SEQ") = CStr(plSeq)
                            paDt.Rows(rows_count - 1).Item("GRA_PA_LIST_NO") = "1"
                            paDt.Rows(rows_count - 1).Item("GRA_ITM_CODE") = keyArray(0)
                            paDt.Rows(rows_count - 1).Item("DSP_ITM_NAME") = itmDict.Item(keys(i)).itmName
                            paDt.Rows(rows_count - 1).Item("GRA_PACK_KEY") = keyArray(1)
                            paDt.Rows(rows_count - 1).Item("GRA_PALLET_NO") = keyArray(2)
                            paDt.Rows(rows_count - 1).Item("GRA_REF_NO") = ""
                            paDt.Rows(rows_count - 1).Item("ITM_SKU_NO") = itmDict.Item(keys(i)).skuNo

                            paDt.Rows(rows_count - 1).Item("GRA_PA_QTY") = 0
                            paDt.Rows(rows_count - 1).Item("GRA_SUG_QTY") = itmDict.Item(keys(i)).qty
                            paDt.Rows(rows_count - 1).Item("GRA_SPLIT_FR") = itmDict.Item(keys(i)).itmSeq
                            'paDt.Rows(rows_count - 1).Item("GRA_WH") = itmDict.Item(keys(i)).whCode
                            paDt.Rows(rows_count - 1).Item("GRA_WH") = ""
                            paDt.Rows(rows_count - 1).Item("GRA_LOC") = ""
                            paDt.Rows(rows_count - 1).Item("mFlag") = "N"

                            paDt.Rows(rows_count - 1).Item("GRA_BATCH_NO") = keyArray(3)

                            paDt.Rows(rows_count - 1).Item("GRA_EXPIRY_DATE") = itmDict.Item(keys(i)).expDate
                            paDt.Rows(rows_count - 1).Item("GRA_MANU_DATE") = itmDict.Item(keys(i)).manuDate

                            paDt.Rows(rows_count - 1).Item("GRA_REJ_QTY") = 0
                            paDt.Rows(rows_count - 1).Item("GRA_REJ_REASON") = ""
                        End If

                    End If

                Next

                paSeq = CStr(plSeq)

                paDt.AcceptChanges()
            End If
        End If

    End Sub

    'Public Sub genPutAway(ByVal regenFlag As Boolean)
    '    Dim rows_count As Integer
    '    Dim i, j As Integer
    '    Dim plSeq As Integer
    '    Dim itmDict, paDict As Dictionary(Of String, itmStruct)
    '    Dim itmKey As String
    '    Dim keys As Dictionary(Of String, itmStruct).KeyCollection
    '    Dim keyArray As String()
    '    Dim tmpStruct As itmStruct
    '    Dim paLocList As List(Of itmLocQty)
    '    Dim delPaList As List(Of String)

    '    If paDt Is Nothing Then
    '        Throw New Exception("Missing Puty Away datatable!")
    '    End If

    '    If grDt Is Nothing Then
    '        Throw New Exception("Missing Good Receive datatable!")
    '    End If

    '    setInspDict()

    '    If regenFlag Then
    '        'Delete all existing put away
    '        For i = paDt.Rows.Count - 1 To 0 Step -1
    '            If paDt.Rows(i).Item("mFlag") = "N" Then
    '                paDt.Rows(i).Delete()
    '            Else
    '                paDelList = gU.appendToList(paDelList, paDt.Rows(i).Item("gra_seq"))
    '                paDt.Rows(i).Delete()
    '            End If
    '        Next
    '        paDt.AcceptChanges()

    '        'Set pa seq to zero
    '        paSeq = "0"

    '        plSeq = 0
    '    Else
    '        plSeq = paSeq
    '    End If

    '    If grDt.Rows.Count > 0 Then
    '        itmDict = New Dictionary(Of String, itmStruct)

    '        'Get items information
    '        For i = 0 To grDt.Rows.Count - 1
    '            If DB.decodeDBNull(grDt.Rows(i).Item("GRD_RCV_QTY"), 0) > 0 AndAlso grDt.Rows(i).Item("mFlag").ToString <> "D" AndAlso grDt.Rows(i).Item("GRD_ON_BEHALF").ToString.Trim <> "Y" Then
    '                itmKey = grDt.Rows(i).Item("grd_itm_code").ToString & "#_#" & _
    '                            grDt.Rows(i).Item("grd_pack_key").ToString & "#_#" & _
    '                            grDt.Rows(i).Item("grd_pallet_no").ToString & "#_#" & _
    '                            grDt.Rows(i).Item("GRD_BATCH_NO").ToString

    '                If Not itmDict.ContainsKey(itmKey) Then
    '                    tmpStruct = New itmStruct
    '                    tmpStruct.itmName = gU.decodeNullOrEmpty(grDt.Rows(i).Item("GRD_ITM_NAME").ToString, "")
    '                    tmpStruct.itmSeq = gU.decodeNullOrEmpty(grDt.Rows(i).Item("GRD_SEQ").ToString, "")
    '                    tmpStruct.qty = gU.decodeEmptyCdbl(grDt.Rows(i).Item("GRD_RCV_QTY"), 0)
    '                    tmpStruct.batchNo = gU.decodeNullOrEmpty(grDt.Rows(i).Item("GRD_BATCH_NO").ToString, "")
    '                    tmpStruct.skuNo = gU.decodeNullOrEmpty(grDt.Rows(i).Item("ITM_SKU_NO").ToString, "")

    '                    tmpStruct.expDate = grDt.Rows(i).Item("GRD_EXPIRY_DATE").ToString.Trim
    '                    tmpStruct.manuDate = grDt.Rows(i).Item("GRD_MANU_DATE").ToString.Trim

    '                    'tmpStruct.whCode = grDt.Rows(i).Item("GRD_WH").ToString.Trim
    '                    tmpStruct.MainWhCode = grDt.Rows(i).Item("GRD_WH").ToString.Trim

    '                    tmpStruct.reqInsp = grDt.Rows(i).Item("GRD_INSP_REQ").ToString.Trim

    '                    itmDict.Add(itmKey, tmpStruct)
    '                Else
    '                    tmpStruct = itmDict.Item(itmKey)
    '                    tmpStruct.qty = tmpStruct.qty + gU.decodeEmptyCdbl(grDt.Rows(i).Item("GRD_RCV_QTY"), 0)

    '                    itmDict(itmKey) = tmpStruct
    '                End If
    '            End If
    '        Next

    '        paDict = New Dictionary(Of String, itmStruct)

    '        delPaList = New List(Of String)

    '        If Not regenFlag Then
    '            'Get current put away information
    '            For i = 0 To paDt.Rows.Count - 1

    '                itmKey = paDt.Rows(i).Item("GRA_ITM_CODE").ToString & "#_#" & _
    '                            paDt.Rows(i).Item("GRA_PACK_KEY").ToString & "#_#" & _
    '                            paDt.Rows(i).Item("GRA_PALLET_NO").ToString & "#_#" & _
    '                            paDt.Rows(i).Item("GRA_BATCH_NO").ToString

    '                If Not paDict.ContainsKey(itmKey) Then
    '                    tmpStruct = New itmStruct
    '                    tmpStruct.itmName = gU.decodeNullOrEmpty(paDt.Rows(i).Item("DSP_ITM_NAME").ToString, "")
    '                    tmpStruct.itmSeq = gU.decodeNullOrEmpty(paDt.Rows(i).Item("GRA_SPLIT_FR").ToString, "")
    '                    tmpStruct.qty = gU.decodeEmptyCdbl(paDt.Rows(i).Item("GRA_SUG_QTY"), 0)
    '                    tmpStruct.batchNo = gU.decodeNullOrEmpty(paDt.Rows(i).Item("GRA_BATCH_NO").ToString, "")
    '                    tmpStruct.skuNo = gU.decodeNullOrEmpty(paDt.Rows(i).Item("ITM_SKU_NO").ToString, "")

    '                    tmpStruct.expDate = paDt.Rows(i).Item("GRA_EXPIRY_DATE").ToString.Trim
    '                    tmpStruct.manuDate = paDt.Rows(i).Item("GRA_MANU_DATE").ToString.Trim

    '                    tmpStruct.whCode = paDt.Rows(i).Item("GRA_WH").ToString.Trim

    '                    tmpStruct.reqInsp = ""

    '                    paDict.Add(itmKey, tmpStruct)
    '                Else
    '                    tmpStruct = paDict.Item(itmKey)
    '                    tmpStruct.qty = tmpStruct.qty + gU.decodeEmptyCdbl(paDt.Rows(i).Item("GRA_SUG_QTY"), 0)

    '                    paDict(itmKey) = tmpStruct
    '                End If
    '            Next

    '            keys = paDict.Keys

    '            For i = 0 To keys.Count - 1
    '                If itmDict.ContainsKey(keys(i)) Then

    '                    tmpStruct = itmDict.Item(keys(i))

    '                    If tmpStruct.qty > paDict.Item(keys(i)).qty Then
    '                        tmpStruct.qty = tmpStruct.qty - paDict.Item(keys(i)).qty
    '                    Else
    '                        'Assume recv qty will not descrease
    '                        itmDict.Remove(keys(i))
    '                    End If
    '                Else
    '                    delPaList.Add(keys(i))
    '                End If
    '            Next
    '        End If


    '        keys = itmDict.Keys

    '        If keys.Count > 0 Then

    '            'Insert put away records
    '            For i = 0 To keys.Count - 1

    '                keyArray = Split(keys(i), "#_#")

    '                If Not regenFlag AndAlso paDict.ContainsKey(keys(i)) Then
    '                    'If item already exists in put away, increase the put away qty
    '                    For j = paDt.Rows.Count - 1 To 0 Step -1
    '                        itmKey = paDt.Rows(j).Item("GRA_ITM_CODE").ToString & "#_#" & _
    '                                    paDt.Rows(j).Item("GRA_PACK_KEY").ToString & "#_#" & _
    '                                    paDt.Rows(j).Item("GRA_PALLET_NO").ToString & "#_#" & _
    '                                    paDt.Rows(j).Item("GRA_BATCH_NO").ToString

    '                        If keys(i) = itmKey Then
    '                            paDt.Rows(j).Item("GRA_SUG_QTY") = paDt.Rows(j).Item("GRA_SUG_QTY") + itmDict.Item(keys(i)).qty

    '                            Exit For
    '                        End If
    '                    Next
    '                Else
    '                    'If it is a new item or re-gen put away, find the preferred location and add row
    '                    'paLocList = getItemLocQty(keyArray(0), keyArray(1), itmDict.Item(keys(i)).qty, itmDict.Item(keys(i)).whCode, itmDict.Item(keys(i)).reqInsp)
    '                    paLocList = getItemLocQty(keyArray(0), keyArray(1), itmDict.Item(keys(i)).qty, itmDict.Item(keys(i)).MainWhCode, itmDict.Item(keys(i)).reqInsp)

    '                    If paLocList.Count > 0 Then
    '                        For j = 0 To paLocList.Count - 1
    '                            paDt.Rows.Add()

    '                            plSeq = plSeq + 1

    '                            rows_count = paDt.Rows.Count

    '                            paDt.Rows(rows_count - 1).Item("GRA_SEQ") = CStr(plSeq)
    '                            paDt.Rows(rows_count - 1).Item("GRA_DISP_SEQ") = CStr(plSeq)
    '                            paDt.Rows(rows_count - 1).Item("GRA_PA_LIST_NO") = "1"
    '                            paDt.Rows(rows_count - 1).Item("GRA_ITM_CODE") = keyArray(0)
    '                            paDt.Rows(rows_count - 1).Item("DSP_ITM_NAME") = itmDict.Item(keys(i)).itmName
    '                            paDt.Rows(rows_count - 1).Item("GRA_PACK_KEY") = keyArray(1)
    '                            paDt.Rows(rows_count - 1).Item("GRA_PALLET_NO") = keyArray(2)
    '                            paDt.Rows(rows_count - 1).Item("GRA_REF_NO") = ""
    '                            paDt.Rows(rows_count - 1).Item("ITM_SKU_NO") = itmDict.Item(keys(i)).skuNo

    '                            paDt.Rows(rows_count - 1).Item("GRA_PA_QTY") = 0
    '                            paDt.Rows(rows_count - 1).Item("GRA_SUG_QTY") = paLocList(j).qty
    '                            paDt.Rows(rows_count - 1).Item("GRA_SPLIT_FR") = itmDict.Item(keys(i)).itmSeq
    '                            'paDt.Rows(rows_count - 1).Item("GRA_WH") = itmDict.Item(keys(i)).whCode
    '                            paDt.Rows(rows_count - 1).Item("GRA_WH") = paLocList(j).wh
    '                            paDt.Rows(rows_count - 1).Item("GRA_LOC") = paLocList(j).loc
    '                            paDt.Rows(rows_count - 1).Item("mFlag") = "N"

    '                            paDt.Rows(rows_count - 1).Item("GRA_BATCH_NO") = keyArray(3)

    '                            paDt.Rows(rows_count - 1).Item("GRA_EXPIRY_DATE") = itmDict.Item(keys(i)).expDate
    '                            paDt.Rows(rows_count - 1).Item("GRA_MANU_DATE") = itmDict.Item(keys(i)).manuDate

    '                            paDt.Rows(rows_count - 1).Item("GRA_REJ_QTY") = 0
    '                            paDt.Rows(rows_count - 1).Item("GRA_REJ_REASON") = ""
    '                        Next

    '                    Else
    '                        paDt.Rows.Add()

    '                        plSeq = plSeq + 1

    '                        rows_count = paDt.Rows.Count

    '                        paDt.Rows(rows_count - 1).Item("GRA_SEQ") = CStr(plSeq)
    '                        paDt.Rows(rows_count - 1).Item("GRA_DISP_SEQ") = CStr(plSeq)
    '                        paDt.Rows(rows_count - 1).Item("GRA_PA_LIST_NO") = "1"
    '                        paDt.Rows(rows_count - 1).Item("GRA_ITM_CODE") = keyArray(0)
    '                        paDt.Rows(rows_count - 1).Item("DSP_ITM_NAME") = itmDict.Item(keys(i)).itmName
    '                        paDt.Rows(rows_count - 1).Item("GRA_PACK_KEY") = keyArray(1)
    '                        paDt.Rows(rows_count - 1).Item("GRA_PALLET_NO") = keyArray(2)
    '                        paDt.Rows(rows_count - 1).Item("GRA_REF_NO") = ""
    '                        paDt.Rows(rows_count - 1).Item("ITM_SKU_NO") = itmDict.Item(keys(i)).skuNo

    '                        paDt.Rows(rows_count - 1).Item("GRA_PA_QTY") = 0
    '                        paDt.Rows(rows_count - 1).Item("GRA_SUG_QTY") = itmDict.Item(keys(i)).qty
    '                        paDt.Rows(rows_count - 1).Item("GRA_SPLIT_FR") = itmDict.Item(keys(i)).itmSeq
    '                        'paDt.Rows(rows_count - 1).Item("GRA_WH") = itmDict.Item(keys(i)).whCode
    '                        paDt.Rows(rows_count - 1).Item("GRA_WH") = ""
    '                        paDt.Rows(rows_count - 1).Item("GRA_LOC") = ""
    '                        paDt.Rows(rows_count - 1).Item("mFlag") = "N"

    '                        paDt.Rows(rows_count - 1).Item("GRA_BATCH_NO") = keyArray(3)

    '                        paDt.Rows(rows_count - 1).Item("GRA_EXPIRY_DATE") = itmDict.Item(keys(i)).expDate
    '                        paDt.Rows(rows_count - 1).Item("GRA_MANU_DATE") = itmDict.Item(keys(i)).manuDate

    '                        paDt.Rows(rows_count - 1).Item("GRA_REJ_QTY") = 0
    '                        paDt.Rows(rows_count - 1).Item("GRA_REJ_REASON") = ""
    '                    End If

    '                End If

    '            Next

    '            paSeq = CStr(plSeq)

    '            paDt.AcceptChanges()
    '        End If
    '    End If

    'End Sub



    Public Property putAwayDelList() As String
        Get
            Return paDelList
        End Get
        Set(value As String)
            paDelList = value
        End Set
    End Property

    Public Property putAwaySeq() As String
        Get
            Return paSeq
        End Get
        Set(value As String)
            paSeq = value
        End Set
    End Property

    Public Property putAwayDataTable() As DataTable
        Get
            Return paDt
        End Get
        Set(value As DataTable)
            paDt = value
        End Set
    End Property

    Public Property goodRecvDataTable() As DataTable
        Get
            Return grDt
        End Get
        Set(value As DataTable)
            grDt = value
        End Set
    End Property

End Class
