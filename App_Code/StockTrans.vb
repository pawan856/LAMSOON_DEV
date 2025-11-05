Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web

Public Class StockTrans

    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils

    Private lSTORER_CODE As String = ""
    Private lITM_CODE As String = ""
    Private lPACK_KEY As String = ""
    Private lIO_DOC As String = ""
    Private lIO_DOC_ID As String = ""
    Private lIO_QTY As Double = 0
    Private lIO_CUST_CODE As String = ""
    Private lIO_LOC As String = ""
    Private lIO_CBM As Double = 0
    Private lIO_KG As Double = 0
    Private lIMP_CODE As String = ""
    Private lIO_WH As String = ""
    Private lIO_FL As String = ""
    Private lIO_AREA As String = ""
    Private lIO_RANK As String = ""
    Private lIO_BIN As String = ""
    Private lIO_SEQ As String = ""
    Private lPALLET_NO As String = ""
    Private lIO_BATCH_NO As String = ""
    Private lIO_VND_CODE As String = ""
    Private lIOS_SERIAL_NO As String = ""
    Private lIOS_DRUM_ID As String = ""
    Private lIOS_DRUM_LEVEL As String = ""
    Private lIOS_UOM2 As String = ""
    Private lIOS_QTY2 As String = ""
    Private lIOS_SL As String = ""
    Private lIOS_ORG_QTY2 As String = ""
    Private lIOS_ORG_SERIAL_NO As String = ""
    Private lILOC_SEQ As Double = Nothing

    Private lIOS_REDRUM_YN As String = "Y"

    Private lIO_SYS_SEQ As String = ""

    Private lIO_EXPIRY_DATE As String = ""
    Private lIO_MANU_DATE As String = ""
    Private lIO_HOLD_QTY As Double = 0


    Public Enum IO_TYPE
        STOCKIN = 0
        STOCKOUT = 1
    End Enum

    Public Property IOS_REDRUM_YN() As String
        Get
            Return lIOS_REDRUM_YN
        End Get
        Set(ByVal Value As String)
            lIOS_REDRUM_YN = Value
        End Set
    End Property

    Public Property IO_SYS_SEQ() As String
        Get
            Return lIO_SYS_SEQ
        End Get
        Set(ByVal Value As String)
            lIO_SYS_SEQ = Value
        End Set
    End Property


    Public Property IOS_ORG_SERIAL_NO() As String
        Get
            Return lIOS_ORG_SERIAL_NO
        End Get
        Set(ByVal Value As String)
            lIOS_ORG_SERIAL_NO = Value
        End Set
    End Property

    Public Property IOS_ORG_QTY2() As String
        Get
            Return lIOS_ORG_QTY2
        End Get
        Set(ByVal Value As String)
            lIOS_ORG_QTY2 = Value
        End Set
    End Property

    Public Property IOS_SL() As String
        Get
            Return lIOS_SL
        End Get
        Set(ByVal Value As String)
            lIOS_SL = Value
        End Set
    End Property

    Public Property IOS_QTY2() As String
        Get
            Return lIOS_QTY2
        End Get
        Set(ByVal Value As String)
            lIOS_QTY2 = Value
        End Set
    End Property

    Public Property IOS_UOM2() As String
        Get
            Return lIOS_UOM2
        End Get
        Set(ByVal Value As String)
            lIOS_UOM2 = Value
        End Set
    End Property

    Public Property IOS_DRUM_LEVEL() As String
        Get
            Return lIOS_DRUM_LEVEL
        End Get
        Set(ByVal Value As String)
            lIOS_DRUM_LEVEL = Value
        End Set
    End Property

    Public Property IOS_DRUM_ID() As String
        Get
            Return lIOS_DRUM_ID
        End Get
        Set(ByVal Value As String)
            lIOS_DRUM_ID = Value
        End Set
    End Property

    Public Property IOS_SERIAL_NO() As String
        Get
            Return lIOS_SERIAL_NO
        End Get
        Set(ByVal Value As String)
            lIOS_SERIAL_NO = Value
        End Set
    End Property

    Public Property lO_BATCH_NO() As String
        Get
            Return lIO_BATCH_NO
        End Get
        Set(ByVal Value As String)
            lIO_BATCH_NO = Value
        End Set
    End Property

    Public Property lO_VND_CODE() As String
        Get
            Return lIO_VND_CODE
        End Get
        Set(ByVal Value As String)
            lIO_VND_CODE = Value
        End Set
    End Property

    Public Property IO_SEQ() As String
        Get
            Return lIO_SEQ
        End Get
        Set(ByVal Value As String)
            lIO_SEQ = Value
        End Set
    End Property

    Public Property IO_CBM() As Double
        Get
            Return lIO_CBM
        End Get
        Set(ByVal Value As Double)
            lIO_CBM = Value
        End Set
    End Property

    Public Property IO_KG() As Double
        Get
            Return lIO_KG
        End Get
        Set(ByVal Value As Double)
            lIO_KG = Value
        End Set
    End Property

    Public Property IO_AREA() As String
        Get
            Return lIO_AREA
        End Get
        Set(ByVal Value As String)
            lIO_AREA = Value
        End Set
    End Property

    Public Property IO_WH() As String
        Get
            Return lIO_WH
        End Get
        Set(ByVal Value As String)
            lIO_WH = Value
        End Set
    End Property

    Public Property IO_LOC() As String
        Get
            Return lIO_LOC
        End Get
        Set(ByVal Value As String)
            lIO_LOC = Value
        End Set
    End Property

    Public Property IO_CUST_CODE() As String
        Get
            Return lIO_CUST_CODE
        End Get
        Set(ByVal Value As String)
            lIO_CUST_CODE = Value
        End Set
    End Property

    Public Property IO_QTY() As Double
        Get
            Return lIO_QTY
        End Get
        Set(ByVal Value As Double)
            lIO_QTY = Value
        End Set
    End Property

    Public Property IO_DOC_ID() As String
        Get
            Return lIO_DOC_ID
        End Get
        Set(ByVal Value As String)
            lIO_DOC_ID = Value
        End Set
    End Property

    Public Property IO_DOC() As String
        Get
            Return lIO_DOC
        End Get
        Set(ByVal Value As String)
            lIO_DOC = Value
        End Set
    End Property

    Public Property STORER_CODE() As String
        Get
            Return lSTORER_CODE
        End Get
        Set(ByVal Value As String)
            lSTORER_CODE = Value
        End Set
    End Property

    Public Property ITM_CODE() As String
        Get
            Return lITM_CODE
        End Get
        Set(ByVal Value As String)
            lITM_CODE = Value
        End Set
    End Property

    Public Property PACK_KEY() As String
        Get
            Return lPACK_KEY
        End Get
        Set(ByVal Value As String)
            lPACK_KEY = Value
        End Set
    End Property

    Public Property PALLET_NO() As String
        Get
            Return lPALLET_NO
        End Get
        Set(ByVal Value As String)
            lPALLET_NO = Value
        End Set
    End Property

    Public Property IO_EXPIRY_DATE() As String
        Get
            Return lIO_EXPIRY_DATE
        End Get
        Set(ByVal Value As String)
            lIO_EXPIRY_DATE = Value
        End Set
    End Property

    Public Property IO_MANU_DATE() As String
        Get
            Return lIO_MANU_DATE
        End Get
        Set(ByVal Value As String)
            lIO_MANU_DATE = Value
        End Set
    End Property

    Public Property IO_HOLD_QTY() As String
        Get
            Return lIO_HOLD_QTY
        End Get
        Set(ByVal Value As String)
            lIO_HOLD_QTY = Value
        End Set
    End Property

    Public Function inCableTrans(ByRef serialNo As String, ByRef drumID As String, ByRef drumLevel As String, ByRef uom2 As String, ByRef qty2 As String, ByVal isFullDrum As Boolean, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As Boolean
        Dim newSerialNo As String = ""

        newSerialNo = getNewSerialNo(serialNo, pConn, pTransaction)

        lIOS_SERIAL_NO = newSerialNo
        If isFullDrum = True Then
            lIOS_DRUM_ID = newSerialNo
            lIOS_DRUM_LEVEL = "1"
            lIOS_UOM2 = uom2
            lIOS_QTY2 = qty2
        Else
            lIOS_DRUM_ID = drumID
            lIOS_DRUM_LEVEL = drumLevel
            lIOS_UOM2 = uom2
            lIOS_QTY2 = qty2
        End If

        Call setOrgSerialInfo(serialNo, pConn, pTransaction)

        Call UpdateStockSerialTrans("IN", pConn, pTransaction)

        Call UpdateStockBalSerialTrans("IN", pConn, pTransaction)

        Return True
    End Function

    Public Function outCableTrans(ByRef serialNo As String, ByRef drumID As String, ByRef drumLevel As String, ByRef uom2 As String, ByRef qty2 As String, ByVal isFullDrum As Boolean, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As Boolean
        Dim newSerialNo As String = ""

        lIOS_SERIAL_NO = serialNo
        If isFullDrum = True Then
            'lIOS_DRUM_ID = serialNo
            'lIOS_DRUM_LEVEL = "1"
            'lIOS_UOM2 = uom2
            lIOS_QTY2 = qty2
        Else
            'lIOS_DRUM_ID = drumID
            'lIOS_DRUM_LEVEL = drumLevel
            'lIOS_UOM2 = uom2
            lIOS_QTY2 = qty2
        End If

        Call setOrgSerialInfo(serialNo, pConn, pTransaction)

        Call UpdateStockSerialTrans("OUT", pConn, pTransaction)

        Call UpdateStockBalSerialTrans("OUT", pConn, pTransaction)

        Return True
    End Function

    Public Function reOrderDrumLevel(Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As Boolean
        Dim sqlString As String
        Dim dt As New DataTable

        sqlString = "select ISNULL(ILBS_QTY2,0) as ILBS_QTY2, ILBS_DRUM_LEVEL, ILBS_DRUM_ID from WMS_ITEM_LOC_BAL_S WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' "

        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            If dt.Rows(0).Item("ILBS_QTY2") <= 0 AndAlso dt.Rows(0).Item("ILBS_DRUM_LEVEL").ToString.Trim <> "" AndAlso dt.Rows(0).Item("ILBS_DRUM_LEVEL") > 0 Then
                sqlString = "UPDATE WMS_ITEM_LOC_BAL_S SET ILBS_DRUM_LEVEL = ILBS_DRUM_LEVEL - 1 "
                sqlString = sqlString & ", SYS_LUB = '" & Session("usr_id") & "' " &
                                        ", SYS_LUD = Getdate() " &
                            "WHERE ILBS_DRUM_ID = '" & gU.dbEncode(dt.Rows(0).Item("ILBS_DRUM_ID")) & "' AND convert(int, ILBS_DRUM_LEVEL) > " & gU.dbEncode(dt.Rows(0).Item("ILBS_DRUM_LEVEL")) & " "
                If pConn Is Nothing Then
                    Call gDB.amendData(sqlString)
                Else
                    If pTransaction Is Nothing Then
                        Call gDB.amendData(sqlString, pConn)
                    Else
                        Call gDB.amendData(sqlString, pConn, pTransaction)
                    End If
                End If

                sqlString = "UPDATE WMS_ITEM_LOC_BAL_S SET ILBS_DRUM_ID = NULL, ILBS_DRUM_LEVEL = NULL"
                sqlString = sqlString & ", SYS_LUB = '" & Session("usr_id") & "' " &
                                        ", SYS_LUD = Getdate() " &
                            "WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' "
                If pConn Is Nothing Then
                    Call gDB.amendData(sqlString)
                Else
                    If pTransaction Is Nothing Then
                        Call gDB.amendData(sqlString, pConn)
                    Else
                        Call gDB.amendData(sqlString, pConn, pTransaction)
                    End If
                End If

                dt = Nothing
            End If
            Return True
        Else
            Return False
        End If

    End Function

    Public Function getSLInfo(ByRef serialNo As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As String
        Dim sqlString As String = ""
        Dim dt As New DataTable
        Dim tempArr() As String
        Dim tempSerialNo As String

        tempArr = serialNo.Split("-")

        If tempArr.Length > 1 Then
            tempSerialNo = tempArr(0)
        Else
            tempSerialNo = serialNo
        End If

        sqlString = "select ILBS_SL from WMS_ITEM_LOC_BAL_S WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
        "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(tempSerialNo) & "' AND ILBS_QTY2 > 0 "

        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            Return dt.Rows(0).Item("ILBS_SL").ToString
        Else
            Return ""
        End If

    End Function

    Public Function setOrgSerialInfo(ByRef serialNo As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As Boolean
        Dim sqlString As String = ""
        Dim dt As New DataTable
        Dim tempArr() As String
        Dim tempSerialNo As String

        tempArr = serialNo.Split("-")

        If tempArr.Length > 1 Then
            tempSerialNo = tempArr(0)
        Else
            tempSerialNo = serialNo
        End If

        sqlString = "select ILBS_ORG_SERIAL_NO, ILBS_ORG_QTY2 from WMS_ITEM_LOC_BAL_S WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
        "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(tempSerialNo) & "'"

        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            lIOS_ORG_QTY2 = dt.Rows(0).Item("ILBS_ORG_QTY2").ToString
            lIOS_ORG_SERIAL_NO = dt.Rows(0).Item("ILBS_ORG_SERIAL_NO").ToString
            Return True
        Else
            sqlString = "select ILBS_ORG_SERIAL_NO, ILBS_ORG_QTY2 from WMS_ITEM_LOC_BAL_S WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
             "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(serialNo) & "'"

            If pConn Is Nothing Then
                dt = gDB.getDataTable(sqlString)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(sqlString, pConn)
                Else
                    dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                End If
            End If

            If dt.Rows.Count > 0 Then
                lIOS_ORG_QTY2 = dt.Rows(0).Item("ILBS_ORG_QTY2").ToString
                lIOS_ORG_SERIAL_NO = dt.Rows(0).Item("ILBS_ORG_SERIAL_NO").ToString
                Return True
            Else
                lIOS_ORG_QTY2 = lIOS_QTY2
                lIOS_ORG_SERIAL_NO = lIOS_SERIAL_NO
                Return False
            End If
        End If

    End Function

    Public Function getNewSerialNo(ByRef serialNo As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As String
        Dim sqlString As String = ""
        Dim dt As New DataTable
        Dim tempArr() As String
        Dim tempSerialNo As String
        Dim iscable As Boolean = False

        tempArr = serialNo.Split("-")

        If tempArr.Length > 1 Then
            tempSerialNo = tempArr(0)
        Else
            tempSerialNo = serialNo
        End If

        sqlString = "select ITM_CODE from WMS_ITEM WHERE ITM_TYPE = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                    "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            iscable = True
        End If

        If iscable Then
            sqlString = "select max(isnull(convert(int, REPLACE(REPLACE(ILBS_SERIAL_NO, '" & gU.dbEncode(tempSerialNo) & "',''),'-','')),0)) + 1 from WMS_ITEM_LOC_BAL_S WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILBS_SERIAL_NO LIKE '" & gU.dbEncode(tempSerialNo) & "%'"

            If pConn Is Nothing Then
                dt = gDB.getDataTable(sqlString)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(sqlString, pConn)
                Else
                    dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                End If
            End If

            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item(0).ToString = "" Then
                    Return tempSerialNo
                Else
                    Return tempSerialNo & "-" & dt.Rows(0).Item(0).ToString
                End If
            Else
                Return tempSerialNo & "-" & 1
            End If
        Else
            Return serialNo
        End If
    End Function

    Public Function getILOC_SEQ(Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As Boolean
        Dim sqlString As String = ""
        Dim dt As New DataTable
        sqlString = "select ILOC_SEQ from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
        "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILOC_LOC = '" & gU.dbEncode(lIO_LOC) & "'  " &
        "AND ISNULL(ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If
        If dt.Rows.Count > 0 Then
            lILOC_SEQ = dt.Rows(0).Item(0)
        Else
            lILOC_SEQ = Nothing
        End If
        Return True
    End Function

    Public Function UpdateStockSerialTrans(ByVal in_out_type As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing, Optional ByRef errorMsg As String = "") As Boolean
        Dim sqlString As String = ""
        Dim dt As New DataTable
        Dim txTable As String = ""
        Dim calType As String = ""
        Dim imTable As String = "WMS_ITEM"
        Dim SrchStr As String = ""
        Dim iscable As Boolean = False
        Dim cableSL As Boolean = False
        Dim isSerial As Boolean = False

        sqlString = "select ITM_CODE from " & imTable & " WHERE ITM_TYPE = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            iscable = True
        End If

        sqlString = "select ITM_CODE from " & imTable & " WHERE ITM_SERIAL_NO_YN = 'Y' AND isnull(ITM_TYPE,'') <> 'CABLE' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            isSerial = True
        End If

        REM Update TX Record
        If in_out_type = "IN" Then
            txTable = "WMS_IN_S_TX"
            calType = "+"
        ElseIf in_out_type = "OUT" Then
            txTable = "WMS_OUT_S_TX"
            calType = "-"
        ElseIf in_out_type = "UNPOSTOUT" Then
            in_out_type = "UNPOST"
            txTable = "WMS_OUT_S_TX"
            calType = "-"
        ElseIf in_out_type = "UNPOSTIN" Then
            in_out_type = "UNPOST"
            txTable = "WMS_IN_S_TX"
            calType = "+"
        End If

        Call getILOC_SEQ(pConn, pTransaction)

        If in_out_type = "OUT" Then
            SrchStr = "select * from WMS_ITEM_LOC_BAL_S WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' "
            If pConn Is Nothing Then
                dt = gDB.getDataTable(SrchStr)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(SrchStr, pConn)
                Else
                    dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
                End If
            End If

            If dt.Rows.Count = 0 Then
                If Session("gLang") = "E" Then
                    errorMsg = "Serial No. " & lIOS_SERIAL_NO & " don\'t exists!"
                Else
                    errorMsg = "Serial No. " & lIOS_SERIAL_NO & "不存在!"
                End If

                Throw New Exception(errorMsg)
                Return False
            End If
        End If

        REM Split LOC code
        'locCode = WH_CODE.Value & "" & Right("0" & FL_NUM.Value.ToString, 2) & "" & AR_CODE.Value & "" & RK_CODE.Value & "" & nDT.Rows(i).Item(i1).ToString

        SrchStr = "select * from V_LOCATION WHERE  LOC = '" & gU.dbEncode(lIO_LOC) & "' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND WH_CODE = '" & lIO_WH & "'"
        If pConn Is Nothing Then
            dt = gDB.getDataTable(SrchStr)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(SrchStr, pConn)
            Else
                dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            lIO_WH = dt.Rows(0).Item("WH_CODE").ToString
            lIO_FL = dt.Rows(0).Item("FL_NUM").ToString
            lIO_AREA = dt.Rows(0).Item("AR_CODE").ToString
            lIO_RANK = dt.Rows(0).Item("RK_CODE").ToString
            lIO_BIN = dt.Rows(0).Item("BN_CODE").ToString
        End If

        Dim holdQty As Double = 0
        holdQty = getHoldQty(pConn, pTransaction)

        sqlString = "select * from WMS_ITEM_LOC_BAL_S WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If iscable = True Then
            If lIOS_SL <> "" Then
                lIOS_SL = lIOS_SL
            Else
                lIOS_SL = "Y"
            End If
        Else
            If isSerial = True Then
                If calType = "+" Then
                    lIOS_QTY2 = 1
                ElseIf calType = "-" Then
                    lIOS_QTY2 = 1
                End If
            Else

            End If
        End If

        If dt.Rows.Count > 0 Then
            sqlString = "INSERT INTO " & txTable & " (" &
                        "IMP_CODE, " &
                        "STORER_CODE, " &
                        "ITM_CODE, " &
                        "PACK_KEY, " &
                        "IOSX_SERIAL_NO, " &
                        "IOSX_DRUM_ID, " &
                        "IOSX_DRUM_LEVEL, " &
                        "IOSX_UOM2, " &
                        "IOSX_QTY2, " &
                        "IOSX_BAL_BEFORE, " &
                        "IOSX_BAL_AFTER, " &
                        "IOSX_ORG_QTY2, " &
                        "IOSX_ORG_SERIAL_NO, " &
                        "SYS_LUB, " &
                        "SYS_LUD, " &
                        "SYS_CD, " &
                        "SYS_CB, " &
                        "IOSX_ILOC_SEQ, " &
                        "IO_SYS_SEQ, " &
                        "IOSX_DATETIME" &
                        ") (SELECT " &
                        "IMP_CODE," &
                        "STORER_CODE," &
                        "ITM_CODE," &
                        "PACK_KEY," &
                        "ILBS_SERIAL_NO, "
            If in_out_type = "OUT" Then
                sqlString = sqlString & "ILBS_DRUM_ID, " &
                        "ILBS_DRUM_LEVEL, " &
                        "ILBS_UOM2, "
            Else
                sqlString = sqlString & "'" & gU.dbEncode(lIOS_DRUM_ID) & "'," &
                            "'" & gU.dbEncode(lIOS_DRUM_LEVEL) & "'," &
                            "'" & gU.dbEncode(lIOS_UOM2) & "',"
            End If

            sqlString = sqlString & gU.decodeNullOrEmpty(gU.dbEncode(lIOS_QTY2), "NULL") & ", " &
                        "ISNULL(ILBS_QTY2,0)," &
                        "ISNULL(ILBS_QTY2,0) " & calType & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0), " &
                        "ILBS_ORG_QTY2, " &
                        "ILBS_ORG_SERIAL_NO, " &
                        "'" & Session("usr_id") & "'," &
                        "Getdate(), " &
                        "Getdate(), " &
                        "'" & Session("usr_id") & "'," &
                        "ILOC_SEQ, " &
                        "'" & gU.dbEncode(lIO_SYS_SEQ) & "'," &
                        "Getdate() " &
                        "FROM WMS_ITEM_LOC_BAL_S WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "')"

            If pConn Is Nothing Then
                Call gDB.amendData(sqlString)
            Else
                If pTransaction Is Nothing Then
                    Call gDB.amendData(sqlString, pConn)
                Else
                    Call gDB.amendData(sqlString, pConn, pTransaction)
                End If
            End If
        Else
            sqlString = "INSERT INTO " & txTable & " (" &
                        "IMP_CODE, " &
                        "STORER_CODE, " &
                        "ITM_CODE, " &
                        "PACK_KEY, " &
                        "IOSX_SERIAL_NO, " &
                        "IOSX_DRUM_ID, " &
                        "IOSX_DRUM_LEVEL, " &
                        "IOSX_UOM2, " &
                        "IOSX_QTY2, " &
                        "IOSX_BAL_BEFORE, " &
                        "IOSX_BAL_AFTER, " &
                        "IOSX_ORG_QTY2, " &
                        "IOSX_ORG_SERIAL_NO, " &
                        "SYS_LUB, " &
                        "SYS_LUD, " &
                        "SYS_CD, " &
                        "SYS_CB, " &
                        "IOSX_ILOC_SEQ, " &
                        "IO_SYS_SEQ, " &
                        "IOSX_DATETIME" &
                        ") VALUES (" &
                        "'" & gU.dbEncode(lIMP_CODE) & "'," &
                        "'" & gU.dbEncode(lSTORER_CODE) & "'," &
                        "'" & gU.dbEncode(lITM_CODE) & "'," &
                        "'" & gU.dbEncode(lPACK_KEY) & "'," &
                        "'" & gU.dbEncode(lIOS_SERIAL_NO) & "'," &
                        "'" & gU.dbEncode(lIOS_DRUM_ID) & "'," &
                        "'" & gU.dbEncode(lIOS_DRUM_LEVEL) & "'," &
                        "'" & gU.dbEncode(lIOS_UOM2) & "'," &
                        gU.decodeNullOrEmpty(gU.dbEncode(lIOS_QTY2), "NULL") & ", " &
                        "0," &
                        "0 " & calType & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0), " &
                        gU.decodeNullOrEmpty(gU.dbEncode(lIOS_ORG_QTY2), "NULL") & ", " &
                        "'" & gU.dbEncode(lIOS_ORG_SERIAL_NO) & "'," &
                        "'" & Session("usr_id") & "'," &
                        "Getdate(), " &
                        "Getdate(), " &
                        "'" & Session("usr_id") & "'," &
                        "'" & gU.dbEncode(lILOC_SEQ) & "'," &
                        "'" & gU.dbEncode(lIO_SYS_SEQ) & "'," &
                        "Getdate() " &
                        ")"

            If pConn Is Nothing Then
                Call gDB.amendData(sqlString)
            Else
                If pTransaction Is Nothing Then
                    Call gDB.amendData(sqlString, pConn)
                Else
                    Call gDB.amendData(sqlString, pConn, pTransaction)
                End If
            End If
        End If

        Return True
    End Function


    Public Function UpdateStockTrans(ByVal in_out_type As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing, Optional ByRef errorMsg As String = "") As Boolean
        Dim sqlString As String = ""
        Dim dt As New DataTable
        Dim txTable As String = ""
        Dim calType As String = ""
        Dim imTable As String = "WMS_ITEM"
        Dim SrchStr As String = ""
        Dim iscable As Boolean = False
        Dim fullCableYN As String = ""
        Dim updateQty As Double = lIO_QTY

        sqlString = "select ITM_CODE from " & imTable & " WHERE ITM_TYPE = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            iscable = True
        End If

        REM Update TX Record
        If in_out_type = "IN" Then
            txTable = "WMS_IN_TX"
            calType = "+"
        ElseIf in_out_type = "OUT" Then
            txTable = "WMS_OUT_TX"
            calType = "-"
        ElseIf in_out_type = "UNPOSTOUT" Then
            in_out_type = "UNPOST"
            txTable = "WMS_OUT_TX"
            calType = "-"
        ElseIf in_out_type = "UNPOSTIN" Then
            in_out_type = "UNPOST"
            txTable = "WMS_IN_TX"
            calType = "+"
        End If

        If in_out_type = "OUT" Then
            SrchStr = "select * from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILOC_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ILOC_WH = '" & gU.dbEncode(lIO_WH) & "'  " &
            "AND ISNULL(ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "
            If pConn Is Nothing Then
                dt = gDB.getDataTable(SrchStr)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(SrchStr, pConn)
                Else
                    dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
                End If
            End If

            If dt.Rows.Count = 0 Then
                If Session("gLang") = "E" Then
                    errorMsg = "Item " & lITM_CODE & "-" & lPACK_KEY & " don\'t exists!"
                Else
                    errorMsg = "物料" & lITM_CODE & "-" & lPACK_KEY & "不存在!"
                End If

                Throw New Exception(errorMsg)
                Return False
            End If

            If iscable = True Then

                Call getILOC_SEQ(pConn, pTransaction)

                sqlString = "select ITM_CODE from WMS_ITEM_LOC_BAL_S WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' AND " &
                "((ISNULL(ILBS_QTY2,0) " & calType & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) > 0 AND ISNULL(ILBS_QTY2,0) = 0) " &
                "OR (ISNULL(ILBS_QTY2,0) " & calType & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) <= 0 AND ISNULL(ILBS_QTY2,0) > 0)) "

                If pConn Is Nothing Then
                    dt = gDB.getDataTable(sqlString)
                Else
                    If pTransaction Is Nothing Then
                        dt = gDB.getDataTable(sqlString, pConn)
                    Else
                        dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                    End If
                End If

                If dt.Rows.Count > 0 Then
                    fullCableYN = "Y"
                    updateQty = 1
                Else
                    updateQty = 0
                End If
            End If
        End If

        REM Split LOC code
        'locCode = WH_CODE.Value & "" & Right("0" & FL_NUM.Value.ToString, 2) & "" & AR_CODE.Value & "" & RK_CODE.Value & "" & nDT.Rows(i).Item(i1).ToString
        SrchStr = "select * from V_LOCATION WHERE LOC = '" & gU.dbEncode(lIO_LOC) & "' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND WH_CODE = '" & gU.dbEncode(lIO_WH) & "'"
        If pConn Is Nothing Then
            dt = gDB.getDataTable(SrchStr)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(SrchStr, pConn)
            Else
                dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            lIO_WH = dt.Rows(0).Item("WH_CODE").ToString
            lIO_FL = dt.Rows(0).Item("FL_NUM").ToString
            lIO_AREA = dt.Rows(0).Item("AR_CODE").ToString
            lIO_RANK = dt.Rows(0).Item("RK_CODE").ToString
            lIO_BIN = dt.Rows(0).Item("BN_CODE").ToString
        End If

        Dim holdQty As Double = 0
        holdQty = getHoldQty(pConn, pTransaction)

        sqlString = "select * from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
        "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILOC_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ILOC_WH = '" & gU.dbEncode(lIO_WH) & "'  " &
        "AND ISNULL(ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If
        If dt.Rows.Count > 0 Then
            sqlString = "INSERT INTO " & txTable & " (" &
            "IMP_CODE, " &
            "STORER_CODE, " &
            "ITM_CODE, " &
            "PACK_KEY, " &
            "IO_PALLET_NO," &
            "IO_CODE, " &
            "IO_TYPE, " &
            "IO_BATCH_NO, " &
            "VND_CODE, " &
            "IO_CUST_CODE, " &
            "IO_WH," &
            "IO_AREA," &
            "IO_LOC, " &
            "IO_DATETIME, " &
            "IO_DOC, " &
            "IO_DOC_ID, " &
            "IO_QTY, " &
            "IO_BAL_BEFORE, " &
            "IO_BAL_AFTER," &
            "IO_CBM," &
            "IO_KG," &
            "IO_EXPIRY_DATE, " &
            "IO_MANU_DATE, " &
            "IO_HOLD_QTY, " &
            "IO_FULL_CABLE_YN," &
            "IO_SYS_SEQ, " &
            "SYS_CB, " &
            "SYS_CD, " &
            "SYS_LUB, " &
            "SYS_LUD" &
            ") (SELECT " &
            "IMP_CODE," &
            "STORER_CODE," &
            "ITM_CODE," &
            "PACK_KEY," &
            "'" & gU.dbEncode(lPALLET_NO) & "'," &
            "REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(datetimeoffset, SYSDATETIMEOFFSET()), '.', ''), '-', ''), ':', ''), ' ', '') + '" & lIO_SEQ & "'," &
            "'" & gU.dbEncode(in_out_type) & "'," &
            "" & gU.convdbNVCData(gU.dbEncode(lIO_BATCH_NO)) & "," &
            "'" & gU.dbEncode(lIO_VND_CODE) & "'," &
            "'" & gU.dbEncode(lIO_CUST_CODE) & "'," &
            "'" & gU.dbEncode(lIO_WH) & "'," &
            "'" & gU.dbEncode(lIO_AREA) & "'," &
            "'" & gU.dbEncode(lIO_LOC) & "'," &
            "Getdate()," &
            "'" & gU.dbEncode(lIO_DOC) & "'," &
            "'" & gU.dbEncode(lIO_DOC_ID) & "'," &
            gU.decodeNullOrEmpty(gU.dbEncode(lIO_QTY), "NULL") & ", " &
            "ISNULL(ILOC_BAL_QTY,0)," &
            "ISNULL(ILOC_BAL_QTY,0) " & calType & " ISNULL(" & gU.dbEncode(updateQty) & ",0), " &
            "" & gU.dbEncode(lIO_CBM) & "," &
            "" & gU.dbEncode(lIO_KG) & "," &
            gU.convdbDate(lIO_EXPIRY_DATE) & "," &
            gU.convdbDate(lIO_MANU_DATE) & "," &
            "" & gU.dbEncode(holdQty) & "," &
            "'" & gU.dbEncode(fullCableYN) & "'," &
            "'" & gU.dbEncode(lIO_SYS_SEQ) & "'," &
            "'" & Session("usr_id") & "', " &
            "Getdate(), " &
            "'" & Session("usr_id") & "'," &
            "Getdate() " &
            "FROM WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ILOC_LOC = '" & gU.dbEncode(lIO_LOC) & "' " &
            "AND ISNULL(ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "')"

            If pConn Is Nothing Then
                Call gDB.amendData(sqlString)
            Else
                If pTransaction Is Nothing Then
                    Call gDB.amendData(sqlString, pConn)
                Else
                    Call gDB.amendData(sqlString, pConn, pTransaction)
                End If
            End If
        Else
            sqlString = "INSERT INTO " & txTable & " (" &
            "IMP_CODE, " &
            "STORER_CODE, " &
            "ITM_CODE, " &
            "PACK_KEY, " &
            "IO_PALLET_NO," &
            "IO_CODE, " &
            "IO_TYPE, " &
            "IO_BATCH_NO, " &
            "VND_CODE, " &
            "IO_CUST_CODE, " &
            "IO_WH," &
            "IO_AREA," &
            "IO_LOC, " &
            "IO_DATETIME, " &
            "IO_DOC, " &
            "IO_DOC_ID, " &
            "IO_QTY, " &
            "IO_BAL_BEFORE, " &
            "IO_BAL_AFTER," &
            "IO_CBM," &
            "IO_KG," &
            "IO_EXPIRY_DATE, " &
            "IO_MANU_DATE, " &
            "IO_HOLD_QTY, " &
            "IO_FULL_CABLE_YN," &
            "IO_SYS_SEQ, " &
            "SYS_CB, " &
            "SYS_CD, " &
            "SYS_LUB, " &
            "SYS_LUD" &
            ") VALUES (" &
            "'" & gU.dbEncode(lIMP_CODE) & "'," &
            "'" & gU.dbEncode(lSTORER_CODE) & "'," &
            "'" & gU.dbEncode(lITM_CODE) & "'," &
            "'" & gU.dbEncode(lPACK_KEY) & "'," &
            "'" & gU.dbEncode(lPALLET_NO) & "'," &
            "REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(datetimeoffset, SYSDATETIMEOFFSET()), '.', ''), '-', ''), ':', ''), ' ', '') + '" & lIO_SEQ & "'," &
            "'" & gU.dbEncode(in_out_type) & "'," &
            "" & gU.convdbNVCData(gU.dbEncode(lIO_BATCH_NO)) & "," &
            "'" & gU.dbEncode(lIO_VND_CODE) & "'," &
            "'" & gU.dbEncode(lIO_CUST_CODE) & "'," &
            "'" & gU.dbEncode(lIO_WH) & "'," &
            "'" & gU.dbEncode(lIO_AREA) & "'," &
            "'" & gU.dbEncode(lIO_LOC) & "'," &
            "Getdate()," &
            "'" & gU.dbEncode(lIO_DOC) & "'," &
            "'" & gU.dbEncode(lIO_DOC_ID) & "'," &
            gU.decodeNullOrEmpty(gU.dbEncode(lIO_QTY), "NULL") & ", " &
            "0," &
            "0 " & calType & " ISNULL(" & gU.dbEncode(updateQty) & ",0), " &
            "" & gU.dbEncode(lIO_CBM) & "," &
            "" & gU.dbEncode(lIO_KG) & "," &
            gU.convdbDate(lIO_EXPIRY_DATE) & "," &
            gU.convdbDate(lIO_MANU_DATE) & "," &
            "" & gU.dbEncode(holdQty) & "," &
            "'" & gU.dbEncode(fullCableYN) & "'," &
            "'" & gU.dbEncode(lIO_SYS_SEQ) & "'," &
            "'" & Session("usr_id") & "', " &
            "Getdate(), " &
            "'" & Session("usr_id") & "'," &
            "Getdate() " &
            ")"

            If pConn Is Nothing Then
                Call gDB.amendData(sqlString)
            Else
                If pTransaction Is Nothing Then
                    Call gDB.amendData(sqlString, pConn)
                Else
                    Call gDB.amendData(sqlString, pConn, pTransaction)
                End If
            End If
        End If

        Dim sqlString1 As String
        sqlString1 = sqlString

        REM Update Item Master Record
        'sqlString = "UPDATE " & imTable & " SET ITM_BALANCE = ISNULL(ITM_BALANCE,0) " & calType & " ISNULL(" & gU.dbEncode(lIO_QTY) & ",0) " & _
        '"WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
        '"AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "'"

        'If pConn Is Nothing Then
        '    Call gDB.amendData(sqlString)
        'Else
        '    If pTransaction Is Nothing Then
        '        Call gDB.amendData(sqlString, pConn)
        '    Else
        '        Call gDB.amendData(sqlString, pConn, pTransaction)
        '    End If
        'End If

        Return True
    End Function

    Public Function UpdateStockBalSerialTrans(ByVal in_out_type As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing, Optional ByRef errorMsg As String = "") As Boolean
        Dim sqlString As String = ""
        Dim dt As New DataTable
        Dim calType As String = ""
        Dim imTable As String = "WMS_ITEM"
        Dim SrchStr As String = ""
        Dim lupdateSL As String = ""
        Dim iscable As Boolean = False
        Dim cableSL As Boolean = False
        Dim isSerial As Boolean = False

        sqlString = "select ITM_CODE from " & imTable & " WHERE ITM_TYPE = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            iscable = True
        End If

        sqlString = "select ITM_CODE from " & imTable & " WHERE ITM_SERIAL_NO_YN = 'Y' AND isnull(ITM_TYPE,'') <> 'CABLE' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            isSerial = True
        End If

        REM Update TX Record
        If in_out_type = "IN" Then
            calType = "+"
        ElseIf in_out_type = "OUT" Then
            calType = "-"
        ElseIf in_out_type = "UNPOSTOUT" Then
            calType = "-"
        ElseIf in_out_type = "UNPOSTIN" Then
            calType = "+"
        End If

        Call getILOC_SEQ(pConn, pTransaction)

        If in_out_type = "OUT" Or in_out_type = "UNPOSTOUT" Then
            SrchStr = "select * from WMS_ITEM_LOC_BAL_S WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' "
            If pConn Is Nothing Then
                dt = gDB.getDataTable(SrchStr)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(SrchStr, pConn)
                Else
                    dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
                End If
            End If

            If dt.Rows.Count = 0 Then
                If Session("gLang") = "E" Then
                    errorMsg = "Serial No. " & lIOS_SERIAL_NO & " don\'t exists!"
                Else
                    errorMsg = "Serial No. " & lIOS_SERIAL_NO & "不存在!"
                End If

                Throw New Exception(errorMsg)
                Return False
            End If
        End If

        REM Split LOC code
        'locCode = WH_CODE.Value & "" & Right("0" & FL_NUM.Value.ToString, 2) & "" & AR_CODE.Value & "" & RK_CODE.Value & "" & nDT.Rows(i).Item(i1).ToString
        SrchStr = "select * from V_LOCATION WHERE LOC = '" & gU.dbEncode(lIO_LOC) & "' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(SrchStr)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(SrchStr, pConn)
            Else
                dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            lIO_WH = dt.Rows(0).Item("WH_CODE").ToString
            lIO_FL = dt.Rows(0).Item("FL_NUM").ToString
            lIO_AREA = dt.Rows(0).Item("AR_CODE").ToString
            lIO_RANK = dt.Rows(0).Item("RK_CODE").ToString
            lIO_BIN = dt.Rows(0).Item("BN_CODE").ToString
        End If

        REM Update Item Location
        If lIO_LOC <> "" Then
            sqlString = "select * from WMS_ITEM_LOC_BAL_S WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' "
            If pConn Is Nothing Then
                dt = gDB.getDataTable(sqlString)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(sqlString, pConn)
                Else
                    dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                End If
            End If
            If dt.Rows.Count > 0 Then
                REM UPDATE SQL
                If iscable = True Then
                    If lIOS_SL <> "" Then
                        lupdateSL = lIOS_SL
                    Else
                        lupdateSL = "Y"
                    End If
                    sqlString = "UPDATE WMS_ITEM_LOC_BAL_S SET ILBS_SL = '" & lupdateSL & "', ILBS_QTY2 = ISNULL(ILBS_QTY2,0) " & calType & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) "
                Else
                    If isSerial = True Then
                        If calType = "+" Then
                            sqlString = "UPDATE WMS_ITEM_LOC_BAL_S SET ILBS_QTY2 = 1 "
                            lIOS_QTY2 = 1
                        ElseIf calType = "-" Then
                            sqlString = "UPDATE WMS_ITEM_LOC_BAL_S SET ILBS_QTY2 = 0 "
                            lIOS_QTY2 = 1
                        End If
                    Else
                        If lIOS_QTY2 = 0 Then
                            lIOS_QTY2 = 1
                        End If
                        sqlString = "UPDATE WMS_ITEM_LOC_BAL_S SET ILBS_SL = '" & lupdateSL & "', ILBS_QTY2 = ISNULL(ILBS_QTY2,0) " & calType & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) "
                    End If
                End If

                cableSL = True
                If in_out_type = "OUT" Then

                Else
                    sqlString = sqlString & ", ILBS_DRUM_ID = '" & gU.dbEncode(lIOS_DRUM_ID) & "' " &
                                            ", ILBS_DRUM_LEVEL = '" & gU.dbEncode(lIOS_DRUM_LEVEL) & "' "
                End If
                sqlString = sqlString & ", SYS_LUB = '" & Session("usr_id") & "' " &
            ", SYS_LUD = Getdate() " &
            "WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' "
            Else
                If iscable = True Then
                    If lIOS_SL <> "" Then
                        lIOS_SL = lIOS_SL
                    Else
                        lIOS_SL = "Y"
                    End If
                Else
                    If isSerial = True Then
                        If calType = "+" Then
                            lIOS_QTY2 = 1
                        ElseIf calType = "-" Then
                            lIOS_QTY2 = 1
                        End If
                    Else

                    End If
                End If

                REM INSERT SQL
                sqlString = "INSERT INTO WMS_ITEM_LOC_BAL_S (" &
                            "IMP_CODE, " &
                            "STORER_CODE, " &
                            "ITM_CODE, " &
                            "PACK_KEY, " &
                            "ILBS_SERIAL_NO, " &
                            "ILBS_DRUM_ID, " &
                            "ILBS_DRUM_LEVEL, " &
                            "ILBS_UOM2, " &
                            "ILBS_QTY2, " &
                            "ILBS_SL, " &
                            "ILBS_ORG_QTY2, " &
                            "ILBS_ORG_SERIAL_NO, " &
                            "SYS_LUB, " &
                            "SYS_LUD, " &
                            "SYS_CD, " &
                            "SYS_CB, " &
                            "ILOC_SEQ" &
                            ") VALUES (" &
                            "'" & gU.dbEncode(lIMP_CODE) & "'," &
                            "'" & gU.dbEncode(lSTORER_CODE) & "'," &
                            "'" & gU.dbEncode(lITM_CODE) & "'," &
                            "'" & gU.dbEncode(lPACK_KEY) & "'," &
                            "'" & gU.dbEncode(lIOS_SERIAL_NO) & "'," &
                            "'" & gU.dbEncode(lIOS_DRUM_ID) & "'," &
                            "'" & gU.dbEncode(lIOS_DRUM_LEVEL) & "'," &
                            "'" & gU.dbEncode(lIOS_UOM2) & "'," &
                            gU.decodeNullOrEmpty(gU.dbEncode(lIOS_QTY2), "NULL") & ", " &
                            "'" & gU.dbEncode(lIOS_SL) & "'," &
                            gU.decodeNullOrEmpty(gU.dbEncode(lIOS_ORG_QTY2), "NULL") & ", " &
                            "'" & gU.dbEncode(lIOS_ORG_SERIAL_NO) & "'," &
                            "'" & Session("usr_id") & "', " &
                            "Getdate(), " &
                            "Getdate(), " &
                            "'" & Session("usr_id") & "'," &
                            "'" & gU.dbEncode(lILOC_SEQ) & "'" &
                            ")"
            End If

            If pConn Is Nothing Then
                Call gDB.amendData(sqlString)
            Else
                If pTransaction Is Nothing Then
                    Call gDB.amendData(sqlString, pConn)
                Else
                    Call gDB.amendData(sqlString, pConn, pTransaction)
                End If
            End If

            dt = Nothing
        End If

        If iscable = True Then
            REM ReOrder drum level
            If lIOS_REDRUM_YN = "Y" Then
                Call reOrderDrumLevel(pConn, pTransaction)
            End If
        End If

        REM Update Item daily Transcation
        If lIO_LOC <> "" Then
            sqlString = "select * from WMS_ITEM_LOC_BAL_S_TX WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ITSX_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' AND convert(varchar, ITX_DATE, 1) = convert(varchar, Getdate(), 1) "

            If pConn Is Nothing Then
                dt = gDB.getDataTable(sqlString)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(sqlString, pConn)
                Else
                    dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                End If
            End If

            Dim holdQty As Double = 0
            holdQty = getHoldQty(pConn, pTransaction)

            If dt.Rows.Count > 0 Then
                REM UPDATE SQL
                sqlString = "UPDATE WMS_ITEM_LOC_BAL_S_TX SET ITSX_BAL_BEFORE2 = ITSX_BAL_AFTER2, ITSX_BAL_AFTER2 = ISNULL(ITSX_BAL_AFTER2,0) " & calType & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) "
                If in_out_type = "OUT" Then

                Else
                    sqlString = sqlString & ", ITSX_DRUM_ID = '" & gU.dbEncode(lIOS_DRUM_ID) & "' " &
                    ", ITSX_DRUM_LEVEL = '" & gU.dbEncode(lIOS_DRUM_LEVEL) & "' "
                End If
                sqlString = sqlString & ", SYS_LUB = '" & Session("usr_id") & "' " &
                ", SYS_LUD = Getdate() " &
                "WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ITSX_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' AND convert(varchar, ITX_DATE, 1) = convert(varchar, Getdate(), 1) "
            Else
                dt = Nothing
                sqlString = "select WMS_ITEM_LOC_BAL_S_TX.*, REPLACE(REPLACE(REPLACE(convert(varchar, ITX_DATE, 20),'-',''),':',''),' ','') as F_ITX_DATE from WMS_ITEM_LOC_BAL_S_TX " &
                    "WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ITSX_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' ORDER BY ITX_DATE DESC"
                If pConn Is Nothing Then
                    dt = gDB.getDataTable(sqlString)
                Else
                    If pTransaction Is Nothing Then
                        dt = gDB.getDataTable(sqlString, pConn)
                    Else
                        dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                    End If
                End If
                If dt.Rows.Count > 0 Then
                    REM INSERT SQL
                    sqlString = "INSERT INTO WMS_ITEM_LOC_BAL_S_TX (" &
                                "IMP_CODE, " &
                                "STORER_CODE, " &
                                "ITM_CODE, " &
                                "PACK_KEY, " &
                                "ITSX_SERIAL_NO, " &
                                "ITSX_DRUM_ID, " &
                                "ITSX_DRUM_LEVEL, " &
                                "ITSX_UOM2, " &
                                "ITSX_BAL_BEFORE2, " &
                                "ITSX_BAL_AFTER2, " &
                                "ITSX_ORG_QTY2, " &
                                "ITSX_ORG_SERIAL_NO, " &
                                "SYS_LUB, " &
                                "SYS_LUD, " &
                                "SYS_CD, " &
                                "SYS_CB, " &
                                "ILOC_SEQ, " &
                                "ITX_DATE," &
                                "ITSX_SL" &
                                ") (SELECT " &
                                "IMP_CODE, " &
                                "STORER_CODE, " &
                                "ITM_CODE, " &
                                "PACK_KEY, " &
                                "ITSX_SERIAL_NO, "
                    If in_out_type = "OUT" Then
                        sqlString = sqlString & "ITSX_DRUM_ID, " &
                                "ITSX_DRUM_LEVEL, "
                    Else
                        sqlString = sqlString & "'" & gU.dbEncode(lIOS_DRUM_ID) & "'," &
                                    "'" & gU.dbEncode(lIOS_DRUM_LEVEL) & "',"
                    End If
                    sqlString = sqlString & "ITSX_UOM2, " &
                                "ISNULL(ITSX_BAL_BEFORE2,0), " &
                                "ISNULL(ITSX_BAL_AFTER2,0) " & calType & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0), " &
                                "ITSX_ORG_QTY2, " &
                                "ITSX_ORG_SERIAL_NO, " &
                                "'" & Session("usr_id") & "', " &
                                "Getdate(), " &
                                "Getdate(), " &
                                "'" & Session("usr_id") & "'," &
                                "ILOC_SEQ, " &
                                "Getdate(),ITSX_SL FROM WMS_ITEM_LOC_BAL_S_TX WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ITSX_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' AND REPLACE(REPLACE(REPLACE(convert(varchar, ITX_DATE, 20),'-',''),':',''),' ','') = '" & dt.Rows(0).Item("F_ITX_DATE").ToString & "') "
                Else
                    REM INSERT SQL
                    sqlString = "INSERT INTO WMS_ITEM_LOC_BAL_S_TX (" &
                                "IMP_CODE, " &
                                "STORER_CODE, " &
                                "ITM_CODE, " &
                                "PACK_KEY, " &
                                "ITSX_SERIAL_NO, " &
                                "ITSX_DRUM_ID, " &
                                "ITSX_DRUM_LEVEL, " &
                                "ITSX_UOM2, " &
                                "ITSX_BAL_BEFORE2, " &
                                "ITSX_BAL_AFTER2, " &
                                "ITSX_ORG_QTY2, " &
                                "ITSX_ORG_SERIAL_NO, " &
                                "SYS_LUB, " &
                                "SYS_LUD, " &
                                "SYS_CD, " &
                                "SYS_CB, " &
                                "ILOC_SEQ, " &
                                "ITX_DATE," &
                                "ITSX_SL" &
                                ") VALUES (" &
                                "'" & gU.dbEncode(lIMP_CODE) & "'," &
                                "'" & gU.dbEncode(lSTORER_CODE) & "'," &
                                "'" & gU.dbEncode(lITM_CODE) & "'," &
                                "'" & gU.dbEncode(lPACK_KEY) & "'," &
                                "'" & gU.dbEncode(lIOS_SERIAL_NO) & "'," &
                                "'" & gU.dbEncode(lIOS_DRUM_ID) & "'," &
                                "'" & gU.dbEncode(lIOS_DRUM_LEVEL) & "'," &
                                "'" & gU.dbEncode(lIOS_UOM2) & "'," &
                                "0," &
                                gU.decodeNullOrEmpty(gU.dbEncode(lIOS_QTY2), "0") & ", " &
                                gU.decodeNullOrEmpty(gU.dbEncode(lIOS_ORG_QTY2), "NULL") & ", " &
                                "'" & gU.dbEncode(lIOS_ORG_SERIAL_NO) & "'," &
                                "'" & Session("usr_id") & "', " &
                                "Getdate(), " &
                                "Getdate(), " &
                                "'" & Session("usr_id") & "'," &
                                "'" & gU.dbEncode(lILOC_SEQ) & "'," &
                                "Getdate(), "
                    If cableSL = True Then
                        sqlString = sqlString & "'Y')"
                    Else
                        sqlString = sqlString & "null)"
                    End If
                End If
            End If
            If pConn Is Nothing Then
                Call gDB.amendData(sqlString)
            Else
                If pTransaction Is Nothing Then
                    Call gDB.amendData(sqlString, pConn)
                Else
                    Call gDB.amendData(sqlString, pConn, pTransaction)
                End If
            End If
            dt = Nothing
        End If

        If isSerial = True Then
            sqlString = "update WMS_ITEM_LOC_BAL set ILOC_BAL_QTY = (SELECT count(*) from WMS_ITEM_LOC_BAL_S " &
                        "WHERE WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ and ILBS_QTY2 > 0) " &
                        "WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                        "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' " &
                        "AND EXISTS (SELECT 1 from WMS_ITEM_LOC_BAL_S " &
                        "WHERE WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ and ILBS_QTY2 > 0)"
            If pConn Is Nothing Then
                Call gDB.amendData(sqlString)
            Else
                If pTransaction Is Nothing Then
                    Call gDB.amendData(sqlString, pConn)
                Else
                    Call gDB.amendData(sqlString, pConn, pTransaction)
                End If
            End If
            dt = Nothing
        End If

        UpdateStockBalSerialTrans = True
    End Function

    Public Function UpdateStockBalTrans(ByVal in_out_type As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing, Optional ByRef errorMsg As String = "") As Boolean
        Dim sqlString As String = ""
        Dim dt As New DataTable
        Dim calType As String = ""
        Dim imTable As String = "WMS_ITEM"
        Dim SrchStr As String = ""
        Dim iscable As Boolean = False
        Dim fullCableYN As String = ""
        Dim updateQty As Double = lIO_QTY
        Dim updateCBM As Double = lIO_CBM
        Dim updateKG As Double = lIO_KG

        sqlString = "select ITM_CODE from " & imTable & " WHERE ITM_TYPE = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            iscable = True
        End If

        REM Update TX Record
        If in_out_type = "IN" Then
            calType = "+"
        ElseIf in_out_type = "OUT" Then
            calType = "-"
        End If

        If iscable = True Then

            Call getILOC_SEQ(pConn, pTransaction)

            sqlString = "select ITM_CODE from WMS_ITEM_LOC_BAL_S WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' AND " &
            "((ISNULL(ILBS_QTY2,0) " & calType & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) > 0 AND ISNULL(ILBS_QTY2,0) = 0) " &
            "OR (ISNULL(ILBS_QTY2,0) " & calType & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) <= 0 AND ISNULL(ILBS_QTY2,0) > 0)) "

            If pConn Is Nothing Then
                dt = gDB.getDataTable(sqlString)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(sqlString, pConn)
                Else
                    dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                End If
            End If

            If dt.Rows.Count > 0 Then
                fullCableYN = "Y"
                updateQty = 1
            Else
                sqlString = "select ITM_CODE from WMS_ITEM_LOC_BAL_S WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' "

                If pConn Is Nothing Then
                    dt = gDB.getDataTable(sqlString)
                Else
                    If pTransaction Is Nothing Then
                        dt = gDB.getDataTable(sqlString, pConn)
                    Else
                        dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                    End If
                End If

                If dt.Rows.Count > 0 Then
                    updateQty = 0
                    updateCBM = 0
                    updateKG = 0
                Else
                    updateQty = 1
                End If
            End If
        End If

        If in_out_type = "OUT" Then
            SrchStr = "select * from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILOC_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ILOC_WH = '" & gU.dbEncode(lIO_WH) & "'  " &
            "AND ISNULL(ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "
            If pConn Is Nothing Then
                dt = gDB.getDataTable(SrchStr)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(SrchStr, pConn)
                Else
                    dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
                End If
            End If

            If dt.Rows.Count = 0 Then
                If Session("gLang") = "E" Then
                    errorMsg = "Item " & lITM_CODE & "-" & lPACK_KEY & " don\'t exists!"
                Else
                    errorMsg = "物料" & lITM_CODE & "-" & lPACK_KEY & "不存在!"
                End If

                Throw New Exception(errorMsg)
                Return False
            End If
        End If

        REM Split LOC code
        'locCode = WH_CODE.Value & "" & Right("0" & FL_NUM.Value.ToString, 2) & "" & AR_CODE.Value & "" & RK_CODE.Value & "" & nDT.Rows(i).Item(i1).ToString
        SrchStr = "select * from V_LOCATION WHERE LOC = '" & gU.dbEncode(lIO_LOC) & "' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND WH_CODE='" & gU.dbEncode(lIO_WH) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(SrchStr)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(SrchStr, pConn)
            Else
                dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            lIO_WH = dt.Rows(0).Item("WH_CODE").ToString
            lIO_FL = dt.Rows(0).Item("FL_NUM").ToString
            lIO_AREA = dt.Rows(0).Item("AR_CODE").ToString
            lIO_RANK = dt.Rows(0).Item("RK_CODE").ToString
            lIO_BIN = dt.Rows(0).Item("BN_CODE").ToString
        End If

        REM Update Item Location
        If lIO_LOC <> "" Then
            sqlString = "select * from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILOC_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ILOC_WH = '" & gU.dbEncode(lIO_WH) & "' " &
            "AND ISNULL(ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "
            If pConn Is Nothing Then
                dt = gDB.getDataTable(sqlString)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(sqlString, pConn)
                Else
                    dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                End If
            End If
            If dt.Rows.Count > 0 Then
                REM UPDATE SQL
                If lIO_MANU_DATE IsNot Nothing AndAlso lIO_MANU_DATE <> "" AndAlso lIO_MANU_DATE <> "01/01/1900" Then
                    sqlString = "UPDATE WMS_ITEM_LOC_BAL SET ILOC_BAL_QTY = ISNULL(ILOC_BAL_QTY,0) " & calType & " ISNULL(" & gU.dbEncode(updateQty) & ",0) " &
                ", ILOC_BAL_CBM = case when ISNULL(ILOC_BAL_CBM,0) " & calType & "" & gU.dbEncode(updateCBM) & " < 0 then 0 else ISNULL(ILOC_BAL_CBM,0) " & calType & "" & gU.dbEncode(updateCBM) & " end " &
                ", ILOC_BAL_KG = case when ISNULL(ILOC_BAL_KG,0) " & calType & "" & gU.dbEncode(updateKG) & " < 0 then 0 else ISNULL(ILOC_BAL_KG,0) " & calType & "" & gU.dbEncode(updateKG) & " end " &
                ", ILOC_EXPIRY_DATE = " & gU.convdbDate(lIO_MANU_DATE) &
                ", ILOC_MANU_DATE = " & gU.convdbDate(lIO_MANU_DATE) &
                ", VND_CODE = '" & gU.dbEncode(lIO_VND_CODE) & "' " &
                ", SYS_LUB = '" & Session("usr_id") & "' " &
                ", SYS_LUD = Getdate() " &
                "WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILOC_LOC = '" & gU.dbEncode(Right(lIO_LOC, 8)) & "' AND ILOC_WH = '" & gU.dbEncode(lIO_WH) & "' " &
                "AND ISNULL(ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "
                Else
                    sqlString = "UPDATE WMS_ITEM_LOC_BAL SET ILOC_BAL_QTY = ISNULL(ILOC_BAL_QTY,0) " & calType & " ISNULL(" & gU.dbEncode(updateQty) & ",0) " &
                ", ILOC_BAL_CBM = case when ISNULL(ILOC_BAL_CBM,0) " & calType & "" & gU.dbEncode(updateCBM) & " < 0 then 0 else ISNULL(ILOC_BAL_CBM,0) " & calType & "" & gU.dbEncode(updateCBM) & " end " &
                ", ILOC_BAL_KG = case when ISNULL(ILOC_BAL_KG,0) " & calType & "" & gU.dbEncode(updateKG) & " < 0 then 0 else ISNULL(ILOC_BAL_KG,0) " & calType & "" & gU.dbEncode(updateKG) & " end " &
                ", ILOC_EXPIRY_DATE = " & gU.convdbDate(lIO_EXPIRY_DATE) &
                ", VND_CODE = '" & gU.dbEncode(lIO_VND_CODE) & "' " &
                ", SYS_LUB = '" & Session("usr_id") & "' " &
                ", SYS_LUD = Getdate() " &
                "WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILOC_LOC = '" & gU.dbEncode(Right(lIO_LOC, 8)) & "' AND ILOC_WH = '" & gU.dbEncode(lIO_WH) & "' " &
                "AND ISNULL(ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "
                End If

            Else
                If lIO_MANU_DATE IsNot Nothing AndAlso lIO_MANU_DATE <> "" Then
                    lIO_EXPIRY_DATE = lIO_MANU_DATE
                End If
                REM INSERT SQL
                sqlString = "INSERT INTO WMS_ITEM_LOC_BAL (" &
            "IMP_CODE, " &
            "STORER_CODE, " &
            "ITM_CODE, " &
            "PACK_KEY, " &
            "ILOC_LOC," &
            "ILOC_BATCH_NO," &
            "VND_CODE," &
            "ILOC_PALLET_NO," &
            "ILOC_WH," &
            "ILOC_FLOOR," &
            "ILOC_AREA," &
            "ILOC_RACK," &
            "ILOC_BIN," &
            "ILOC_BAL_QTY," &
            "ILOC_BAL_CBM," &
            "ILOC_BAL_KG," &
            "ILOC_EXPIRY_DATE," &
            "ILOC_MANU_DATE," &
            "SYS_CB, " &
            "SYS_CD, " &
            "SYS_LUB, " &
            "SYS_LUD" &
            ") VALUES (" &
            "'" & gU.dbEncode(lIMP_CODE) & "'," &
            "'" & gU.dbEncode(lSTORER_CODE) & "'," &
            "'" & gU.dbEncode(lITM_CODE) & "'," &
            "'" & gU.dbEncode(lPACK_KEY) & "'," &
            "'" & gU.dbEncode(Right(lIO_LOC, 8)) & "'," &
            "" & gU.convdbNVCData(gU.dbEncode(lIO_BATCH_NO)) & "," &
            "'" & gU.dbEncode(lIO_VND_CODE) & "'," &
            "'" & gU.dbEncode(lPALLET_NO) & "'," &
            "'" & gU.dbEncode(lIO_WH) & "'," &
            "'" & gU.decodeNullOrEmpty(gU.dbEncode(lIO_FL), "NULL") & "', " &
            "'" & gU.dbEncode(lIO_AREA) & "'," &
            "'" & gU.dbEncode(lIO_RANK) & "'," &
            "'" & gU.dbEncode(lIO_BIN) & "'," &
            gU.decodeNullOrEmpty(gU.dbEncode(lIO_QTY), "NULL") & ", " &
            "" & gU.dbEncode(lIO_CBM) & "," &
            "" & gU.dbEncode(lIO_KG) & "," &
            gU.convdbDate(lIO_EXPIRY_DATE) & "," &
            gU.convdbDate(lIO_MANU_DATE) & "," &
            "'" & Session("usr_id") & "', " &
            "Getdate(), " &
            "'" & Session("usr_id") & "'," &
            "Getdate() " &
            ")"
            End If

            If pConn Is Nothing Then
                Call gDB.amendData(sqlString)
            Else
                If pTransaction Is Nothing Then
                    Call gDB.amendData(sqlString, pConn)
                Else
                    Call gDB.amendData(sqlString, pConn, pTransaction)
                End If
            End If

            dt = Nothing
        End If

        REM Update Item daily Transcation
        If lIO_LOC <> "" Then
            sqlString = "select * from WMS_ITEM_LOC_BAL_TX WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ITX_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ITX_WH = '" & gU.dbEncode(lIO_WH) & "' AND ITX_PALLET_NO = '" & gU.dbEncode(lPALLET_NO) & "' AND convert(varchar, ITX_DATE, 1) = convert(varchar, Getdate(), 1) " &
            "AND ISNULL(ITX_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "

            If pConn Is Nothing Then
                dt = gDB.getDataTable(sqlString)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(sqlString, pConn)
                Else
                    dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                End If
            End If

            Dim holdQty As Double = 0
            holdQty = getHoldQty(pConn, pTransaction)

            If dt.Rows.Count > 0 Then
                REM UPDATE SQL
                sqlString = "UPDATE WMS_ITEM_LOC_BAL_TX SET ITX_BAL_BEFORE = ITX_BAL_AFTER, ITX_BAL_AFTER = ISNULL(ITX_BAL_AFTER,0) " & calType & " ISNULL(" & gU.dbEncode(updateQty) & ",0) " &
                ", ITX_BAL_CBM = case when ISNULL(ITX_BAL_CBM,0) " & calType & "" & gU.dbEncode(updateCBM) & " < 0 then 0 else ISNULL(ITX_BAL_CBM,0) " & calType & "" & gU.dbEncode(updateCBM) & " end " &
                ", ITX_BAL_KG = case when ISNULL(ITX_BAL_KG,0) " & calType & "" & gU.dbEncode(updateKG) & " < 0 then 0 else ISNULL(ITX_BAL_KG,0) " & calType & "" & gU.dbEncode(updateKG) & " end " &
                ", ITX_HOLD_QTY = " & gU.dbEncode(holdQty) & " " &
                ", ITX_EXPIRY_DATE = " & gU.convdbDate(lIO_EXPIRY_DATE) &
                ", VND_CODE = '" & gU.dbEncode(lIO_VND_CODE) & "' " &
                ", SYS_LUB = '" & Session("usr_id") & "' " &
                ", SYS_LUD = Getdate() " &
                "WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ITX_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ITX_WH = '" & gU.dbEncode(lIO_WH) & "' AND ITX_PALLET_NO = '" & gU.dbEncode(lPALLET_NO) & "' AND convert(varchar, ITX_DATE, 1) = convert(varchar, Getdate(), 1) " &
                "AND ISNULL(ITX_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "
            Else
                dt = Nothing
                sqlString = "select WMS_ITEM_LOC_BAL_TX.*, REPLACE(REPLACE(REPLACE(convert(varchar, ITX_DATE, 20),'-',''),':',''),' ','') as F_ITX_DATE from WMS_ITEM_LOC_BAL_TX WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ITX_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ITX_WH = '" & gU.dbEncode(lIO_WH) & "' AND ITX_PALLET_NO = '" & gU.dbEncode(lPALLET_NO) & "' " &
                 "AND ISNULL(ITX_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' ORDER BY ITX_DATE DESC"
                If pConn Is Nothing Then
                    dt = gDB.getDataTable(sqlString)
                Else
                    If pTransaction Is Nothing Then
                        dt = gDB.getDataTable(sqlString, pConn)
                    Else
                        dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                    End If
                End If
                If dt.Rows.Count > 0 Then
                    REM INSERT SQL
                    sqlString = "INSERT INTO WMS_ITEM_LOC_BAL_TX (" &
                    "IMP_CODE, " &
                    "STORER_CODE, " &
                    "ITM_CODE, " &
                    "PACK_KEY, " &
                    "ITX_PALLET_NO," &
                    "ITX_DATE," &
                    "ITX_LOC," &
                    "ITX_BATCH_NO," &
                    "VND_CODE," &
                    "ITX_WH," &
                    "ITX_FLOOR," &
                    "ITX_AREA," &
                    "ITX_RACK," &
                    "ITX_BIN," &
                    "ITX_BAL_BEFORE," &
                    "ITX_BAL_AFTER," &
                    "ITX_BAL_CBM," &
                    "ITX_BAL_KG," &
                    "ITX_HOLD_QTY, " &
                    "ITX_EXPIRY_DATE, " &
                    "ITX_MANU_DATE, " &
                    "SYS_CB, " &
                    "SYS_CD, " &
                    "SYS_LUB, " &
                    "SYS_LUD" &
                    ") (SELECT " &
                    "'" & gU.dbEncode(lIMP_CODE) & "'," &
                    "'" & gU.dbEncode(lSTORER_CODE) & "'," &
                    "'" & gU.dbEncode(lITM_CODE) & "'," &
                    "'" & gU.dbEncode(lPACK_KEY) & "'," &
                    "'" & gU.dbEncode(lPALLET_NO) & "'," &
                    "Getdate(), " &
                    "'" & gU.dbEncode(lIO_LOC) & "'," &
                    "" & gU.convdbNVCData(gU.dbEncode(lIO_BATCH_NO)) & "," &
                    "'" & gU.dbEncode(lIO_VND_CODE) & "'," &
                    "'" & gU.dbEncode(lIO_WH) & "'," &
                    "'" & gU.decodeNullOrEmpty(gU.dbEncode(lIO_FL), "NULL") & "', " &
                    "'" & gU.dbEncode(lIO_AREA) & "'," &
                    "'" & gU.dbEncode(lIO_RANK) & "'," &
                    "'" & gU.dbEncode(lIO_BIN) & "'," &
                    "ISNULL(ITX_BAL_BEFORE,0), " &
                    "ISNULL(ITX_BAL_AFTER,0) " & calType & " ISNULL(" & gU.dbEncode(updateQty) & ",0), " &
                    "ISNULL(ITX_BAL_CBM,0) " & calType & " ISNULL(" & gU.dbEncode(updateCBM) & ",0), " &
                    "ISNULL(ITX_BAL_KG,0) " & calType & " ISNULL(" & gU.dbEncode(updateKG) & ",0), " &
                    "ISNULL(" & gU.dbEncode(holdQty) & ",0)," &
                    gU.convdbDate(lIO_EXPIRY_DATE) & "," &
                    gU.convdbDate(lIO_MANU_DATE) & "," &
                    "'" & Session("usr_id") & "', " &
                    "Getdate(), " &
                    "'" & Session("usr_id") & "'," &
                    "Getdate() FROM WMS_ITEM_LOC_BAL_TX WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                    "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ITX_LOC = '" & gU.dbEncode(lIO_LOC) & "' " &
                    "AND ITX_PALLET_NO = '" & gU.dbEncode(lPALLET_NO) & "' AND REPLACE(REPLACE(REPLACE(convert(varchar, ITX_DATE, 20),'-',''),':',''),' ','') = '" & dt.Rows(0).Item("F_ITX_DATE").ToString & "' " &
                    "AND ISNULL(ITX_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "')"
                Else
                    REM INSERT SQL
                    sqlString = "INSERT INTO WMS_ITEM_LOC_BAL_TX (" &
                    "IMP_CODE, " &
                    "STORER_CODE, " &
                    "ITM_CODE, " &
                    "PACK_KEY, " &
                    "ITX_PALLET_NO," &
                    "ITX_DATE," &
                    "ITX_LOC," &
                    "ITX_BATCH_NO," &
                    "VND_CODE," &
                    "ITX_WH," &
                    "ITX_FLOOR," &
                    "ITX_AREA," &
                    "ITX_RACK," &
                    "ITX_BIN," &
                    "ITX_BAL_BEFORE," &
                    "ITX_BAL_AFTER," &
                    "ITX_BAL_CBM," &
                    "ITX_BAL_KG," &
                    "ITX_HOLD_QTY," &
                    "ITX_EXPIRY_DATE, " &
                    "ITX_MANU_DATE, " &
                    "SYS_CB, " &
                    "SYS_CD, " &
                    "SYS_LUB, " &
                    "SYS_LUD" &
                    ") VALUES (" &
                    "'" & gU.dbEncode(lIMP_CODE) & "'," &
                    "'" & gU.dbEncode(lSTORER_CODE) & "'," &
                    "'" & gU.dbEncode(lITM_CODE) & "'," &
                    "'" & gU.dbEncode(lPACK_KEY) & "'," &
                    "'" & gU.dbEncode(lPALLET_NO) & "'," &
                    "Getdate(), " &
                    "'" & gU.dbEncode(lIO_LOC) & "'," &
                    "" & gU.convdbNVCData(gU.dbEncode(lIO_BATCH_NO)) & "," &
                    "'" & gU.dbEncode(lIO_VND_CODE) & "'," &
                    "'" & gU.dbEncode(lIO_WH) & "'," &
                    "'" & gU.decodeNullOrEmpty(gU.dbEncode(lIO_FL), "NULL") & "', " &
                    "'" & gU.dbEncode(lIO_AREA) & "'," &
                    "'" & gU.dbEncode(lIO_RANK) & "'," &
                    "'" & gU.dbEncode(lIO_BIN) & "'," &
                    "0," &
                    gU.decodeNullOrEmpty(gU.dbEncode(lIO_QTY), "0") & ", " &
                    gU.decodeNullOrEmpty(gU.dbEncode(lIO_CBM), "0") & ", " &
                    gU.decodeNullOrEmpty(gU.dbEncode(lIO_KG), "0") & ", " &
                    "ISNULL(" & gU.dbEncode(holdQty) & ",0)," &
                    gU.convdbDate(lIO_EXPIRY_DATE) & "," &
                    gU.convdbDate(lIO_MANU_DATE) & "," &
                    "'" & Session("usr_id") & "', " &
                    "Getdate(), " &
                    "'" & Session("usr_id") & "'," &
                    "Getdate() " &
                    ")"
                End If
            End If
            If pConn Is Nothing Then
                Call gDB.amendData(sqlString)
            Else
                If pTransaction Is Nothing Then
                    Call gDB.amendData(sqlString, pConn)
                Else
                    Call gDB.amendData(sqlString, pConn, pTransaction)
                End If
            End If
            dt = Nothing
        End If

        UpdateStockBalTrans = True
    End Function

    Public Function UnPostStocks(ByVal type As IO_TYPE, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing, Optional ByRef errorMsg As String = "") As Boolean
        Dim SrchStr As String = ""
        Dim txSrchStr As String = ""
        Dim sqlString As String = ""
        Dim txsqlString As String = ""
        Dim dt As New DataTable
        Dim imTable As String = "WMS_ITEM"
        Dim oper As String = ""
        Dim unpostType As String = ""
        Dim iscable As Boolean = False
        Dim fullCableYN As String = ""
        Dim updateQty As Double = lIO_QTY

        sqlString = "select ITM_CODE from " & imTable & " WHERE ITM_TYPE = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            iscable = True
        End If

        SrchStr = "select * from V_LOCATION WHERE LOC = '" & gU.dbEncode(lIO_LOC) & "' AND WH_CODE = '" & gU.dbEncode(lIO_WH) & "' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(SrchStr)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(SrchStr, pConn)
            Else
                dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            lIO_WH = dt.Rows(0).Item("WH_CODE").ToString
            lIO_FL = dt.Rows(0).Item("FL_NUM").ToString
            lIO_AREA = dt.Rows(0).Item("AR_CODE").ToString
            lIO_RANK = dt.Rows(0).Item("RK_CODE").ToString
            lIO_BIN = dt.Rows(0).Item("BN_CODE").ToString
        End If

        If type = IO_TYPE.STOCKOUT Then
            oper = "-"
            unpostType = "UNPOSTOUT"
        ElseIf type = IO_TYPE.STOCKIN Then
            oper = "+"
            unpostType = "UNPOSTIN"
        End If

        If iscable = True Then

            Call getILOC_SEQ(pConn, pTransaction)

            sqlString = "select ITM_CODE from WMS_ITEM_LOC_BAL_S WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ILBS_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' AND " &
            "((ISNULL(ILBS_QTY2,0) " & oper & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) > 0 AND ISNULL(ILBS_QTY2,0) = 0) " &
            "OR (ISNULL(ILBS_QTY2,0) " & oper & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) <= 0 AND ISNULL(ILBS_QTY2,0) > 0)) "

            If pConn Is Nothing Then
                dt = gDB.getDataTable(sqlString)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(sqlString, pConn)
                Else
                    dt = gDB.getDataTable(sqlString, pConn, pTransaction)
                End If
            End If

            If dt.Rows.Count > 0 Then
                fullCableYN = "Y"
                updateQty = 1
            Else
                updateQty = 0
            End If
        End If

        REM Update Item Location
        If lIO_LOC <> "" Then
            SrchStr = "select * from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILOC_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ILOC_WH = '" & gU.dbEncode(lIO_WH) & "'  " &
            "AND ISNULL(ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "
            If pConn Is Nothing Then
                dt = gDB.getDataTable(SrchStr)
            Else
                If pTransaction Is Nothing Then
                    dt = gDB.getDataTable(SrchStr, pConn)
                Else
                    dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
                End If
            End If

            If dt.Rows.Count > 0 Then
                Dim stQTY As Double = gU.decodeEmptyCInt(dt.Rows(0).Item("ILOC_BAL_QTY").ToString, 0)

                If type = IO_TYPE.STOCKOUT Then
                    If stQTY < lIO_QTY Then
                        If Session("gLang") = "E" Then
                            errorMsg = "Item " & lITM_CODE & "-" & lPACK_KEY & " don\'t have enough stock balance to Un-Post!"
                        Else
                            errorMsg = "物料" & lITM_CODE & "-" & lPACK_KEY & "沒有足夠貨存去取消發布!"
                        End If

                        Return False
                    End If
                End If

                UpdateStockTrans(unpostType, pConn, pTransaction)

                sqlString = "UPDATE WMS_ITEM_LOC_BAL SET ILOC_BAL_QTY = ISNULL(ILOC_BAL_QTY,0)  " & oper & " ISNULL(" & gU.decodeEmptyCInt(gU.dbEncode(updateQty), 0) & ",0) " &
                                ", ILOC_BAL_CBM = case when ISNULL(ILOC_BAL_CBM,0) " & oper & " " & gU.dbEncode(lIO_CBM) & " < 0 then 0 else ISNULL(ILOC_BAL_CBM,0) " & oper & " " & gU.dbEncode(lIO_CBM) & " end " &
                                ", ILOC_BAL_KG = case when ISNULL(ILOC_BAL_KG,0) " & oper & " " & gU.dbEncode(lIO_KG) & " < 0 then 0 else ISNULL(ILOC_BAL_KG,0) " & oper & " " & gU.dbEncode(lIO_KG) & " end " &
                                ", VND_CODE = '" & gU.dbEncode(lIO_VND_CODE) & "' " &
                                ", SYS_LUB = '" & Session("usr_id") & "' " &
                                ", SYS_LUD = Getdate() " &
                                "WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                                "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ILOC_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ILOC_WH = '" & gU.dbEncode(lIO_WH) & "'  " &
                                "AND ISNULL(ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "

                txSrchStr = "select * from WMS_ITEM_LOC_BAL_TX WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                            "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ITX_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ITX_PALLET_NO = '" & gU.dbEncode(lPALLET_NO) & "' AND convert(varchar, ITX_DATE, 1) = convert(varchar, Getdate(), 1) " &
                            "AND ISNULL(ITX_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "

                If pConn Is Nothing Then
                    dt = gDB.getDataTable(txSrchStr)
                Else
                    If pTransaction Is Nothing Then
                        dt = gDB.getDataTable(txSrchStr, pConn)
                    Else
                        dt = gDB.getDataTable(txSrchStr, pConn, pTransaction)
                    End If
                End If

                Dim holdQty As Double = 0
                holdQty = getHoldQty(pConn, pTransaction)

                If dt.Rows.Count > 0 Then
                    txsqlString = "UPDATE WMS_ITEM_LOC_BAL_TX SET ITX_BAL_BEFORE = ITX_BAL_AFTER, ITX_BAL_AFTER = ISNULL(ITX_BAL_AFTER,0)  " & oper & " ISNULL(" & gU.dbEncode(updateQty) & ",0) " &
                    ", ITX_BAL_CBM = case when ISNULL(ITX_BAL_CBM,0) " & oper & " " & gU.dbEncode(lIO_CBM) & " < 0 then 0 else ISNULL(ITX_BAL_CBM,0) " & oper & " " & gU.dbEncode(lIO_CBM) & " end " &
                    ", ITX_BAL_KG = case when ISNULL(ITX_BAL_KG,0) " & oper & " " & gU.dbEncode(lIO_KG) & " < 0 then 0 else ISNULL(ITX_BAL_KG,0) " & oper & " " & gU.dbEncode(lIO_KG) & " end " &
                    ", ITX_HOLD_QTY = " & gU.dbEncode(holdQty) & " " &
                    ", VND_CODE = '" & gU.dbEncode(lIO_VND_CODE) & "' " &
                    ", SYS_LUB = '" & Session("usr_id") & "' " &
                    ", SYS_LUD = Getdate() " &
                    "WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                    "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ITX_LOC = '" & gU.dbEncode(lIO_LOC) & "'  AND ITX_WH = '" & gU.dbEncode(lIO_WH) & "' AND ITX_PALLET_NO = '" & gU.dbEncode(lPALLET_NO) & "' AND convert(varchar, ITX_DATE, 1) = convert(varchar, Getdate(), 1) " &
                    "AND ISNULL(ITX_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' "
                Else
                    dt = Nothing

                    txSrchStr = "select WMS_ITEM_LOC_BAL_TX.*, REPLACE(REPLACE(REPLACE(convert(varchar, ITX_DATE, 20),'-',''),':',''),' ','') as F_ITX_DATE " &
                    "from WMS_ITEM_LOC_BAL_TX WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                    "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ITX_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND ITX_WH = '" & gU.dbEncode(lIO_WH) & "' AND ITX_PALLET_NO = '" & gU.dbEncode(lPALLET_NO) & "' " &
                    "AND ISNULL(ITX_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' ORDER BY ITX_DATE DESC"

                    If pConn Is Nothing Then
                        dt = gDB.getDataTable(txSrchStr)
                    Else
                        If pTransaction Is Nothing Then
                            dt = gDB.getDataTable(txSrchStr, pConn)
                        Else
                            dt = gDB.getDataTable(txSrchStr, pConn, pTransaction)
                        End If
                    End If

                    If dt.Rows.Count > 0 Then
                        txsqlString = "INSERT INTO WMS_ITEM_LOC_BAL_TX (" &
                                        "IMP_CODE, " &
                                        "STORER_CODE, " &
                                        "ITM_CODE, " &
                                        "PACK_KEY, " &
                                        "ITX_PALLET_NO," &
                                        "ITX_DATE," &
                                        "ITX_LOC," &
                                        "ITX_BATCH_NO," &
                                        "VND_CODE," &
                                        "ITX_WH," &
                                        "ITX_FLOOR," &
                                        "ITX_AREA," &
                                        "ITX_RACK," &
                                        "ITX_BIN," &
                                        "ITX_BAL_BEFORE," &
                                        "ITX_BAL_AFTER," &
                                        "ITX_BAL_CBM," &
                                        "ITX_BAL_KG," &
                                        "ITX_HOLD_QTY," &
                                        "ITX_EXPIRY_DATE, " &
                                        "ITX_MANU_DATE, " &
                                        "SYS_CB, " &
                                        "SYS_CD, " &
                                        "SYS_LUB, " &
                                        "SYS_LUD" &
                                        ") (SELECT " &
                                        "'" & gU.dbEncode(lIMP_CODE) & "'," &
                                        "'" & gU.dbEncode(lSTORER_CODE) & "'," &
                                        "'" & gU.dbEncode(lITM_CODE) & "'," &
                                        "'" & gU.dbEncode(lPACK_KEY) & "'," &
                                        "'" & gU.dbEncode(lPALLET_NO) & "'," &
                                        "Getdate(), " &
                                        "'" & gU.dbEncode(lIO_LOC) & "'," &
                                        "" & gU.convdbNVCData(gU.dbEncode(lIO_BATCH_NO)) & "," &
                                        "'" & gU.dbEncode(lIO_VND_CODE) & "'," &
                                        "'" & gU.dbEncode(lIO_WH) & "'," &
                                        gU.decodeNullOrEmpty(gU.dbEncode(lIO_FL), "NULL") & ", " &
                                        "'" & gU.dbEncode(lIO_AREA) & "'," &
                                        "'" & gU.dbEncode(lIO_RANK) & "'," &
                                        "'" & gU.dbEncode(lIO_BIN) & "'," &
                                        "ISNULL(ITX_BAL_BEFORE,0), " &
                                        "ISNULL(ITX_BAL_AFTER,0)  " & oper & " ISNULL(" & gU.decodeEmptyCInt(gU.dbEncode(updateQty), 0) & ",0), " &
                                        "ISNULL(ITX_BAL_CBM,0)  " & oper & " ISNULL(" & gU.decodeEmptyCdbl(gU.dbEncode(lIO_CBM), 0) & ",0), " &
                                        "ISNULL(ITX_BAL_KG,0)  " & oper & " ISNULL(" & gU.decodeEmptyCdbl(gU.dbEncode(lIO_KG), 0) & ",0), " &
                                        "ISNULL(" & gU.dbEncode(holdQty) & ",0)," &
                                        gU.convdbDate(lIO_EXPIRY_DATE) & "," &
                                        gU.convdbDate(lIO_MANU_DATE) & "," &
                                        "'" & Session("usr_id") & "', " &
                                        "Getdate(), " &
                                        "'" & Session("usr_id") & "'," &
                                        "Getdate() FROM WMS_ITEM_LOC_BAL_TX WHERE IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                                        "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND ITX_LOC = '" & gU.dbEncode(lIO_LOC) & "' " &
                                        "AND ITX_PALLET_NO = '" & gU.dbEncode(lPALLET_NO) & "' AND REPLACE(REPLACE(REPLACE(convert(varchar, ITX_DATE, 20),'-',''),':',''),' ','') = '" & dt.Rows(0).Item("F_ITX_DATE").ToString & "'" &
                                        ")"
                    Else
                        txsqlString = "INSERT INTO WMS_ITEM_LOC_BAL_TX (" &
                                        "IMP_CODE, " &
                                        "STORER_CODE, " &
                                        "ITM_CODE, " &
                                        "PACK_KEY, " &
                                        "ITX_PALLET_NO," &
                                        "ITX_DATE," &
                                        "ITX_LOC," &
                                        "ITX_BATCH_NO," &
                                        "VND_CODE," &
                                        "ITX_WH," &
                                        "ITX_FLOOR," &
                                        "ITX_AREA," &
                                        "ITX_RACK," &
                                        "ITX_BIN," &
                                        "ITX_BAL_BEFORE," &
                                        "ITX_BAL_AFTER," &
                                        "ITX_BAL_CBM," &
                                        "ITX_BAL_KG," &
                                        "ITX_HOLD_QTY," &
                                        "ITX_EXPIRY_DATE, " &
                                        "ITX_MANU_DATE, " &
                                        "SYS_CB, " &
                                        "SYS_CD, " &
                                        "SYS_LUB, " &
                                        "SYS_LUD" &
                                        ") VALUES (" &
                                        "'" & gU.dbEncode(lIMP_CODE) & "'," &
                                        "'" & gU.dbEncode(lSTORER_CODE) & "'," &
                                        "'" & gU.dbEncode(lITM_CODE) & "'," &
                                        "'" & gU.dbEncode(lPACK_KEY) & "'," &
                                        "'" & gU.dbEncode(lPALLET_NO) & "'," &
                                        "Getdate(), " &
                                        "'" & gU.dbEncode(lIO_LOC) & "'," &
                                        "" & gU.convdbNVCData(gU.dbEncode(lIO_BATCH_NO)) & "," &
                                        "'" & gU.dbEncode(lIO_VND_CODE) & "'," &
                                        "'" & gU.dbEncode(lIO_WH) & "'," &
                                        gU.decodeNullOrEmpty(gU.dbEncode(lIO_FL), "NULL") & ", " &
                                        "'" & gU.dbEncode(lIO_AREA) & "'," &
                                        "'" & gU.dbEncode(lIO_RANK) & "'," &
                                        "'" & gU.dbEncode(lIO_BIN) & "'," &
                                        "0," &
                                        gU.decodeEmptyCInt(gU.dbEncode(updateQty), 0) & ", " &
                                        gU.decodeEmptyCdbl(gU.dbEncode(lIO_CBM), 0) & ", " &
                                        gU.decodeEmptyCdbl(gU.dbEncode(lIO_KG), 0) & ", " &
                                        "ISNULL(" & gU.dbEncode(holdQty) & ",0)," &
                                        gU.convdbDate(lIO_EXPIRY_DATE) & "," &
                                        gU.convdbDate(lIO_MANU_DATE) & "," &
                                        "'" & Session("usr_id") & "', " &
                                        "Getdate(), " &
                                        "'" & Session("usr_id") & "'," &
                                        "Getdate() " &
                                        ")"
                    End If
                End If

                If pConn Is Nothing Then
                    Call gDB.amendData(sqlString)
                    Call gDB.amendData(txsqlString)
                Else
                    If pTransaction Is Nothing Then
                        Call gDB.amendData(sqlString, pConn)
                        Call gDB.amendData(txsqlString, pConn)
                    Else
                        Call gDB.amendData(sqlString, pConn, pTransaction)
                        Call gDB.amendData(txsqlString, pConn, pTransaction)
                    End If
                End If

                dt = Nothing
            Else
                If Session("gLang") = "E" Then
                    errorMsg = "Item " & lITM_CODE & "-" & lPACK_KEY & " not found in location " & lIO_LOC & "!"
                Else
                    errorMsg = "貨倉位置" & lIO_LOC & "找不到物料" & lITM_CODE & "-" & lPACK_KEY & "!"
                End If

                Return False
            End If
        Else
            If Session("gLang") = "E" Then
                errorMsg = "Item " & lITM_CODE & "-" & lPACK_KEY & " didn\'t assign location!"
            Else
                errorMsg = "物料" & lITM_CODE & "-" & lPACK_KEY & "沒有配置貨倉位置!"
            End If

            Return False
        End If

        UnPostStocks = True
    End Function

    Public Function UnPostSERIAL(ByVal type As IO_TYPE, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing, Optional ByRef errorMsg As String = "") As Boolean
        Dim SrchStr As String = ""
        Dim txSrchStr As String = ""
        Dim sqlString As String = ""
        Dim txsqlString As String = ""
        Dim dt As New DataTable
        Dim imTable As String = "WMS_ITEM"
        Dim oper As String = ""
        Dim unpostType As String = ""
        Dim iscable As Boolean = False
        Dim isSerial As Boolean = False
        Dim fILBS_SEQ As Integer

        SrchStr = "select * from V_LOCATION WHERE LOC = '" & gU.dbEncode(lIO_LOC) & "' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(SrchStr)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(SrchStr, pConn)
            Else
                dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            lIO_WH = dt.Rows(0).Item("WH_CODE").ToString
            lIO_FL = dt.Rows(0).Item("FL_NUM").ToString
            lIO_AREA = dt.Rows(0).Item("AR_CODE").ToString
            lIO_RANK = dt.Rows(0).Item("RK_CODE").ToString
            lIO_BIN = dt.Rows(0).Item("BN_CODE").ToString
        End If

        SrchStr = "select ITM_CODE from " & imTable & " WHERE isnull(ITM_TYPE,'') = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                    "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(SrchStr)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(SrchStr, pConn)
            Else
                dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            iscable = True
        End If

        SrchStr = "select ITM_CODE from " & imTable & " WHERE ITM_SERIAL_NO_YN = 'Y' AND isnull(ITM_TYPE,'') <> 'CABLE' AND IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                    "AND ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "
        If pConn Is Nothing Then
            dt = gDB.getDataTable(SrchStr)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(SrchStr, pConn)
            Else
                dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
            End If
        End If

        If dt.Rows.Count > 0 Then
            isSerial = True
        End If


        If isSerial = True Or iscable = True Then
            REM Update Item Location
            If lIO_LOC <> "" Then
                SrchStr = "select * from WMS_ITEM_LOC_BAL, WMS_ITEM_LOC_BAL_S WHERE WMS_ITEM_LOC_BAL.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND WMS_ITEM_LOC_BAL.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                "AND WMS_ITEM_LOC_BAL.ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND WMS_ITEM_LOC_BAL.PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' AND WMS_ITEM_LOC_BAL.ILOC_LOC = '" & gU.dbEncode(lIO_LOC) & "' AND WMS_ITEM_LOC_BAL.ILOC_PALLET_NO = '" & gU.dbEncode(lPALLET_NO) & "' " &
                "AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, '') = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, " ")) & "' AND WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " &
                "AND WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIOS_SERIAL_NO, " ")) & "' "
                If pConn Is Nothing Then
                    dt = gDB.getDataTable(SrchStr)
                Else
                    If pTransaction Is Nothing Then
                        dt = gDB.getDataTable(SrchStr, pConn)
                    Else
                        dt = gDB.getDataTable(SrchStr, pConn, pTransaction)
                    End If
                End If

                If dt.Rows.Count > 0 Then
                    Dim stQTY2 As Double = gU.decodeEmptyCInt(dt.Rows(0).Item("ILBS_QTY2").ToString, 0)
                    If iscable = True Then

                    Else
                        If isSerial = True Then
                            lIOS_QTY2 = 1
                        End If
                    End If

                    fILBS_SEQ = gU.decodeEmptyCInt(dt.Rows(0).Item("ILBS_SEQ"), 0)
                    lILOC_SEQ = gU.decodeEmptyCInt(dt.Rows(0).Item("ILOC_SEQ").ToString, 0)

                    If type = IO_TYPE.STOCKOUT Then
                        If stQTY2 < lIOS_QTY2 Then
                            If Session("gLang") = "E" Then
                                errorMsg = "Item " & lITM_CODE & "-" & lPACK_KEY & ", Serial " & lIOS_SERIAL_NO & " don\'t have enough stock balance to Un-Post!"
                            Else
                                errorMsg = "物料" & lITM_CODE & "-" & lPACK_KEY & ", Serial " & lIOS_SERIAL_NO & " 沒有足夠貨存去取消發布!"
                            End If

                            Return False
                        End If

                        oper = "-"
                        unpostType = "UNPOSTOUT"
                    ElseIf type = IO_TYPE.STOCKIN Then
                        oper = "+"
                        unpostType = "UNPOSTIN"
                    End If

                    UpdateStockSerialTrans(unpostType, pConn, pTransaction)

                    REM Gen UPDATE SQL
                    If iscable = True Then
                        sqlString = "UPDATE WMS_ITEM_LOC_BAL_S SET ILBS_SL = case when ISNULL(ILBS_QTY2,0) " & oper & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) = ILBS_ORG_QTY2 AND ILBS_SERIAL_NO = ISNULL(ILBS_DRUM_ID,'" & gU.dbEncode(lIOS_DRUM_ID) & "') THEN 'N' else 'Y' end, ILBS_QTY2 = ISNULL(ILBS_QTY2,0) " & oper & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) " &
                                    ", ILBS_DRUM_ID = '" & gU.dbEncode(lIOS_DRUM_ID) & "' " &
                                    ", ILBS_DRUM_LEVEL = '" & gU.dbEncode(lIOS_DRUM_LEVEL) & "' "
                    Else
                        If isSerial = True Then
                            If oper = "+" Then
                                sqlString = "UPDATE WMS_ITEM_LOC_BAL_S SET ILBS_QTY2 = 1 "
                                lIOS_QTY2 = 1
                            ElseIf oper = "-" Then
                                sqlString = "UPDATE WMS_ITEM_LOC_BAL_S SET ILBS_QTY2 = 0 "
                                lIOS_QTY2 = 1
                            End If
                        Else
                            If lIOS_QTY2 = 0 Then
                                lIOS_QTY2 = 1
                            End If
                            sqlString = "UPDATE WMS_ITEM_LOC_BAL_S SET ILBS_QTY2 = ISNULL(ILBS_QTY2,0) " & oper & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) "
                        End If
                    End If

                    sqlString = sqlString & ", SYS_LUB = '" & Session("usr_id") & "' " &
                                            ", SYS_LUD = Getdate() " &
                                            "WHERE ILBS_SEQ = '" & gU.dbEncode(fILBS_SEQ) & "'"

                    REM Gen Tx Update SQL
                    txSrchStr = "select * from WMS_ITEM_LOC_BAL_S_TX WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ITSX_SERIAL_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIOS_SERIAL_NO, " ")) & "' "

                    If pConn Is Nothing Then
                        dt = gDB.getDataTable(txSrchStr)
                    Else
                        If pTransaction Is Nothing Then
                            dt = gDB.getDataTable(txSrchStr, pConn)
                        Else
                            dt = gDB.getDataTable(txSrchStr, pConn, pTransaction)
                        End If
                    End If

                    If dt.Rows.Count > 0 Then
                        If iscable = True Then
                            txsqlString = "UPDATE WMS_ITEM_LOC_BAL_S_TX SET ITSX_BAL_BEFORE2 = ITSX_BAL_AFTER2, ITSX_BAL_AFTER2 = ISNULL(ITSX_BAL_AFTER2,0)  " & oper & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0) "
                        Else
                            If isSerial = True Then
                                If oper = "+" Then
                                    txsqlString = "UPDATE WMS_ITEM_LOC_BAL_S_TX SET ITSX_BAL_BEFORE2 = 0, ITSX_BAL_AFTER2 = 1 "
                                    lIOS_QTY2 = 1
                                ElseIf oper = "-" Then
                                    txsqlString = "UPDATE WMS_ITEM_LOC_BAL_S_TX SET ITSX_BAL_BEFORE2 = 1, ITSX_BAL_AFTER2 = 0 "
                                    lIOS_QTY2 = 1
                                End If
                            End If
                        End If

                        txsqlString = txsqlString & ", SYS_LUB = '" & Session("usr_id") & "' " &
                        ", SYS_LUD = Getdate() " &
                        "WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ITSX_SERIAL_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIOS_SERIAL_NO, " ")) & "' "
                    Else
                        dt = Nothing

                        txSrchStr = "select WMS_ITEM_LOC_BAL_S_TX.*, REPLACE(REPLACE(REPLACE(convert(varchar, ITX_DATE, 20),'-',''),':',''),' ','') as F_ITX_DATE from WMS_ITEM_LOC_BAL_S_TX " &
                            "WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ITSX_SERIAL_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIOS_SERIAL_NO, " ")) & "' ORDER BY ITX_DATE DESC"
                        If pConn Is Nothing Then
                            dt = gDB.getDataTable(txSrchStr)
                        Else
                            If pTransaction Is Nothing Then
                                dt = gDB.getDataTable(txSrchStr, pConn)
                            Else
                                dt = gDB.getDataTable(txSrchStr, pConn, pTransaction)
                            End If
                        End If
                        If dt.Rows.Count > 0 Then
                            REM INSERT SQL
                            txsqlString = "INSERT INTO WMS_ITEM_LOC_BAL_S_TX (" &
                                        "IMP_CODE, " &
                                        "STORER_CODE, " &
                                        "ITM_CODE, " &
                                        "PACK_KEY, " &
                                        "ITSX_SERIAL_NO, " &
                                        "ITSX_DRUM_ID, " &
                                        "ITSX_DRUM_LEVEL, " &
                                        "ITSX_UOM2, " &
                                        "ITSX_BAL_BEFORE2, " &
                                        "ITSX_BAL_AFTER2, " &
                                        "ITSX_ORG_QTY2, " &
                                        "ITSX_ORG_SERIAL_NO, " &
                                        "SYS_LUB, " &
                                        "SYS_LUD, " &
                                        "SYS_CD, " &
                                        "SYS_CB, " &
                                        "ILOC_SEQ, " &
                                        "ITX_DATE," &
                                        "ITSX_SL" &
                                        ") (SELECT " &
                                        "IMP_CODE, " &
                                        "STORER_CODE, " &
                                        "ITM_CODE, " &
                                        "PACK_KEY, " &
                                        "ITSX_SERIAL_NO, "
                            If unpostType = "UNPOSTOUT" Then
                                txsqlString = txsqlString & "ITSX_DRUM_ID, " &
                                        "ITSX_DRUM_LEVEL, "
                            Else
                                txsqlString = txsqlString & "'" & gU.dbEncode(lIOS_DRUM_ID) & "'," &
                                            "'" & gU.dbEncode(lIOS_DRUM_LEVEL) & "',"
                            End If
                            txsqlString = txsqlString & "ITSX_UOM2, " &
                                        "ISNULL(ITSX_BAL_BEFORE2,0), " &
                                        "ISNULL(ITSX_BAL_AFTER2,0) " & oper & " ISNULL(" & gU.dbEncode(lIOS_QTY2) & ",0), " &
                                        "ITSX_ORG_QTY2, " &
                                        "ITSX_ORG_SERIAL_NO, " &
                                        "'" & Session("usr_id") & "', " &
                                        "Getdate(), " &
                                        "Getdate(), " &
                                        "'" & Session("usr_id") & "'," &
                                        "ILOC_SEQ, " &
                                        "Getdate(),ITSX_SL FROM WMS_ITEM_LOC_BAL_S_TX WHERE ILOC_SEQ = '" & gU.dbEncode(lILOC_SEQ) & "' AND ITSX_SERIAL_NO = '" & gU.dbEncode(lIOS_SERIAL_NO) & "' AND REPLACE(REPLACE(REPLACE(convert(varchar, ITX_DATE, 20),'-',''),':',''),' ','') = '" & dt.Rows(0).Item("F_ITX_DATE").ToString & "') "
                        Else
                            REM INSERT SQL
                            txsqlString = "INSERT INTO WMS_ITEM_LOC_BAL_S_TX (" &
                                        "IMP_CODE, " &
                                        "STORER_CODE, " &
                                        "ITM_CODE, " &
                                        "PACK_KEY, " &
                                        "ITSX_SERIAL_NO, " &
                                        "ITSX_DRUM_ID, " &
                                        "ITSX_DRUM_LEVEL, " &
                                        "ITSX_UOM2, " &
                                        "ITSX_BAL_BEFORE2, " &
                                        "ITSX_BAL_AFTER2, " &
                                        "ITSX_ORG_QTY2, " &
                                        "ITSX_ORG_SERIAL_NO, " &
                                        "SYS_LUB, " &
                                        "SYS_LUD, " &
                                        "SYS_CD, " &
                                        "SYS_CB, " &
                                        "ILOC_SEQ, " &
                                        "ITX_DATE," &
                                        "ITSX_SL" &
                                        ") VALUES (" &
                                        "'" & gU.dbEncode(lIMP_CODE) & "'," &
                                        "'" & gU.dbEncode(lSTORER_CODE) & "'," &
                                        "'" & gU.dbEncode(lITM_CODE) & "'," &
                                        "'" & gU.dbEncode(lPACK_KEY) & "'," &
                                        "'" & gU.dbEncode(lIOS_SERIAL_NO) & "'," &
                                        "'" & gU.dbEncode(lIOS_DRUM_ID) & "'," &
                                        "'" & gU.dbEncode(lIOS_DRUM_LEVEL) & "'," &
                                        "'" & gU.dbEncode(lIOS_UOM2) & "'," &
                                        "0," &
                                        gU.decodeNullOrEmpty(gU.dbEncode(lIOS_QTY2), "0") & ", " &
                                        gU.decodeNullOrEmpty(gU.dbEncode(lIOS_ORG_QTY2), "NULL") & ", " &
                                        "'" & gU.dbEncode(lIOS_ORG_SERIAL_NO) & "'," &
                                        "'" & Session("usr_id") & "', " &
                                        "Getdate(), " &
                                        "Getdate(), " &
                                        "'" & Session("usr_id") & "'," &
                                        "'" & gU.dbEncode(lILOC_SEQ) & "'," &
                                        "Getdate(), "
                            If lIOS_SL <> "" Then
                                txsqlString = txsqlString & "'" & gU.dbEncode(lIOS_SL) & "')"
                            Else
                                txsqlString = txsqlString & "null)"
                            End If
                        End If
                    End If

                    If pConn Is Nothing Then
                        Call gDB.amendData(sqlString)
                        Call gDB.amendData(txsqlString)
                    Else
                        If pTransaction Is Nothing Then
                            Call gDB.amendData(sqlString, pConn)
                            Call gDB.amendData(txsqlString, pConn)
                        Else
                            Call gDB.amendData(sqlString, pConn, pTransaction)
                            Call gDB.amendData(txsqlString, pConn, pTransaction)
                        End If
                    End If

                    dt = Nothing
                Else
                    If Session("gLang") = "E" Then
                        errorMsg = "Item " & lITM_CODE & "-" & lPACK_KEY & ", Serial " & lIOS_SERIAL_NO & " not found in location " & lIO_LOC & "!"
                    Else
                        errorMsg = "貨倉位置" & lIO_LOC & "找不到物料" & lITM_CODE & "-" & lPACK_KEY & ", Serial " & lIOS_SERIAL_NO & "!"
                    End If

                    Return False
                End If
            Else
                If Session("gLang") = "E" Then
                    errorMsg = "Item " & lITM_CODE & "-" & lPACK_KEY & ", Serial " & lIOS_SERIAL_NO & " didn\'t assign location!"
                Else
                    errorMsg = "物料" & lITM_CODE & "-" & lPACK_KEY & ", Serial " & lIOS_SERIAL_NO & " 沒有配置貨倉位置!"
                End If

                Return False
            End If
        End If

        UnPostSERIAL = True
    End Function

    Private Function getHoldQty(Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As Double
        Dim HoldQty As Double = 0
        Dim SQLstring As String = ""
        Dim dt As DataTable
        SQLstring = "select sum(ISNULL(h.coh_in_stock_qty,0) - ISNULL(h.coh_rel_qty,0)) as hold_qty " &
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
                            "and ISNULL(h.coh_in_stock_qty, 0) > 0 " &
                            "and ISNULL(h.COH_PALLET_NO, '000') = '" & gU.dbEncode(lPALLET_NO) & "' " &
                            "and h.COH_BATCH_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, "")) & "' " &
                            "and h.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND h.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " &
                            "AND h.COH_ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND h.COH_PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "

        '"and d.COD_BATCH_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(lIO_BATCH_NO, "")) & "' " & _
        '                    "and d.IMP_CODE = '" & gU.dbEncode(lIMP_CODE) & "' AND d.STORER_CODE = '" & gU.dbEncode(lSTORER_CODE) & "' " & _
        '                    "AND d.COD_ITM_CODE = '" & gU.dbEncode(lITM_CODE) & "' AND d.COD_PACK_KEY = '" & gU.dbEncode(lPACK_KEY) & "' "

        If pConn Is Nothing Then
            dt = gDB.getDataTable(SQLstring)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(SQLstring, pConn)
            Else
                dt = gDB.getDataTable(SQLstring, pConn, pTransaction)
            End If
        End If
        If dt.Rows.Count > 0 Then
            HoldQty = gU.decodeEmptyCdbl(dt.Rows(0).Item(0).ToString, 0)
        End If
        Return HoldQty
    End Function

    Public Function cableReDrum(ByVal toDrum As String, ByVal CableSeqList As String, ByVal tIloc_seq As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As Boolean
        Dim successFlag As Boolean = False

        If CableSeqList <> "" Then
            Dim seqArr() As String = gU.listToArray(CableSeqList)

            Dim oriILOC_SEQ As String = ""
            Dim cableDT As DataTable
            Dim selectSQL As String = ""
            Dim updateSQL As String = ""
            Dim maxDrumLv As Integer = 0
            Dim tempDT As DataTable

            Dim ilbsArr() As String
            Dim fIlbs_seq As String = ""
            Dim fDrumLv As Integer = 0

            For i = 0 To seqArr.Length - 1
                fDrumLv = 0
                fIlbs_seq = ""

                ilbsArr = Split(seqArr(i), "||")

                fIlbs_seq = ilbsArr(0)
                If ilbsArr.Length > 1 Then
                    fDrumLv = gU.decodeEmptyCInt(ilbsArr(1), 0)
                End If

                selectSQL = "Select IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILBS_SERIAL_NO, ILBS_DRUM_ID, ILBS_DRUM_LEVEL, ILBS_UOM2, ILBS_QTY2, ILBS_ORG_QTY2, " &
                            " ILBS_ORG_SERIAL_NO, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, ILOC_SEQ, ILBS_SEQ, ILBS_KG, ILBS_PO_NO, ILBS_SL from wms_item_loc_bal_s" &
                            " Where ILBS_SEQ=" & gU.dbEncode(fIlbs_seq) & " and ILBS_QTY2 > 0 "
                cableDT = gDB.getDataTable(selectSQL, pConn, pTransaction)

                If cableDT.Rows.Count > 0 Then
                    selectSQL = "SELECT WMS_ITEM_LOC_BAL.itm_code FROM WMS_ITEM_LOC_BAL INNER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ where " &
                                " WMS_ITEM_LOC_BAL.ILOC_SEQ='" & gU.dbEncode(tIloc_seq) & "' "
                    '" WMS_ITEM_LOC_BAL.imp_code='" & cableDT.Rows(0).Item("imp_code").ToString.Trim & "' and WMS_ITEM_LOC_BAL.storer_code='" & cableDT.Rows(0).Item("storer_code").ToString.Trim & "' and WMS_ITEM_LOC_BAL.itm_code='" & cableDT.Rows(0).Item("itm_code").ToString.Trim & "' and WMS_ITEM_LOC_BAL.pack_key='" & cableDT.Rows(0).Item("pack_key").ToString.Trim & "' " & _

                    tempDT = gDB.getDataTable(selectSQL, pConn, pTransaction)

                    If tempDT.Rows.Count > 0 Then

                        Dim lItm_code As String = tempDT.Rows(0).Item("itm_code").ToString.Trim

                        If lItm_code <> "EMPTY_DRUM" And lItm_code <> "" Then

                            'selectSQL = "Select max(ILBS_DRUM_LEVEL) from wms_item_loc_bal_s where ILBS_DRUM_ID='" & gU.dbEncode(toDrum) & "' and ILBS_QTY2 > 0 "
                            'maxDrumLv = gU.decodeEmptyCInt(DB.getValueFromSQL(selectSQL, pConn, pTransaction), 0) + 1

                            maxDrumLv = fDrumLv

                            updateSQL = "Update wms_item_loc_bal_s set " &
                                        " ILOC_SEQ='" & gU.dbEncode(tIloc_seq) & "'," &
                                        " ILBS_DRUM_ID='" & gU.dbEncode(toDrum) & "'," &
                                        " ILBS_DRUM_LEVEL='" & gU.dbEncode(maxDrumLv) & "'," &
                                        " sys_lud=getdate(), sys_lub='" & Session("usr_id") & "'" &
                                        " Where ILBS_SEQ='" & gU.dbEncode(fIlbs_seq) & "'"

                            gDB.amendData(updateSQL, pConn, pTransaction)

                            updateSQL = "update wms_item_loc_bal set " &
                                        " ILOC_BAL_QTY=ILOC_BAL_QTY - 1, " &
                                        " sys_lud=getdate(), sys_lub='" & Session("usr_id") & "'" &
                                        " where ILOC_SEQ='" & gU.dbEncode(cableDT.Rows(0).Item("ILOC_SEQ").ToString.Trim) & "'"

                            gDB.amendData(updateSQL, pConn, pTransaction)

                            updateSQL = "update wms_item_loc_bal set " &
                                        " ILOC_BAL_QTY=ILOC_BAL_QTY + 1, " &
                                        " sys_lud=getdate(), sys_lub='" & Session("usr_id") & "'" &
                                        " where ILOC_SEQ='" & gU.dbEncode(tIloc_seq) & "'"

                            gDB.amendData(updateSQL, pConn, pTransaction)

                            successFlag = True
                        Else
                            maxDrumLv = fDrumLv

                            updateSQL = "Update wms_item_loc_bal_s set " &
                                        " ILOC_SEQ='" & gU.dbEncode(tIloc_seq) & "'," &
                                        " ILBS_DRUM_ID='" & gU.dbEncode(toDrum) & "'," &
                                        " ILBS_DRUM_LEVEL='" & gU.dbEncode(maxDrumLv) & "'," &
                                        " sys_lud=getdate(), sys_lub='" & Session("usr_id") & "'" &
                                        " Where ILBS_SEQ='" & gU.dbEncode(fIlbs_seq) & "'"

                            gDB.amendData(updateSQL, pConn, pTransaction)


                            updateSQL = " UPDATE WMS_ITEM_LOC_BAL " &
                                        " SET IMP_CODE = b.IMP_CODE, STORER_CODE =b.STORER_CODE, ITM_CODE =b.ITM_CODE, PACK_KEY =b.PACK_KEY, ILOC_PALLET_NO =b.ILOC_PALLET_NO, ILOC_LOC =b.ILOC_LOC, ILOC_WH =b.ILOC_WH, ILOC_FLOOR =b.ILOC_FLOOR, ILOC_AREA =b.ILOC_AREA,  " &
                                        " ILOC_RACK =b.ILOC_RACK, ILOC_BIN =b.ILOC_BIN, ILOC_BAL_QTY =1, ILOC_BAL_CBM =b.ILOC_BAL_CBM,ILOC_BAL_KG =b.ILOC_BAL_KG, SYS_LUB ='" & Session("usr_id") & "', SYS_LUD =getdate(), ILOC_BATCH_NO =b.ILOC_BATCH_NO, MFG_ITM_CODE =b.MFG_ITM_CODE, ILOC_EXPIRY_DATE =b.ILOC_EXPIRY_DATE,  " &
                                        " ILOC_MANU_DATE =b.ILOC_MANU_DATE, VND_CODE =b.VND_CODE, ILOC_PO_NO = b.ILOC_PO_NO " &
                                        " FROM (SELECT  IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILOC_PALLET_NO, ILOC_LOC, ILOC_WH, ILOC_FLOOR, ILOC_AREA, ILOC_RACK, ILOC_BIN, " &
                                        " ILOC_BAL_CBM, ILOC_BAL_KG, ILOC_BATCH_NO, MFG_ITM_CODE, ILOC_EXPIRY_DATE, ILOC_MANU_DATE, VND_CODE,  " &
                                        " ITM_KEY2, ITM_KEY, ITM_KEY3, ILOC_SEQ, ILOC_PO_NO " &
                                        " FROM WMS_ITEM_LOC_BAL WHERE ILOC_SEQ = '" & gU.dbEncode(cableDT.Rows(0).Item("ILOC_SEQ").ToString.Trim) & "') b  " &
                                        " WHERE (WMS_ITEM_LOC_BAL.ILOC_SEQ = '" & gU.dbEncode(tIloc_seq) & "') "
                            gDB.amendData(updateSQL, pConn, pTransaction)

                            updateSQL = "update wms_item_loc_bal set " &
                                        " ILOC_BAL_QTY=ILOC_BAL_QTY - 1, " &
                                        " sys_lud=getdate(), sys_lub='" & Session("usr_id") & "'" &
                                        " where ILOC_SEQ='" & gU.dbEncode(cableDT.Rows(0).Item("ILOC_SEQ").ToString.Trim) & "'"

                            gDB.amendData(updateSQL, pConn, pTransaction)


                            updateSQL = "Delete from wms_item_loc_bal_s where itm_code='EMPTY_DRUM' and ILOC_SEQ='" & gU.dbEncode(tIloc_seq) & "'"
                            gDB.amendData(updateSQL, pConn, pTransaction)

                            successFlag = True
                        End If

                    End If
                End If

            Next


        End If

        Return successFlag
    End Function


    Public Sub New()
        Dim sqlString As String = "select IMP_CODE from WMS_SETTINGS_APPLICATION"
        Dim dt As New DataTable
        dt = gDB.getDataTable(sqlString)
        lIMP_CODE = dt.Rows(0).Item(0).ToString
        dt = Nothing
    End Sub
End Class
