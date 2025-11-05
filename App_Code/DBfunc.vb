Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient


Public Class DBfunc

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils

    Public Function UpdateStockTrans(ByVal in_out_type As String, ByVal STORER_CODE As String, ByVal ITM_CODE As String, ByVal PACK_KEY As String, ByVal IO_DOC As String, ByVal IO_DOC_ID As String, ByVal IO_QTY As String, ByVal IO_CUST_CODE As String, ByVal IO_LOC As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As String
        Dim sqlString As String = "select IMP_CODE from WMS_SETTINGS_APPLICATION"
        Dim IMP_CODE As String
        Dim dt As New DataTable
        Dim txTable As String = ""
        Dim calType As String = ""
        Dim imTable As String = "WMS_ITEM"
        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        IMP_CODE = dt.Rows(0).Item(0).ToString
        dt = Nothing
        REM Update TX Record
        If in_out_type = "IN" Then
            txTable = "WMS_IN_TX"
            calType = "+"
        ElseIf in_out_type = "OUT" Then
            txTable = "WMS_OUT_TX"
            calType = "-"
        End If

        If IO_QTY = "" Then
            IO_QTY = 0
        End If
        sqlString = "INSERT INTO " & txTable & " (" & _
        "IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, IO_CODE, IO_TYPE, IO_CUST_CODE, IO_LOC, IO_DATETIME, IO_DOC, IO_DOC_ID, IO_QTY, IO_BAL_BEFORE, IO_BAL_AFTER" & _
        ") (SELECT " & _
        "IMP_CODE," & _
        "STORER_CODE," & _
        "ITM_CODE," & _
        "PACK_KEY," & _
        "getdate()," & _
        "N'" & gU.dbEncode(in_out_type) & "'," & _
        "N'" & gU.dbEncode(IO_CUST_CODE) & "'," & _
        "N'" & gU.dbEncode(IO_LOC) & "'," & _
        "getdate()," & _
        "N'" & gU.dbEncode(IO_DOC) & "'," & _
        "N'" & gU.dbEncode(IO_DOC_ID) & "'," & _
        "" & gU.dbEncode(IO_QTY) & "," & _
        "ITM_BALANCE," & _
        "ITM_BALANCE " & calType & "" & gU.dbEncode(IO_QTY) & " " & _
        "FROM " & imTable & " WHERE IMP_CODE = N'" & gU.dbEncode(IMP_CODE) & "' AND STORER_CODE = N'" & gU.dbEncode(STORER_CODE) & "' " & _
        "AND ITM_CODE = N'" & gU.dbEncode(ITM_CODE) & "' AND PACK_KEY = N'" & gU.dbEncode(PACK_KEY) & "')"

        If pConn Is Nothing Then
            Call gDB.amendData(sqlString)
        Else
            If pTransaction Is Nothing Then
                Call gDB.amendData(sqlString, pConn)
            Else
                Call gDB.amendData(sqlString, pConn, pTransaction)
            End If
        End If
        REM Update Item Master Record
        sqlString = "UPDATE " & imTable & " SET ITM_BALANCE = ITM_BALANCE " & calType & "" & gU.dbEncode(IO_QTY) & " " & _
        "WHERE IMP_CODE = N'" & gU.dbEncode(IMP_CODE) & "' AND STORER_CODE = N'" & gU.dbEncode(STORER_CODE) & "' " & _
        "AND ITM_CODE = N'" & gU.dbEncode(ITM_CODE) & "' AND PACK_KEY = N'" & gU.dbEncode(PACK_KEY) & "'"

        If pConn Is Nothing Then
            Call gDB.amendData(sqlString)
        Else
            If pTransaction Is Nothing Then
                Call gDB.amendData(sqlString, pConn)
            Else
                Call gDB.amendData(sqlString, pConn, pTransaction)
            End If
        End If

        UpdateStockTrans = True
    End Function

    Public Function getDocNo(ByVal dot_type As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As String
        Dim sqlString As String = "select DOC_NO_LEN, DOC_PAD_LENGTH,DOC_PREFIX from WMS_DOC_NO where DOC_TYPE= '" & dot_type & "'"
        Dim nextNo As String = ""
        Dim dt As New DataTable
        Dim padLength As String = ""

        If pConn Is Nothing Then
            dt = gDB.getDataTable(sqlString)
        Else
            If pTransaction Is Nothing Then
                dt = gDB.getDataTable(sqlString, pConn)
            Else
                dt = gDB.getDataTable(sqlString, pConn, pTransaction)
            End If
        End If

        'nextNo = Right("000000" & dt.Rows(0).Item(0).ToString, 6)

        If dt.Rows.Count > 0 Then
            padLength = gU.decodeNullOrEmpty(dt.Rows(0).Item("DOC_PAD_LENGTH").ToString, "6")
            nextNo = dt.Rows(0).Item(0).ToString.PadLeft(CInt(padLength), "0")

            sqlString = "update WMS_DOC_NO set DOC_NO_LEN = DOC_NO_LEN + 1 where DOC_TYPE = '" & dot_type & "'"


            If pConn Is Nothing Then
                Call gDB.amendData(sqlString)
            Else
                If pTransaction Is Nothing Then
                    Call gDB.amendData(sqlString, pConn)
                Else
                    Call gDB.amendData(sqlString, pConn, pTransaction)
                End If
            End If
            If dt.Rows(0).Item(2).ToString <> "" Then
                nextNo = dt.Rows(0).Item(2).ToString & nextNo
            End If
        Else
            nextNo = "#ERROR"
        End If

        dt.Dispose()

        getDocNo = nextNo

    End Function

    'Public Function getDocNo(ByVal comp_code As String, ByVal doc_type As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As String
    '    Dim selectSql, updateSql As String
    '    Dim nextDocNo As String = ""
    '    Dim nextNo As Long
    '    Dim dt As New DataTable
    '    Dim padLength As String = ""
    '    Dim selCmdPa, updtCmdPa As GlobalDBFunc.DBCmdPara
    '    Dim docPrefix As String = ""

    '    selCmdPa = New GlobalDBFunc.DBCmdPara

    '    selectSql = "select docn_next_no, docn_prefix " & _
    '                "from lm_doc_no " & _
    '                "where comp_code = " & selCmdPa.AP(comp_code) & " " & _
    '                "and doc_type = " & selCmdPa.AP(doc_type) & " "

    '    dt = gDB.getDataTable(selectSql, pConn, pTransaction, , selCmdPa)

    '    'nextNo = Right("000000" & dt.Rows(0).Item(0).ToString, 6)

    '    If dt.Rows.Count > 0 Then
    '        padLength = "6"
    '        nextNo = dt.Rows(0).Item("docn_next_no")
    '        nextDocNo = nextNo.ToString.PadLeft(CInt(padLength), "0")

    '        docPrefix = dt.Rows(0).Item("docn_prefix").ToString.Trim

    '        updateSql = "update lm_doc_no " & _
    '                    "set docn_next_no = docn_next_no + 1 " & _
    '                    "where comp_code = :p_comp_code " & _
    '                    "and doc_type = :p_doc_type " & _
    '                    "and docn_next_no = :p_docn_next_no "

    '        updtCmdPa = New GlobalDBFunc.DBCmdPara
    '        updtCmdPa.Add("p_comp_code", sqlDbType.Varchar, comp_code)
    '        updtCmdPa.Add("p_doc_type", sqlDbType.Varchar, doc_type)
    '        updtCmdPa.Add("p_docn_next_no", sqlDbType.Decimal, nextNo)

    '        Do While (gDB.amendData2(updateSql, pConn, pTransaction, updtCmdPa) = 0)
    '            'If update row = 0, that means the doc no is changed by other process, so need get the doc no again
    '            dt = gDB.getDataTable(selectSql, pConn, pTransaction, , selCmdPa)

    '            'padLength = gU.decodeNullOrEmpty(dt.Rows(0).Item("DOC_PAD_LENGTH").ToString, "6")
    '            nextNo = dt.Rows(0).Item("docn_next_no")
    '            nextDocNo = nextNo.ToString.PadLeft(CInt(padLength), "0")

    '            updtCmdPa = New GlobalDBFunc.DBCmdPara
    '            updtCmdPa.Add("p_comp_code", sqlDbType.Varchar, comp_code)
    '            updtCmdPa.Add("p_doc_type", sqlDbType.Varchar, doc_type)
    '            updtCmdPa.Add("p_docn_next_no", sqlDbType.Decimal, nextNo)
    '        Loop

    '        If docPrefix <> "" Then
    '            nextDocNo = docPrefix & nextDocNo
    '        End If
    '    Else
    '        Throw New Exception("No document number found! " & doc_type)
    '    End If

    '    dt.Dispose()

    '    getDocNo = nextDocNo

    'End Function

    Public Function getValueFromSQL(ByVal SQLString As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByRef pTransaction As SqlTransaction = Nothing, Optional ByRef cmdPara As GlobalDBFunc.DBCmdPara = Nothing, Optional ByRef pLog As PrgmLog = Nothing) As String
        Dim ldt As DataTable
        SQLString = SQLString

        'If pConn Is Nothing Then
        '    ldt = gDB.getDataTable(SQLString)
        'Else
        '    If pTransaction Is Nothing Then
        '        ldt = gDB.getDataTable(SQLString, pConn)
        '    Else
        '        ldt = gDB.getDataTable(SQLString, pConn, pTransaction)
        '    End If
        'End If

        ldt = gDB.getDataTable(SQLString, pConn, pTransaction, , cmdPara, pLog)

        'ldt = gDB.getDataTable(SQLString)

        If ldt.Rows.Count > 0 Then
            getValueFromSQL = decodeDBNull(ldt.Rows(0).Item(0).ToString, "")
        Else
            getValueFromSQL = ""
        End If
    End Function

    Public Function decodeDBNull(ByVal aValue As Object, ByVal defaultVal As String) As String
        If IsDBNull(aValue) Then
            decodeDBNull = defaultVal
        ElseIf CStr(aValue) = "#|MAP_NULL|#" Then
            decodeDBNull = defaultVal
        Else
            decodeDBNull = CStr(aValue)
        End If
    End Function

    Public Function chkItmListExist(ByVal imp_code As String, ByVal storer_code As String, ByRef itmList As List(Of String), ByRef invalidItm As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As Boolean
        Dim selectSql As String
        Dim i As Integer
        Dim itmArray As String()

        For i = 0 To itmList.Count - 1
            itmArray = Split(itmList(i), "#_#")

            selectSql = "select count(*) from WMS_ITEM " & _
                        "where imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                        "and storer_code = '" & gU.dbEncode(storer_code) & "' " & _
                        "and itm_code = '" & gU.dbEncode(itmArray(0)) & "' " & _
                        "and pack_key = '" & gU.dbEncode(itmArray(1)) & "' "

            If getValueFromSQL(selectSql, pConn, pTransaction) = "0" Then
                invalidItm = itmList(i)
                Return False
            End If
        Next

        Return True
    End Function

    Public Function chkItmExist(ByVal imp_code As String, ByVal storer_code As String, ByVal itm_code As String, ByVal pack_key As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByVal pTransaction As SqlTransaction = Nothing) As Boolean
        Dim selectSql As String

        selectSql = "select count(*) from WMS_ITEM " & _
                    "where imp_code = '" & gU.dbEncode(imp_code) & "' " & _
                    "and storer_code = '" & gU.dbEncode(storer_code) & "' " & _
                    "and itm_code = '" & gU.dbEncode(itm_code) & "' " & _
                    "and pack_key = '" & gU.dbEncode(pack_key) & "' "

        If getValueFromSQL(selectSql, pConn, pTransaction) = "0" Then
            Return False
        End If

        Return True
    End Function

    Public Function getAddr(ByVal ID As String, ByVal seq As String) As String
        Dim fullAddr As String
        Dim sqlString As String
        Dim dt As DataTable
        Dim cmdPara As GlobalDBFunc.DBCmdPara

        sqlString = "select tms_streets.*,  " & _
                    "wms_region.region_name, wms_region.region_chi_name, " & _
                    "wms_district.dis_name, wms_district.dis_chi_name, " & _
                    "wms_area.area_name, wms_area.area_chi_name " & _
                    "from tms_streets, wms_region, wms_district, wms_area where " & _
                    "tms_streets.st_region_code = wms_region.region_code (+) " & _
                    "and tms_streets.st_district_code = wms_district.dis_code (+) " & _
                    "and tms_streets.st_area_code = wms_area.area_code (+) " & _
                    "and tms_streets.st_id = :p_st_id " & _
                    "and tms_streets.st_seq = :p_st_seq "

        cmdPara = New GlobalDBFunc.DBCmdPara
        cmdPara.Add("p_st_id", SqlDbType.VarChar, ID)
        cmdPara.Add("p_st_seq", sqlDbType.Decimal, seq)

        dt = gDB.getDataTable(sqlString, , , , cmdPara)

        If dt.Rows.Count > 0 Then
            Dim st_Address As String = ""

            If dt.Rows(0).Item("st_bldg_name").ToString.Trim <> "" Then
                st_Address = st_Address & dt.Rows(0).Item("st_bldg_name").ToString.Trim
            ElseIf dt.Rows(0).Item("st_chi_bldg_name").ToString.Trim <> "" Then
                st_Address = st_Address & dt.Rows(0).Item("st_chi_bldg_name").ToString.Trim
            End If

            If dt.Rows(0).Item("st_estate_name").ToString.Trim <> "" Then
                If st_Address <> "" Then st_Address = st_Address & ", "
                st_Address = st_Address & dt.Rows(0).Item("st_estate_name").ToString.Trim
            ElseIf dt.Rows(0).Item("st_chi_estate_name").ToString.Trim <> "" Then
                If st_Address <> "" Then st_Address = st_Address & ", "
                st_Address = st_Address & dt.Rows(0).Item("st_chi_estate_name").ToString.Trim
            End If

            If dt.Rows(0).Item("st_no").ToString.Trim <> "" Then
                If st_Address <> "" Then st_Address = st_Address & "<br>"

                st_Address = st_Address & dt.Rows(0).Item("st_no").ToString.Trim

                If dt.Rows(0).Item("st_name").ToString.Trim <> "" Then
                    st_Address = st_Address & " " & dt.Rows(0).Item("st_name").ToString.Trim
                ElseIf dt.Rows(0).Item("st_chi_name").ToString.Trim <> "" Then
                    st_Address = st_Address & " " & dt.Rows(0).Item("st_chi_name").ToString.Trim
                End If

            ElseIf dt.Rows(0).Item("st_name").ToString.Trim <> "" Then
                If st_Address <> "" Then st_Address = st_Address & "<br>"
                st_Address = st_Address & dt.Rows(0).Item("st_name").ToString.Trim

            ElseIf dt.Rows(0).Item("st_chi_name").ToString.Trim <> "" Then
                If st_Address <> "" Then st_Address = st_Address & "<br>"
                st_Address = st_Address & dt.Rows(0).Item("st_chi_name").ToString.Trim

            End If

            If dt.Rows(0).Item("area_name").ToString.Trim <> "" Then
                If st_Address <> "" Then st_Address = st_Address & "<br>"
                st_Address = st_Address & dt.Rows(0).Item("area_name").ToString.Trim
            ElseIf dt.Rows(0).Item("area_chi_name").ToString.Trim <> "" Then
                If st_Address <> "" Then st_Address = st_Address & "<br>"
                st_Address = st_Address & dt.Rows(0).Item("area_chi_name").ToString.Trim
            End If

            'If dt.Rows(0).Item("dis_name").ToString.Trim <> "" Then
            '    If st_Address <> "" Then st_Address = st_Address & "<br>"
            '    st_Address = st_Address & dt.Rows(0).Item("dis_name").ToString.Trim
            'ElseIf dt.Rows(0).Item("dis_chi_name").ToString.Trim <> "" Then
            '    If st_Address <> "" Then st_Address = st_Address & "<br>"
            '    st_Address = st_Address & dt.Rows(0).Item("dis_chi_name").ToString.Trim
            'End If

            If dt.Rows(0).Item("region_name").ToString.Trim <> "" Then
                If st_Address <> "" Then st_Address = st_Address & "<br>"
                st_Address = st_Address & dt.Rows(0).Item("region_name").ToString.Trim
            ElseIf dt.Rows(0).Item("region_chi_name").ToString.Trim <> "" Then
                If st_Address <> "" Then st_Address = st_Address & "<br>"
                st_Address = st_Address & dt.Rows(0).Item("region_chi_name").ToString.Trim
            End If

            fullAddr = st_Address
        Else
            fullAddr = ""
        End If

        Return fullAddr
    End Function

End Class
