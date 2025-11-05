Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class CtrlDataUpdt
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private gU As New GeneralUtils

    Private pConn As SqlConnection
    Private pTrans As SqlTransaction
    Private hasDirtyList As Boolean = False
    Private dirtyList As String = ""
    Private hasPrevDt As Boolean = False
    Private prevDt As DataTable
    Private dataDt As DataTable
    Private prevDtRowNo As Integer = 0
    Private tableCdm As CtrlDataMap
    Private custUpdtDict As New Dictionary(Of String, custField)
    Private keyDict As New Dictionary(Of String, custField)
    Private failReason As String = ""
    Private DDFORMAT As String = System.Configuration.ConfigurationManager.AppSettings.Item("DDFORMAT")
    Private nextNo As Integer = -1
    Private formHolderID As String = "MainContent"
    Private updtSysUserDate As Boolean = True

    Private Const custRefPrefix = "#["
    Private Const custRefSuffix = "]#"
    Private Const pageNoStr = "_#P#_"

    Public Sub New(ByRef ctrlMapping As CtrlDataMap, Optional ByRef cnn As SqlConnection = Nothing, Optional ByRef transaction As SqlTransaction = Nothing)
        tableCdm = ctrlMapping
        pConn = cnn
        pTrans = transaction
    End Sub

    Public Function getDirtyDict() As Dictionary(Of Integer, List(Of String))
        Dim dirtyArray, dirtyStrArray As String()
        Dim dirtyDict As Dictionary(Of Integer, List(Of String))
        Dim tmpDtlNo As Integer
        Dim tmpDtlFieldName As String
        Dim tmpDirtyFieldList As List(Of String)
        Dim ctrlPrefix As String = ""
        Dim formPrefix As String
        Dim tmpDirtyArrayStr As String

        If formHolderID <> "" Then
            formPrefix = formHolderID & "_"
        Else
            formPrefix = ""
        End If

        dirtyDict = New Dictionary(Of Integer, List(Of String))

        If dirtyList.Trim = "" Then
            Return dirtyDict
        End If

        Dim pageidxCnt As Integer = 1
        Dim pageSize As Integer = 0

        Select Case UCase(TypeName(tableCdm.DataParentControl))

            Case "REPEATER"
                Dim tmpRepeater As Repeater = CType(tableCdm.DataParentControl, Repeater)

                If tmpRepeater.Items.Count <= 0 Then
                    Return dirtyDict
                Else
                    If tmpRepeater.Items(0).FindControl("_T_PREFIX_T_") IsNot Nothing Then
                        ctrlPrefix = Left(tmpRepeater.Items(0).FindControl("_T_PREFIX_T_").ClientID, InStr(tmpRepeater.Items(0).FindControl("_T_PREFIX_T_").ClientID, "_T_PREFIX_T_") - 1)
                    Else
                        ctrlPrefix = formPrefix & tableCdm.DataParentControl.ID & "_"
                    End If
                End If

            Case "GRIDVIEW"
                Dim tmpGV As GridView = CType(tableCdm.DataParentControl, GridView)
                Dim hasPrefixCtrl As Boolean = False

                If tmpGV.Rows.Count <= 0 Then
                    Return dirtyDict
                Else
                    For Each itemCell As TableCell In tmpGV.Rows(0).Cells
                        If itemCell.FindControl("_T_PREFIX_T_") IsNot Nothing Then
                            ctrlPrefix = Left(itemCell.FindControl("_T_PREFIX_T_").ClientID, InStr(itemCell.FindControl("_T_PREFIX_T_").ClientID, "_T_PREFIX_T_") - 1)
                            hasPrefixCtrl = True
                            Exit For
                        End If
                    Next
                    If Not hasPrefixCtrl Then
                        ctrlPrefix = formPrefix & tableCdm.DataParentControl.ID & "_"
                    End If
                End If

                pageidxCnt = tmpGV.PageCount
                pageSize = tmpGV.PageSize

            Case Else
                Throw New Exception("Only GridView or Repeater control can generate dirty dictionary!")

        End Select

        'dtlNoList = New List(Of Integer)

        Dim pageNo As Integer

        If hasDirtyList AndAlso dirtyList.Trim <> "" Then
            dirtyArray = gU.jsListToArray(dirtyList)

            For i = 0 To UBound(dirtyArray)
                If InStr(dirtyArray(i), pageNoStr) > 0 Then
                    dirtyStrArray = Split(dirtyArray(i), pageNoStr)

                    tmpDirtyArrayStr = dirtyStrArray(0)
                    pageNo = CInt(dirtyStrArray(1))
                Else
                    tmpDirtyArrayStr = dirtyArray(i)
                    pageNo = -1
                End If

                tmpDtlNo = CInt(Mid(tmpDirtyArrayStr, InStrRev(tmpDirtyArrayStr, "_") + 1))

                If pageNo > 0 Then
                    tmpDtlNo += pageSize * pageNo

                    If Left(tmpDirtyArrayStr, Len(ctrlPrefix)) = ctrlPrefix Then
                        tmpDtlFieldName = Mid(Left(tmpDirtyArrayStr, InStrRev(tmpDirtyArrayStr, "_") - 1), Len(ctrlPrefix) + 1)

                        If tmpDtlFieldName.Contains("__") Then
                            tmpDtlFieldName = tmpDtlFieldName.Substring(tmpDtlFieldName.IndexOf("__") + 2, tmpDtlFieldName.Length - tmpDtlFieldName.IndexOf("__") - 2)
                        End If

                        If dirtyDict.ContainsKey(tmpDtlNo) Then
                            tmpDirtyFieldList = dirtyDict.Item(tmpDtlNo)
                            If Not tmpDirtyFieldList.Contains(tmpDtlFieldName) Then
                                tmpDirtyFieldList.Add(tmpDtlFieldName)
                            End If
                        Else
                            tmpDirtyFieldList = New List(Of String)
                            tmpDirtyFieldList.Add(tmpDtlFieldName)
                            dirtyDict.Add(tmpDtlNo, tmpDirtyFieldList)
                        End If
                    End If
                Else
                    For cnt As Integer = 1 To pageidxCnt
                        If Left(tmpDirtyArrayStr, Len(ctrlPrefix)) = ctrlPrefix Then
                            tmpDtlFieldName = Mid(Left(tmpDirtyArrayStr, InStrRev(tmpDirtyArrayStr, "_") - 1), Len(ctrlPrefix) + 1)

                            If tmpDtlFieldName.Contains("__") Then
                                tmpDtlFieldName = tmpDtlFieldName.Substring(tmpDtlFieldName.IndexOf("__") + 2, tmpDtlFieldName.Length - tmpDtlFieldName.IndexOf("__") - 2)
                            End If

                            If dirtyDict.ContainsKey(tmpDtlNo) Then
                                tmpDirtyFieldList = dirtyDict.Item(tmpDtlNo)
                                If Not tmpDirtyFieldList.Contains(tmpDtlFieldName) Then
                                    tmpDirtyFieldList.Add(tmpDtlFieldName)
                                End If
                            Else
                                tmpDirtyFieldList = New List(Of String)
                                tmpDirtyFieldList.Add(tmpDtlFieldName)
                                dirtyDict.Add(tmpDtlNo, tmpDirtyFieldList)
                            End If
                        End If

                        tmpDtlNo += pageSize
                    Next
                End If
            Next
        End If

        Return dirtyDict
    End Function

    Private Function getCustRefValue(ByVal custValue As String, ByRef dtRow As DataRow, ByRef cmdPa As GlobalDBFunc.DBCmdPara) As String
        Dim tmpCustValue As String
        Dim tmpRefFieldName As String
        Dim preLoc, suffLoc As Integer

        tmpCustValue = custValue

        preLoc = InStr(tmpCustValue, custRefPrefix)
        suffLoc = InStr(tmpCustValue, custRefSuffix)

        Do While preLoc > 0 AndAlso suffLoc > 0
            tmpRefFieldName = Mid(tmpCustValue, preLoc + Len(custRefPrefix), suffLoc - preLoc - Len(custRefPrefix))

            If tmpRefFieldName = "|SEQ_NO|" Then
                tmpCustValue = Replace(tmpCustValue, custRefPrefix & tmpRefFieldName & custRefSuffix, cmdPa.AP(nextNo))
            Else
                tmpCustValue = Replace(tmpCustValue, custRefPrefix & tmpRefFieldName & custRefSuffix, cmdPa.AP(dtRow.Item(tmpRefFieldName).ToString.Trim))
            End If

            preLoc = InStr(tmpCustValue, custRefPrefix)
            suffLoc = InStr(tmpCustValue, custRefSuffix)
        Loop

        Return tmpCustValue
    End Function

    Public Function saveDatatable(ByVal tableName As String, ByRef dirtyDict As Dictionary(Of Integer, List(Of String)), Optional ByVal keyList As String = "") As Boolean
        Dim cnn As SqlConnection
        Dim transaction As SqlTransaction
        Dim cmdPa As GlobalDBFunc.DBCmdPara
        Dim insertSql, insSqlPart1, insSqlPart2, updateSql, deleteSql As String
        Dim dictItem As KeyValuePair(Of String, custField)
        Dim cdmItem As KeyValuePair(Of String, CtrlDataMap.DataFieldCtrl)
        Dim fieldList As List(Of String)
        Dim keyArray As String()
        Dim i, j As Integer
        Dim tmpCustValue As String
        'Dim formPrefix As String

        If dataDt Is Nothing Then
            failReason = "No data table!"
            Return False
        End If

        If pConn Is Nothing Then
            cnn = gDB.getConnection()
        Else
            cnn = pConn
        End If

        transaction = pTrans

        Try
            For i = 0 To dataDt.Rows.Count - 1
                Select Case dataDt.Rows(i).Item("mFlag").ToString.Trim
                    Case "N"
                        'Insert records ---------------------------------------------------
                        fieldList = New List(Of String)

                        cmdPa = New GlobalDBFunc.DBCmdPara
                        insSqlPart1 = "insert into " & tableName & " ("
                        insSqlPart2 = "values ("

                        For Each dictItem In custUpdtDict
                            If Not dictItem.Value.updateOnly Then
                                If Not fieldList.Contains(UCase(dictItem.Value.fieldName)) Then
                                    insSqlPart1 += dictItem.Value.fieldName & ", "

                                    If dictItem.Value.fieldType = "SQL" AndAlso InStr(dictItem.Value.fieldValue, custRefPrefix) > 0 AndAlso InStr(dictItem.Value.fieldValue, custRefSuffix) > 0 Then
                                        tmpCustValue = getCustRefValue(dictItem.Value.fieldValue, dataDt.Rows(i), cmdPa)
                                    Else
                                        tmpCustValue = dictItem.Value.fieldValue
                                    End If

                                    insSqlPart2 += getSqlStr(tmpCustValue, dictItem.Value.fieldType, cmdPa) & ", "

                                    fieldList.Add(UCase(dictItem.Value.fieldName))
                                End If
                            End If
                        Next

                        For Each cdmItem In tableCdm.DataCtrlDict
                            If cdmItem.Value.saveFlag AndAlso Not fieldList.Contains(UCase(cdmItem.Value.fieldName)) Then
                                insSqlPart1 += cdmItem.Value.fieldName & ", "

                                insSqlPart2 += getSqlStr(dataDt.Rows(i).Item(cdmItem.Value.fieldName).ToString.Trim, cdmItem.Value.fieldType, cmdPa) & ", "

                                fieldList.Add(UCase(cdmItem.Value.fieldName))
                            End If
                        Next

                        If keyList.Trim <> "" Then
                            keyArray = gU.listToArray(keyList)

                            For j = 0 To UBound(keyArray)
                                If Not fieldList.Contains(UCase(keyArray(j))) Then
                                    If keyDict.ContainsKey(UCase(keyArray(j))) Then
                                        If keyDict.Item(UCase(keyArray(j))).fieldValue <> "" Then
                                            insSqlPart1 += keyArray(j) & ", "
                                            insSqlPart2 += getSqlStr(keyDict.Item(UCase(keyArray(j))).fieldValue, keyDict.Item(UCase(keyArray(j))).fieldType, cmdPa) & ", "
                                        Else
                                            insSqlPart1 += keyArray(j) & ", "
                                            insSqlPart2 += getSqlStr(dataDt.Rows(i).Item(keyArray(j)).ToString.Trim, keyDict.Item(UCase(keyArray(j))).fieldType, cmdPa) & ", "
                                        End If
                                    ElseIf dataDt.Columns.Contains(keyArray(j)) Then
                                        insSqlPart1 += keyArray(j) & ", "
                                        insSqlPart2 += getSqlStr(dataDt.Rows(i).Item(keyArray(j)).ToString.Trim, "", cmdPa) & ", "
                                    Else
                                        Throw New Exception("Key value not found: " & keyArray(j))
                                    End If
                                End If
                            Next
                        Else
                            For Each dictItem In keyDict
                                If Not fieldList.Contains(UCase(dictItem.Value.fieldName)) Then
                                    If dictItem.Value.fieldValue <> "" Then
                                        insSqlPart1 += dictItem.Value.fieldName & ", "
                                        insSqlPart2 += getSqlStr(dictItem.Value.fieldValue, dictItem.Value.fieldType, cmdPa) & ", "
                                    Else
                                        insSqlPart1 += dictItem.Value.fieldName & ", "
                                        insSqlPart2 += getSqlStr(dataDt.Rows(i).Item(dictItem.Value.fieldName).ToString.Trim, dictItem.Value.fieldType, cmdPa) & ", "
                                    End If
                                End If
                            Next
                        End If

                        If updtSysUserDate Then
                            If Not fieldList.Contains("SYS_CD") Then
                                insSqlPart1 += "sys_cd, "
                                insSqlPart2 += "Getdate(), "
                            End If

                            If Not fieldList.Contains("SYS_CB") Then
                                insSqlPart1 += "sys_cb, "
                                insSqlPart2 += cmdPa.AP(System.Web.HttpContext.Current.Session("userid")) & ", "
                            End If

                            If Not fieldList.Contains("SYS_UD") Then
                                insSqlPart1 += "sys_ud, "
                                insSqlPart2 += "Getdate(), "
                            End If

                            If Not fieldList.Contains("SYS_UB") Then
                                insSqlPart1 += "sys_ub, "
                                insSqlPart2 += cmdPa.AP(System.Web.HttpContext.Current.Session("userid")) & ", "
                            End If
                        End If

                        insSqlPart1 = Left(insSqlPart1, Len(insSqlPart1) - 2) & ") "

                        insSqlPart2 = Left(insSqlPart2, Len(insSqlPart2) - 2) & ") "

                        insertSql = insSqlPart1 & insSqlPart2

                        gDB.amendData(insertSql, cnn, transaction, cmdPa)

                        If nextNo > 0 Then
                            nextNo += 1
                        End If

                    Case "D"
                        'Delete records ---------------------------------------------------
                        cmdPa = New GlobalDBFunc.DBCmdPara
                        deleteSql = "delete from " & tableName & " " & _
                                    "where "

                        If keyList.Trim <> "" Then
                            keyArray = gU.listToArray(keyList)

                            For j = 0 To UBound(keyArray)
                                If j > 0 Then
                                    deleteSql += " and "
                                End If
                                If tableCdm.DataCtrlDict.ContainsKey(UCase(keyArray(j))) Then
                                    deleteSql += keyArray(j) & " = " & getSqlStr(dataDt.Rows(i).Item(keyArray(j)).ToString.Trim, tableCdm.DataCtrlDict.Item(keyArray(j)).fieldType, cmdPa)
                                ElseIf keyDict.ContainsKey(UCase(keyArray(j))) Then
                                    If keyDict.Item(UCase(keyArray(j))).fieldValue <> "" Then
                                        deleteSql += keyArray(j) & " = " & getSqlStr(keyDict.Item(UCase(keyArray(j))).fieldValue, keyDict.Item(UCase(keyArray(j))).fieldType, cmdPa)
                                    Else
                                        deleteSql += keyArray(j) & " = " & getSqlStr(dataDt.Rows(i).Item(keyArray(j)).ToString.Trim, keyDict.Item(UCase(keyArray(j))).fieldType, cmdPa)
                                    End If
                                ElseIf dataDt.Columns.Contains(keyArray(j)) Then
                                    deleteSql += keyArray(j) & " = " & getSqlStr(dataDt.Rows(i).Item(keyArray(j)).ToString.Trim, "", cmdPa)
                                Else
                                    Throw New Exception("Key value not found: " & keyArray(j))
                                End If
                            Next
                        Else
                            j = 0
                            For Each dictItem In keyDict
                                If j > 0 Then
                                    deleteSql += "and "
                                End If
                                If dictItem.Value.fieldValue <> "" Then
                                    deleteSql += dictItem.Value.fieldName & " = " & getSqlStr(dictItem.Value.fieldValue, dictItem.Value.fieldType, cmdPa)
                                Else
                                    deleteSql += dictItem.Value.fieldName & " = " & getSqlStr(dataDt.Rows(i).Item(dictItem.Value.fieldName).ToString.Trim, dictItem.Value.fieldType, cmdPa)
                                End If
                                j += 1
                            Next
                        End If

                        gDB.amendData(deleteSql, cnn, transaction, cmdPa)

                    Case Else
                        'Update records ---------------------------------------------------
                        If hasDirtyList AndAlso Not dirtyDict.ContainsKey(i) Then
                            Continue For
                        End If

                        fieldList = New List(Of String)

                        cmdPa = New GlobalDBFunc.DBCmdPara
                        updateSql = "update " & tableName & " " & _
                                    "set "

                        For Each dictItem In custUpdtDict
                            If Not dictItem.Value.insertOnly Then
                                If Not fieldList.Contains(UCase(dictItem.Value.fieldName)) Then

                                    If dictItem.Value.fieldType = "SQL" AndAlso InStr(dictItem.Value.fieldValue, custRefPrefix) > 0 AndAlso InStr(dictItem.Value.fieldValue, custRefSuffix) > 0 Then
                                        tmpCustValue = getCustRefValue(dictItem.Value.fieldValue, dataDt.Rows(i), cmdPa)
                                    Else
                                        tmpCustValue = dictItem.Value.fieldValue
                                    End If

                                    updateSql += dictItem.Value.fieldName & " = " & getSqlStr(tmpCustValue, dictItem.Value.fieldType, cmdPa) & ", "

                                    fieldList.Add(UCase(dictItem.Value.fieldName))
                                End If
                            End If
                        Next

                        For Each cdmItem In tableCdm.DataCtrlDict
                            If cdmItem.Value.saveFlag AndAlso (Not hasDirtyList OrElse dirtyDict.Item(i).Contains(UCase(cdmItem.Value.fieldName))) Then
                                If Not hasPrevDt OrElse isFieldChanged(UCase(cdmItem.Value.fieldName), dataDt.Rows(i).Item(cdmItem.Value.fieldName).ToString.Trim, prevDt.Rows(i)) Then
                                    If Not fieldList.Contains(UCase(cdmItem.Value.fieldName)) Then
                                        updateSql += cdmItem.Value.fieldName & " = " & getSqlStr(dataDt.Rows(i).Item(cdmItem.Value.fieldName).ToString.Trim, cdmItem.Value.fieldType, cmdPa) & ", "

                                        fieldList.Add(UCase(cdmItem.Value.fieldName))
                                    End If
                                End If
                            End If
                        Next

                        If fieldList.Count > 0 Then
                            If updtSysUserDate Then
                                If Not fieldList.Contains("SYS_UD") Then
                                    updateSql += "sys_ud = Getdate(), "
                                End If

                                If Not fieldList.Contains("SYS_UB") Then
                                    updateSql += "sys_ub = " & cmdPa.AP(System.Web.HttpContext.Current.Session("userid")) & ", "
                                End If
                            End If

                            updateSql = Left(updateSql, Len(updateSql) - 2)

                            updateSql += " where "

                            If keyList.Trim <> "" Then
                                keyArray = gU.listToArray(keyList)

                                For j = 0 To UBound(keyArray)
                                    If j > 0 Then
                                        updateSql += " and "
                                    End If
                                    If tableCdm.DataCtrlDict.ContainsKey(UCase(keyArray(j))) Then
                                        If dataDt.Columns.Contains("ORG_" & keyArray(j)) Then
                                            If dataDt.Rows(i).Item("ORG_" & keyArray(j)).ToString.Trim <> "" Then
                                                updateSql += keyArray(j) & " = " & getSqlStr(dataDt.Rows(i).Item("ORG_" & keyArray(j)).ToString.Trim, tableCdm.DataCtrlDict.Item(keyArray(j)).fieldType, cmdPa)
                                            Else
                                                updateSql += keyArray(j) & " = " & getSqlStr(dataDt.Rows(i).Item(keyArray(j)).ToString.Trim, tableCdm.DataCtrlDict.Item(keyArray(j)).fieldType, cmdPa)
                                            End If
                                        Else
                                            updateSql += keyArray(j) & " = " & getSqlStr(dataDt.Rows(i).Item(keyArray(j)).ToString.Trim, tableCdm.DataCtrlDict.Item(keyArray(j)).fieldType, cmdPa)
                                        End If

                                    ElseIf keyDict.ContainsKey(UCase(keyArray(j))) Then
                                        If keyDict.Item(UCase(keyArray(j))).fieldValue <> "" Then
                                            updateSql += keyArray(j) & " = " & getSqlStr(keyDict.Item(UCase(keyArray(j))).fieldValue, keyDict.Item(UCase(keyArray(j))).fieldType, cmdPa)
                                        Else
                                            updateSql += keyArray(j) & " = " & getSqlStr(dataDt.Rows(i).Item(keyArray(j)).ToString.Trim, keyDict.Item(UCase(keyArray(j))).fieldType, cmdPa)
                                        End If
                                    ElseIf dataDt.Columns.Contains(keyArray(j)) Then
                                        updateSql += keyArray(j) & " = " & getSqlStr(dataDt.Rows(i).Item(keyArray(j)).ToString.Trim, "", cmdPa)
                                    Else
                                        Throw New Exception("Key value not found: " & keyArray(j))
                                    End If
                                Next
                            Else
                                j = 0
                                For Each dictItem In keyDict
                                    If j > 0 Then
                                        updateSql += "and "
                                    End If
                                    If dictItem.Value.fieldValue <> "" Then
                                        updateSql += dictItem.Value.fieldName & " = " & getSqlStr(dictItem.Value.fieldValue, dictItem.Value.fieldType, cmdPa)
                                    Else
                                        updateSql += dictItem.Value.fieldName & " = " & getSqlStr(dataDt.Rows(i).Item(dictItem.Value.fieldName).ToString.Trim, dictItem.Value.fieldType, cmdPa)
                                    End If
                                    'updateSql += dictItem.Value.fieldName & " = " & getSqlStr(dictItem.Value.fieldValue, dictItem.Value.fieldType, cmdPa)
                                    j += 1
                                Next
                            End If

                            gDB.amendData(updateSql, cnn, transaction, cmdPa)
                        End If

                End Select
            Next
        Catch ex As Exception
            failReason = ex.Message
            Return False
        Finally
            If pConn Is Nothing Then
                cnn.Close()
                cnn.Dispose()
            End If
        End Try

        Return True
    End Function

    Public Function saveData(ByVal tableName As String, ByVal pageMode As String, Optional ByVal keyList As String = "") As Boolean
        Dim cnn As SqlConnection
        Dim transaction As SqlTransaction
        Dim cmdPa As GlobalDBFunc.DBCmdPara
        Dim insertSql, insSqlPart1, insSqlPart2, updateSql As String
        'Dim insertSql, insSqlPart1, insSqlPart2, updateSql, deleteSql As String
        Dim dictItem As KeyValuePair(Of String, custField)
        Dim cdmItem As KeyValuePair(Of String, CtrlDataMap.DataFieldCtrl)
        Dim fieldList As List(Of String)
        Dim keyArray As String()
        Dim i, j As Integer
        Dim formPrefix As String
        Dim dirtyDict As Dictionary(Of Integer, List(Of String))

        If formHolderID <> "" Then
            formPrefix = formHolderID & "_"
        Else
            formPrefix = ""
        End If

        failReason = ""

        If tableCdm Is Nothing Then
            failReason = "No data mapping!"
            Return False
        End If

        If tableName = "" Then
            failReason = "No table name!"
            Return False
        End If

        If keyList.Trim = "" AndAlso keyDict.Count = 0 Then
            failReason = "No matching key!"
            Return False
        End If

        If pConn Is Nothing Then
            cnn = gDB.getConnection()
        Else
            cnn = pConn
        End If

        transaction = pTrans

        Try
            Select Case UCase(TypeName(tableCdm.DataParentControl))
                Case "REPEATER"
                    If dataDt Is Nothing Then
                        failReason = "No data table!"
                        Return False
                    End If

                    Dim tmpRepeater As Repeater
                    'Dim ctrlPrefix As String

                    tmpRepeater = CType(tableCdm.DataParentControl, Repeater)

                    If tmpRepeater.Items.Count <= 0 Then
                        Return True
                    End If

                    dirtyDict = getDirtyDict()

                    Return saveDatatable(tableName, dirtyDict, keyList)

                Case "GRIDVIEW"
                    If dataDt Is Nothing Then
                        failReason = "No data table!"
                        Return False
                    End If

                    Dim tmpGridView As GridView
                    'Dim ctrlPrefix As String

                    tmpGridView = CType(tableCdm.DataParentControl, GridView)

                    If tmpGridView.Rows.Count <= 0 Then
                        Return True
                    End If

                    dirtyDict = getDirtyDict()

                    Return saveDatatable(tableName, dirtyDict, keyList)

                Case Else
                    'Handle form save ===========================================================================================
                    If pageMode = "N" Then
                        'Insert records -----------------------------------------------------------------------------------------
                        fieldList = New List(Of String)

                        cmdPa = New GlobalDBFunc.DBCmdPara
                        insSqlPart1 = "insert into " & tableName & " ("
                        insSqlPart2 = "values ("

                        For Each dictItem In custUpdtDict
                            If Not dictItem.Value.updateOnly Then
                                If Not fieldList.Contains(UCase(dictItem.Value.fieldName)) Then
                                    insSqlPart1 += dictItem.Value.fieldName & ", "

                                    insSqlPart2 += getSqlStr(dictItem.Value.fieldValue, dictItem.Value.fieldType, cmdPa) & ", "

                                    fieldList.Add(UCase(dictItem.Value.fieldName))
                                End If
                            End If
                        Next

                        For Each cdmItem In tableCdm.DataCtrlDict
                            If cdmItem.Value.saveFlag AndAlso Not fieldList.Contains(UCase(cdmItem.Value.fieldName)) Then
                                insSqlPart1 += cdmItem.Value.fieldName & ", "

                                insSqlPart2 += getSqlStr(tableCdm.getFormCtrlValue(cdmItem.Value.fieldName), cdmItem.Value.fieldType, cmdPa, cdmItem.Value.fieldFormat) & ", "

                                fieldList.Add(UCase(cdmItem.Value.fieldName))
                            End If
                        Next

                        If keyList.Trim <> "" Then
                            keyArray = gU.listToArray(keyList)

                            For i = 0 To UBound(keyArray)
                                If Not fieldList.Contains(UCase(keyArray(i))) Then
                                    If tableCdm.DataCtrlDict.ContainsKey(UCase(keyArray(i))) Then
                                        insSqlPart1 += keyArray(i) & ", "
                                        insSqlPart2 += getSqlStr(tableCdm.getFormCtrlValue(UCase(keyArray(i))), tableCdm.DataCtrlDict.Item(keyArray(i)).fieldType, cmdPa, tableCdm.DataCtrlDict.Item(keyArray(i)).fieldFormat) & ", "
                                    ElseIf keyDict.ContainsKey(UCase(keyArray(i))) Then
                                        insSqlPart1 += keyArray(i) & ", "
                                        insSqlPart2 += getSqlStr(keyDict.Item(UCase(keyArray(i))).fieldValue, keyDict.Item(UCase(keyArray(i))).fieldType, cmdPa) & ", "
                                    Else
                                        Throw New Exception("Key value not found: " & keyArray(i))
                                    End If
                                End If
                            Next
                        Else
                            For Each dictItem In keyDict
                                If Not fieldList.Contains(UCase(dictItem.Value.fieldName)) Then
                                    insSqlPart1 += dictItem.Value.fieldName & ", "
                                    insSqlPart2 += getSqlStr(dictItem.Value.fieldValue, dictItem.Value.fieldType, cmdPa) & ", "
                                End If
                            Next
                        End If

                        If updtSysUserDate Then
                            If Not fieldList.Contains("SYS_CD") Then
                                insSqlPart1 += "sys_cd, "
                                insSqlPart2 += "Getdate(), "
                            End If

                            If Not fieldList.Contains("SYS_CB") Then
                                insSqlPart1 += "sys_cb, "
                                insSqlPart2 += cmdPa.AP(System.Web.HttpContext.Current.Session("userid")) & ", "
                            End If

                            If Not fieldList.Contains("SYS_UD") Then
                                insSqlPart1 += "sys_ud, "
                                insSqlPart2 += "Getdate(), "
                            End If

                            If Not fieldList.Contains("SYS_UB") Then
                                insSqlPart1 += "sys_ub, "
                                insSqlPart2 += cmdPa.AP(System.Web.HttpContext.Current.Session("userid")) & ", "
                            End If
                        End If

                        insSqlPart1 = Left(insSqlPart1, Len(insSqlPart1) - 2) & ") "

                        insSqlPart2 = Left(insSqlPart2, Len(insSqlPart2) - 2) & ") "

                        insertSql = insSqlPart1 & insSqlPart2

                        gDB.amendData(insertSql, cnn, transaction, cmdPa)

                    Else
                        'Update records -----------------------------------------------------------------------------------------
                        If hasPrevDt AndAlso prevDt Is Nothing Then
                            failReason = "No old record to compare!"
                            Return False
                        End If

                        fieldList = New List(Of String)

                        cmdPa = New GlobalDBFunc.DBCmdPara
                        updateSql = "update " & tableName & " " & _
                                    "set "

                        For Each dictItem In custUpdtDict
                            If Not dictItem.Value.insertOnly Then
                                If Not fieldList.Contains(UCase(dictItem.Value.fieldName)) Then
                                    updateSql += dictItem.Value.fieldName & " = " & getSqlStr(dictItem.Value.fieldValue, dictItem.Value.fieldType, cmdPa) & ", "

                                    fieldList.Add(UCase(dictItem.Value.fieldName))
                                End If
                            End If
                        Next

                        For Each cdmItem In tableCdm.DataCtrlDict
                            If cdmItem.Value.saveFlag AndAlso (Not hasDirtyList OrElse gU.jsInList(UCase(dirtyList), UCase(formPrefix & gU.decodeNullOrEmpty(UCase(cdmItem.Value.fieldCtrlID), UCase(cdmItem.Value.fieldName))))) Then
                                If Not hasPrevDt OrElse isFieldChanged(UCase(cdmItem.Value.fieldName), tableCdm.getFormCtrlValue(cdmItem.Value.fieldName), prevDt.Rows(prevDtRowNo)) Then
                                    If Not fieldList.Contains(UCase(cdmItem.Value.fieldName)) Then
                                        updateSql += cdmItem.Value.fieldName & " = " & getSqlStr(tableCdm.getFormCtrlValue(cdmItem.Value.fieldName), cdmItem.Value.fieldType, cmdPa, cdmItem.Value.fieldFormat) & ", "

                                        fieldList.Add(UCase(cdmItem.Value.fieldName))
                                    End If
                                End If
                            End If
                        Next

                        If fieldList.Count > 0 Then
                            If updtSysUserDate Then
                                If Not fieldList.Contains("SYS_UD") Then
                                    updateSql += "sys_ud = Getdate(), "
                                End If

                                If Not fieldList.Contains("SYS_UB") Then
                                    updateSql += "sys_ub = " & cmdPa.AP(System.Web.HttpContext.Current.Session("userid")) & ", "
                                End If
                            End If

                            updateSql = Left(updateSql, Len(updateSql) - 2)

                            updateSql += " where "

                            If keyList.Trim <> "" Then
                                keyArray = gU.listToArray(keyList)

                                For i = 0 To UBound(keyArray)
                                    If i > 0 Then
                                        updateSql += " and "
                                    End If

                                    If tableCdm.DataCtrlDict.ContainsKey(UCase(keyArray(i))) Then
                                        If tableCdm.DataCtrlDict.ContainsKey(UCase("ORG_" & keyArray(i))) Then
                                            updateSql += keyArray(i) & " = " & getSqlStr(tableCdm.getFormCtrlValue(UCase("ORG_" & keyArray(i))), tableCdm.DataCtrlDict.Item("ORG_" & keyArray(i)).fieldType, cmdPa, tableCdm.DataCtrlDict.Item("ORG_" & keyArray(i)).fieldFormat)
                                        Else
                                            updateSql += keyArray(i) & " = " & getSqlStr(tableCdm.getFormCtrlValue(UCase(keyArray(i))), tableCdm.DataCtrlDict.Item(keyArray(i)).fieldType, cmdPa, tableCdm.DataCtrlDict.Item(keyArray(i)).fieldFormat)
                                        End If

                                    ElseIf keyDict.ContainsKey(UCase(keyArray(i))) Then
                                        updateSql += keyArray(i) & " = " & getSqlStr(keyDict.Item(UCase(keyArray(i))).fieldValue, keyDict.Item(UCase(keyArray(i))).fieldType, cmdPa)
                                    ElseIf hasPrevDt Then
                                        updateSql += keyArray(i) & " = " & getSqlStr(prevDt.Rows(prevDtRowNo).Item(keyArray(i)).ToString.Trim, "", cmdPa)
                                    Else
                                        Throw New Exception("Key value not found: " & keyArray(i))
                                    End If
                                Next
                            Else
                                i = 0
                                For Each dictItem In keyDict
                                    If i > 0 Then
                                        updateSql += "and "
                                    End If
                                    updateSql += dictItem.Value.fieldName & " = " & getSqlStr(dictItem.Value.fieldValue, dictItem.Value.fieldType, cmdPa)
                                    i += 1
                                Next
                            End If

                            gDB.amendData(updateSql, cnn, transaction, cmdPa)
                        End If
                    End If
            End Select

        Catch ex As Exception
            failReason = ex.Message
            Return False
        Finally
            If pConn Is Nothing Then
                cnn.Close()
                cnn.Dispose()
            End If
        End Try

        Return True
    End Function

    Public Sub addCustUpdt(ByVal fieldName As String, ByVal fieldValue As String, Optional fieldType As String = "", Optional insertOnly As Boolean = False, Optional updateOnly As Boolean = False)
        Dim tmpCustField As custField

        If Not custUpdtDict.ContainsKey(UCase(fieldName)) Then
            tmpCustField = New custField
            tmpCustField.fieldName = UCase(fieldName)
            tmpCustField.fieldValue = fieldValue
            tmpCustField.fieldType = fieldType
            tmpCustField.insertOnly = insertOnly
            tmpCustField.updateOnly = updateOnly

            custUpdtDict.Add(UCase(fieldName), tmpCustField)
        Else
            Throw New Exception("Customerize field already exists: " & fieldName)
        End If
    End Sub

    Public Sub addKey(ByVal fieldName As String, ByVal fieldValue As String, Optional fieldType As String = "")
        Dim tmpCustField As custField

        If Not keyDict.ContainsKey(UCase(fieldName)) Then
            tmpCustField = New custField
            tmpCustField.fieldName = UCase(fieldName)
            tmpCustField.fieldValue = fieldValue
            tmpCustField.fieldType = fieldType
            tmpCustField.insertOnly = False
            tmpCustField.updateOnly = False

            keyDict.Add(UCase(fieldName), tmpCustField)
        Else
            Throw New Exception("Key already exists: " & fieldName)
        End If
    End Sub

    Public Function isFieldChanged(ByVal fieldName As String, ByVal fieldValue As String, ByRef fieldRow As DataRow, Optional ByVal fieldType As String = "") As Boolean
        Dim newValue, oldValue As String

        'newValue = fieldValue.Trim
        newValue = fieldValue

        'oldValue = fieldRow.Item(fieldName).ToString.Trim
        oldValue = fieldRow.Item(fieldName).ToString

        Select Case (fieldType)
            Case "COMMA_NUM"
                If newValue = "" And oldValue = "" Then
                    Return False
                ElseIf newValue = "" OrElse oldValue = "" Then
                    Return True
                Else
                    newValue = Replace(newValue, ",", "")
                    oldValue = Replace(oldValue, ",", "")

                    If CDbl(newValue) = CDbl(oldValue) Then
                        Return False
                    Else
                        Return True
                    End If
                End If

            Case Else
                If newValue = oldValue Then
                    Return False
                Else
                    Return True
                End If
        End Select
    End Function

    Private Function getSqlStr(ByVal fieldValue As String, ByVal fieldType As String, ByRef cmdPa As GlobalDBFunc.DBCmdPara, _
                               Optional numFormat As GeneralUtils.NumberFormat = Nothing) As String
        Select Case fieldType
            Case "DATE"
                Return "to_date(" & cmdPa.AP(fieldValue.Trim) & ", '" & DDFORMAT & "')"

            Case "NUMERIC"
                If fieldValue = "" Then
                    Return "null"
                Else
                    If numFormat IsNot Nothing Then
                        Return cmdPa.AP(gU.formatNumRev(fieldValue, numFormat), SqlDbType.Decimal)
                    Else
                        Return cmdPa.AP(Replace(fieldValue, ",", ""), SqlDbType.Decimal)
                    End If
                End If

            Case "SQL"
                Return fieldValue

            Case "SEQ_NO"
                Return cmdPa.AP(nextNo)
                'nextNo += 1
                'Return cmdPa.AP(nextNo - 1)

            Case Else
                If fieldValue = "" Then
                    Return "null"
                Else
                    Return cmdPa.AP(fieldValue.Trim)
                End If

        End Select
    End Function

    Public Property DataDirtyList() As String
        Get
            Return dirtyList
        End Get
        Set(value As String)
            dirtyList = value
        End Set
    End Property

    Public Property OldDataTable() As DataTable
        Get
            Return prevDt
        End Get
        Set(value As DataTable)
            prevDt = value
        End Set
    End Property

    Public Property UpdtDataTable() As DataTable
        Get
            Return dataDt
        End Get
        Set(value As DataTable)
            dataDt = value
        End Set
    End Property

    Public Property OldDataRowNo() As Integer
        Get
            Return prevDtRowNo
        End Get
        Set(value As Integer)
            prevDtRowNo = value
        End Set
    End Property

    Public Property CtrlMapping() As CtrlDataMap
        Get
            Return tableCdm
        End Get
        Set(value As CtrlDataMap)
            tableCdm = value
        End Set
    End Property

    Public ReadOnly Property ReasonDesc() As String
        Get
            Return failReason
        End Get
    End Property

    Public Property NextSeqNo() As Integer
        Get
            Return nextNo
        End Get
        Set(value As Integer)
            nextNo = value
        End Set
    End Property

    Public Property UseDirtyList() As Boolean
        Get
            Return hasDirtyList
        End Get
        Set(value As Boolean)
            hasDirtyList = value
        End Set
    End Property

    Public Property CompareOldData() As Boolean
        Get
            Return hasPrevDt
        End Get
        Set(value As Boolean)
            hasPrevDt = value
        End Set
    End Property

    Public Property ContentPlaceHolderID() As String
        Get
            Return formHolderID
        End Get
        Set(value As String)
            formHolderID = value
        End Set
    End Property

    Public Property UpdateSysUserDate() As Boolean
        Get
            Return updtSysUserDate
        End Get
        Set(value As Boolean)
            updtSysUserDate = value
        End Set
    End Property

    Private Structure custField
        Dim fieldName As String
        Dim fieldValue As String
        Dim fieldType As String
        Dim insertOnly As Boolean
        Dim updateOnly As Boolean
    End Structure
End Class
