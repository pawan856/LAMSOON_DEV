Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class GlobalDBFunc
    Private Shared gU As New GeneralUtils

    Public Function getConnection() As SqlConnection
        Dim l_conn As SqlConnection
        Dim con As String = ""

        If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("l_conn") IsNot Nothing AndAlso HttpContext.Current.Session("l_conn").ToString = "Archived" Then
            con = "ConnectionStringArch"
        Else
            con = "ConnectionString"
        End If

        l_conn = New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(con).ConnectionString)
        l_conn.Open()
        Return l_conn

    End Function

    Public Function getDataSet(ByVal as_sql As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByRef pTransaction As SqlTransaction = Nothing) As DataSet
        'Retrieve data into a DataSet
        'return DataSet

        Dim lds_1 As New DataSet
        Dim lda_1 As SqlDataAdapter
        Dim l_conn As SqlConnection
        Dim l_SqlCmd As SqlCommand

        If pConn Is Nothing Then
            l_conn = getConnection()
        Else
            l_conn = pConn
        End If
        'as_sql = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; " & as_sql
        l_SqlCmd = New SqlCommand(as_sql, l_conn)
        l_SqlCmd.CommandTimeout = System.Configuration.ConfigurationManager.AppSettings.Item("cmdTimeOut")
        If Not pTransaction Is Nothing Then
            l_SqlCmd.Transaction = pTransaction
        End If

        lda_1 = New SqlDataAdapter(l_SqlCmd)

        lda_1.Fill(lds_1)

        If pConn Is Nothing Then
            l_conn.Close()
            l_conn.Dispose()
        End If

        Return lds_1

    End Function

    Public Function getDataSetProc(ByVal as_sql As String, Optional ByRef pConn As SqlConnection = Nothing, Optional ByRef pTransaction As SqlTransaction = Nothing) As DataSet
        'Retrieve data into a DataSet
        'return DataSet

        Dim lds_1 As New DataSet
        Dim lda_1 As SqlDataAdapter
        Dim l_conn As SqlConnection
        Dim l_SqlCmd As SqlCommand

        If pConn Is Nothing Then
            l_conn = getConnection()
        Else
            l_conn = pConn
        End If

        l_SqlCmd = New SqlCommand(as_sql, l_conn)
        l_SqlCmd.CommandTimeout = System.Configuration.ConfigurationManager.AppSettings.Item("cmdTimeOut")
        If Not pTransaction Is Nothing Then
            l_SqlCmd.Transaction = pTransaction
        End If

        lda_1 = New SqlDataAdapter(l_SqlCmd)

        lda_1.Fill(lds_1)

        If pConn Is Nothing Then
            l_conn.Close()
            l_conn.Dispose()
        End If

        Return lds_1

    End Function

    Public Function getValueFromSQLByDR(ByVal SQLString As String, Optional ByRef pConn As SqlConnection = Nothing, _
                                        Optional ByRef pTransaction As SqlTransaction = Nothing, _
                                        Optional ByRef cmdPara As GlobalDBFunc.DBCmdPara = Nothing, _
                                        Optional ByRef pLog As PrgmLog = Nothing, _
                                        Optional ByVal callFromWCF As Boolean = False) As String

        Dim oCmd As SqlCommand
        Dim l_conn As SqlConnection
        Dim drExp As SqlDataReader = Nothing
        Dim tmpValue As String = ""
        Dim msgLog As PrgmLog = Nothing

        If pConn Is Nothing Then
            l_conn = getConnection()
        Else
            l_conn = pConn
        End If

        oCmd = New SqlCommand(SQLString, l_conn)

        If cmdPara IsNot Nothing Then
            setCmdWithPara(oCmd, cmdPara)
        End If

        Try
            drExp = oCmd.ExecuteReader

            If drExp.Read Then
                tmpValue = drExp.Item(0).ToString
            End If

        Catch ex As Exception
            If pLog IsNot Nothing Then
                msgLog = pLog
            Else
                If callFromWCF Then
                    If System.IO.Directory.Exists(gU.getConfig("SYSP_LOG_DIR")) Then
                        msgLog = New PrgmLog(gU.getConfig("SYSP_LOG_DIR"), "AppLog" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")
                    End If
                Else
                    If System.IO.Directory.Exists(gU.getConfig("SYSP_LOG_DIR")) Then
                        msgLog = New PrgmLog(gU.getConfig("SYSP_LOG_DIR"), "AppLog" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")
                    End If
                End If
            End If

            If msgLog IsNot Nothing Then
                If cmdPara IsNot Nothing Then
                    msgLog.writeLog("[" & ex.Message & "] [SQL] : " & getCmdSql(oCmd.CommandText, cmdPara))
                Else
                    msgLog.writeLog("[" & ex.Message & "] [SQL] : " & oCmd.CommandText)
                End If
            End If

            Throw ex
        Finally
            If drExp IsNot Nothing Then
                drExp.Dispose()
                drExp = Nothing
            End If

            oCmd.Dispose()

            If pConn Is Nothing Then
                l_conn.Close()
                l_conn.Dispose()
            End If
        End Try

        Return tmpValue
    End Function

    Public Sub setCmdWithPara(ByRef oCmd As SqlCommand, ByRef cmdPara As GlobalDBFunc.DBCmdPara)
        Dim i As Integer

        'oCmd.BindByName = cmdPara.BindByName

        For i = 0 To cmdPara.paraList.Count - 1
            oCmd.Parameters.Add(cmdPara.paraList(i).paraID, cmdPara.paraList(i).paraType).Value = cmdPara.paraList(i).paraDBValue
        Next
    End Sub

    Public Function getDataTable(ByVal as_sql As String, Optional ByRef pConn As SqlConnection = Nothing,
                                 Optional ByRef pTransaction As SqlTransaction = Nothing,
                                 Optional ByRef TableSchema As DataTable = Nothing,
                                 Optional ByRef cmdPara As DBCmdPara = Nothing,
                                 Optional ByRef pLog As PrgmLog = Nothing,
                                 Optional ByVal callFromWCF As Boolean = False,
                                 Optional ByVal cmdTimeOut As Integer = 240) As DataTable
        cmdTimeOut = System.Configuration.ConfigurationManager.AppSettings.Item("cmdTimeOut")

        Dim lds_1 As New DataSet
        Dim lda_1 As SqlDataAdapter
        Dim l_conn As SqlConnection
        Dim l_SqlCmd As SqlCommand
        Dim i As Integer
        Dim msgLog As PrgmLog = Nothing

        If pConn Is Nothing Then
            l_conn = getConnection()
        Else
            l_conn = pConn
        End If
        'as_sql = "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; " & as_sql
        l_SqlCmd = New SqlCommand(as_sql, l_conn)

        If cmdPara IsNot Nothing Then
            'l_SqlCmd.BindByName = cmdPara.BindByName

            For i = 0 To cmdPara.paraList.Count - 1
                l_SqlCmd.Parameters.Add(cmdPara.paraList(i).paraID, cmdPara.paraList(i).paraType).Value = cmdPara.paraList(i).paraDBValue
            Next
        End If

        If cmdTimeOut > 0 Then
            l_SqlCmd.CommandTimeout = cmdTimeOut
        End If
        If Not pTransaction Is Nothing Then
            l_SqlCmd.Transaction = pTransaction
        End If

        lda_1 = New SqlDataAdapter(l_SqlCmd)

        Try
            If Not TableSchema Is Nothing Then lda_1.FillSchema(TableSchema, SchemaType.Source)

            lda_1.Fill(lds_1)
        Catch ex As Exception
            If pLog IsNot Nothing Then
                msgLog = pLog
            Else
                If callFromWCF Then
                    If System.IO.Directory.Exists(gU.getConfig("SYSP_LOG_DIR")) Then
                        msgLog = New PrgmLog(gU.getConfig("SYSP_LOG_DIR"), "AppLog" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")
                    End If
                Else
                    If System.IO.Directory.Exists(gU.getConfig("SYSP_LOG_DIR")) Then
                        msgLog = New PrgmLog(gU.getConfig("SYSP_LOG_DIR"), "AppLog" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")
                    End If
                End If
            End If

            If msgLog IsNot Nothing Then
                If cmdPara IsNot Nothing Then
                    msgLog.writeLog("[" & ex.Message & "] [SQL] : " & getCmdSql(l_SqlCmd.CommandText, cmdPara))
                Else
                    msgLog.writeLog("[" & ex.Message & "] [SQL] : " & l_SqlCmd.CommandText)
                End If
            End If

            Throw ex
        Finally
            l_SqlCmd.Dispose()

            If pConn Is Nothing Then
                l_conn.Close()
                l_conn.Dispose()
            End If
        End Try

        Return lds_1.Tables(0)

    End Function

    Public Function amendData(ByVal sql As String, _
                         Optional ByRef pConn As SqlConnection = Nothing, _
                         Optional ByRef pTransaction As SqlTransaction = Nothing, _
                         Optional ByRef cmdPara As DBCmdPara = Nothing, _
                         Optional ByRef pLog As PrgmLog = Nothing, _
                         Optional ByVal callFromWCF As Boolean = False) As Long

        Dim affectedRow As Integer
        Dim l_conn As SqlConnection
        Dim i As Integer
        Dim msgLog As PrgmLog = Nothing

        If pConn Is Nothing Then
            l_conn = getConnection()
        Else
            l_conn = pConn
        End If

        Dim sqlCmd As SqlCommand = New SqlCommand(sql, l_conn)

        If cmdPara IsNot Nothing Then
            'sqlCmd.BindByName = cmdPara.BindByName

            For i = 0 To cmdPara.paraList.Count - 1
                sqlCmd.Parameters.Add(cmdPara.paraList(i).paraID, cmdPara.paraList(i).paraType).Value = cmdPara.paraList(i).paraDBValue
            Next
        End If

        sqlCmd.CommandTimeout = 120

        If Not pTransaction Is Nothing Then
            sqlCmd.Transaction = pTransaction
        End If

        Try
            affectedRow = sqlCmd.ExecuteNonQuery()
        Catch ex As Exception
            If pLog IsNot Nothing Then
                msgLog = pLog
            Else
                If callFromWCF Then
                    If System.IO.Directory.Exists(gU.getConfig("SYSP_LOG_DIR")) Then
                        msgLog = New PrgmLog(gU.getConfig("SYSP_LOG_DIR"), "AppLog" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")
                    End If
                Else
                    If System.IO.Directory.Exists(gU.getConfig("SYSP_LOG_DIR")) Then
                        msgLog = New PrgmLog(gU.getConfig("SYSP_LOG_DIR"), "AppLog" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")
                    End If
                End If
            End If

            If msgLog IsNot Nothing Then
                If cmdPara IsNot Nothing Then
                    msgLog.writeLog("[" & ex.Message & "] [SQL] : " & getCmdSql(sqlCmd.CommandText, cmdPara))
                Else
                    msgLog.writeLog("[" & ex.Message & "] [SQL] : " & sqlCmd.CommandText)
                End If
            End If

            Throw ex

        Finally
            sqlCmd.Dispose()

            If pConn Is Nothing Then
                l_conn.Close()
                l_conn.Dispose()
            End If
        End Try

        Return affectedRow
    End Function

    Public Function amendData2(ByVal sql As String, _
                               Optional ByRef pConn As SqlConnection = Nothing, _
                               Optional ByRef pTransaction As SqlTransaction = Nothing, _
                               Optional ByRef cmdPara As DBCmdPara = Nothing, _
                               Optional ByRef pLog As PrgmLog = Nothing, _
                               Optional ByVal callFromWCF As Boolean = False) As Integer

        Return amendData(sql, pConn, pTransaction, cmdPara, pLog, callFromWCF)

        'Dim affectedRow As Integer
        'Dim l_conn As SqlConnection
        'Dim i As Integer
        'Dim msgLog As PrgmLog = Nothing

        'If pConn Is Nothing Then
        '    l_conn = getConnection()
        'Else
        '    l_conn = pConn
        'End If

        'Dim sqlCmd As SqlCommand = New SqlCommand(sql, l_conn)

        'If cmdPara IsNot Nothing Then
        '    sqlCmd.BindByName = cmdPara.BindByName

        '    For i = 0 To cmdPara.paraList.Count - 1
        '        sqlCmd.Parameters.Add(cmdPara.paraList(i).paraID, cmdPara.paraList(i).paraType).Value = cmdPara.paraList(i).paraDBValue
        '    Next
        'End If

        'sqlCmd.CommandTimeout = 120

        'If Not pTransaction Is Nothing Then
        '    'sqlCmd.Transaction = pTransaction
        'End If

        'Try
        '    affectedRow = sqlCmd.ExecuteNonQuery()
        'Catch ex As Exception
        '    If pLog IsNot Nothing Then
        '        msgLog = pLog
        '    Else
        '        If callFromWCF Then
        '            If System.IO.Directory.Exists(gU.getConfig("SYSP_LOG_DIR")) Then
        '                msgLog = New PrgmLog(gU.getConfig("SYSP_LOG_DIR"), "AppLog" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")
        '            End If
        '        Else
        '            If System.IO.Directory.Exists(gU.getConfig("SYSP_LOG_DIR")) Then
        '                msgLog = New PrgmLog(gU.getConfig("SYSP_LOG_DIR"), "AppLog" & Now.Year & Right("0" & Now.Month, 2) & Right("0" & Now.Day, 2) & ".txt")
        '            End If
        '        End If
        '    End If

        '    If msgLog IsNot Nothing Then
        '        If cmdPara IsNot Nothing Then
        '            msgLog.writeLog("[" & ex.Message & "] [SQL] : " & getCmdSql(sqlCmd.CommandText, cmdPara))
        '        Else
        '            msgLog.writeLog("[" & ex.Message & "] [SQL] : " & sqlCmd.CommandText)
        '        End If
        '    End If

        '    Throw ex

        'Finally
        '    sqlCmd.Dispose()

        '    If pConn Is Nothing Then
        '        l_conn.Close()
        '        l_conn.Dispose()
        '    End If
        'End Try

        'Return affectedRow
    End Function

    Public Function getCmdSql(ByVal sqlString As String, ByRef cmdPara As DBCmdPara) As String
        Dim i As Integer
        Dim tmpParaList As New List(Of CmdParaStruct)(cmdPara.paraList)
        Dim commandSql As String = sqlString

        Do While tmpParaList.Count > 0
            i = 0

            Do While paraLikeInList(tmpParaList, tmpParaList(i).paraID)
                i = i + 1
            Loop

            If tmpParaList(i).paraType = sqlDbType.Decimal Then
                commandSql = Replace(commandSql, tmpParaList(i).paraID, gU.dbEncode(tmpParaList(i).paraStrValue))
            Else
                commandSql = Replace(commandSql, tmpParaList(i).paraID, "'" & gU.dbEncode(tmpParaList(i).paraValue) & "'")
            End If

            tmpParaList.RemoveAt(i)
        Loop

        Return commandSql
    End Function

    Private Function paraLikeInList(ByRef paraList As List(Of CmdParaStruct), ByVal paraID As String) As Boolean
        Dim i As Integer

        For i = 0 To paraList.Count - 1
            If paraList(i).paraID <> paraID AndAlso InStr(paraList(i).paraID, paraID) > 0 Then
                Return True
            End If
        Next

        Return False
    End Function


    'Class for Command Parameters defination
    Public Class DBCmdPara
        Public paraList As New List(Of CmdParaStruct)
        Public BindByName As Boolean = True
        Private paraIDList As New List(Of String)
        Private paraCnt As Integer = 0

        Public Sub Add(ByVal paraID As String, ByVal paratype As sqlDbType, ByVal paraValue As String)
            Dim newCmdParaStruct = New CmdParaStruct

            If paraID.Trim = "" Then
                Throw New Exception("Parameter ID cannot be empty.")
            End If

            If Exists(paraID) Then
                Throw New Exception("Parameter already existed.")
            End If

            If Left(paraID, 1) = "@" Then
                newCmdParaStruct.paraID = LCase(paraID)
            Else
                newCmdParaStruct.paraID = "@" & LCase(paraID)
            End If

            paraIDList.Add(newCmdParaStruct.paraID)

            newCmdParaStruct.paraType = paratype

            newCmdParaStruct.paraValue = paraValue

            paraList.Add(newCmdParaStruct)
        End Sub

        Public Sub AddNew(ByVal paraID As String, ByVal paratype As SqlDbType, ByVal paraValue As String)
            If Not Exists(paraID) Then
                Add(paraID, paratype, paraValue)
            End If
        End Sub

        Public Function AddPara(ByVal paraValue As Byte(), Optional ByVal paratype As SqlDbType = SqlDbType.VarChar) As String
            Dim newCmdParaStruct = New CmdParaStruct

            paraCnt = paraCnt + 1

            newCmdParaStruct.paraID = "@para_" & Right("000" & paraCnt, 4)

            paraIDList.Add(newCmdParaStruct.paraID)
            newCmdParaStruct.paraType = paratype
            newCmdParaStruct.paraByte = paraValue

            paraList.Add(newCmdParaStruct)

            Return newCmdParaStruct.paraID
        End Function

        Public Function AddPara(ByVal paraValue As DBNull, Optional ByVal paratype As SqlDbType = SqlDbType.VarChar) As String
            Dim newCmdParaStruct = New CmdParaStruct

            paraCnt = paraCnt + 1

            newCmdParaStruct.paraID = "@para_" & Right("000" & paraCnt, 4)

            paraIDList.Add(newCmdParaStruct.paraID)
            newCmdParaStruct.paraType = paratype
            newCmdParaStruct.paraNull = paraValue

            paraList.Add(newCmdParaStruct)

            Return newCmdParaStruct.paraID
        End Function

        Public Function AddPara(ByVal paraValue As String, Optional ByVal paratype As SqlDbType = SqlDbType.VarChar) As String
            Dim newCmdParaStruct = New CmdParaStruct

            paraCnt = paraCnt + 1

            newCmdParaStruct.paraID = "@para_" & Right("000" & paraCnt, 4)

            paraIDList.Add(newCmdParaStruct.paraID)
            newCmdParaStruct.paraType = paratype
            newCmdParaStruct.paraValue = paraValue

            paraList.Add(newCmdParaStruct)

            Return newCmdParaStruct.paraID
        End Function

        Public Function AddList(ByVal paraList As String) As String
            Dim tmpList As String
            Dim tmpArray As String()
            Dim i As Integer
            Dim sqlList As String = ""

            tmpList = gU.formatList(paraList)

            tmpArray = Split(tmpList, ", ")

            For i = 0 To UBound(tmpArray)
                If i > 0 Then
                    sqlList = sqlList & ", "
                End If

                sqlList = sqlList & AddPara(tmpArray(i))
            Next

            Return sqlList
        End Function

        Public Function AP(ByVal paraValue As Byte(), Optional ByVal paratype As SqlDbType = SqlDbType.VarBinary) As String
            Return AddPara(paraValue, paratype)
        End Function

        Public Function AP(ByVal paraValue As DBNull, Optional ByVal paratype As SqlDbType = SqlDbType.VarChar) As String
            Return AddPara(paraValue, paratype)
        End Function

        Public Function AP(ByVal paraValue As String, Optional ByVal paratype As SqlDbType = SqlDbType.VarChar) As String
            Return AddPara(paraValue, paratype)
        End Function

        Public Function Exists(ByVal paraID As String) As Boolean
            If Left(paraID, 1) = "@" Then
                Return paraIDList.Contains(LCase(paraID))
            Else
                Return paraIDList.Contains("@" & LCase(paraID))
            End If
        End Function
    End Class

    'Class for storing parameter
    Public Class CmdParaStruct
        Public paraID, paraValue As String
        Public paraByte() As Byte
        Public paraType As SqlDbType
        Public paraNull As DBNull
        Public setNullForNum As Boolean = True

        Public ReadOnly Property paraDBValue() As Object
            Get
                If setNullForNum AndAlso paraType = SqlDbType.Decimal AndAlso paraValue.Trim = "" Then
                    Return DBNull.Value
                ElseIf paraType = SqlDbType.VarBinary Then
                    If paraByte Is Nothing Then
                        Return DBNull.Value
                    Else
                        Return paraByte
                    End If
                Else
                    If paraValue Is Nothing OrElse paraValue = "" Then
                        Return DBNull.Value
                    Else
                        Return paraValue
                    End If
                End If
            End Get
        End Property

        Public ReadOnly Property paraStrValue() As Object
            Get
                If setNullForNum AndAlso paraType = SqlDbType.Decimal AndAlso paraValue.Trim = "" Then
                    Return "null"
                ElseIf paraType = SqlDbType.VarBinary Then
                    If paraByte Is Nothing Then
                        Return "null"
                    Else
                        Return paraByte
                    End If
                Else
                    Return paraValue
                End If
            End Get
        End Property
    End Class

    Private Function decodeDBNull(ByVal aValue As Object, ByVal defaultVal As String) As String
        If IsDBNull(aValue) Then
            decodeDBNull = defaultVal
        ElseIf CStr(aValue) = "#|MAP_NULL|#" Then
            decodeDBNull = defaultVal
        Else
            decodeDBNull = CStr(aValue)
        End If
    End Function

    Public Function getColValue(ByVal colc_code As String, ByVal cold_tabcol As String,
                                Optional ByRef pConn As SqlConnection = Nothing, _
                                Optional ByRef pTransaction As SqlTransaction = Nothing, Optional gLang As String = "E") As String
        Dim selectSql As String
        Dim cmdPa As New GlobalDBFunc.DBCmdPara
        Dim tempdt As DataTable
        Dim returnString As String = ""

        selectSql = "select COLC_ENG_VALUE, COLC_CHI_VALUE " & _
                    "from wms_col_code " & _
                    "where COLC_CODE = " & cmdPa.AP(colc_code) & " " & _
                    "and COLC_TABCOL = " & cmdPa.AP(cold_tabcol) & " "

        tempdt = getDataTable(selectSql, pConn, pTransaction, , cmdPa)

        If tempdt.Rows.Count > 0 Then
            If gLang = "E" Then
                returnString = tempdt.Rows(0).Item("COLC_ENG_VALUE").ToString.Trim
            Else
                returnString = tempdt.Rows(0).Item("COLC_CHI_VALUE").ToString.Trim
            End If
        End If

        Return returnString
    End Function
End Class
