Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.OleDb
Imports System.Web
Imports System.Data.SqlClient
Imports System.Net

Public Class securityUtil
    Public fvFieldList As String
    Private sender As Page
    Private dt As New GeneralUtils
    Private gDB As New GlobalDBFunc

    Public Sub setSenderPage(ByVal s As Page)
        sender = New Page
        sender = s
    End Sub

    Public Function getPostedData(ByVal name As String, Optional ByVal defaultValue As String = "", Optional ByVal isRequestForm As Boolean = False) As String
        Dim temp As String

        If isRequestForm Then
            temp = Convert.ToString(sender.Request.Form(Trim(name)))
        Else
            temp = Convert.ToString(sender.Request(Trim(name)))
        End If

        If Trim(temp) = "" Then
            getPostedData = defaultValue
        Else
            getPostedData = temp
        End If
    End Function

    Public Function getPostedData_int(ByVal postName As String, Optional ByVal defaultValue As Integer = 0, Optional ByVal isRequestForm As Boolean = False) As Integer
        getPostedData_int = getPostedDataChk_int(postName, defaultValue, False, "", isRequestForm)
    End Function

    Private Function getPostedDataAlert_int(ByVal postName As String, Optional ByVal defaultValue As Integer = 0, Optional ByVal isRequestForm As Boolean = False) As Integer
        getPostedDataAlert_int = getPostedDataChk_int(postName, defaultValue, True, "", isRequestForm)
    End Function

    Private Function getPostedDataChk_int(ByVal postName As String, ByVal defaultValue As Integer, ByVal isAlert As Boolean, ByVal chkValue As String, ByVal isRequestForm As Boolean) As Integer
        Dim temp, name As String
        Dim maxLength As Integer
        Dim nameArray As String()

        If InStr(postName, ":") > 0 Then
            nameArray = postName.Split(":")
            name = nameArray(0)
            maxLength = Convert.ToInt64(nameArray(1))
        Else
            name = postName
            maxLength = 0
        End If

        If chkValue <> "" Then
            temp = chkValue
        Else
            If isRequestForm Then
                temp = Convert.ToString(sender.Request.Form(Trim(name)))
            Else
                temp = Convert.ToString(sender.Request(Trim(name)))

            End If
        End If

        If trim(temp) = "" Then
            fvFieldList = appendToList(fvFieldList, UCase(Trim(name)))
            getPostedDataChk_int = defaultValue
        Else
            If IsNumeric(Trim(temp)) Then
                If maxLength <> 0 And Len(Trim(temp)) > maxLength Then
                    If isAlert Then
                        sender.Response.Write("<scr" & "ipt>alert('" & jsString(name) & " has exceeded the maxlength of " & jsString(maxLength) & "! - " & jsString(temp) & "');</scr" & "ipt>")
                    Else
                        sender.Response.Write(sender.Server.HtmlEncode(name) & " has exceeded the maxlength of " & maxLength & "! - " & sender.Server.HtmlEncode(temp))
                    End If
                    sender.Response.End()
                Else
                    fvFieldList = appendToList(fvFieldList, UCase(Trim(name)))
                    getPostedDataChk_int = temp
                End If
            Else
                If isAlert Then
                    sender.Response.Write("<scr" & "ipt>alert('" & jsString(name) & " - " & jsString(temp) & " : Incorrect data type! Only integer is allowed.');</scr" & "ipt>")
                Else
                    sender.Response.Write(sender.Server.HtmlEncode(name) & " - " & sender.Server.HtmlEncode(temp) & " : Incorrect data type! Only integer is allowed.")
                End If
                sender.Response.End()
                'getPostedDataChk_int = temp
            End If
        End If
    End Function

    Public Function getPostedData_float(ByVal postName As String, Optional ByVal defaultValue As Double = 0, Optional ByVal isRequestForm As Boolean = False) As Double
        getPostedData_float = getPostedDataChk_float(postName, defaultValue, False, "", isRequestForm)
    End Function

    Private Function getPostedDataAlert_float(ByVal postName As String, Optional ByVal defaultValue As Double = 0, Optional ByVal isRequestForm As Boolean = False) As Double
        getPostedDataAlert_float = getPostedDataChk_float(postName, defaultValue, True, "", isRequestForm)
    End Function

    Private Function getPostedDataChk_float(ByVal postName As String, ByVal defaultValue As Double, ByVal isAlert As Boolean, ByVal chkValue As String, ByVal isRequestForm As Boolean) As Double
        Dim temp, name As String
        Dim maxLength, digitLen As Integer
        Dim nameArray As String()
        Dim lenArray As String()
        Dim tempArray As String()

        If InStr(postName, ":") > 0 Then
            nameArray = Split(postName, ":")
            name = nameArray(0)

            If InStr(nameArray(1), ".") > 0 Then
                lenArray = Split(nameArray(1), ".")
                maxLength = Convert.ToInt64(lenArray(0)) - Convert.ToInt64(lenArray(1))
            Else
                maxLength = Convert.ToInt64(nameArray(1))
            End If
        Else
            name = postName
            maxLength = 0
        End If

        If chkValue <> "" Then
            temp = chkValue
        Else
            If isRequestForm Then
                temp = Convert.ToString(sender.Request.Form(Trim(name)))
            Else
                temp = Convert.ToString(sender.Request(Trim(name)))
            End If
        End If

        If Trim(temp) = "" Then
            fvFieldList = appendToList(fvFieldList, UCase(Trim(name)))
            getPostedDataChk_float = defaultValue
        Else
            If IsNumeric(Trim(temp)) Then
                If InStr(Trim(temp), ".") > 0 Then
                    tempArray = Trim(temp).Split(".")
                    digitLen = tempArray(0).Length
                Else
                    digitLen = Len(Trim(temp))
                End If

                If maxLength <> 0 And digitLen > maxLength Then
                    If isAlert Then
                        sender.Response.Write("<scr" & "ipt>alert('" & jsString(name) & " has exceeded the maxlength of " & jsString(maxLength) & "! - " & jsString(temp) & "');</scr" & "ipt>")
                    Else
                        sender.Response.Write(sender.Server.HtmlEncode(name) & " has exceeded the maxlength of " & maxLength & "! - " & sender.Server.HtmlEncode(temp))
                    End If
                    sender.Response.End()
                Else
                    fvFieldList = appendToList(fvFieldList, UCase(Trim(name)))
                    getPostedDataChk_float = temp
                End If
            Else
                If isAlert Then
                    sender.Response.Write("<scr" & "ipt>alert('" & jsString(name) & " - " & jsString(temp) & " : Incorrect data type! Only float is allowed.');</scr" & "ipt>")
                Else
                    sender.Response.Write(sender.Server.HtmlEncode(name) & " - " & sender.Server.HtmlEncode(temp) & " : Incorrect data type! Only float is allowed.")
                End If
                sender.Response.End()
                'getPostedDataChk_float = temp
            End If
        End If
    End Function

    Public Function getPostedData_date(ByVal postName As String, Optional ByVal defaultValue As Date = Nothing, Optional ByVal isRequestForm As Boolean = False) As Date
        getPostedData_date = getPostedDataChk_date(postName, defaultValue, False, "", isRequestForm)
    End Function

    Private Function getPostedDataAlert_date(ByVal postName As String, Optional ByVal defaultValue As Date = Nothing, Optional ByVal isRequestForm As Boolean = False) As Date
        getPostedDataAlert_date = getPostedDataChk_date(postName, defaultValue, True, "", isRequestForm)
    End Function

    Private Function getPostedDataChk_date(ByVal name As String, ByVal defaultValue As String, ByVal isAlert As Boolean, ByVal chkValue As String, ByVal isRequestForm As Boolean) As Date
        Dim temp As String

        If chkValue <> "" Then
            temp = chkValue
        Else
            If isRequestForm Then
                temp = Convert.ToString(sender.Request.Form(Trim(name)))
            Else
                temp = Convert.ToString(sender.Request(Trim(name)))
            End If
        End If

        If Trim(temp) = "" Or IsNothing(temp) Then
            fvFieldList = appendToList(fvFieldList, UCase(Trim(name)))
            getPostedDataChk_date = defaultValue
        Else
            If IsDate(Trim(temp)) Then
                fvFieldList = appendToList(fvFieldList, UCase(Trim(name)))
                getPostedDataChk_date = temp
            Else
                If isAlert Then
                    sender.Response.Write("<scr" & "ipt>alert('" & jsString(name) & " - " & jsString(temp) & " : Incorrect data type! Only date is allowed.');</scr" & "ipt>")
                Else
                    sender.Response.Write(sender.Server.HtmlEncode(name) & " - " & sender.Server.HtmlEncode(temp) & " : Incorrect data type! Only date is allowed.")
                End If
                sender.Response.End()
                'getPostedDataChk_date = temp
            End If
        End If
    End Function

    Public Function getPostedData_text(ByVal postName As String, Optional ByVal defaultValue As String = "", Optional ByVal isRequestForm As Boolean = False) As String
        getPostedData_text = getPostedDataChk_text(postName, defaultValue, False, "", isRequestForm)
    End Function

    Private Function getPostedDataAlert_text(ByVal postName As String, Optional ByVal defaultValue As String = "", Optional ByVal isRequestForm As Boolean = False) As String
        getPostedDataAlert_text = getPostedDataChk_text(postName, defaultValue, True, "", isRequestForm)
    End Function

    Private Function getPostedDataChk_text(ByVal postName As String, ByVal defaultValue As String, ByVal isAlert As Boolean, ByVal chkValue As String, ByVal isRequestForm As Boolean) As String
        Dim name As String = ""
        Dim blackList As String = ""
        Dim temp As String = ""
        Dim maxLength As Integer = 0
        Dim tempLength As Integer = 0
        Dim nameArray As String()

        If InStr(postName, ":") > 0 Then
            nameArray = postName.Split(":")
            name = nameArray(0)
            maxLength = Convert.ToInt64(nameArray(1))
            'Response.Write "name = " & name & "<br>"
            'Response.Write "maxLength = " & maxLength & "<br>"
            'Response.Write "Request = " & Request(trim(name)) & "<br>"
            'Response.Write "Len = " & Len(Request(trim(name))) & "<br>"
            'Response.End
        Else
            name = postName
            maxLength = 0
        End If

        If chkValue <> "" Then
            temp = chkValue
        Else
            If isRequestForm Then
                temp = Convert.ToString(sender.Request.Form(Trim(name)))
            Else
                temp = Convert.ToString(sender.Request(Trim(name)))
            End If
        End If

        If Not CheckStringForSQL(temp) Then
            blackList = "'--', '/*', '*/', '@@', 'nchar', 'varchar', 'nvarchar', 'alter', 'begin', 'cast', 'create', 'cursor', 'declare', 'delete', 'drop', 'exec', 'execute', 'fetch', 'insert', 'kill', 'select', 'sysobjects', 'syscolumns', 'update'"
            'blackList = "'--', ';', '/*', '*/', '@@', '@', 'char', 'nchar', 'varchar', 'nvarchar', 'alter', 'begin', 'cast', 'create', 'cursor', 'declare', 'delete', 'drop', 'end', 'exec', 'execute', 'fetch', 'insert', 'kill', 'open', 'select', 'sys', 'sysobjects', 'syscolumns', 'table', 'update'"
            If isAlert Then
                sender.Response.Write("<scr" & "ipt>alert('" & jsString(name) & " - " & jsString(temp) & " : Incorrect data type! (" & jsString(blackList) & ") is not allowed.');</scr" & "ipt>")
            Else
                sender.Response.Write(sender.Server.HtmlEncode(name) & " - " & sender.Server.HtmlEncode(temp) & " : Incorrect data type! (" & sender.Server.HtmlEncode(blackList) & ") is not allowed.")
            End If
            sender.Response.End()
        End If

        If temp <> "" Or IsNothing(temp) Then
            tempLength = Len(temp)
        Else
            tempLength = 0
        End If


        If maxLength <> 0 And tempLength > maxLength Then
            If isAlert Then
                sender.Response.Write("<scr" & "ipt>alert('" & jsString(name) & " has exceeded the maxlength of " & jsString(maxLength) & "! - " & jsString(temp) & "');</scr" & "ipt>")
            Else
                sender.Response.Write(sender.Server.HtmlEncode(name) & " has exceeded the maxlength of " & maxLength & "! - " & sender.Server.HtmlEncode(temp))
            End If
            sender.Response.End()
        End If

        fvFieldList = appendToList(fvFieldList, UCase(Trim(name)))

        If Trim(temp) = "" Then
            getPostedDataChk_text = defaultValue
        Else
            getPostedDataChk_text = temp
        End If
    End Function

    Public Function getPostedDataCheck(ByVal postName As String, Optional ByVal defaultValue As String = "", Optional ByVal isRequestForm As Boolean = False) As String
        getPostedDataCheck = getPostedDataCheckOpt(postName, defaultValue, False, "", isRequestForm)
    End Function

    Private Function getPostedDataAlertCheck(ByVal postName As String, Optional ByVal defaultValue As String = "", Optional ByVal isRequestForm As Boolean = False) As String
        getPostedDataAlertCheck = getPostedDataCheckOpt(postName, defaultValue, True, "", isRequestForm)
    End Function

    Private Function getPostedDataCheckOpt(ByVal name As String, ByVal defaultValue As String, ByVal isAlert As Boolean, ByVal chkValue As String, ByVal isRequestForm As Boolean) As String
        Dim temp As String

        If chkValue <> "" Then
            temp = chkValue
        Else
            If isRequestForm Then
                temp = Convert.ToString(dt.decodeNull(sender.Request.Form(Trim(name)).ToString, ""))
            Else
                temp = Convert.ToString(dt.decodeNull(sender.Request(Trim(name)).ToString, ""))
            End If
        End If
        If InStr(", " & fvFieldList.ToUpper & ", ", ", " & name.ToUpper & ", ") <= 0 Then
            If isAlert Then
                sender.Response.Write("<scr" & "ipt>alert('" & jsString(name) & " does not exist in the validation list (" & jsString(fvFieldList) & "), Please add!');</scr" & "ipt>")
            Else
                sender.Response.Write(sender.Server.HtmlEncode(name) & " does not exist in the validation list (" & fvFieldList & "), Please add!")
            End If

            sender.Response.End()
        End If

        If Trim(temp) = "" Then
            getPostedDataCheckOpt = defaultValue
        Else
            getPostedDataCheckOpt = temp
        End If
    End Function

    Private Function CheckStringForSQL(ByVal value As String) As Boolean
        Dim BlackList As String()
        Dim stringLst As String
        Dim s As String

        stringLst = "--, /*, */, @@, nchar, varchar, nvarchar, alter, begin, cast, create, cursor, declare, delete, drop, exec, " & _
                    "execute, fetch, insert, kill, select, sysobjects, syscolumns, update"

        BlackList = stringLst.Split(", ")

        Dim lstr As String = ""

        If value = "" Or value = Nothing Then
            CheckStringForSQL = True
            Exit Function
        ElseIf (StrComp(value, "") = 0) Then
            CheckStringForSQL = True
            Exit Function
        End If

        lstr = LCase(value)

        ' Check if the string contains any patterns in our
        ' black list
        For Each s In BlackList
            If InStr(value.ToUpper, s.ToUpper.Trim) <> 0 Then
                CheckStringForSQL = False
                Exit Function
            End If
        Next

        CheckStringForSQL = True

    End Function

    Public Sub fieldsValidation(ByVal programName As String, ByVal fieldsarrayList As String, Optional ByVal isRequestForm As Boolean = False)
        Dim fieldsArray As String()
        Dim fvCount As Integer
        fieldsArray = Split(fieldsarrayList, ", ")
        For fvCount = 0 To UBound(fieldsArray)
            'execute("call " & programName & "(""" & fieldsArray(fvCount) & ""","""")")

            Select Case programName
                Case "getPostedData_int"
                    Call getPostedData_int(fieldsArray(fvCount), isRequestForm)
                Case "getPostedData_text"
                    Call getPostedData_text(fieldsArray(fvCount), isRequestForm)
                Case "getPostedData_date"
                    Call getPostedData_text(fieldsArray(fvCount), isRequestForm)
                Case "getPostedData_float"
                    Call getPostedData_text(fieldsArray(fvCount), isRequestForm)
            End Select

        Next
    End Sub

    Public Sub fvAppendFormFields(ByVal text_List As String, ByVal Int_List As String, ByVal Float_List As String, ByVal Date_List As String)
        If text_List <> "" Then
            sender.Response.Write("<input type=""hidden"" name=""fvTextList"" value=""" & sender.Server.HtmlEncode(text_List) & """>")
        End If
        If Int_List <> "" Then
            sender.Response.Write("<input type=""hidden"" name=""fvIntList"" value=""" & sender.Server.HtmlEncode(Int_List) & """>")
        End If
        If Float_List <> "" Then
            sender.Response.Write("<input type=""hidden"" name=""fvFloatList"" value=""" & sender.Server.HtmlEncode(Float_List) & """>")
        End If
        If Date_List <> "" Then
            sender.Response.Write("<input type=""hidden"" name=""fvDateList"" value=""" & sender.Server.HtmlEncode(Date_List) & """>")
        End If
    End Sub

    '========================================================================
    'append a new value to a list
    '========================================================================
    Public Function appendToList(ByVal parentList As String, ByVal itemToBeAppended As String) As String
        Dim aList As String
        aList = parentList
        If itemToBeAppended <> "" Then
            If Not inList(aList, itemToBeAppended) Then
                If Trim(aList) = "" Or IsNothing(aList) Then
                    aList = itemToBeAppended
                Else
                    aList = aList & ", " & itemToBeAppended
                End If
            End If
        End If
        appendToList = aList
    End Function

    '========================================================================
    'check whether checkValue is in the given list
    '========================================================================
    Public Function inList(ByVal aList As String, ByVal checkValue As String) As Boolean

        Dim tempArray As String()
        inList = False
        If aList <> "" And checkValue <> "" Then
            tempArray = Split(aList, ", ")
            For i As Integer = 0 To UBound(tempArray)
                If tempArray(i) = checkValue Then
                    inList = True
                    Exit For
                End If
            Next
        End If
    End Function

    Public Function jsString(ByVal aText As String) As String
        '*** escape the special characters for Javascript
        Dim tempText As String

        If IsNothing(aText) Then
            tempText = ""
        Else
            tempText = aText
        End If

        If Trim(tempText).Length > 0 Then
            tempText = Replace(tempText, "\", "\\")
            tempText = Replace(tempText, "'", "\'")
            tempText = Replace(tempText, """", "\""")
            tempText = Replace(tempText, vbCrLf, "\n")
            tempText = Replace(tempText, Chr(10), "\n")
            tempText = Replace(tempText, Chr(13), "\n")
            jsString = tempText
        Else
            jsString = tempText
        End If
    End Function

    Public Function checkValueByDouble(ByVal objname As String, ByVal checkValue As Double, Optional ByVal defautlvalue As Double = 0) As Double

        checkValueByDouble = getPostedDataChk_float(objname, defautlvalue, False, checkValue, False)

    End Function

    Public Function checkValueByInt(ByVal objname As String, ByVal checkValue As Integer, Optional ByVal defautlvalue As Integer = 0) As Integer

        checkValueByInt = getPostedDataChk_int(objname, defautlvalue, False, checkValue, False)

    End Function

    Public Function checkValueByString(ByVal objname As String, ByVal checkValue As String, Optional ByVal defautlvalue As String = "") As String

        checkValueByString = getPostedDataChk_text(objname, defautlvalue, False, checkValue, False)

    End Function

    Public Function checkValueByDate(ByVal objname As String, ByVal checkValue As Date, Optional ByVal defautlvalue As Date = Nothing) As Date

        checkValueByDate = getPostedDataChk_date(objname, defautlvalue, False, checkValue, False)

    End Function

    Public Sub OnSessionIdCheck(ByVal s As Page, Optional ByVal isRequestForm As Boolean = True)
        Try
            Dim ctx As HttpContext = HttpContext.Current
            Dim session As SessionState.HttpSessionState = ctx.Session
            'session.Timeout = 120

            Dim y As Integer = session.Count

            Dim secl_sessionid As String = ""
            Dim session_SQL As String
            Dim session_table As New DataTable
            Dim tempSessionId As String

            If isRequestForm Then
                tempSessionId = s.Request.Form("secl_sessionid")
            Else
                tempSessionId = s.Request("secl_sessionid")
            End If

            If session("secl_sessionid") Is Nothing Then
                If tempSessionId Is Nothing Then
                    secl_sessionid = ""

                ElseIf tempSessionId = "" Then
                    secl_sessionid = ""
                Else
                    secl_sessionid = tempSessionId
                    session("secl_sessionid") = secl_sessionid
                End If
            ElseIf session("secl_sessionid") = "" Then
                If tempSessionId Is Nothing Then
                    secl_sessionid = ""
                ElseIf tempSessionId = "" Then
                    secl_sessionid = ""
                Else
                    secl_sessionid = tempSessionId
                    session("secl_sessionid") = secl_sessionid
                End If
            Else
                secl_sessionid = session("secl_sessionid")
            End If

            If secl_sessionid <> "" Then
                session_SQL = "select * from LM_SECUR_LOG where secl_sessionid = '" & secl_sessionid & "'"
                session_table = gDB.getDataTable(session_SQL)

                If session_table.Rows.Count = 0 Then
                    s.Response.Write("<script language='javascript'>alert('Server Update. Please login again.');</script>")
                    session.Clear()
                    s.Response.End()
                End If
            Else
                s.Response.Write("<script language='javascript'>alert('Server Update. Please login again.');</script>")
                session.Clear()
                s.Response.End()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Function aspSessionXfer() As String
        Dim tempStringList As String

        tempStringList = ""

        'sender.Session.Timeout = 120

        For i As Integer = 0 To sender.Request.Form.Count - 1
            If sender.Request.Form.GetKey(i).ToString.ToUpper <> "SQL_TEXT" Then
                If tempStringList <> "" Then
                    tempStringList = tempStringList & ", " & sender.Request.Form.GetKey(i).ToString
                Else
                    tempStringList = sender.Request.Form.GetKey(i).ToString
                End If

                sender.Session(sender.Request.Form.GetKey(i)) = sender.Request.Form(i).ToString()
            End If
        Next

        aspSessionXfer = tempStringList

    End Function

    Public Function htmlEncodeNull(ByVal input As String) As String

        Dim aStr As String

        aStr = input.Replace(";", "_000_SEMICOLON_000_")
        aStr = aStr.Replace("#", "_000_HASH_000_")
        'aStr = sender.Server.HtmlEncode(decodeNull(aStr, ""))

        aStr = aStr.Replace("(", "&#40;")
        aStr = aStr.Replace(")", "&#41;")
        aStr = aStr.Replace("'", "&#39;")
        aStr = aStr.Replace("_000_HASH_000_", "&#35;")
        aStr = aStr.Replace("/", "&#47;")
        aStr = aStr.Replace("\", "&#92;")
        aStr = aStr.Replace("_000_SEMICOLON_000_", "&#59;")
        aStr = aStr.Replace(":", "&#58;")
        aStr = aStr.Replace("=", "&#61;")

        htmlEncodeNull = aStr

    End Function

    Private Function decodeNull(ByVal aValue As String, ByVal defaultVal As String) As String
        If aValue = Nothing Then
            decodeNull = defaultVal
        Else
            decodeNull = aValue
        End If
    End Function

    Private Function dbEncode(ByVal aStr As String) As String
        dbEncode = Replace(decodeNull(aStr, ""), "'", "")
    End Function

    Public Function get_SqlAppInTemp(ByVal tmpkey As String) As String
        Dim SQL_APP As String = ""

        'sender.Response.Write("<script language='javascript'>alert('key=" & tmpkey & "');</script>")
        Dim nSQL As String = "select * from lm_tmp_dot_net where tmp_key = " & tmpkey
        Dim nTable As New DataTable

        If tmpkey <> "" Then
            nTable = gDB.getDataTable(nSQL)

            If nTable.Rows.Count <> 0 Then
                SQL_APP = Convert.ToString(nTable.Rows(0).Item("SRCL_SQL"))
            End If

            'Dim dSQL As String = "delete from lm_tmp_dot_net where tmp_key = " & tmpkey
            'dt.amendData(dSQL)

            nTable.Dispose()

            Return SQL_APP
        Else
            SQL_APP = ""

            Return SQL_APP
        End If

    End Function

    Public Sub stor_session_key(Optional ByVal default_request_name As String = "tmp_key")
        sender.Session("tmpkey") = Convert.ToString(sender.Request.Form(Trim(default_request_name)))
    End Sub
End Class
