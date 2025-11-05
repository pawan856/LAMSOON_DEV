
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Net
Imports System.IO
Imports System

Public Class GeneralUtils
    Public Const FIELD_SEPERATOR = "|#|"
    Public Const USE_CACHE = False

    Public Function getConfig(ByVal cacheName As String) As String
        If USE_CACHE Then
            Return HttpContext.Current.Cache(cacheName)
        Else
            Return System.Configuration.ConfigurationManager.AppSettings.Item(cacheName)
        End If
    End Function

    Public Function convdbDate(ByVal aValue As String) As String
        If aValue = "" Then
            convdbDate = "NULL"
        Else
            'convdbDate = "to_date('" & aValue & "', '" & getConfig("DDFORMAT") & "')"
            convdbDate = "CONVERT(datetime, '" & aValue & "', " & getConfig("DDFORMATNO") & ")"
        End If
    End Function

    Public Function convdbNumData(ByVal aValue As String) As String
        If aValue = "" Then
            convdbNumData = "NULL"
        Else
            convdbNumData = aValue
        End If
    End Function

    Public Function convdbVCData(ByVal aValue As String) As String
        If aValue = "" Then
            convdbVCData = "NULL"
        Else
            convdbVCData = "'" & aValue & "'"
        End If
    End Function

    Public Function convdbNVCData(ByVal aValue As String) As String
        If aValue = "" Then
            convdbNVCData = "NULL"
        Else
            convdbNVCData = "N'" & aValue & "'"
        End If
    End Function

    Public Function decodeNull(ByVal aValue As String, ByVal defaultVal As String) As String
        If aValue Is Nothing Then
            decodeNull = defaultVal
        Else
            decodeNull = aValue
        End If
    End Function

    Public Function decodeNullOrEmpty(ByVal aValue As String, ByVal defaultVal As String) As String
        If aValue Is Nothing OrElse aValue = "" Then
            decodeNullOrEmpty = defaultVal
        Else
            decodeNullOrEmpty = aValue
        End If
    End Function

    Public Function dbEncode(ByVal aStr As String) As String
        dbEncode = Replace(decodeNull(aStr, ""), "'", "''")
    End Function

    Public Function decodeEmptyCdbl(ByVal aValue As String, ByVal defaultVal As Double) As Double
        If aValue Is Nothing OrElse aValue = "" Then
            decodeEmptyCdbl = defaultVal.ToString
        Else
            If Not IsNumeric(aValue) Then
                aValue = 0
            End If

            decodeEmptyCdbl = aValue
        End If

        decodeEmptyCdbl = CDbl(decodeEmptyCdbl)
    End Function

    Public Function decodeEmptyCInt(ByVal aValue As String, ByVal defaultVal As Integer) As Integer
        If aValue Is Nothing OrElse aValue = "" Then
            decodeEmptyCInt = defaultVal.ToString
        Else
            If Not IsNumeric(aValue) Then
                aValue = 0
            End If

            decodeEmptyCInt = aValue
        End If

        decodeEmptyCInt = CInt(decodeEmptyCInt)
    End Function

    Public Function checkPhoneNumber(ByVal Value As String) As Boolean
        Dim tempString As String = Value.Replace("+", "").Replace("-", "").Replace(" ", "").Trim

        If Not Value.Trim = "" Then
            If IsNumeric(tempString) Then
                checkPhoneNumber = True
            Else
                checkPhoneNumber = False
            End If
        Else
            checkPhoneNumber = True
        End If

    End Function

    Public Function isValidEmail(ByVal EmailValue As String) As Boolean
        Dim isValid As Boolean = False
        Dim regEx As New Regex("^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$", RegexOptions.IgnoreCase)
        Dim tmpEmailValue As String = EmailValue
        Dim emailArray As String()
        Dim i As Integer

        tmpEmailValue = Replace(tmpEmailValue, ";", ",")

        If tmpEmailValue <> "" Then
            emailArray = listToArray(tmpEmailValue)

            For i = 0 To UBound(emailArray)
                If Not regEx.IsMatch(emailArray(i)) Then
                    Return False
                End If
            Next
        End If

        Return True
    End Function

    Public Function isValidPWD(ByVal pwdValue As String) As Boolean
        Dim regEx1 As New Regex("[\d]")
        Dim regEx2 As New Regex("[a-z]")
        Dim regEx3 As New Regex("[A-Z]")
        Dim regEx4 As New Regex("[\W]")
        Dim matchCnt As Integer = 0

        If regEx1.IsMatch(pwdValue) Then
            matchCnt += 1
        End If
        If regEx2.IsMatch(pwdValue) Then
            matchCnt += 1
        End If
        If regEx3.IsMatch(pwdValue) Then
            matchCnt += 1
        End If
        If regEx4.IsMatch(pwdValue) Then
            matchCnt += 1
        End If

        If matchCnt >= 3 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function IsWholeNumber(ByVal str As String) As Boolean
        Dim regEx As New Regex("\D", RegexOptions.IgnoreCase)

        IsWholeNumber = Not regEx.IsMatch(str)
    End Function

    Public Function IsSignedInteger(ByVal strVal As [String]) As Boolean
        Dim reg As New Regex("[^0-9-]")
        Dim reg2 As New Regex("^-[0-9]+$|^[0-9]+$")
        Return (Not reg.IsMatch(strVal) AndAlso reg2.IsMatch(strVal))
    End Function

    Public Function isDecimal(ByVal str As String) As Boolean
        If Trim(str) <> "" Then
            If Left(str, 1) = "." Then
                Return False
            Else
                Return IsNumeric(str)
            End If
        Else
            Return True
        End If
    End Function

    Public Function isTime(ByVal str As String) As Boolean
        Dim tempArray As String()

        isTime = True

        If Trim(str) <> "" Then
            tempArray = Split(str, ":")

            If UBound(tempArray) = 0 Then
                If IsWholeNumber(str) Then
                    If CInt(str) > 23 Then
                        isTime = False
                    End If
                Else
                    isTime = False
                End If
            ElseIf UBound(tempArray) = 1 Then
                If IsWholeNumber(tempArray(0)) And IsWholeNumber(tempArray(1)) Then
                    If CInt(tempArray(0)) > 23 Then
                        isTime = False
                    ElseIf CInt(tempArray(1)) > 59 Then
                        isTime = False
                    End If
                Else
                    isTime = False
                End If
            Else
                isTime = False
            End If
        End If
    End Function

    Public Function jsString(ByVal aText As String) As String
        '*** escape the special characters for Javascript
        Dim tempText As String

        If (String.IsNullOrEmpty(aText)) Then
            tempText = ""
        Else
            tempText = aText
        End If

        If Len(Trim(tempText)) > 0 Then
            tempText = Replace(tempText, "\", "\\")
            tempText = Replace(tempText, "'", "\'")
            tempText = Replace(tempText, """", "\""")
            tempText = Replace(tempText, vbCrLf, "\n")
            tempText = Replace(tempText, Chr(10), "\n")
            tempText = Replace(tempText, Chr(13), "\n")
            jsString = tempText
        Else
            jsString = ""
        End If
    End Function

    Public Function jsURLEncode(ByVal aText As String) As String
        Dim aStr As String
        aStr = aText
        'aStr = Server.UrlEncode(decodeNull(aText, ""))
        jsURLEncode = jsString(aStr)
    End Function

    Public Function jsHTMLEncode(ByVal aText As String) As String
        Dim aStr As String
        aStr = decodeNull(aText, "")
        aStr = Replace(aStr, ";", "_000_SEMICOLON_000_")
        aStr = Replace(aStr, "#", "_000_HASH_000_")
        'aStr = Server.HtmlEncode(aStr)
        aStr = Replace(aStr, "(", "&#40;")
        aStr = Replace(aStr, ")", "&#41;")
        aStr = Replace(aStr, "'", "&#39;")
        aStr = Replace(aStr, "_000_HASH_000_", "&#35;")
        aStr = Replace(aStr, "/", "&#47;")
        aStr = Replace(aStr, "\", "&#92;")
        aStr = Replace(aStr, "_000_SEMICOLON_000_", "&#59;")
        aStr = Replace(aStr, ":", "&#58;")
        aStr = Replace(aStr, "=", "&#61;")
        jsHTMLEncode = jsString(aStr)
    End Function

    Public Function htmlEncodeNullDef(ByVal aText As String, ByVal defaultVal As String) As String
        If String.IsNullOrEmpty(aText) Then
            htmlEncodeNullDef = defaultVal
        Else
            htmlEncodeNullDef = htmlEncodeNull(aText)
        End If
    End Function

    Public Function htmlEncodeNull(ByVal aText As String) As String
        Dim aStr As String
        aStr = decodeNull(aText, "")
        aStr = Replace(aStr, ";", "_000_SEMICOLON_000_")
        aStr = Replace(aStr, "#", "_000_HASH_000_")
        'aStr = Server.HtmlEncode(aStr)
        aStr = Replace(aStr, "(", "&#40;")
        aStr = Replace(aStr, ")", "&#41;")
        aStr = Replace(aStr, "'", "&#39;")
        aStr = Replace(aStr, "_000_HASH_000_", "&#35;")
        aStr = Replace(aStr, "/", "&#47;")
        aStr = Replace(aStr, "\", "&#92;")
        aStr = Replace(aStr, "_000_SEMICOLON_000_", "&#59;")
        aStr = Replace(aStr, ":", "&#58;")
        aStr = Replace(aStr, "=", "&#61;")
        htmlEncodeNull = aStr
    End Function

    Public Function urlEncodeNull(ByVal aText As String) As String
        Dim aStr As String
        aStr = aText
        'aStr = Server.UrlEncode(decodeNull(aText, ""))
        urlEncodeNull = aStr
    End Function

    '========================================================================
    'append a new value to a list, no duplicate
    '========================================================================
    Public Function appendToList(ByVal parentList As String, ByVal itemToBeAppended As String) As String
        Dim aList As String
        aList = parentList
        If (itemToBeAppended <> "") Then
            If (Not inList(aList, itemToBeAppended)) Then
                If (Trim(aList) = "") Then
                    aList = itemToBeAppended
                Else
                    aList = aList & ", " & itemToBeAppended
                End If
            End If
        End If
        appendToList = aList
    End Function

    '========================================================================
    'append a new value to a list, duplicate
    '========================================================================
    Public Function appendToList2(ByVal parentList As String, ByVal itemToBeAppended As String) As String
        Dim aList As String
        aList = parentList

        If (Trim(aList) <> "") Then
            aList = aList & ", " & itemToBeAppended
        Else
            aList = itemToBeAppended
        End If
        appendToList2 = aList
    End Function

    '========================================================================
    'append a new value to a jsList, no duplicate
    '========================================================================
    Public Function jsAppendToList(ByVal parentList As String, ByVal itemToBeAppended As String) As String
        Dim aList As String
        aList = parentList
        If (itemToBeAppended <> "") Then
            If (InStr(aList, "|" & itemToBeAppended & "|") <= 0) Then
                aList = aList & "|" & itemToBeAppended & "|"
            End If
        End If
        jsAppendToList = aList
    End Function

    '========================================================================
    'check whether checkValue is in the given jsList
    '========================================================================
    Public Function jsInList(ByVal aList As String, ByVal checkValue As String) As Boolean
        If (aList <> "" And checkValue <> "") Then
            If InStr(aList, "|" & checkValue & "|") > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    '========================================================================
    'check whether checkValue is in the given list
    '========================================================================
    Public Function inList(ByVal aList As String, ByVal checkValue As String) As Boolean

        Dim tempArray As String()
        inList = False
        If (aList <> "" And checkValue <> "") Then
            tempArray = Split(formatList(aList), ", ")
            For i As Integer = 0 To UBound(tempArray)
                If (tempArray(i) = checkValue) Then
                    inList = True
                    Exit For
                End If
            Next
        End If
    End Function

    '========================================================================
    'combine two lists into one, eliminate duplicates
    '========================================================================
    Public Function listUnion(ByVal listA As String, ByVal listB As String) As String
        Dim tempArray As String()
        Dim finalList As String
        Dim i As Integer

        If (listB = "") Then
            listUnion = listA
        ElseIf (listA = "") Then
            listUnion = listB
        Else
            finalList = listA

            tempArray = listToArray(listB)
            For i = 0 To UBound(tempArray)
                If (Not inList(finalList, tempArray(i))) Then
                    finalList = appendToList(finalList, tempArray(i))
                End If
            Next
            listUnion = finalList
        End If
    End Function

    '========================================================================
    'combine two jsLists into one, eliminate duplicates
    '========================================================================
    Public Function jsListUnion(ByVal listA As String, ByVal listB As String) As String
        Dim tempArray As String()
        Dim finalList As String
        Dim i As Integer

        If (listB = "") Then
            jsListUnion = listA
        ElseIf (listA = "") Then
            jsListUnion = listB
        Else
            finalList = listA

            tempArray = Split(Mid(listB, 2, Len(listB) - 2), "||")
            For i = 0 To UBound(tempArray)
                If (InStr(finalList, "|" & tempArray(i) & "|") <= 0) Then
                    finalList = finalList & "|" & tempArray(i) & "|"
                End If
            Next
            jsListUnion = finalList
        End If
    End Function

    '========================================================================
    'subtract listB from listA
    '========================================================================
    Public Function listSubtract(ByVal listA As String, ByVal listB As String) As String
        Dim jsListA, jsListB As String

        jsListA = vbListToJsList(listA)
        jsListB = vbListToJsList(listB)

        listSubtract = jsListToVbList(jsListSubtract(jsListA, jsListB))
    End Function

    '========================================================================
    'subtract listB from listA
    '========================================================================
    Public Function jsListSubtract(ByVal listA As String, ByVal listB As String) As String
        Dim tempArrayB As String()
        Dim i As Integer
        Dim finalList As String

        finalList = listA

        If listB = "" Then
            Return finalList
        Else
            tempArrayB = jsListToArray(listB)
            For i = 0 To UBound(tempArrayB)
                finalList = jsRemoveFromList(finalList, "" & tempArrayB(i))
            Next
            Return finalList
        End If
    End Function

    '========================================================================
    'combine two lists into one, only take duplicates
    '========================================================================
    Public Function listIntercept(ByVal listA As String, ByVal listB As String) As String
        Dim tempArray As String()
        Dim finalList As String
        Dim i As Integer

        If (listB = "" Or listA = "") Then
            listIntercept = ""
        Else
            finalList = ""

            tempArray = listToArray(listB)
            For i = 0 To UBound(tempArray)
                If (inList(listA, tempArray(i))) Then
                    finalList = appendToList(finalList, tempArray(i))
                End If
            Next
            listIntercept = finalList
        End If
    End Function

    '========================================================================
    'Format comma in list, make sure the items are seperated by ", "
    '========================================================================
    Public Function formatList(ByVal aList As String) As String
        Dim tempList As String

        tempList = Replace(aList, ",", ", ")
        Do
            tempList = Replace(tempList, ",  ", ", ")
        Loop While InStr(1, tempList, ",  ") > 0
        formatList = tempList
    End Function

    '========================================================================
    'convert a list to array
    '========================================================================
    Public Function listToArray(ByVal aList As String) As String()
        Dim tempList As String

        tempList = formatList(aList)
        listToArray = Split(tempList, ", ")
    End Function

    '========================================================================
    'convert a jsList to array
    '========================================================================
    Public Function jsListToArray(ByVal aList As String) As String()
        Dim tempList As String

        tempList = aList

        If (Len(tempList) - 2 <= 0) Then
            jsListToArray = Split("", "||")
        Else
            tempList = Mid(tempList, 2, Len(tempList) - 2)
            jsListToArray = Split(tempList, "||")
        End If
    End Function

    '========================================================================
    'convert a jsList to vbList
    '========================================================================
    Public Function jsListToVbList(ByVal aList As String) As String
        Dim tempList As String
        tempList = aList

        If (Len(tempList) - 2 <= 0) Then
            jsListToVbList = ""
        Else
            tempList = Mid(tempList, 2, Len(tempList) - 2)
            jsListToVbList = Replace(tempList, "||", ", ")
        End If
    End Function

    '========================================================================
    'convert a vbList to jsList
    '========================================================================
    Public Function vbListToJsList(ByVal aList As String) As String
        If (aList = "") Then
            vbListToJsList = ""
        Else
            vbListToJsList = "|" & Replace(aList, ", ", "||") & "|"
        End If
    End Function

    '========================================================================
    'remove an item from the list if exists
    '========================================================================
    Public Function removeFromList(ByVal aList As String, ByVal itemToBeRemoved As String) As String
        Dim tempList As String

        tempList = formatList(aList)

        If (tempList <> "" AndAlso itemToBeRemoved <> "") Then
            If (inList(tempList, itemToBeRemoved)) Then
                If tempList = itemToBeRemoved Then
                    tempList = ""
                Else
                    tempList = Replace(tempList, ", " & itemToBeRemoved & ", ", ", ")

                    If Len(tempList) > Len(itemToBeRemoved) Then
                        If Left(tempList, Len(itemToBeRemoved) + 2) = itemToBeRemoved & ", " Then
                            tempList = Mid(tempList, Len(itemToBeRemoved) + 3)
                        End If

                        If Right(tempList, Len(itemToBeRemoved) + 2) = ", " & itemToBeRemoved Then
                            tempList = Left(tempList, Len(tempList) - Len(itemToBeRemoved) - 2)
                        End If
                    End If
                End If
            End If
        End If

        removeFromList = tempList
    End Function

    '========================================================================
    'remove an item from the jsList if exists
    '========================================================================
    Public Function jsRemoveFromList(ByVal aList As String, ByVal itemToBeRemoved As String) As String
        jsRemoveFromList = Replace(aList, "|" & itemToBeRemoved & "|", "")
    End Function

    Public Function isEndDateLaterThanStart(ByVal startDate As String, ByVal endDate As String, Optional ByVal falseForEqual As Boolean = False) As Boolean
        Dim startDateYMD, endDateYmd As String

        If startDate.Trim = "" OrElse endDate.Trim = "" Then
            Return False
        End If

        startDateYMD = chgDateFormat(startDate, getConfig("DDFORMAT"), "YYYYMMDD")
        endDateYmd = chgDateFormat(endDate, getConfig("DDFORMAT"), "YYYYMMDD")

        If endDateYmd > startDateYMD Then
            Return True
        ElseIf Not falseForEqual AndAlso endDateYmd = startDateYMD Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function convertToDate(ByVal Value As String, Optional ByRef isSuccess As Boolean = False,
                                  Optional ByVal dateFormat As String = "") As DateTime
        Dim tempDate As DateTime
        Dim myculture As New System.Globalization.CultureInfo("en-US")
        Dim datePattern As String = dateFormat
        Dim inputDate As String = Value

        'If datePattern = "" Then
        '    isSuccess = DateTime.TryParse(Value, tempDate)
        'Else
        '    myculture.DateTimeFormat.ShortDatePattern = datePattern
        '    myculture.DateTimeFormat.LongDatePattern = datePattern
        '    isSuccess = DateTime.TryParse(inputDate, myculture, Globalization.DateTimeStyles.None, tempDate)
        'End If

        If dateFormat = "" Then
            datePattern = getConfig("DDFORMAT")
        ElseIf InStr(dateFormat, "/") <= 0 Then
            inputDate = chgDateFormat(inputDate, dateFormat, getConfig("DDFORMAT"))
            datePattern = getConfig("DDFORMAT")
        Else
            datePattern = dateFormat
        End If

        datePattern = Replace(Replace(UCase(datePattern), "DD", "dd"), "YY", "yy")

        myculture.DateTimeFormat.ShortDatePattern = datePattern
        myculture.DateTimeFormat.LongDatePattern = datePattern

        isSuccess = DateTime.TryParse(inputDate, myculture, Globalization.DateTimeStyles.None, tempDate)

        If isSuccess Then
            Return tempDate
        Else
            Return Nothing
        End If
    End Function

    Public Function isValidDate(ByVal aDate As String, Optional ByVal dateFormat As String = "") As Boolean
        Dim myculture As New System.Globalization.CultureInfo("en-US")
        Dim datePattern As String
        Dim inputDate As String = aDate

        If dateFormat = "" Then
            datePattern = getConfig("DDFORMAT2")
        ElseIf InStr(dateFormat, "/") <= 0 AndAlso InStr(UCase(dateFormat), "DD") > 0 AndAlso InStr(UCase(dateFormat), "MM") > 0 AndAlso InStr(UCase(dateFormat), "YY") > 0 Then
            inputDate = chgDateFormat(inputDate, dateFormat, getConfig("DDFORMAT"))
            datePattern = getConfig("DDFORMAT2")
        Else
            datePattern = dateFormat
        End If

        datePattern = Replace(Replace(UCase(datePattern), "D", "d"), "YY", "yy")

        myculture.DateTimeFormat.ShortDatePattern = datePattern
        myculture.DateTimeFormat.LongDatePattern = datePattern

        'If DateTime.TryParse(inputDate, myculture, Globalization.DateTimeStyles.None, New DateTime) Then
        '    Return True
        'Else
        '    Return False
        'End If

        If DateTime.TryParseExact(inputDate, datePattern, myculture, Globalization.DateTimeStyles.None, New DateTime) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function getChkBoxListValue(ByRef chkBoxListCtrl As CheckBoxList) As String
        Dim i As Integer
        Dim chkBoxValue As String = ""

        For i = 0 To chkBoxListCtrl.Items.Count - 1
            If chkBoxListCtrl.Items(i).Selected Then
                chkBoxValue = appendToList(chkBoxValue, chkBoxListCtrl.Items(i).Value)
            End If
        Next

        Return chkBoxValue
    End Function

    Public Sub setChkBoxListByValue(ByRef chkBoxListCtrl As CheckBoxList, ByVal chkBoxValue As String)
        Dim chkBoxArray As String() = listToArray(chkBoxValue)
        Dim i, j As Integer

        chkBoxListCtrl.ClearSelection()

        For i = 0 To chkBoxListCtrl.Items.Count - 1
            For j = 0 To UBound(chkBoxArray)
                If chkBoxListCtrl.Items(i).Value = chkBoxArray(j) Then
                    chkBoxListCtrl.Items(i).Selected = True
                End If
            Next
        Next
    End Sub

    Public Function chgDateFormat(ByVal aDate As String, ByVal orgFormat As String, ByVal tarFormat As String) As String
        Dim yearStr, monthStr, dayStr As String
        Dim orgDateFormat As String = UCase(orgFormat.Trim)
        Dim tarDateFormat As String = UCase(tarFormat.Trim)
        Dim inputDate As String = UCase(aDate.Trim)
        Dim resultDate As String

        If inputDate <> "" Then
            resultDate = tarDateFormat

            If InStr(orgDateFormat, "YYYY") Then
                yearStr = Mid(inputDate, InStr(orgDateFormat, "YYYY"), 4)
                resultDate = Replace(resultDate, "YYYY", yearStr)
            Else
                yearStr = Mid(inputDate, InStr(orgDateFormat, "YY"), 2)
                resultDate = Replace(resultDate, "YY", yearStr)
            End If

            monthStr = Mid(inputDate, InStr(orgDateFormat, "MM"), 2)
            resultDate = Replace(resultDate, "MM", monthStr)

            dayStr = Mid(inputDate, InStr(orgDateFormat, "DD"), 2)
            resultDate = Replace(resultDate, "DD", dayStr)
        Else
            resultDate = ""
        End If

        Return resultDate
    End Function

    'Public Function gvValidate(ByRef inPage As Page, ByVal gvDataTable As DataTable, ByVal fieldNmae As String, ByVal ctrlLabel As String, ByVal ctrlValue As String) As Boolean
    '    Select Case gvDataTable.Columns(fieldNmae).DataType.FullName
    '        Case "System.DateTime"
    '            If Not IsDate(ctrlValue) Then
    '                gvValidate = False

    '                uiFun.displayMsg(inPage, "1010", "", HttpContext.Current.Session("gLang"), ctrlLabel)
    '            End If
    '        Case "System.Decimal"
    '            If Not gU.isDecimal(ctrlValue) Then
    '                gvValidate = False

    '                uiFun.displayMsg(inPage, "1009", "", HttpContext.Current.Session("gLang"), ctrlLabel)
    '            End If

    '        Case "System.DBNull"

    '        Case "System.String"

    '    End Select
    'End Function

    Public Function splitRngOfString(ByVal val As String, ByVal FieldName As String, Optional upperEntry As Boolean = False) As String
        Try
            Dim tmpStr As String = ""
            If val <> "" Then
                Dim nList() As String
                Dim tempstring As String = ""
                Dim rangeStr As String = ""
                Dim entry As String

                nList = Regex.Split(val, "("")|(,)")

                For Each entry In nList
                    If entry.Trim <> "" Then
                        If Not entry.Trim.Contains("-") AndAlso entry.Trim <> "," Then
                            If tempstring <> "" Then tempstring += ","
                            tempstring += convdbVCData(decodeNullOrEmpty(entry.Trim, ""))
                        ElseIf entry.Trim.Contains("-") Then
                            If entry.Substring(0, 1) <> "-" Then
                                Dim toList() As String
                                Dim tmpRStr As String = ""
                                toList = Split(entry, "-")

                                If toList.Length > 0 Then
                                    If upperEntry Then
                                        tmpRStr += "UPPER(" & FieldName & ") >= " & convdbVCData(decodeNullOrEmpty(toList(0).Trim.ToUpper, ""))
                                    Else
                                        tmpRStr += FieldName & " >= " & convdbVCData(decodeNullOrEmpty(toList(0).Trim, ""))
                                    End If

                                    If toList.Length > 1 Then
                                        If upperEntry Then
                                            tmpRStr += " AND UPPER(" & FieldName & ") <= " & convdbVCData(decodeNullOrEmpty(toList(toList.Length - 1).Trim.ToUpper, ""))
                                        Else
                                            tmpRStr += " AND " & FieldName & " <= " & convdbVCData(decodeNullOrEmpty(toList(toList.Length - 1).Trim, ""))
                                        End If

                                    End If
                                End If

                                If tmpRStr <> "" Then
                                    If rangeStr <> "" Then rangeStr += " OR "
                                    rangeStr += "(" & tmpRStr & ")"
                                End If

                            End If
                        End If
                    End If

                Next entry

                If upperEntry Then
                    If tempstring <> "" Then tmpStr += "UPPER(" & FieldName & ") IN (" & tempstring.ToUpper & ")"
                Else
                    If tempstring <> "" Then tmpStr += FieldName & " IN (" & tempstring & ")"
                End If

                If rangeStr <> "" Then
                    If tmpStr <> "" Then tmpStr += " OR "
                    tmpStr += rangeStr
                End If

            Else
                tmpStr = val
            End If

            If tmpStr <> "" Then
                If tmpStr.Contains(" OR (") Then
                    tmpStr = "(" & tmpStr & ")"
                End If
            End If

            Return tmpStr

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function isAlphaNumueric(ByVal strInputText As String) As Boolean
        Dim intCounter As Integer
        Dim strCompare As String
        Dim strInput As String
        isAlphaNumueric = False

        For intCounter = 1 To Len(strInputText)
            strCompare = Mid$(strInputText, intCounter, 1)
            strInput = Mid$(strInputText, intCounter + 1, Len(strInputText))
            If strCompare Like ("[A-Z]") Or _
                strCompare Like ("[a-z]") Or _
                strCompare Like ("#") Or _
                strCompare = "_" Then
                isAlphaNumueric = True
            Else
                isAlphaNumueric = False
                Exit Function
            End If
        Next intCounter

    End Function

    Public Function isPhoneFaxNo(ByVal strInputText As String) As Boolean
        Dim intCounter As Integer
        Dim strCompare As String
        Dim strInput As String
        isPhoneFaxNo = False

        For intCounter = 1 To Len(strInputText)
            strCompare = Mid$(strInputText, intCounter, 1)
            strInput = Mid$(strInputText, intCounter + 1, Len(strInputText))
            If strCompare Like ("#") Or _
                strCompare = "-" Or _
                strCompare = " " Or _
                strCompare = "(" Or _
                strCompare = ")" _
            Then
                isPhoneFaxNo = True
            Else
                isPhoneFaxNo = False
                Exit Function
            End If
        Next intCounter

    End Function

    Public Function sendEmail(ByVal emailTo As String, ByVal emailTitle As String, ByVal emailBody As String, ByRef msgLog As PrgmLog) As Boolean
        Dim eMsg As System.Net.Mail.MailMessage
        Dim frAddr As System.Net.Mail.MailAddress
        Dim smtp As System.Net.Mail.SmtpClient
        Dim email_subject, email_content As String
        Dim trySendCnt As Integer
        Dim preLogLoc As String

        preLogLoc = msgLog.PrgmLocation
        msgLog.PrgmLocation = "GeneralUtils.sendEmail"

        Try
            email_subject = emailTitle
            email_content = emailBody

            frAddr = New System.Net.Mail.MailAddress(System.Configuration.ConfigurationManager.AppSettings.Item("From_Email"), "")

            eMsg = New System.Net.Mail.MailMessage()
            eMsg.From = frAddr

            eMsg.To.Add(emailTo)
            'eMsg.CC.Add(ccEmailList)

            eMsg.Subject = email_subject

            eMsg.SubjectEncoding = Text.Encoding.UTF8
            eMsg.Body = email_content
            eMsg.BodyEncoding = Text.Encoding.UTF8
            eMsg.IsBodyHtml = False

            smtp = New System.Net.Mail.SmtpClient()
            smtp.Host = System.Configuration.ConfigurationManager.AppSettings.Item("SMTP_HOST")

            If System.Configuration.ConfigurationManager.AppSettings.Item("SMTP_PORT") <> "" Then
                smtp.Port = System.Configuration.ConfigurationManager.AppSettings.Item("SMTP_PORT")
            End If

            If System.Configuration.ConfigurationManager.AppSettings.Item("SMTP_SENDUSERNAME") <> "" Then
                smtp.Credentials = New System.Net.NetworkCredential(System.Configuration.ConfigurationManager.AppSettings.Item("SMTP_SENDUSERNAME"), System.Configuration.ConfigurationManager.AppSettings.Item("SMTP_SENDPASSWORD"))
            End If

            trySendCnt = 0
            Do While trySendCnt < 3
                Try
                    smtp.Send(eMsg)
                    trySendCnt = 3
                Catch ex As System.Net.Mail.SmtpException
                    trySendCnt = trySendCnt + 1
                    If trySendCnt >= 3 Then
                        Throw ex
                    End If
                End Try
            Loop

            frAddr = Nothing
            eMsg = Nothing
            smtp = Nothing

            Return True

        Catch ex As Exception
            msgLog.writeLog("Email sending error: " & ex.Message)
            Return False
        Finally
            msgLog.PrgmLocation = preLogLoc
        End Try
    End Function

    Function isMMYYYY(ByVal dateValues As String, ByRef inbox As TextBox) As Boolean
        Dim intCounter As Integer
        Dim strCompare As String
        Dim strInput As String

        Dim temp_year As String
        Dim temp_month As String

        For intCounter = 1 To Len(dateValues)
            strCompare = Mid$(dateValues, intCounter, 1)
            strInput = Mid$(dateValues, intCounter + 1, Len(dateValues))
            If strCompare Like ("#") Or _
                strCompare = "/" _
            Then
                'Return True
            Else
                Return False
            End If
        Next intCounter

        If dateValues.IndexOf("/") = -1 Then
            Return False
        End If

        temp_year = dateValues.Substring(dateValues.IndexOf("/") + 1, dateValues.Length - (dateValues.IndexOf("/") + 1))
        temp_month = dateValues.Substring(0, dateValues.IndexOf("/"))

        If CInt(temp_year) < 1950 OrElse CInt(temp_year) = 0 OrElse temp_year.Length = 0 OrElse temp_year.Length > 4 OrElse temp_month.Length = 0 OrElse temp_month.Length > 2 OrElse CInt(temp_month) > 12 OrElse CInt(temp_month) = 0 Then
            Return False
        End If

        temp_month = Right("00" & temp_month, 2)

        If Not inbox Is Nothing Then
            inbox.Text = temp_month + "/" + temp_year
        End If
        Return True

    End Function

    '========================================================================
    'Quote the list items with "'", eg. "a, b" to "'a', 'b'"
    '========================================================================
    Function dbConvList(ByVal aList As String) As String

        Dim listArray(), tmpList As String

        tmpList = ""

        listArray = listToArray(formatList(aList))

        If UBound(listArray) >= 0 Then
            tmpList = "'" & listArray(0) & "'"

            For i = 1 To UBound(listArray)
                tmpList = tmpList & ", '" & listArray(i) & "'"
            Next
        End If

        Return tmpList
    End Function

    Public Function isValidPWD2(ByVal pwdValue As String, Optional ByVal criteria As Integer = 4) As Boolean
        Dim regEx1 As New Regex("[\d]")
        Dim regEx2 As New Regex("[a-z]")
        Dim regEx3 As New Regex("[A-Z]")
        Dim regEx4 As New Regex("[\W]")
        Dim matchCnt As Integer = 0

        If regEx1.IsMatch(pwdValue) Then
            matchCnt += 1
        End If
        If regEx2.IsMatch(pwdValue) Then
            matchCnt += 1
        End If
        If regEx3.IsMatch(pwdValue) Then
            matchCnt += 1
        End If
        If regEx4.IsMatch(pwdValue) Then
            matchCnt += 1
        End If

        If matchCnt >= criteria Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Sub setCtrlValue(ByRef fieldCtrl As Control, ByVal fieldValue As String, Optional ByRef ctrlNumFormat As NumberFormat = Nothing)
        Dim vwCtrl As Control
        Dim vwText As String = ""
        Dim ctrlFieldValue As String = fieldValue

        Try
            If ctrlNumFormat IsNot Nothing Then
                ctrlFieldValue = formatNum(ctrlFieldValue, ctrlNumFormat)
            End If

            Select Case TypeName(fieldCtrl).ToUpper
                Case "LABEL"
                    CType(fieldCtrl, Label).Text = ctrlFieldValue

                Case "TEXTBOX"
                    CType(fieldCtrl, TextBox).Text = ctrlFieldValue

                Case "DROPDOWNLIST"
                    If CType(fieldCtrl, DropDownList).Items.FindByValue(ctrlFieldValue) IsNot Nothing Then
                        CType(fieldCtrl, DropDownList).SelectedValue = ctrlFieldValue
                        vwText = CType(fieldCtrl, DropDownList).SelectedItem.Text
                    ElseIf ctrlFieldValue = "" Then
                        vwText = ""
                    End If

                Case "CHECKBOXLIST"
                    If CType(fieldCtrl, CheckBoxList).Items.FindByValue(ctrlFieldValue) IsNot Nothing Then
                        CType(fieldCtrl, CheckBoxList).SelectedValue = ctrlFieldValue
                        vwText = CType(fieldCtrl, CheckBoxList).SelectedItem.Text
                    ElseIf ctrlFieldValue = "" Then
                        vwText = ""
                    End If

                Case "RADIOBUTTONLIST"
                    If CType(fieldCtrl, RadioButtonList).Items.FindByValue(ctrlFieldValue) IsNot Nothing Then
                        CType(fieldCtrl, RadioButtonList).SelectedValue = ctrlFieldValue
                        vwText = CType(fieldCtrl, RadioButtonList).SelectedItem.Text
                    ElseIf ctrlFieldValue = "" Then
                        vwText = ""
                    End If

                Case "HIDDENFIELD"
                    CType(fieldCtrl, HiddenField).Value = ctrlFieldValue
                Case "LINKBUTTON"
                    CType(fieldCtrl, LinkButton).Text = ctrlFieldValue
                Case Else
                    Throw New Exception(TypeName(fieldCtrl).ToUpper & " control type is not handled yet.")

            End Select

            vwCtrl = fieldCtrl.Parent.FindControl("vw_" & fieldCtrl.ID)

            If vwCtrl IsNot Nothing AndAlso (UCase(TypeName(vwCtrl)) = "LABEL" OrElse UCase(TypeName(vwCtrl)) = "TEXTBOX") Then
                If vwText <> "" Then
                    setCtrlValue(vwCtrl, vwText)
                    'CType(vwCtrl, Label).Text = vwText
                Else
                    setCtrlValue(vwCtrl, ctrlFieldValue)
                    'CType(vwCtrl, Label).Text = ctrlFieldValue
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Function getCtrlValue(ByRef fieldCtrl As Control) As String
        Select Case TypeName(fieldCtrl).ToUpper
            Case "TEXTBOX"
                Return DirectCast(fieldCtrl, TextBox).Text

            Case "DROPDOWNLIST"
                Return DirectCast(fieldCtrl, DropDownList).SelectedValue

            Case "RADIOBUTTONLIST"
                Return DirectCast(fieldCtrl, RadioButtonList).SelectedValue

            Case "CHECKBOXLIST"
                Return DirectCast(fieldCtrl, CheckBoxList).SelectedValue

            Case "HIDDENFIELD"
                Return DirectCast(fieldCtrl, HiddenField).Value

            Case "LABEL"
                Return DirectCast(fieldCtrl, Label).Text

            Case Else
                Return ""

        End Select
    End Function

    Public Function FindControlRecursive(ByVal parentCtrl As Control, ByVal ctrlID As String) As Control
        Dim tmpCtrl As Control = Nothing
        Dim resCtrl As Control

        If parentCtrl.ID = ctrlID Then
            tmpCtrl = parentCtrl
        Else
            For Each childCtrl In parentCtrl.Controls
                resCtrl = FindControlRecursive(childCtrl, ctrlID)
                If resCtrl IsNot Nothing Then
                    tmpCtrl = resCtrl
                End If
            Next
        End If

        Return tmpCtrl
    End Function

    Public Function formatNum(ByVal numValue As String, ByRef numFormat As NumberFormat) As String
        If numValue.Trim = "" Then
            Return ""
        Else
            Return FormatNumber(numValue, numFormat.decimalDigit, numFormat.useLeadDigit, numFormat.useParensNeg, numFormat.useGroupDigit)
        End If
    End Function

    Public Function formatNumRev(ByVal numValue As String, ByRef numFormat As NumberFormat) As String
        Dim tmpValue As String = numValue.Trim

        If tmpValue = "" OrElse numFormat Is Nothing Then
            Return ""
        Else
            If numFormat.useGroupDigit = TriState.True Then
                tmpValue = Replace(tmpValue, ",", "")
            End If

            If numFormat.useParensNeg = TriState.True Then
                tmpValue = "-" & Replace(Replace(tmpValue, "(", ""), ")", "")
            End If

            If numFormat.useLeadDigit = TriState.False Then
                If Left(tmpValue, 1) = "." Then
                    tmpValue = "0" & tmpValue
                End If
            End If

            Return tmpValue
        End If
    End Function

    Public Function formatNum(ByVal numValue As String, Optional ByVal decDig As Integer = -1,
                       Optional ByVal leadDig As TriState = TriState.UseDefault,
                       Optional ByVal parenNeg As TriState = TriState.UseDefault,
                       Optional ByVal groupDig As TriState = TriState.UseDefault) As String

        If numValue.Trim = "" Then
            Return ""
        Else
            Return FormatNumber(numValue, decDig, leadDig, parenNeg, groupDig)
        End If
    End Function

    Public Function csvEncode(ByVal value As String) As String
        If InStr(value, ",") > 0 Then
            Return """" & Replace(value, """", """""") & """"
        Else
            Return value
        End If
    End Function

    Public Function convLineBrk(ByVal aText As String) As String
        Return Replace(Replace(aText, Chr(13) & Chr(10), "<br />"), Chr(10), "<br />")
    End Function

    Public Class NumberFormat
        Public decimalDigit As Integer
        Public useLeadDigit, useParensNeg, useGroupDigit As TriState

        Public Sub New(Optional ByVal decDig As Integer = -1,
                       Optional ByVal leadDig As TriState = TriState.UseDefault,
                       Optional ByVal parenNeg As TriState = TriState.UseDefault,
                       Optional ByVal groupDig As TriState = TriState.UseDefault)
            decimalDigit = decDig
            useLeadDigit = leadDig
            useParensNeg = parenNeg
            useGroupDigit = groupDig
        End Sub
    End Class

    Public Sub buildCtrlDict(ByRef ctl As Control, ByVal fieldList As String, _
                              ByRef clientIDDict As Dictionary(Of String, String), _
                              ByRef fieldNameDict As Dictionary(Of String, String))
        Dim tmpArray As String()
        Dim i As Integer
        Dim tmpCtrl As Control

        Select Case TypeName(ctl).ToUpper
            Case "REPEATER"
                For Each repItemCtrl As RepeaterItem In ctl.Controls
                    Dim repCtlArray As New ArrayList

                    If fieldList = "" Then
                        Exit For
                    Else
                        tmpArray = listToArray(fieldList)

                        For i = 0 To UBound(tmpArray)
                            tmpCtrl = FindControlRecursive(repItemCtrl, tmpArray(i))

                            If tmpCtrl IsNot Nothing Then
                                clientIDDict.Add(tmpCtrl.ClientID, tmpArray(i) & FIELD_SEPERATOR & repItemCtrl.ItemIndex)

                                fieldNameDict.Add(tmpArray(i) & FIELD_SEPERATOR & repItemCtrl.ItemIndex, tmpCtrl.ClientID)
                            End If
                        Next
                    End If
                Next

            Case Else
                If fieldList = "" Then
                    clientIDDict.Add(ctl.ClientID, ctl.ID)

                    fieldNameDict.Add(ctl.ID, ctl.ClientID)
                Else
                    tmpArray = listToArray(fieldList)

                    For i = 0 To UBound(tmpArray)
                        tmpCtrl = FindControlRecursive(ctl, tmpArray(i))

                        If tmpCtrl IsNot Nothing Then
                            clientIDDict.Add(tmpCtrl.ClientID, tmpArray(i))

                            fieldNameDict.Add(tmpArray(i), tmpCtrl.ClientID)
                        End If
                    Next
                End If
        End Select


    End Sub

    Public Function getFieldIndex(ByVal clientID As String, ByRef clientIDDict As Dictionary(Of String, String)) As String
        Dim tmpItem As String

        If clientIDDict Is Nothing OrElse clientID = "" Then
            Return ""
        Else
            If clientIDDict.ContainsKey(clientID) Then
                tmpItem = clientIDDict.Item(clientID)

                If InStr(tmpItem, FIELD_SEPERATOR) > 0 Then
                    Return Split(tmpItem, FIELD_SEPERATOR)(1)
                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End If
    End Function

    Public Function getFieldClientID(ByVal fieldName As String, ByVal itemIndex As String, _
                                     ByRef fieldNameDict As Dictionary(Of String, String)) As String
        Dim tmpKey As String

        If fieldNameDict Is Nothing OrElse fieldName = "" Then
            Return ""
        Else
            If itemIndex <> "" Then
                tmpKey = fieldName & FIELD_SEPERATOR & itemIndex
            Else
                tmpKey = fieldName
            End If

            If fieldNameDict.ContainsKey(tmpKey) Then
                Return fieldNameDict.Item(tmpKey)
            Else
                Return ""
            End If
        End If
    End Function

    Public Function decodeEmptyCLng(ByVal aValue As String, ByVal defaultVal As Long) As Long
        If aValue = Nothing Or aValue = "" Then
            decodeEmptyCLng = defaultVal.ToString
        Else
            If Not IsNumeric(aValue) Then
                aValue = 0
            End If

            decodeEmptyCLng = aValue
        End If

        decodeEmptyCLng = CLng(decodeEmptyCLng)
    End Function

    Public Sub setSessionTempData(ByVal moduleAbb As String, ByVal dataName As String, ByRef dataValue As Object)
        Dim prefix As String = "_M_" & moduleAbb & "_TMP_"

        System.Web.HttpContext.Current.Session(prefix & dataName) = dataValue
    End Sub

    Function getSessionTempData(ByVal moduleAbb As String, ByVal dataName As String, ByRef defValue As Object) As Object
        Dim prefix As String = "_M_" & moduleAbb & "_TMP_"

        If System.Web.HttpContext.Current.Session(prefix & dataName) Is Nothing Then
            If defValue Is Nothing Then
                Return Nothing
            Else
                Return defValue
            End If
        Else
            Return System.Web.HttpContext.Current.Session(prefix & dataName)
        End If
    End Function

    Public Sub clearSessionTempData(ByVal moduleAbb As String)
        Dim prefix As String = "_M_" & moduleAbb & "_TMP_"
        Dim sessionItem As Object
        Dim removeItemList As New List(Of Object)

        For Each sessionItem In System.Web.HttpContext.Current.Session.Contents
            If InStr(sessionItem, prefix, CompareMethod.Text) > 0 Then
                removeItemList.Add(sessionItem)
            End If
        Next

        For Each sessionItem In removeItemList
            System.Web.HttpContext.Current.Session.Contents.Remove(sessionItem)
        Next
    End Sub

    Public Sub clearAllSessionTempData(Optional ByVal exceptModuleAbb As String = "")
        Dim prefix As String = "_M_" & exceptModuleAbb & "_TMP_"
        Dim sessionItem As Object
        Dim removeItemList As New List(Of Object)

        For Each sessionItem In System.Web.HttpContext.Current.Session.Contents
            If Left(sessionItem, 3) = "_M_" AndAlso InStr(sessionItem, "_TMP_", CompareMethod.Text) > 0 Then
                If exceptModuleAbb = "" OrElse InStr(sessionItem, prefix, CompareMethod.Text) <= 0 Then
                    removeItemList.Add(sessionItem)
                End If
            End If
        Next

        For Each sessionItem In removeItemList
            System.Web.HttpContext.Current.Session.Contents.Remove(sessionItem)
        Next
    End Sub


End Class
