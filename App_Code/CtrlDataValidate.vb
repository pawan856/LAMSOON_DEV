Imports Microsoft.VisualBasic

Public Class CtrlDataValidate
    Private gU As New GeneralUtils
    Private dataCheckList As List(Of DataFieldCheck)
    Private uiFun As New UIfunc
    Private invalidCtrlID As String
    Private invalidDesc As String

    Private formPage As System.Web.UI.Page

    Private updtCtrl As Control

    Public Sub New(ByRef aspxPage As System.Web.UI.Page)
        formPage = aspxPage

        dataCheckList = New List(Of DataFieldCheck)
    End Sub

    'Validate all fields
    Public Function validate() As Boolean
        Dim tmpFieldCheck As DataFieldCheck
        Dim tmpLabel As String
        Dim tmpCtrl, tmpLabelCtrl As Control
        Dim tmpCtrlArray As String()
        Dim i As Integer

        For Each tmpFieldCheck In dataCheckList
            If tmpFieldCheck.fieldCtrl Is Nothing Then

                If tmpFieldCheck.parentCtrl Is Nothing Then
                    Throw New Exception("Field control and field parent control is not set!")
                End If

                If tmpFieldCheck.fieldID = "" Then
                    Throw New Exception("Field ID not found for validation!")
                End If

                tmpCtrlArray = gU.listToArray(tmpFieldCheck.fieldID)

                Select Case UCase(TypeName(tmpFieldCheck.parentCtrl))
                    Case "REPEATER"
                        Dim tmpRepeater As Repeater = DirectCast(tmpFieldCheck.parentCtrl, Repeater)
                        Dim rptrItem As RepeaterItem

                        For Each rptrItem In tmpRepeater.Items
                            If rptrItem.Visible Then
                                tmpLabel = ""
                                tmpCtrl = rptrItem

                                For i = 0 To UBound(tmpCtrlArray)
                                    If tmpFieldCheck.fieldLabel <> "" Then
                                        tmpLabel = tmpFieldCheck.fieldLabel

                                        If InStr(tmpLabel, "#NUM#") > 0 Then
                                            tmpLabel = Replace(tmpLabel, "#NUM#", rptrItem.ItemIndex + 1)
                                        End If
                                    Else
                                        tmpLabelCtrl = tmpCtrl.FindControl("lbl_" & tmpCtrlArray(i))
                                        If tmpLabelCtrl IsNot Nothing AndAlso UCase(TypeName(tmpLabelCtrl)) = "LABEL" Then
                                            tmpLabel = DirectCast(tmpLabelCtrl, Label).Text
                                        End If
                                    End If
                                    'tmpCtrl = tmpCtrl.FindControl(tmpCtrlArray(i))
                                    tmpCtrl = gU.FindControlRecursive(tmpCtrl, tmpCtrlArray(i))
                                Next

                                If tmpCtrl Is Nothing Then
                                    Throw New Exception("Field control not found for validation! " & tmpFieldCheck.fieldID)
                                End If

                                If Not validateField(tmpCtrl, tmpLabel, tmpFieldCheck) Then
                                    Return False
                                End If
                            End If
                        Next

                    Case "GRIDVIEW"
                        Dim tmpGV As GridView = DirectCast(tmpFieldCheck.parentCtrl, GridView)
                        Dim tmpGVRow As GridViewRow
                        Dim tmpGVCell As TableCell

                        For Each tmpGVRow In tmpGV.Rows
                            If Not tmpGVRow.Visible Then
                                Continue For
                            End If

                            tmpLabel = ""
                            tmpCtrl = Nothing

                            For Each tmpGVCell In tmpGVRow.Cells
                                If Not tmpGVCell.Visible Then
                                    Continue For
                                End If

                                tmpCtrl = tmpGVCell

                                For i = 0 To UBound(tmpCtrlArray)
                                    If tmpFieldCheck.fieldLabel <> "" Then
                                        tmpLabel = tmpFieldCheck.fieldLabel

                                        If InStr(tmpLabel, "#NUM#") > 0 Then
                                            tmpLabel = Replace(tmpLabel, "#NUM#", tmpGVRow.RowIndex + 1)
                                        End If
                                    Else
                                        tmpLabelCtrl = tmpCtrl.FindControl("lbl_" & tmpCtrlArray(i))
                                        If tmpLabelCtrl IsNot Nothing AndAlso UCase(TypeName(tmpLabelCtrl)) = "LABEL" Then
                                            tmpLabel = DirectCast(tmpLabelCtrl, Label).Text
                                        End If
                                    End If
                                    'tmpCtrl = tmpCtrl.FindControl(tmpCtrlArray(i))
                                    tmpCtrl = gU.FindControlRecursive(tmpCtrl, tmpCtrlArray(i))
                                Next

                                If tmpCtrl IsNot Nothing Then
                                    Exit For
                                End If
                            Next

                            If tmpCtrl Is Nothing Then
                                Throw New Exception("Field control not found for validation! " & tmpFieldCheck.fieldID)
                            End If

                            If Not validateField(tmpCtrl, tmpLabel, tmpFieldCheck) Then
                                Return False
                            End If
                        Next

                    Case Else
                        tmpLabel = ""
                        tmpCtrl = tmpFieldCheck.parentCtrl

                        For i = 0 To UBound(tmpCtrlArray)
                            If tmpFieldCheck.fieldLabel <> "" Then
                                tmpLabel = tmpFieldCheck.fieldLabel
                            Else
                                tmpLabelCtrl = tmpCtrl.FindControl("lbl_" & tmpCtrlArray(i))
                                If tmpLabelCtrl IsNot Nothing AndAlso UCase(TypeName(tmpLabelCtrl)) = "LABEL" Then
                                    tmpLabel = DirectCast(tmpLabelCtrl, Label).Text
                                End If
                            End If
                            'tmpCtrl = tmpCtrl.FindControl(tmpCtrlArray(i))
                            tmpCtrl = gU.FindControlRecursive(tmpCtrl, tmpCtrlArray(i))
                        Next

                        If tmpCtrl Is Nothing Then
                            Throw New Exception("Field control not found for validation! " & tmpFieldCheck.fieldID)
                        End If

                        If Not validateField(tmpCtrl, tmpLabel, tmpFieldCheck) Then
                            Return False
                        End If

                End Select
            Else
                tmpCtrl = tmpFieldCheck.fieldCtrl

                tmpLabel = ""

                If tmpFieldCheck.fieldLabel <> "" Then
                    tmpLabel = tmpFieldCheck.fieldLabel
                ElseIf tmpFieldCheck.parentCtrl IsNot Nothing Then
                    tmpLabelCtrl = tmpCtrl.FindControl("lbl_" & tmpCtrl.ID)
                    If tmpLabelCtrl IsNot Nothing AndAlso UCase(TypeName(tmpLabelCtrl)) = "LABEL" Then
                        tmpLabel = DirectCast(tmpLabelCtrl, Label).Text
                    End If
                End If

                If Not validateField(tmpCtrl, tmpLabel, tmpFieldCheck) Then
                    Return False
                End If
            End If
        Next

        Return True
    End Function

    Private Sub displayMsg(ByRef fieldCtrl As Control, ByVal alertMsg As String)
        invalidCtrlID = fieldCtrl.ID
        invalidDesc = Replace(Replace(Replace(alertMsg, vbNewLine, "\n"), "<br>", "\n"), "<br />", "\n")

        If updtCtrl IsNot Nothing Then
            ScriptManager.RegisterStartupScript(updtCtrl, updtCtrl.GetType, "WARN_" & updtCtrl.ID, "alert('" & gU.jsString(invalidDesc) & "');", True)
        Else
            uiFun.displayMsg(formPage, "", invalidDesc, HttpContext.Current.Session("gLang"))
        End If

        If UCase(TypeName(fieldCtrl)) <> "HIDDENFIELD" Then
            fieldCtrl.Focus()
        End If
    End Sub

    Private Function validateField(ByRef fieldCtrl As Control, ByVal fieldLabel As String, ByRef fieldCheck As DataFieldCheck) As Boolean
        Dim fieldValue As String
        Dim dotPos As Integer
        Dim fieldDesc As String

        If fieldCtrl.Visible = False OrElse fieldCtrl.Parent.Visible = False Then
            Return True
        End If

        If fieldCtrl Is Nothing Then
            Throw New Exception("Field control is nothing!")
        End If

        fieldDesc = Trim(fieldLabel)

        If Right(fieldDesc, 1) = ":" Then
            fieldDesc = Trim(Left(fieldDesc, Len(fieldDesc) - 1))
        End If

        fieldValue = gU.getCtrlValue(fieldCtrl)

        If fieldCheck.isNotEmpty Then
            If TypeName(fieldCtrl).ToUpper = "FILEUPLOAD" Then
                If Not DirectCast(fieldCtrl, FileUpload).HasFile Then
                    displayMsg(fieldCtrl, "Entry required! " & fieldDesc)
                    Return False
                End If
            ElseIf fieldValue.Trim = "" Then
                displayMsg(fieldCtrl, "Entry required! " & fieldDesc)
                Return False
            End If

            'invalidCtrlID = fieldCtrl.ID
            'invalidDesc = "Entry required! " & fieldDesc

            'uiFun.displayMsg(formPage, "", invalidDesc, HttpContext.Current.Session("gLang"))
            'If UCase(TypeName(fieldCtrl)) <> "HIDDENFIELD" Then
            '    fieldCtrl.Focus()
            'End If

        End If

        If fieldCheck.maxLength > 0 AndAlso Len(fieldValue) > fieldCheck.maxLength Then
            'invalidCtrlID = fieldCtrl.ID
            'invalidDesc = "Field value too large, maximum length is " & fieldCheck.maxLength & "! " & fieldDesc

            'uiFun.displayMsg(formPage, "", invalidDesc, HttpContext.Current.Session("gLang"))
            'If UCase(TypeName(fieldCtrl)) <> "HIDDENFIELD" Then
            '    fieldCtrl.Focus()
            'End If
            displayMsg(fieldCtrl, "Field value too large, maximum length is " & fieldCheck.maxLength & "! " & fieldDesc)
            Return False
        End If

        If fieldValue <> "" Then
            Select Case UCase(fieldCheck.dataType)
                Case "EMAIL"
                    If Not gU.isValidEmail(fieldValue) Then
                        displayMsg(fieldCtrl, "Invalid email! " & fieldDesc)
                        Return False
                    End If

                Case "DATE"
                    If Not gU.isValidDate(fieldValue) Then
                        'invalidCtrlID = fieldCtrl.ID
                        'invalidDesc = "Invalid date! " & fieldDesc

                        'uiFun.displayMsg(formPage, "", invalidDesc, HttpContext.Current.Session("gLang"))
                        'If UCase(TypeName(fieldCtrl)) <> "HIDDENFIELD" Then
                        '    fieldCtrl.Focus()
                        'End If
                        displayMsg(fieldCtrl, "Invalid date! " & fieldDesc)
                        Return False
                    End If

                Case "INTEGER"
                    fieldValue = Replace(fieldValue, ",", "")

                    If Not gU.IsWholeNumber(fieldValue) Then
                        'invalidCtrlID = fieldCtrl.ID
                        'invalidDesc = "Invalid whole number! " & fieldDesc

                        'uiFun.displayMsg(formPage, "", invalidDesc, HttpContext.Current.Session("gLang"))
                        'If UCase(TypeName(fieldCtrl)) <> "HIDDENFIELD" Then
                        '    fieldCtrl.Focus()
                        'End If
                        displayMsg(fieldCtrl, "Invalid whole number! " & fieldDesc)
                        Return False
                    End If

                    dotPos = InStr(fieldValue, ".")

                    If dotPos > 1 Then
                        fieldValue = Left(fieldValue, dotPos - 1)
                    ElseIf dotPos = 1 Then
                        fieldValue = "0"
                    End If

                    If fieldCheck.maxLength > 0 AndAlso Len(fieldValue) Then
                        'invalidCtrlID = fieldCtrl.ID
                        'invalidDesc = "Field value too large, maximum precision is " & fieldCheck.maxLength & "! " & fieldDesc

                        'uiFun.displayMsg(formPage, "", invalidDesc, HttpContext.Current.Session("gLang"))
                        'If UCase(TypeName(fieldCtrl)) <> "HIDDENFIELD" Then
                        '    fieldCtrl.Focus()
                        'End If
                        displayMsg(fieldCtrl, "Field value too large, maximum precision is " & fieldCheck.maxLength & "! " & fieldDesc)
                        Return False
                    End If

                Case "FLOAT"
                    fieldValue = Replace(fieldValue, ",", "")

                    If Not gU.isDecimal(fieldValue) Then
                        'invalidCtrlID = fieldCtrl.ID
                        'invalidDesc = "Invalid number! " & fieldDesc

                        'uiFun.displayMsg(formPage, "", invalidDesc, HttpContext.Current.Session("gLang"))
                        'If UCase(TypeName(fieldCtrl)) <> "HIDDENFIELD" Then
                        '    fieldCtrl.Focus()
                        'End If
                        displayMsg(fieldCtrl, "Invalid number! " & fieldDesc)
                        Return False
                    End If

                    dotPos = InStr(fieldValue, ".")

                    If dotPos > 1 Then
                        fieldValue = Left(fieldValue, dotPos - 1)
                    ElseIf dotPos = 1 Then
                        fieldValue = "0"
                    End If

                    If fieldCheck.maxLength > 0 AndAlso Len(fieldValue) Then
                        'invalidCtrlID = fieldCtrl.ID
                        'invalidDesc = "Field value too large, maximum precision is " & fieldCheck.maxLength & "! " & fieldDesc

                        'uiFun.displayMsg(formPage, "", invalidDesc, HttpContext.Current.Session("gLang"))
                        'If UCase(TypeName(fieldCtrl)) <> "HIDDENFIELD" Then
                        '    fieldCtrl.Focus()
                        'End If
                        displayMsg(fieldCtrl, "Field value too large, maximum precision is " & fieldCheck.maxLength & "! " & fieldDesc)
                        Return False
                    End If

            End Select
        End If

        Return True
    End Function

    'Build JavaScript checking string for field
    Private Function getValidateStr(ByRef fieldCtrl As Control, ByVal fieldLabel As String, ByRef fieldCheck As DataFieldCheck) As String
        Dim jsStr As String
        Dim fieldDesc As String

        If fieldCtrl Is Nothing Then
            Throw New Exception("Field control is nothing!")
        End If

        jsStr = ""

        fieldDesc = Trim(fieldLabel)

        If Right(fieldDesc, 1) = ":" Then
            fieldDesc = Trim(Left(fieldDesc, Len(fieldDesc) - 1))
        End If

        If UCase(TypeName(fieldCtrl)) = "HIDDENFIELD" Then
            Return ""
        End If

        If fieldCheck.isNotEmpty Then
            jsStr += "if(!validateExistField(document.getElementById(""" & fieldCtrl.ClientID & """), """ & fieldDesc & """, ""|NOT_EMPTY|"")) return false; "
        End If

        If fieldCheck.maxLength > 0 Then
            jsStr += "if(!validateExistField(document.getElementById(""" & fieldCtrl.ClientID & """), """ & fieldDesc & """, ""|MAX_LENGTH:" & fieldCheck.maxLength & "|"")) return false; "
        End If

        Select Case UCase(fieldCheck.dataType)
            Case "EMAIL"
                jsStr += "if(!validateExistField(document.getElementById(""" & fieldCtrl.ClientID & """), """ & fieldDesc & """, ""|EMAIL|"")) return false; "

            Case "DATE"
                jsStr += "if(!validateExistField(document.getElementById(""" & fieldCtrl.ClientID & """), """ & fieldDesc & """, ""|DATE|"")) return false; "

            Case "INTEGER"
                jsStr += "if(!validateExistField(document.getElementById(""" & fieldCtrl.ClientID & """), """ & fieldDesc & """, ""|INTEGER2|"")) return false; "

                If fieldCheck.maxLength > 0 Then
                    jsStr += "if(!validateExistField(document.getElementById(""" & fieldCtrl.ClientID & """), """ & fieldDesc & """, ""|FLOAT2:" & fieldCheck.maxLength & "|"")) return false; "
                End If

            Case "FLOAT"
                If fieldCheck.maxLength > 0 Then
                    jsStr += "if(!validateExistField(document.getElementById(""" & fieldCtrl.ClientID & """), """ & fieldDesc & """, ""|FLOAT2:" & fieldCheck.maxLength & "|"")) return false; "
                Else
                    jsStr += "if(!validateExistField(document.getElementById(""" & fieldCtrl.ClientID & """), """ & fieldDesc & """, ""|FLOAT2|"")) return false; "
                End If

        End Select

        Return jsStr
    End Function


    'Generate JavaScript checking function
    Public Function getValidateJS(Optional ByVal funcName As String = "") As String
        Dim jsFuncStr As String = ""

        jsFuncStr = "function "

        If funcName = "" Then
            jsFuncStr += "validateFieldFormat"
        Else
            jsFuncStr += funcName
        End If

        jsFuncStr += "() {"

        '************************************************************************************************************************
        Dim tmpFieldCheck As DataFieldCheck
        Dim tmpLabel As String
        Dim tmpCtrl, tmpLabelCtrl As Control
        Dim tmpCtrlArray As String()
        Dim i As Integer

        For Each tmpFieldCheck In dataCheckList
            If tmpFieldCheck.fieldCtrl Is Nothing Then

                If tmpFieldCheck.parentCtrl Is Nothing Then
                    Throw New Exception("Field control and field parent control is not set!")
                End If

                If tmpFieldCheck.fieldID = "" Then
                    Throw New Exception("Field ID not found for validation!")
                End If

                tmpCtrlArray = gU.listToArray(tmpFieldCheck.fieldID)

                Select Case UCase(TypeName(tmpFieldCheck.parentCtrl))
                    Case "REPEATER"
                        Dim tmpRepeater As Repeater = DirectCast(tmpFieldCheck.parentCtrl, Repeater)
                        Dim rptrItem As RepeaterItem

                        For Each rptrItem In tmpRepeater.Items
                            tmpLabel = ""
                            tmpCtrl = rptrItem

                            For i = 0 To UBound(tmpCtrlArray)
                                If tmpFieldCheck.fieldLabel <> "" Then
                                    tmpLabel = tmpFieldCheck.fieldLabel

                                    If InStr(tmpLabel, "#NUM#") > 0 Then
                                        tmpLabel = Replace(tmpLabel, "#NUM#", rptrItem.ItemIndex + 1)
                                    End If
                                Else
                                    tmpLabelCtrl = tmpCtrl.FindControl("lbl_" & tmpCtrlArray(i))
                                    If tmpLabelCtrl IsNot Nothing AndAlso UCase(TypeName(tmpLabelCtrl)) = "LABEL" Then
                                        tmpLabel = DirectCast(tmpLabelCtrl, Label).Text
                                    End If
                                End If
                                'tmpCtrl = tmpCtrl.FindControl(tmpCtrlArray(i))
                                tmpCtrl = gU.FindControlRecursive(tmpCtrl, tmpCtrlArray(i))
                            Next

                            If tmpCtrl Is Nothing Then
                                Throw New Exception("Field control not found for validation! " & tmpFieldCheck.fieldID)
                            End If

                            jsFuncStr += getValidateStr(tmpCtrl, tmpLabel, tmpFieldCheck)
                            'If Not validateField(tmpCtrl, tmpLabel, tmpFieldCheck) Then
                            '    Return False
                            'End If
                        Next

                    Case "GRIDVIEW"
                        Dim tmpGV As GridView = DirectCast(tmpFieldCheck.parentCtrl, GridView)
                        Dim tmpGVRow As GridViewRow
                        Dim tmpGVCell As TableCell

                        For Each tmpGVRow In tmpGV.Rows
                            For Each tmpGVCell In tmpGVRow.Cells
                                tmpLabel = ""
                                tmpCtrl = tmpGVCell

                                For i = 0 To UBound(tmpCtrlArray)
                                    If tmpFieldCheck.fieldLabel <> "" Then
                                        tmpLabel = tmpFieldCheck.fieldLabel

                                        If InStr(tmpLabel, "#NUM#") > 0 Then
                                            tmpLabel = Replace(tmpLabel, "#NUM#", tmpGVRow.RowIndex + 1)
                                        End If
                                    Else
                                        tmpLabelCtrl = tmpCtrl.FindControl("lbl_" & tmpCtrlArray(i))
                                        If tmpLabelCtrl IsNot Nothing AndAlso UCase(TypeName(tmpLabelCtrl)) = "LABEL" Then
                                            tmpLabel = DirectCast(tmpLabelCtrl, Label).Text
                                        End If
                                    End If
                                    'tmpCtrl = tmpCtrl.FindControl(tmpCtrlArray(i))
                                    tmpCtrl = gU.FindControlRecursive(tmpCtrl, tmpCtrlArray(i))
                                Next

                                If tmpCtrl Is Nothing Then
                                    Throw New Exception("Field control not found for validation! " & tmpFieldCheck.fieldID)
                                End If

                                jsFuncStr += getValidateStr(tmpCtrl, tmpLabel, tmpFieldCheck)
                            Next
                        Next

                    Case Else
                        tmpLabel = ""
                        tmpCtrl = tmpFieldCheck.parentCtrl

                        For i = 0 To UBound(tmpCtrlArray)
                            If tmpFieldCheck.fieldLabel <> "" Then
                                tmpLabel = tmpFieldCheck.fieldLabel
                            Else
                                tmpLabelCtrl = tmpCtrl.FindControl("lbl_" & tmpCtrlArray(i))
                                If tmpLabelCtrl IsNot Nothing AndAlso UCase(TypeName(tmpLabelCtrl)) = "LABEL" Then
                                    tmpLabel = DirectCast(tmpLabelCtrl, Label).Text
                                End If
                            End If
                            'tmpCtrl = tmpCtrl.FindControl(tmpCtrlArray(i))
                            tmpCtrl = gU.FindControlRecursive(tmpCtrl, tmpCtrlArray(i))
                        Next

                        If tmpCtrl Is Nothing Then
                            Throw New Exception("Field control not found for validation! " & tmpFieldCheck.fieldID)
                        End If

                        jsFuncStr += getValidateStr(tmpCtrl, tmpLabel, tmpFieldCheck)

                End Select
            Else
                tmpCtrl = tmpFieldCheck.fieldCtrl

                tmpLabel = ""

                If tmpFieldCheck.fieldLabel <> "" Then
                    tmpLabel = tmpFieldCheck.fieldLabel
                ElseIf tmpFieldCheck.parentCtrl IsNot Nothing Then
                    tmpLabelCtrl = tmpCtrl.FindControl("lbl_" & tmpCtrl.ID)
                    If tmpLabelCtrl IsNot Nothing AndAlso UCase(TypeName(tmpLabelCtrl)) = "LABEL" Then
                        tmpLabel = DirectCast(tmpLabelCtrl, Label).Text
                    End If
                End If

                jsFuncStr += getValidateStr(tmpCtrl, tmpLabel, tmpFieldCheck)
            End If
        Next
        '************************************************************************************************************************


        jsFuncStr += " return true;}"

        Return jsFuncStr
    End Function

    Public Function validateSingleField(ByRef fieldCtrl As Control, fieldLabel As String, ByVal checkType As String,
                                 Optional ByVal isMandatory As Boolean = False, Optional ByVal maxLength As Integer = 0) As Boolean

        Dim tmpFieldCheck As New DataFieldCheck

        tmpFieldCheck.dataType = UCase(checkType.Trim)
        tmpFieldCheck.fieldCtrl = fieldCtrl
        If Right(fieldLabel.Trim, 1) = ":" AndAlso fieldLabel.Trim.Length > 1 Then
            tmpFieldCheck.fieldLabel = Left(fieldLabel.Trim, fieldLabel.Trim.Length - 1).Trim
        Else
            tmpFieldCheck.fieldLabel = fieldLabel.Trim
        End If
        tmpFieldCheck.parentCtrl = Nothing
        tmpFieldCheck.fieldID = fieldCtrl.ID
        tmpFieldCheck.isNotEmpty = isMandatory
        tmpFieldCheck.maxLength = maxLength

        Return validateField(fieldCtrl, tmpFieldCheck.fieldLabel, tmpFieldCheck)
    End Function

    Public Sub addFieldCheck(ByVal checkType As String, Optional ByRef fieldCtrl As Control = Nothing, Optional fieldLabel As String = "",
                             Optional ByRef parentCtrl As Control = Nothing, Optional ByVal fieldID As String = "",
                             Optional ByVal isMandatory As Boolean = False, Optional ByVal maxLength As Integer = 0)

        Dim tmpFieldCheck As New DataFieldCheck

        tmpFieldCheck.dataType = UCase(checkType.Trim)
        tmpFieldCheck.fieldCtrl = fieldCtrl
        If Right(fieldLabel.Trim, 1) = ":" AndAlso fieldLabel.Trim.Length > 1 Then
            tmpFieldCheck.fieldLabel = Left(fieldLabel.Trim, fieldLabel.Trim.Length - 1).Trim
        Else
            tmpFieldCheck.fieldLabel = fieldLabel.Trim
        End If
        tmpFieldCheck.parentCtrl = parentCtrl
        tmpFieldCheck.fieldID = fieldID
        tmpFieldCheck.isNotEmpty = isMandatory
        tmpFieldCheck.maxLength = maxLength

        dataCheckList.Add(tmpFieldCheck)

    End Sub


    '============================================================================================================================
    'Properties =================================================================================================================
    Public Property DataFieldCheckList() As List(Of DataFieldCheck)
        Get
            Return dataCheckList
        End Get
        Set(value As List(Of DataFieldCheck))
            dataCheckList = value
        End Set
    End Property

    Public Property AspxPage() As System.Web.UI.Page
        Get
            Return formPage
        End Get
        Set(value As System.Web.UI.Page)
            formPage = value
        End Set
    End Property

    Public Property UpdateControl() As Control
        Get
            Return updtCtrl
        End Get
        Set(value As Control)
            updtCtrl = value
        End Set
    End Property



    Public ReadOnly Property InvalidControlID() As String
        Get
            Return invalidCtrlID
        End Get
    End Property

    Public ReadOnly Property InvalidMessage() As String
        Get
            Return invalidDesc
        End Get
    End Property



    'End of Properties ==========================================================================================================
    '============================================================================================================================

    Public Structure DataFieldCheck
        Dim fieldCtrl As Control
        Dim fieldLabel As String
        Dim parentCtrl As Control
        Dim fieldID As String
        Dim dataType As String
        Dim isNotEmpty As Boolean
        Dim maxLength As Integer
    End Structure
End Class
