Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

Public Class CtrlDataMap
    Private gU As New GeneralUtils

    Private dataParentCtrl As Control

    Private ctrlPrefix As String = ""

    'Use field name as key, for edit field only
    Private ctrlDict As Dictionary(Of String, DataFieldCtrl)

    'Use control ID as key, for edit field only
    Private fieldDict As Dictionary(Of String, String)

    'For display only
    Private ctrlList As List(Of DataFieldCtrl)

    Public Sub New(Optional ByRef dataParentControl As Control = Nothing)
        ctrlDict = New Dictionary(Of String, DataFieldCtrl)
        fieldDict = New Dictionary(Of String, String)
        ctrlList = New List(Of DataFieldCtrl)
        dataParentCtrl = dataParentControl
    End Sub

    Public Sub addCtrl(ByVal fieldName As String, ByVal fieldID As String, ByVal isEditField As Boolean,
                       Optional ByVal fieldType As String = "", Optional ByRef fieldFormat As GeneralUtils.NumberFormat = Nothing,
                       Optional ByVal saveFlag As Boolean = True)
        Dim tmpFieldCtrl As New DataFieldCtrl

        tmpFieldCtrl.fieldName = UCase(fieldName)
        tmpFieldCtrl.fieldCtrlID = fieldID
        tmpFieldCtrl.fieldType = fieldType
        tmpFieldCtrl.fieldFormat = fieldFormat
        tmpFieldCtrl.saveFlag = saveFlag

        If isEditField Then
            If Not ctrlDict.ContainsKey(UCase(fieldName)) Then
                ctrlDict.Add(UCase(fieldName), tmpFieldCtrl)

                If Not fieldDict.ContainsKey(fieldID) Then
                    fieldDict.Add(fieldID, UCase(fieldName))
                End If
            End If
        Else
            ctrlList.Add(tmpFieldCtrl)
        End If
    End Sub

    Public Sub addCtrl(ByVal fieldName As String, ByRef fieldContrl As Control, ByVal isEditField As Boolean,
                       Optional ByVal fieldType As String = "", Optional ByRef fieldFormat As GeneralUtils.NumberFormat = Nothing,
                       Optional ByVal saveFlag As Boolean = True)
        Dim tmpFieldCtrl As New DataFieldCtrl

        tmpFieldCtrl.fieldName = UCase(fieldName)
        tmpFieldCtrl.fieldCtrl = fieldContrl
        tmpFieldCtrl.fieldCtrlID = fieldContrl.ID
        tmpFieldCtrl.fieldType = fieldType
        tmpFieldCtrl.fieldFormat = fieldFormat
        tmpFieldCtrl.saveFlag = saveFlag

        If isEditField Then
            If Not ctrlDict.ContainsKey(UCase(fieldName)) Then
                ctrlDict.Add(UCase(fieldName), tmpFieldCtrl)

                If Not fieldDict.ContainsKey(fieldContrl.ID) Then
                    fieldDict.Add(fieldContrl.ID, UCase(fieldName))
                End If
            End If
        Else
            ctrlList.Add(tmpFieldCtrl)
        End If
    End Sub

    'Public Sub setCtrlValue(ByRef dtRow As Data.DataRow)
    '    Dim dictItem As KeyValuePair(Of String, DataFieldCtrl)
    '    Dim i As Integer
    '    Dim tmpFieldID As String

    '    For Each dictItem In ctrlDict
    '        If dictItem.Value.fieldCtrl Is Nothing Then
    '            If dataParentCtrl IsNot Nothing Then
    '                If ctrlPrefix <> "" Then
    '                    tmpFieldID = ctrlPrefix & dictItem.Value.fieldCtrlID
    '                Else
    '                    tmpFieldID = dictItem.Value.fieldCtrlID
    '                End If
    '                gU.setCtrlValue(dataParentCtrl.FindControl(tmpFieldID), dtRow.Item(dictItem.Value.fieldName).ToString.Trim)
    '            End If
    '        Else
    '            gU.setCtrlValue(dictItem.Value.fieldCtrl, dtRow.Item(dictItem.Value.fieldName).ToString.Trim)
    '        End If
    '    Next

    '    For i = 0 To ctrlList.Count - 1
    '        If ctrlList(i).fieldCtrl Is Nothing Then
    '            If dataParentCtrl IsNot Nothing Then
    '                If ctrlPrefix <> "" Then
    '                    tmpFieldID = ctrlPrefix & ctrlList(i).fieldCtrlID
    '                Else
    '                    tmpFieldID = ctrlList(i).fieldCtrlID
    '                End If
    '                gU.setCtrlValue(dataParentCtrl.FindControl(tmpFieldID), dtRow.Item(ctrlList(i).fieldName).ToString.Trim)
    '            End If
    '        Else
    '            gU.setCtrlValue(ctrlList(i).fieldCtrl, dtRow.Item(ctrlList(i).fieldName).ToString.Trim)
    '        End If
    '    Next
    'End Sub

    Public Function getFormCtrlValue(ByVal fieldName As String) As String
        Dim tmpFieldID As String
        Dim tmpCtrl As Control

        If ctrlDict.ContainsKey(fieldName) Then
            If ctrlDict.Item(fieldName).fieldCtrl IsNot Nothing Then
                Return gU.getCtrlValue(ctrlDict.Item(fieldName).fieldCtrl)
            Else
                If ctrlPrefix <> "" Then
                    tmpFieldID = ctrlPrefix & ctrlDict.Item(fieldName).fieldCtrlID
                Else
                    tmpFieldID = ctrlDict.Item(fieldName).fieldCtrlID
                End If
                'tmpCtrl = dataParentCtrl.FindControl(tmpFieldID)
                tmpCtrl = gU.FindControlRecursive(dataParentCtrl, tmpFieldID)
                Return gU.getCtrlValue(tmpCtrl)
            End If
        Else
            Throw New Exception("Field name not found: " & fieldName)
        End If
    End Function

    Public Sub setCtrlValue(ByRef dtItem As Object)
        Dim i As Integer
        'Dim keys As Dictionary(Of String, DataFieldCtrl).KeyCollection
        Dim dictItem As KeyValuePair(Of String, DataFieldCtrl)
        Dim tmpFieldID As String
        Dim tmpCtrl As Control

        'keys = ctrlDict.Keys

        If dataParentCtrl Is Nothing Then
            Exit Sub
        End If

        Select Case UCase(TypeName(dataParentCtrl))
            Case "REPEATER"
                Dim tmpRepItem As RepeaterItem

                tmpRepItem = CType(dtItem, RepeaterItem)

                For Each dictItem In ctrlDict
                    If ctrlPrefix <> "" Then
                        tmpFieldID = ctrlPrefix & dictItem.Value.fieldCtrlID
                    Else
                        tmpFieldID = dictItem.Value.fieldCtrlID
                    End If
                    'tmpCtrl = tmpRepItem.FindControl(tmpFieldID)
                    tmpCtrl = gU.FindControlRecursive(tmpRepItem, tmpFieldID)
                    gU.setCtrlValue(tmpCtrl, DataBinder.Eval(tmpRepItem.DataItem, dictItem.Value.fieldName).ToString.Trim, dictItem.Value.fieldFormat)
                Next

                'For i = 0 To keys.Count - 1
                '    gU.setCtrlValue(tmpRepItem.FindControl(ctrlDict.Item(keys(i)).fieldCtrlID), DataBinder.Eval(tmpRepItem.DataItem, ctrlDict.Item(keys(i)).fieldName).ToString.Trim)
                'Next

                For i = 0 To ctrlList.Count - 1
                    If ctrlPrefix <> "" Then
                        tmpFieldID = ctrlPrefix & ctrlList(i).fieldCtrlID
                    Else
                        tmpFieldID = ctrlList(i).fieldCtrlID
                    End If
                    'tmpCtrl = tmpRepItem.FindControl(tmpFieldID)
                    tmpCtrl = gU.FindControlRecursive(tmpRepItem, tmpFieldID)
                    gU.setCtrlValue(tmpCtrl, DataBinder.Eval(tmpRepItem.DataItem, ctrlList(i).fieldName).ToString.Trim, ctrlList(i).fieldFormat)
                Next

            Case "GRIDVIEW"
                Dim tmpGVRow As GridViewRow

                tmpGVRow = CType(dtItem, GridViewRow)

                For Each dictItem In ctrlDict
                    If ctrlPrefix <> "" Then
                        tmpFieldID = ctrlPrefix & dictItem.Value.fieldCtrlID
                    Else
                        tmpFieldID = dictItem.Value.fieldCtrlID
                    End If
                    'tmpCtrl = tmpGVRow.FindControl(tmpFieldID)
                    tmpCtrl = gU.FindControlRecursive(tmpGVRow, tmpFieldID)
                    gU.setCtrlValue(tmpCtrl, DataBinder.Eval(tmpGVRow.DataItem, dictItem.Value.fieldName).ToString.Trim, dictItem.Value.fieldFormat)
                Next

                For i = 0 To ctrlList.Count - 1
                    If ctrlPrefix <> "" Then
                        tmpFieldID = ctrlPrefix & ctrlList(i).fieldCtrlID
                    Else
                        tmpFieldID = ctrlList(i).fieldCtrlID
                    End If
                    'tmpCtrl = tmpGVRow.FindControl(tmpFieldID)
                    tmpCtrl = gU.FindControlRecursive(tmpGVRow, tmpFieldID)
                    gU.setCtrlValue(tmpCtrl, DataBinder.Eval(tmpGVRow.DataItem, ctrlList(i).fieldName).ToString.Trim, ctrlList(i).fieldFormat)
                Next

            Case Else
                Dim dtRow As Data.DataRow

                dtRow = CType(dtItem, Data.DataRow)

                For Each dictItem In ctrlDict
                    If dictItem.Value.fieldCtrl Is Nothing Then
                        If dataParentCtrl IsNot Nothing Then
                            If ctrlPrefix <> "" Then
                                tmpFieldID = ctrlPrefix & dictItem.Value.fieldCtrlID
                            Else
                                tmpFieldID = dictItem.Value.fieldCtrlID
                            End If
                            'tmpCtrl = dataParentCtrl.FindControl(tmpFieldID)
                            tmpCtrl = gU.FindControlRecursive(dataParentCtrl, tmpFieldID)
                            gU.setCtrlValue(tmpCtrl, dtRow.Item(dictItem.Value.fieldName).ToString.Trim, dictItem.Value.fieldFormat)
                        End If
                    Else
                        gU.setCtrlValue(dictItem.Value.fieldCtrl, dtRow.Item(dictItem.Value.fieldName).ToString.Trim, dictItem.Value.fieldFormat)
                    End If
                Next

                For i = 0 To ctrlList.Count - 1
                    If ctrlList(i).fieldCtrl Is Nothing Then
                        If dataParentCtrl IsNot Nothing Then
                            If ctrlPrefix <> "" Then
                                tmpFieldID = ctrlPrefix & ctrlList(i).fieldCtrlID
                            Else
                                tmpFieldID = ctrlList(i).fieldCtrlID
                            End If
                            'tmpCtrl = dataParentCtrl.FindControl(tmpFieldID)
                            tmpCtrl = gU.FindControlRecursive(dataParentCtrl, tmpFieldID)
                            gU.setCtrlValue(tmpCtrl, dtRow.Item(ctrlList(i).fieldName).ToString.Trim, ctrlList(i).fieldFormat)
                        End If
                    Else
                        gU.setCtrlValue(ctrlList(i).fieldCtrl, dtRow.Item(ctrlList(i).fieldName).ToString.Trim, ctrlList(i).fieldFormat)
                    End If
                Next
        End Select
    End Sub

    Public Sub disableAutoComplete()
        Dim cmdDataDict As Dictionary(Of String, CtrlDataMap.DataFieldCtrl)

        cmdDataDict = ctrlDict

        'Reset control class first
        For Each tmpDataFieldCtrl In cmdDataDict
            Select Case TypeName(tmpDataFieldCtrl.Value.fieldCtrl).ToUpper
                Case "TEXTBOX"
                    DirectCast(tmpDataFieldCtrl.Value.fieldCtrl, TextBox).AutoCompleteType = AutoCompleteType.Disabled

            End Select
        Next
    End Sub


    'Public Function saveCtrlValue(ByVal tableName As String, ByRef dtItem As Object, Optional ByVal dirtyList As String = "", _
    '                              Optional ByRef cnn As SqlConnection = Nothing, Optional ByRef transaction As SqlTransaction = Nothing) As Boolean


    '    Return True
    'End Function


    Public Property ControlPrefix() As String
        Get
            Return ctrlPrefix
        End Get
        Set(value As String)
            ctrlPrefix = value
        End Set
    End Property

    Public Property DataParentControl() As Control
        Get
            Return dataParentCtrl
        End Get
        Set(value As Control)
            dataParentCtrl = value
        End Set
    End Property

    'Use field name as key, for edit field only
    Public Property DataCtrlDict() As Dictionary(Of String, DataFieldCtrl)
        Get
            Return ctrlDict
        End Get
        Set(value As Dictionary(Of String, DataFieldCtrl))
            ctrlDict = value
        End Set
    End Property

    'Use control ID as key, for edit field only
    Public Property fieldCtrlDict() As Dictionary(Of String, String)
        Get
            Return fieldDict
        End Get
        Set(value As Dictionary(Of String, String))
            fieldDict = value
        End Set
    End Property

    Public Structure DataFieldCtrl
        Dim fieldCtrl As Control
        Dim fieldName As String
        Dim fieldCtrlID As String
        Dim fieldType As String
        Dim fieldFormat As GeneralUtils.NumberFormat

        'Set saveFlag as false for fields which need to update datatable but no need to save database.
        'For example, changed display only vendor name from vendor lookup control. 
        'Only vendor no need save but vendor name should be update to datatable for postback
        'saveFlag default value = true
        Dim saveFlag As Boolean
    End Structure
End Class
