Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.OleDb
Imports System.Web
Imports Spire.Barcode

Public Class CommonUtils
    Inherits System.Web.UI.Page

    Private dtUpdateFields As DataTable
    Private mvstrFieldName As String
    Private mvstrFieldValue As String
    Private mvctlFieldCtrl As Control
    Private mvbooIsDBF As Boolean

    Private mvintNo As Integer
    Private strSQLStatement1 As New StringBuilder
    Private strSQLStatement2 As New StringBuilder

    'Shared dbfunc1 As DBfunc = New DBfunc
    Private strScreenId As String
    Private strCheckExistLogIn As String
    Private mvstrUserId As String
    Private mvstrCurrentTime As String
    Private sSql As String
    Private mvstrId As String
    Private strReturnString As String

    Private strCtrlValue As String
    Private strFieldName As String
    Private ctlFieldCtrl As Control
    Private booIsDBF As Boolean

    Private gmNo As Integer = 0
    Private mvRowNo As Integer

    Public Function chgToYYYYMMDD(ByVal lchkdate As String) As String
        Dim gu As New GeneralUtils
        Dim DDFORMAT2 As String = gu.getConfig("DDFORMAT2")

        If gu.isValidDate(lchkdate, "MM/dd/yyyy hh:ss:mm") OrElse gu.isValidDate(lchkdate, "dd/MM/yyyy hh:ss:mm") OrElse gu.isValidDate(lchkdate) Then
            'chgToYYYYMMDD = Format(DateTime.Parse(lchkdate), "yyyy/MM/dd")
            chgToYYYYMMDD = Format(DateTime.Parse(lchkdate), DDFORMAT2)
        Else
            chgToYYYYMMDD = lchkdate
        End If
    End Function

    Public Function chgToFullDF(ByVal lchkdate As String) As String
        Dim gu As New GeneralUtils

        If gu.isValidDate(lchkdate) Then
            chgToFullDF = Format(DateTime.Parse(lchkdate), "yyyy/MM/dd hh:mm:ss")
        Else
            chgToFullDF = lchkdate
        End If
    End Function

    Public Function chgToTime(ByVal lchkdate As String) As String
        Dim gu As New GeneralUtils

        If gu.isValidDate(lchkdate) Then
            chgToTime = Format(DateTime.Parse(lchkdate), "HH:mm")
        Else
            chgToTime = lchkdate
        End If
    End Function

    Public Function getHeader(ByVal mName As String, ByVal ation As String) As String
        Select Case ation
            Case "SEARCH"
                Return mName & " (Search)"
            Case "NEW"
                Return mName & " (New)"
            Case "EDIT"
                Return mName
            Case Else
                Return mName
        End Select
    End Function
    Public Function gfNewDataTable() As DataTable
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Create new datatable
        '=============================================
        Dim dtDataTable As New DataTable
        dtDataTable.Columns.Add("FieldName")
        dtDataTable.Columns.Add("Value")
        dtDataTable.Columns.Add("Ctrl")
        dtDataTable.Columns.Add("IsDBF")
        Return dtDataTable
    End Function
    Public Function gfCheckifDate(ByVal vstrValue As String) As String

        If IsDate(vstrValue) = True Then
            vstrValue = "to_char(" & vstrValue & ", 'dd/mm/yy')"
        End If

        Return vstrValue
    End Function
    Public Function gfBuiltSQLSearchString(ByVal vpnlPanel As Panel, ByVal vstrSqlSelect As String, ByVal vstrWhereClause As String) As String
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Build Sql string for search
        '=============================================


        dtUpdateFields = New DataTable
        dtUpdateFields = gfBuildDataTableforMaster(vpnlPanel, gfNewDataTable)


        Dim strSQLStatement As New StringBuilder

        Dim strSQLStatement2 As New StringBuilder

        Dim strSqlFinalStatement As New StringBuilder

        Dim i As Integer
        For i = 0 To dtUpdateFields.Rows.Count - 1
            Dim strFieldName As String = dtUpdateFields.Rows(i).Item("FieldName")
            Dim strValue As String = dtUpdateFields.Rows(i).Item("value")
            Dim booIsDBF As String = dtUpdateFields.Rows(i).Item("IsDBF")
            If booIsDBF Then
                'strValue = gfCheckifDate(strValue)
                If strValue <> "" Then
                    REM Special Handling
                    If UCase(strFieldName.Substring(0, 3)) = "PRE" Then
                        strFieldName = gfGetFieldNames(strFieldName)
                        strSQLStatement.Append(strFieldName & " LIKE '" & strValue & "%'")
                        strSQLStatement.Append(" and ")
                    ElseIf UCase(strFieldName.Substring(0, 3)) = "SUF" Then
                        strFieldName = gfGetFieldNames(strFieldName)
                        strSQLStatement.Append(strFieldName & " LIKE '%" & strValue & "'")
                        strSQLStatement.Append(" and ")
                    ElseIf UCase(strFieldName.Substring(0, 3)) = "LIK" Then
                        strFieldName = gfGetFieldNames(strFieldName)
                        strSQLStatement.Append(strFieldName & " LIKE '%" & strValue & "%'")
                        strSQLStatement.Append(" and ")
                    ElseIf UCase(strFieldName.Substring(0, 3)) = "FRM" Then

                        strFieldName = gfGetFieldNames(strFieldName)
                        strSQLStatement.Append(strFieldName & " >= " & gfCheckifDate("'" & strValue & "'"))
                        strSQLStatement.Append(" and ")

                    ElseIf UCase(strFieldName.Substring(0, 3)) = "STO" Then

                        strFieldName = gfGetFieldNames(strFieldName)

                        strSQLStatement.Append(strFieldName & " <= " & gfCheckifDate("'" & strValue & "'"))
                        strSQLStatement.Append(" and ")
                    ElseIf strFieldName.Contains("SA_MEID") Then
                        strSQLStatement.Append("ISNULL(SA_MEID1,'')+ISNULL(SA_MEID2,'')" & " = '" & strValue & "'")
                        strSQLStatement.Append(" and ")
                    ElseIf strFieldName.Contains("SA_REG_NAME") Then
                        strSQLStatement.Append("ISNULL(SA_REG_NAME1,'')+ISNULL(SA_REG_NAME2,'')+ISNULL(SA_REG_NAME3,'')" & " = '" & strValue & "'")
                        strSQLStatement.Append(" and ")
                    Else
                        strSQLStatement.Append(strFieldName & " = " & gfCheckifDate("'" & strValue & "'"))
                        strSQLStatement.Append(" and ")
                    End If
                End If
            End If
        Next i

        strSQLStatement2.Append(vstrSqlSelect)
        strSQLStatement2.Append(" where ")
        strSQLStatement2.Append(strSQLStatement.ToString)

        If vstrWhereClause <> "" Then
            strSQLStatement2.Append(vstrWhereClause)
            strSqlFinalStatement.Append(strSQLStatement2.ToString)
        Else
            Dim strSql As String = strSQLStatement2.ToString

            If strSQLStatement.ToString <> "" Then
                REM trim " and "
                Dim intSqlLength As Integer = strSql.Length
                Dim intSqlStartRemove As Integer = intSqlLength - 5
                Dim strSqlFinal As String = strSql.Remove(intSqlStartRemove, 5)
                strSqlFinalStatement.Append(strSqlFinal)
            Else
                REM trim " where "
                Dim intSqlLength As Integer = strSql.Length
                Dim intSqlStartRemove As Integer = intSqlLength - 7
                Dim strSqlFinal As String = strSql.Remove(intSqlStartRemove, 7)
                strSqlFinalStatement.Append(strSqlFinal)
            End If
        End If

        Dim sdfsa As String = strSqlFinalStatement.ToString
        Return strSqlFinalStatement.ToString

    End Function

    Public Function gfAddGridViewRowwithBoundColumn(ByVal dtDataTable As DataTable, ByVal gvGridView As GridView) As DataTable
        Dim drNewRow As DataRow = dtDataTable.NewRow()
        Dim intDTRowCount As Integer = dtDataTable.Rows.Count
        Dim i As Integer
        Dim j As Integer

        Dim sdf As String = gvGridView.Rows(i).Cells(1).Text.ToString
        'force dtKEng to remember the values just added
        For i = 0 To gvGridView.Rows.Count - 1
            For j = 1 To gvGridView.Columns.Count - 1

                dtDataTable.Rows(i).Item(j) = gvGridView.Rows(i).Cells(j).Text.ToString
            Next
        Next


        dtDataTable.Rows.Add(drNewRow)
        dtDataTable.AcceptChanges()
        Return dtDataTable
    End Function
    Public Sub gfAddGridViewMultipleRow(ByVal dtDataTable As DataTable, ByVal gvGridView As GridView, ByVal strControlName1 As String, ByVal strControlName2 As String, ByVal strControlName3 As String)
        Dim drNewRow As DataRow = dtDataTable.NewRow()
        Dim intDTRowCount As Integer = dtDataTable.Rows.Count
        Dim i As Integer
        Dim j As Integer



        Dim total As Integer = gvGridView.Rows.Count

        'force dtKEng to remember the values just added

        For j = 0 To gvGridView.Columns.Count - 2
            'For i = gvGridView.Rows.Count - 1 To 0 Step -1
            For i = 0 To gvGridView.Rows.Count - 1
                If strControlName1 <> "" Then
                    If j = 0 Then
                        If CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName1), TextBox).Text() = "" Then
                            dtDataTable.Rows(i).Item(j) = System.DBNull.Value
                        Else
                            dtDataTable.Rows(i).Item(j) = CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName1), TextBox).Text()
                        End If
                    End If

                End If
                If strControlName2 <> "" Then
                    If j = 1 Then
                        If CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName2), TextBox).Text() = "" Then
                            dtDataTable.Rows(i).Item(j) = System.DBNull.Value
                        Else
                            dtDataTable.Rows(i).Item(j) = CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName2), TextBox).Text()
                        End If
                    End If
                End If
                If strControlName3 <> "" Then
                    If j = 2 Then
                        If CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName3), TextBox).Text() = "" Then
                            dtDataTable.Rows(i).Item(j) = System.DBNull.Value
                        Else
                            dtDataTable.Rows(i).Item(j) = CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName3), TextBox).Text()
                        End If
                    End If
                End If


            Next i
        Next

        dtDataTable.Rows.Add(drNewRow)
        dtDataTable.AcceptChanges()

        gvGridView.DataSource = dtDataTable
        gvGridView.DataBind()
        mvRowNo = gvGridView.Rows.Count
    End Sub
    Public Sub gfDeleteGirdViewMultipleRow(ByVal dtDataTable As DataTable, ByVal gvGridView As GridView, ByVal strControlName1 As String, ByVal strControlName2 As String, ByVal strControlName3 As String, ByVal strCheckBoxName As String)
        Dim intNoRows As Integer = gvGridView.Rows.Count
        Dim j As Integer
        Dim i As Integer
        Dim sdf As String = CType(gvGridView.Rows(0).Cells(1).FindControl(strControlName1), TextBox).Text
        'For i = gvGridView.Rows.Count - 1 To 0 Step -1
        For i = 0 To gvGridView.Rows.Count - 1
            For j = 0 To gvGridView.Columns.Count - 2
                If j = 0 Then
                    If CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName1), TextBox).Text = "" Then
                        dtDataTable.Rows(i).Item(j) = System.DBNull.Value
                    Else
                        dtDataTable.Rows(i).Item(j) = CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName1), TextBox).Text
                    End If
                End If
                If j = 1 Then
                    If CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName2), TextBox).Text = "" Then
                        dtDataTable.Rows(i).Item(j) = System.DBNull.Value
                    Else
                        dtDataTable.Rows(i).Item(j) = CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName2), TextBox).Text
                    End If
                End If
                If j = 2 Then
                    If CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName3), TextBox).Text = "" Then
                        dtDataTable.Rows(i).Item(j) = System.DBNull.Value
                    Else
                        dtDataTable.Rows(i).Item(j) = CType(gvGridView.Rows(i).Cells(j + 1).FindControl(strControlName3), TextBox).Text
                    End If
                End If
            Next
            Dim lcb_del As CheckBox = CType(gvGridView.Rows(i).Cells(0).FindControl(strCheckBoxName), CheckBox)
            If lcb_del.Checked = True Then


                dtDataTable.Rows(i).Delete()


                intNoRows = intNoRows - 1
            End If
        Next
        dtDataTable.AcceptChanges()
        gvGridView.DataSource = dtDataTable
        gvGridView.DataBind()
    End Sub

    Public Sub gfAddGridViewRow(ByVal dtDataTable As DataTable, ByVal gvGridView As GridView, ByVal strColName As String, ByVal strControlName As String)
        Dim drNewRow As DataRow = dtDataTable.NewRow()
        Dim intDTRowCount As Integer = dtDataTable.Rows.Count
        Dim i As Integer

        Dim sdf As String = CType(gvGridView.Rows(i).Cells(1).FindControl(strControlName), TextBox).Text
        'force dtKEng to remember the values just added
        For i = 0 To gvGridView.Rows.Count - 1
            dtDataTable.Rows(i).Item(strColName) = CType(gvGridView.Rows(i).Cells(1).FindControl(strControlName), TextBox).Text
        Next

        dtDataTable.Rows.Add(drNewRow)
        dtDataTable.AcceptChanges()
        gvGridView.DataSource = dtDataTable
        gvGridView.DataBind()
    End Sub

    Public Sub gfDeleteGirdViewRow(ByVal dtDataTable As DataTable, ByVal gvGridView As GridView, ByVal strColName As String, ByVal strControlName As String, ByVal strCheckBoxName As String)
        Dim intNoRows As Integer = gvGridView.Rows.Count

        Dim i As Integer
        For i = gvGridView.Rows.Count - 1 To 0 Step -1


            dtDataTable.Rows(i).Item(strColName) = CType(gvGridView.Rows(i).Cells(1).FindControl(strControlName), TextBox).Text


            Dim lcb_del As CheckBox = CType(gvGridView.Rows(i).Cells(0).FindControl(strCheckBoxName), CheckBox)
            If lcb_del.Checked = True Then


                dtDataTable.Rows(i).Delete()


                intNoRows = intNoRows - 1
            End If
        Next i
        dtDataTable.AcceptChanges()
        gvGridView.DataSource = dtDataTable
        gvGridView.DataBind()
    End Sub

    Public Function gfBuildDataTableforGridView(ByRef ldt As DataTable, ByVal GV As GridView, Optional ByVal setWithFieldName As Boolean = False, _
                                                Optional ByVal chkboxValue() As String = Nothing) As Boolean
        gfBuildDataTableforGridView = gfSetDataTableforGV(ldt, GV, setWithFieldName, chkboxValue)
    End Function

    Public Function gfBuildDataTableforGridView(ByRef ldt As DataTable, ByVal GV As GridView, ByVal ParamArray chkboxValueParam() As String) As Boolean
        gfBuildDataTableforGridView = gfSetDataTableforGV(ldt, GV, False, chkboxValueParam)
    End Function

    Public Function gfBuildDataRowsforGridView(ByRef org_dt As DataTable, ByRef lRow As DataRow(), ByVal GV As GridView, Optional ByVal setWithFieldName As Boolean = False, _
                                                Optional ByVal chkboxValue() As String = Nothing) As Boolean
        gfBuildDataRowsforGridView = gfSetDataRowsforGV(org_dt, lRow, GV, setWithFieldName)
    End Function

    Private Function gfSetDataTableforGV(ByRef ldt As DataTable, ByVal gvGridView As GridView, ByVal setWithName As Boolean, Optional ByVal chkboxValue() As String = Nothing) As Boolean
        Dim n As Integer
        Dim rowCollection As GridViewRowCollection = gvGridView.Rows
        Dim pageIndex As Integer = gvGridView.PageIndex
        Dim iRowIndex As Integer
        Dim cellIndex As Integer = 0
        'Try
        If Not pageIndex = 0 Then
            For i As Integer = 1 To pageIndex
                iRowIndex = iRowIndex + 10
            Next
        End If

        For Each gridRow As GridViewRow In rowCollection
            n = 0
            If gridRow.Visible = False Then
                iRowIndex += 1
                Continue For
            End If
            Dim rowCell As TableCellCollection = gridRow.Cells
            For Each itemCell As TableCell In rowCell
                Dim strBoundField As String = itemCell.Text

                If n >= 0 Then
                    If strBoundField <> "" Then
                        'Set FieldValue
                        If itemCell.Enabled = True Then
                            gpstrFieldValue = strBoundField
                        End If

                        If setWithName Then
                            If Not itemCell.ID Is Nothing Then
                                set_dataset_value_with_name(ldt, iRowIndex, itemCell.ID)
                            End If
                        Else
                            set_dataset_value_with_index(ldt, iRowIndex, n)
                        End If

                        n = n + 1
                    Else
                        For Each ctl As Control In itemCell.Controls
                            If TypeOf ctl Is TextBox Then
                                Dim tbCtrlTextBox As TextBox = CType(ctl, TextBox)
                                If tbCtrlTextBox.Enabled = True Then

                                    gpstrFieldValue = tbCtrlTextBox.Text.Trim

                                    If setWithName Then
                                        set_dataset_value_with_name(ldt, iRowIndex, tbCtrlTextBox.ID)
                                    Else
                                        set_dataset_value_with_index(ldt, iRowIndex, n)
                                    End If

                                    n = n + 1
                                End If
                            ElseIf UCase(TypeName(ctl)) = "CHECKBOX" Then
                                Dim ckCtrlCheckBox As CheckBox = CType(ctl, CheckBox)
                                Dim checkValue As String = "1"
                                Dim unCheckValue As String = "0"

                                If ckCtrlCheckBox.Enabled = True OrElse ckCtrlCheckBox.Visible = True Then

                                    If Not IsNothing(chkboxValue) Then
                                        If chkboxValue.Length > 1 Then
                                            checkValue = chkboxValue(0).ToString()
                                            unCheckValue = chkboxValue(1).ToString()
                                        End If
                                    End If

                                    If ckCtrlCheckBox.Checked Then gpstrFieldValue = checkValue Else gpstrFieldValue = unCheckValue

                                    If setWithName Then
                                        set_dataset_value_with_name(ldt, iRowIndex, ckCtrlCheckBox.ID)
                                    Else
                                        set_dataset_value_with_index(ldt, iRowIndex, n)
                                    End If

                                    n = n + 1

                                End If

                            ElseIf TypeOf ctl Is Label Then
                                n = n + 1
                            ElseIf TypeOf ctl Is DropDownList Then
                                Dim dplCtrlDropDown As DropDownList = CType(ctl, DropDownList)
                                If dplCtrlDropDown.Enabled = True Then
                                    gpstrFieldValue = gfRemoveInvalidChar(dplCtrlDropDown.SelectedValue.ToString)

                                    If setWithName Then
                                        set_dataset_value_with_name(ldt, iRowIndex, dplCtrlDropDown.ID)
                                    Else
                                        set_dataset_value_with_index(ldt, iRowIndex, n)
                                    End If

                                    n = n + 1

                                End If
                            ElseIf TypeOf ctl Is HiddenField Then
                                Dim tbCtrlHiddenField As HiddenField = CType(ctl, HiddenField)

                                gpstrFieldValue = gfRemoveInvalidChar(tbCtrlHiddenField.Value)

                                If setWithName Then
                                    set_dataset_value_with_name(ldt, iRowIndex, tbCtrlHiddenField.ID)
                                Else
                                    set_dataset_value_with_index(ldt, iRowIndex, n)
                                End If

                                n = n + 1
                            ElseIf TypeOf ctl Is AjaxControlToolkit.ComboBox Then
                                Dim dplCtrlDropDown As AjaxControlToolkit.ComboBox = CType(ctl, AjaxControlToolkit.ComboBox)
                                If dplCtrlDropDown.Enabled = True Then
                                    gpstrFieldValue = gfRemoveInvalidChar(dplCtrlDropDown.SelectedValue.ToString.Trim)

                                    If setWithName Then
                                        set_dataset_value_with_name(ldt, iRowIndex, dplCtrlDropDown.ID)
                                    Else
                                        set_dataset_value_with_index(ldt, iRowIndex, n)
                                    End If

                                    n = n + 1

                                End If
                            End If
                        Next
                    End If
                Else
                    n += 1
                End If

            Next
            iRowIndex += 1
        Next

        gfSetDataTableforGV = True

        'Catch ex As Exception
        '    gfBuildDataTableforGridView = False
        'End Try
    End Function

    Private Function gfSetDataRowsforGV(ByRef org_dt As DataTable, ByRef lrow As DataRow(), ByVal gvGridView As GridView, ByVal setWithName As Boolean, Optional ByVal chkboxValue() As String = Nothing) As Boolean
        Dim n As Integer
        Dim rowCollection As GridViewRowCollection = gvGridView.Rows
        Dim pageIndex As Integer = gvGridView.PageIndex
        Dim iRowIndex As Integer
        Dim cellIndex As Integer = 0

        Dim ldt As New DataTable

        ldt = lrow.CopyToDataTable()

        For Each r As DataRow In lrow
            org_dt.Rows.Remove(r)
        Next

        org_dt.AcceptChanges()

        'Try
        If Not pageIndex = 0 Then
            For i As Integer = 1 To pageIndex
                iRowIndex = iRowIndex + 10
            Next
        End If

        Dim nextContinue As Boolean = False

        For Each gridRow As GridViewRow In rowCollection
            n = 0
            If gridRow.Visible = False Then
                Continue For
            End If
            Dim rowCell As TableCellCollection = gridRow.Cells
            For Each itemCell As TableCell In rowCell
                Dim strBoundField As String = itemCell.Text

                If n >= 0 Then
                    If strBoundField <> "" Then
                        'Set FieldValue
                        If itemCell.Enabled = True Then
                            gpstrFieldValue = strBoundField
                        End If

                        If setWithName Then
                            If Not itemCell.ID Is Nothing Then
                                set_dataset_value_with_name(ldt, iRowIndex, itemCell.ID)
                            End If
                        Else
                            set_dataset_value_with_index(ldt, iRowIndex, n)
                        End If

                        n = n + 1
                    Else
                        For Each ctl As Control In itemCell.Controls
                            If TypeOf ctl Is TextBox Then
                                Dim tbCtrlTextBox As TextBox = CType(ctl, TextBox)

                                If tbCtrlTextBox.Enabled = True Then

                                    gpstrFieldValue = gfRemoveInvalidChar(tbCtrlTextBox.Text)

                                    If setWithName Then
                                        set_dataset_value_with_name(ldt, iRowIndex, tbCtrlTextBox.ID)
                                    Else
                                        set_dataset_value_with_index(ldt, iRowIndex, n)
                                    End If

                                    n = n + 1
                                Else
                                    nextContinue = True
                                    Exit For
                                End If
                            ElseIf TypeOf ctl Is CheckBox Then
                                Dim ckCtrlCheckBox As CheckBox = CType(ctl, CheckBox)
                                Dim checkValue As String = "1"
                                Dim unCheckValue As String = "0"

                                If ckCtrlCheckBox.Enabled = True OrElse ckCtrlCheckBox.Visible = True Then

                                    If Not IsNothing(chkboxValue) Then
                                        If chkboxValue.Length > 1 Then
                                            checkValue = chkboxValue(0).ToString()
                                            unCheckValue = chkboxValue(1).ToString()
                                        End If
                                    End If

                                    If ckCtrlCheckBox.Checked Then gpstrFieldValue = checkValue Else gpstrFieldValue = unCheckValue

                                    If setWithName Then
                                        set_dataset_value_with_name(ldt, iRowIndex, ckCtrlCheckBox.ID)
                                    Else
                                        set_dataset_value_with_index(ldt, iRowIndex, n)
                                    End If

                                    n = n + 1
                                Else
                                    nextContinue = True
                                    Exit For
                                End If

                            ElseIf TypeOf ctl Is Label Then
                                n = n + 1
                            ElseIf TypeOf ctl Is DropDownList Then
                                Dim dplCtrlDropDown As DropDownList = CType(ctl, DropDownList)
                                If dplCtrlDropDown.Enabled = True Then
                                    gpstrFieldValue = gfRemoveInvalidChar(dplCtrlDropDown.SelectedValue.ToString)

                                    If setWithName Then
                                        set_dataset_value_with_name(ldt, iRowIndex, dplCtrlDropDown.ID)
                                    Else
                                        set_dataset_value_with_index(ldt, iRowIndex, n)
                                    End If

                                    n = n + 1
                                Else
                                    nextContinue = True
                                    Exit For
                                End If
                            ElseIf TypeOf ctl Is HiddenField Then
                                Dim tbCtrlHiddenField As HiddenField = CType(ctl, HiddenField)

                                gpstrFieldValue = gfRemoveInvalidChar(tbCtrlHiddenField.Value)

                                If setWithName Then
                                    set_dataset_value_with_name(ldt, iRowIndex, tbCtrlHiddenField.ID)
                                Else
                                    set_dataset_value_with_index(ldt, iRowIndex, n)
                                End If

                                n = n + 1

                            End If
                        Next
                    End If
                Else
                    n += 1
                End If

            Next

            If Not nextContinue Then
                iRowIndex += 1
            End If

            nextContinue = False
        Next

        lrow = ldt.Select("")

        For Each r As DataRow In lrow
            org_dt.ImportRow(r)
        Next

        org_dt.AcceptChanges()

        gfSetDataRowsforGV = True

        'Catch ex As Exception
        '    gfBuildDataTableforGridView = False
        'End Try
    End Function

    Public Sub set_dataset_value_with_index(ByRef ldt As DataTable, ByVal rows As Integer, ByVal cols As Integer)
        Dim tempVal As String

        If IsDBNull(ldt.Rows(rows).Item(cols).ToString) Then
            tempVal = ""
        Else
            tempVal = ldt.Rows(rows).Item(cols).ToString
        End If

        If mvstrFieldValue <> tempVal Then
            If mvstrFieldValue <> "" Then
                ldt.Rows(rows).Item(cols) = mvstrFieldValue
            Else
                ldt.Rows(rows).Item(cols) = DBNull.Value
            End If

            ldt.AcceptChanges()
        End If
    End Sub

    Public Sub set_dataset_value_with_name(ByRef ldt As DataTable, ByVal rows As Integer, ByVal columnName As String)
        Dim tempVal As String

        If IsDBNull(ldt.Rows(rows).Item(columnName).ToString) Then
            tempVal = ""
        Else
            tempVal = ldt.Rows(rows).Item(columnName).ToString
        End If

        Select Case ldt.Columns(columnName).DataType.FullName
            Case "System.DateTime"
                If Not IsDate(mvstrFieldValue) Then
                    mvstrFieldValue = ""
                End If
        End Select

        If mvstrFieldValue <> tempVal Then
            If mvstrFieldValue <> "" Then
                ldt.Rows(rows).Item(columnName) = mvstrFieldValue
            Else
                ldt.Rows(rows).Item(columnName) = DBNull.Value
            End If

            ldt.AcceptChanges()
        End If
    End Sub

    Public Sub changeGVLabel(ByRef oGridViewRow As GridViewRow, ByRef e As System.Web.UI.WebControls.GridViewRowEventArgs, ByVal engName As String, ByVal chiName As String, Optional ByVal align As HorizontalAlign = HorizontalAlign.Left)
        e.Row.Cells.RemoveAt(0)
        Dim oTableCell As New TableCell()
        'oTableCell.Font.Size = 10
        'oTableCell.Font.Bold = True
        oTableCell.CssClass = "DtlLabel"
        oTableCell.HorizontalAlign = align
        If Session("gLang") = "C" Then oTableCell.Text = chiName Else oTableCell.Text = engName
        oGridViewRow.Cells.Add(oTableCell)
    End Sub

    Public Sub newChangeGVLabel(ByRef gv As GridView, ByVal engName As String, ByVal chiName As String, Optional ByVal align As HorizontalAlign = HorizontalAlign.Left)
        CType(gv.Controls(0).Controls(0), GridViewRow).Cells.RemoveAt(0)
        Dim oTableCell As New TableCell()
        oTableCell.CssClass = "DtlLabel"
        oTableCell.HorizontalAlign = align
        If Session("gLang") = "C" Then oTableCell.Text = chiName Else oTableCell.Text = engName
        CType(gv.Controls(0).Controls(0), GridViewRow).Cells.Add(oTableCell)
    End Sub

    Public Sub changeCellByIndex(ByRef e As System.Web.UI.WebControls.GridViewRowEventArgs, ByVal index As Integer, ByVal engName As String, ByVal chiName As String, Optional ByVal align As HorizontalAlign = HorizontalAlign.Left)
        Dim lbl As String

        If Session("gLang") = "C" Then lbl = chiName Else lbl = engName

        e.Row.Cells(index).Font.Size = 10
        e.Row.Cells(index).Font.Bold = True
        e.Row.Cells(index).CssClass = "TITLE"
        e.Row.Cells(index).HorizontalAlign = align

        If lbl <> "" Then
            e.Row.Cells(index).Text = lbl
        End If

    End Sub


    Public Sub gfBuildDataTableforMaster2(ByVal root As Control)

        'Dim DataRow As DataRow
        'DataRow = dtUpdateTables.NewRow()

        Dim j As Integer = 0
        For Each subctrl As Control In root.Controls
            If Not TypeOf root Is GridView Then

                Dim strCtrlID As String = subctrl.ID


                If TypeOf subctrl Is TextBox Then
                    Dim tbCtrlTextBox As TextBox = CType(subctrl, TextBox)
                    If tbCtrlTextBox.Enabled = True Then
                        strFieldName = gfGetFieldNames(strCtrlID)
                        strCtrlValue = tbCtrlTextBox.Text
                    End If

                ElseIf TypeOf subctrl Is RadioButton Then
                    Dim rbCtrlRadioButton As RadioButton = CType(subctrl, RadioButton)
                    If rbCtrlRadioButton.Checked = True Then
                        If rbCtrlRadioButton.Enabled = True Then
                            strFieldName = gfGetFieldNames(rbCtrlRadioButton.GroupName)
                            strCtrlValue = gfGetFieldNames(strCtrlID)
                        End If
                    End If

                ElseIf TypeOf subctrl Is CheckBox Then
                    Dim ckCtrlCheckBox As CheckBox = CType(subctrl, CheckBox)
                    If ckCtrlCheckBox.Enabled = True Then
                        strFieldName = gfGetFieldNames(strCtrlID)
                        strCtrlValue = ckCtrlCheckBox.Text
                    End If

                ElseIf TypeOf subctrl Is DropDownList Then
                    Dim dplCtrlDropDown As DropDownList = CType(subctrl, DropDownList)
                    If dplCtrlDropDown.Enabled = True Then
                        strFieldName = gfGetFieldNames(strCtrlID)
                        strCtrlValue = dplCtrlDropDown.SelectedValue.ToString
                    End If

                End If



                strFieldName = ""
                strCtrlValue = ""
            End If
        Next



    End Sub
    Public Sub gfDisableGridViewControl(ByVal gvGridView As GridView)

        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Disable GridView Control
        '=============================================

        Dim rowCollection As GridViewRowCollection = gvGridView.Rows

        For Each gridRow As GridViewRow In rowCollection
            Dim rowCell As TableCellCollection = gridRow.Cells
            For Each itemCell As TableCell In rowCell

                For Each ctl As Control In itemCell.Controls

                    If TypeOf ctl Is TextBox Then
                        Dim tbCtrlTextBox As TextBox = CType(ctl, TextBox)
                        tbCtrlTextBox.Enabled = False
                    ElseIf TypeOf ctl Is CheckBox Then
                        Dim ckCtrlCheckBox As CheckBox = CType(ctl, CheckBox)
                        ckCtrlCheckBox.Enabled = False
                    ElseIf TypeOf ctl Is DropDownList Then
                        Dim dplCtrlDropDown As DropDownList = CType(ctl, DropDownList)
                        dplCtrlDropDown.Enabled = False

                    End If



                Next


            Next
        Next

    End Sub
    Public Sub gfDisableControls(ByVal root As Control)

        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Disable Master Control
        '=============================================
        'Dim DataRow As DataRow
        'DataRow = dtUpdateTables.NewRow()

        Dim j As Integer = 0

        For Each subctrl As Control In root.Controls
            gfDisableControls(subctrl)
            If Not TypeOf root Is GridView Then

                Dim strCtrlID As String = subctrl.ID

                If TypeOf subctrl Is TextBox Then
                    Dim tbCtrlTextBox As TextBox = CType(subctrl, TextBox)
                    tbCtrlTextBox.Enabled = False

                ElseIf TypeOf subctrl Is RadioButtonList Then
                    Dim rbCtrlRadioButtonList As RadioButtonList = CType(subctrl, RadioButtonList)
                    rbCtrlRadioButtonList.Enabled = False


                ElseIf TypeOf subctrl Is CheckBox Then
                    Dim ckCtrlCheckBox As CheckBox = CType(subctrl, CheckBox)
                    ckCtrlCheckBox.Enabled = False

                ElseIf TypeOf subctrl Is DropDownList Then
                    Dim dplCtrlDropDown As DropDownList = CType(subctrl, DropDownList)
                    dplCtrlDropDown.Enabled = False
                ElseIf TypeOf subctrl Is Button Then

                    Dim btnCtrlButton As Button = CType(subctrl, Button)
                    If btnCtrlButton.ID = "btnEditDelete" Or btnCtrlButton.ID = "btnEditSave" Then
                        btnCtrlButton.Enabled = False
                    ElseIf btnCtrlButton.ID = "btnEditCancel" Then
                        btnCtrlButton.Enabled = True
                    Else
                        btnCtrlButton.Enabled = False
                    End If


                End If


            End If


        Next



    End Sub
    Public Sub gfEnableControls(ByVal root As Control)
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Enable Master Control
        '=============================================

        'Dim DataRow As DataRow
        'DataRow = dtUpdateTables.NewRow()

        Dim j As Integer = 0

        For Each subctrl As Control In root.Controls
            gfEnableControls(subctrl)
            If Not TypeOf root Is GridView Then

                Dim strCtrlID As String = subctrl.ID

                If TypeOf subctrl Is TextBox Then
                    Dim tbCtrlTextBox As TextBox = CType(subctrl, TextBox)
                    tbCtrlTextBox.Enabled = True

                ElseIf TypeOf subctrl Is RadioButtonList Then
                    Dim rbCtrlRadioButtonList As RadioButtonList = CType(subctrl, RadioButtonList)
                    rbCtrlRadioButtonList.Enabled = True


                ElseIf TypeOf subctrl Is CheckBox Then
                    Dim ckCtrlCheckBox As CheckBox = CType(subctrl, CheckBox)
                    ckCtrlCheckBox.Enabled = True

                ElseIf TypeOf subctrl Is DropDownList Then
                    Dim dplCtrlDropDown As DropDownList = CType(subctrl, DropDownList)
                    dplCtrlDropDown.Enabled = True
                ElseIf TypeOf subctrl Is Button Then

                    Dim btnCtrlButton As Button = CType(subctrl, Button)

                    btnCtrlButton.Enabled = True

                End If


            End If





        Next



    End Sub
    Public Function gfRemoveInvalidChar(ByVal strValue As String)
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Remove Invalid Character for string
        '=============================================
        If strValue.Contains("'") Then
            Dim intStart As Integer = strValue.IndexOf("'")
            strValue = strValue.Replace("'", "''")

        End If
        Return strValue
    End Function
    Public Function gfBuildDataTableforMaster(ByVal root As Control, ByVal vdtDataTable As DataTable) As DataTable
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Build DataTable for Master
        '=============================================
        For Each subctrl As Control In root.Controls
            If Not TypeOf root Is GridView Then
                gfBuildDataTableforMaster(subctrl, vdtDataTable)
                Dim strCtrlID As String = subctrl.ID
                If TypeOf subctrl Is TextBox Then
                    Dim tbCtrlTextBox As TextBox = CType(subctrl, TextBox)
                    If tbCtrlTextBox.Enabled = True Then
                        strFieldName = gfGetFieldNames(strCtrlID)
                        ctlFieldCtrl = subctrl
                        strCtrlValue = gfRemoveInvalidChar(tbCtrlTextBox.Text)
                        booIsDBF = IsDBField(strCtrlID)
                    End If
                ElseIf TypeOf subctrl Is RadioButtonList Then
                    Dim rbCtrlRadioButton As RadioButtonList = CType(subctrl, RadioButtonList)
                    If rbCtrlRadioButton.Enabled = True Then
                        strFieldName = gfGetFieldNames(strCtrlID)
                        ctlFieldCtrl = subctrl
                        strCtrlValue = rbCtrlRadioButton.SelectedValue.ToString
                        booIsDBF = IsDBField(strCtrlID)
                    End If
                ElseIf TypeOf subctrl Is CheckBox Then
                    Dim ckCtrlCheckBox As CheckBox = CType(subctrl, CheckBox)
                    If ckCtrlCheckBox.Enabled = True Then
                        strFieldName = gfGetFieldNames(strCtrlID)
                        ctlFieldCtrl = subctrl
                        strCtrlValue = ckCtrlCheckBox.Text
                        booIsDBF = IsDBField(strCtrlID)
                    End If

                ElseIf TypeOf subctrl Is DropDownList Then
                    Dim dplCtrlDropDown As DropDownList = CType(subctrl, DropDownList)
                    If dplCtrlDropDown.Enabled = True Then
                        strFieldName = gfGetFieldNames(strCtrlID)
                        ctlFieldCtrl = subctrl
                        strCtrlValue = dplCtrlDropDown.SelectedValue.ToString
                        booIsDBF = IsDBField(strCtrlID)
                    End If
                End If

                If strFieldName <> "" Then
                    gpstrFieldName = strFieldName
                    gpstrFieldValue = strCtrlValue
                    gpctlFieldCtrl = ctlFieldCtrl
                    gpbooIsDBF = booIsDBF
                    gpintNo = gmNo
                    gfAddDataToDataTable(vdtDataTable, gpintNo)
                    gmNo = gmNo + 1
                End If

                strFieldName = ""
                strCtrlValue = ""
                ctlFieldCtrl = Nothing
            End If
        Next

        Return vdtDataTable

    End Function
    Public Function gvGetDeleteGVDataSql(ByVal vstrTable As String, ByVal vstrWhereClause As String) As String
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Delete Sql for GridView 
        '=============================================
        Dim strSql As String
        If vstrWhereClause <> "" Then
            strSql = "Delete from " & vstrTable & " where " & vstrWhereClause
        Else
            strSql = "Delete from " & vstrTable
        End If
        Return strSql
    End Function
    Public Sub gvSaveGridView(ByVal vstrTable As String, ByVal vstrIdField As String, ByVal vdtUpdateDataTable As DataTable, ByVal vstrId As String)
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Save GridView
        '=============================================


        Dim k As Integer
        For k = 0 To vdtUpdateDataTable.Rows.Count - 1
            strSQLStatement2.Append("Insert into ")
            strSQLStatement2.Append(vstrTable + " ")

            Dim i As Integer
            Dim j As Integer
            Dim l As Integer



            strSQLStatement2.Append("(")
            For j = 0 To vdtUpdateDataTable.Columns.Count - 1
                Dim strFieldName As String = vdtUpdateDataTable.Columns(j).ColumnName.ToString

                If j = vdtUpdateDataTable.Columns.Count - 1 Then
                    strSQLStatement2.Append("last_update_by" & ", ")
                    strSQLStatement2.Append("last_update_date" & ", ")
                    strSQLStatement2.Append(vstrIdField & ", ")
                    strSQLStatement2.Append(strFieldName & ") ")
                Else
                    strSQLStatement2.Append(strFieldName & ", ")
                End If
            Next j


            strSQLStatement2.Append("values ")
            strSQLStatement2.Append("(")


            For l = 0 To vdtUpdateDataTable.Columns.Count - 1
                Dim strValue As String = vdtUpdateDataTable.Rows(k).Item(l).ToString
                If l = vdtUpdateDataTable.Columns.Count - 1 Then

                    strSQLStatement2.Append("'" & mvstrUserId & "'" & ", ")
                    strSQLStatement2.Append(mvstrCurrentTime & ", ")
                    strSQLStatement2.Append(vstrId & ", ")
                    strSQLStatement2.Append("'" & strValue & "'" & ") ")

                Else
                    strSQLStatement2.Append("'" & strValue & "'" & ", ")
                End If
            Next l
            sSql = strSQLStatement2.ToString()

            'dbfunc1.CreateTransactionScope(vstrSQL1, sSql)

            'dbfunc1.SaveTranData2(sSql, vtrnTransaction)
            strSQLStatement2.Replace(sSql, "")

            'k = k + 1

        Next k


    End Sub
    Public Function gvGetMasterSql(ByVal vstrTable As String, ByVal vdtUpdateDataTable As DataTable, ByVal vstrWhereClause As String, ByVal vstrUpdateFlag As String) As String
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Create sql for Master
        '=============================================
        'Dim strIdJustInsert As String
        Dim strSql As String
        If UCase(vstrUpdateFlag) = "UPDATE" Then


            strSQLStatement1.Append("Update ")
            strSQLStatement1.Append(vstrTable + " ")
            strSQLStatement1.Append("Set ")

            Dim i As Integer

            For i = 0 To vdtUpdateDataTable.Rows.Count - 1


                Dim strFieldName As String = vdtUpdateDataTable.Rows(i).Item("FieldName")
                Dim strValue As String = vdtUpdateDataTable.Rows(i).Item("value")
                strValue = gfCheckifDate(strValue)

                If i = vdtUpdateDataTable.Rows.Count - 1 Then
                    strSQLStatement1.Append("last_update_by=" & "'" & mvstrUserId & "'" & ", ")
                    strSQLStatement1.Append("last_update_date=" & mvstrCurrentTime & ", ")
                    If vstrWhereClause <> "" Then
                        strSQLStatement1.Append(strFieldName & "=" & "'" & strValue & "'" & " ")
                        strSQLStatement1.Append(vstrWhereClause)
                    Else
                        strSQLStatement1.Append(strFieldName & "=" & "'" & strValue & "'")
                    End If
                Else
                    strSQLStatement1.Append(strFieldName & "=" & "'" & strValue & "'" & ", ")
                End If
            Next i

            strSql = strSQLStatement1.ToString()

            'dbfunc.amendData(strSql)
            'strIdJustInsert = ""

        ElseIf UCase(vstrUpdateFlag) = "DELETE" Then
            strSql = "Delete from " & vstrTable & " where " & vstrWhereClause
        Else

            strSQLStatement1.Append("Insert into ")
            strSQLStatement1.Append(vstrTable + " ")

            Dim i As Integer
            Dim j As Integer
            strSQLStatement1.Append("(")
            For i = 0 To vdtUpdateDataTable.Rows.Count - 1
                Dim strFieldName As String = vdtUpdateDataTable.Rows(i).Item("FieldName")

                If i = vdtUpdateDataTable.Rows.Count - 1 Then
                    strSQLStatement1.Append("last_update_by" & ", ")
                    strSQLStatement1.Append("last_update_date" & ", ")
                    strSQLStatement1.Append(strFieldName & ") ")
                Else
                    strSQLStatement1.Append(strFieldName & ", ")
                End If
            Next i

            strSQLStatement1.Append("values ")
            strSQLStatement1.Append("(")
            For j = 0 To vdtUpdateDataTable.Rows.Count - 1

                Dim strValue As String = vdtUpdateDataTable.Rows(j).Item("value")
                If j = vdtUpdateDataTable.Rows.Count - 1 Then

                    strSQLStatement1.Append("'" & mvstrUserId & "'" & ", ")
                    strSQLStatement1.Append(mvstrCurrentTime & ", ")
                    strSQLStatement1.Append("'" & strValue & "'" & ") ")

                Else
                    strSQLStatement1.Append("'" & strValue & "'" & ", ")
                End If
            Next j

            'strSQLStatement1.Append("select @@Identity")

            strSql = strSQLStatement1.ToString()

            'strIdJustInsert = dbfunc.gfInsertDatawithReturnValue(strSQL)


        End If
        'Return strIdJustInsert
        Return strSql
    End Function
    Public Function gfGetFieldNames(ByVal vstrObjName As String) As String
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Get FieldNames from Control
        '=============================================
        Dim strFieldName As String
        Dim intcount As Integer
        intcount = vstrObjName.Length - 3
        strFieldName = vstrObjName.Substring(3, intcount)
        Return strFieldName
    End Function
    Public Function IsDBField(ByVal vstrObjName As String) As String
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Get FieldNames from Control
        '=============================================
        If UCase(Left(vstrObjName.ToString, 3)) = "CUS" Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function gfAddDataToDataTable(ByVal vsdtFields As DataTable, ByVal vintNo As Integer) As DataTable
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Add columns to data table
        '=============================================
        Dim drNewRow As DataRow = vsdtFields.NewRow()
        vsdtFields.Rows.Add(drNewRow)
        vsdtFields.Rows(vintNo).Item("FieldName") = mvstrFieldName
        vsdtFields.Rows(vintNo).Item("Ctrl") = mvctlFieldCtrl
        vsdtFields.Rows(vintNo).Item("Value") = mvstrFieldValue
        vsdtFields.Rows(vintNo).Item("IsDBF") = mvbooIsDBF
        vsdtFields.AcceptChanges()
        Dim sdf As Integer = vsdtFields.Rows.Count

        Return vsdtFields
    End Function
    Public Function gfBuildDataTable2(ByVal vsdtFields As DataTable, ByVal vintRow As Integer, ByVal vintColumn As Integer) As DataTable


        If vintColumn = 0 Then
            Dim drNewRow As DataRow = vsdtFields.NewRow()
            vsdtFields.Rows.Add(drNewRow)
        End If
        vsdtFields.Rows(vintRow).Item(vintColumn) = mvstrFieldValue
        vsdtFields.AcceptChanges()
        Dim sdf As Integer = vsdtFields.Rows.Count

        Return vsdtFields

    End Function

    Public Function appendSql(ByVal sql As String, ByVal appSql As String) As String
        Dim returnSql As String
        returnSql = ""
        If appSql <> "" Then
            returnSql = sql
        End If
        Return returnSql
    End Function

    Public Function gfHasDuplicateEntry(ByVal vgvGridView As GridView, ByVal vstrTextBox As String) As Boolean
        Dim i As Integer
        Dim k As Integer
        Dim blnDuplicateEntry As Boolean = False

        For i = 0 To vgvGridView.Rows.Count - 1
            Dim strKChi As String = CType(vgvGridView.Rows(i).Cells(1).FindControl(vstrTextBox), TextBox).Text

            For k = 0 To vgvGridView.Rows.Count - 1
                If i <> k Then
                    Dim strKChi1 As String = CType(vgvGridView.Rows(k).Cells(1).FindControl(vstrTextBox), TextBox).Text
                    If strKChi = strKChi1 Then
                        blnDuplicateEntry = True
                    End If
                End If
            Next
        Next
        Return blnDuplicateEntry
    End Function

    Public Function gfISNumeric(ByVal vgvGridView As GridView, ByVal vstrTypeTextBox As String, ByVal vstrTextBox As String) As Boolean
        Dim i As Integer
        Dim blnISNumeric As Boolean = True

        For i = 0 To vgvGridView.Rows.Count - 1
            Dim val_field As String = CType(vgvGridView.Rows(i).Cells(1).FindControl(vstrTypeTextBox), DropDownList).SelectedValue.ToString.Trim
            Dim val_field2 As String = CType(vgvGridView.Rows(i).Cells(1).FindControl(vstrTextBox), TextBox).Text

            If IsNumeric(val_field2) = False Then
                blnISNumeric = False
            Else
                If val_field = "YR" Then
                    If val_field2 <= 0 Or val_field2 > 7 Then
                        blnISNumeric = False
                    End If
                Else
                    If val_field2 < 0 Or val_field2 > 12 Then
                        blnISNumeric = False
                    End If
                End If
            End If
        Next
        Return blnISNumeric
    End Function

    Public Sub gfResetControls(ByVal root As Control)

        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  6 March 2007
        ' Description:	Reset Master Control
        '=============================================


        Dim j As Integer = 0

        For Each subctrl As Control In root.Controls
            gfResetControls(subctrl)
            If Not TypeOf root Is GridView Then

                Dim strCtrlID As String = subctrl.ID

                If TypeOf subctrl Is TextBox Then
                    Dim tbCtrlTextBox As TextBox = CType(subctrl, TextBox)
                    tbCtrlTextBox.Text = ""

                ElseIf TypeOf subctrl Is RadioButtonList Then
                    Dim rbCtrlRadioButtonList As RadioButtonList = CType(subctrl, RadioButtonList)
                    rbCtrlRadioButtonList.SelectedIndex = 0

                ElseIf TypeOf subctrl Is RadioButton Then
                    Dim rbCtrlRadioButton As RadioButton = CType(subctrl, RadioButton)
                    If rbCtrlRadioButton.ID = "rdbAnd" Then
                        rbCtrlRadioButton.Checked = True
                    End If


                ElseIf TypeOf subctrl Is CheckBox Then
                    Dim ckCtrlCheckBox As CheckBox = CType(subctrl, CheckBox)
                    ckCtrlCheckBox.Checked = False

                ElseIf TypeOf subctrl Is DropDownList Then
                    Dim dplCtrlDropDown As DropDownList = CType(subctrl, DropDownList)
                    dplCtrlDropDown.SelectedIndex = 0


                End If


            End If


        Next



    End Sub

    Public Property gpstrFieldName() As String
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  14 Feb 2007
        ' Description:	Get and set FieldName
        '=============================================
        Get
            Return mvstrFieldName
        End Get
        Set(ByVal value As String)
            mvstrFieldName = value
        End Set
    End Property
    Public Property gpstrFieldValue() As String
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  14 Feb 2007
        ' Description:	Get and set FieldValue
        '=============================================
        Get
            Return mvstrFieldValue
        End Get
        Set(ByVal value As String)
            mvstrFieldValue = value
        End Set
    End Property
    Public Property gpctlFieldCtrl() As Control
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  14 Feb 2007
        ' Description:	Get and set FieldValue
        '=============================================
        Get
            Return mvctlFieldCtrl
        End Get
        Set(ByVal value As Control)
            mvctlFieldCtrl = value
        End Set
    End Property
    Public Property gpbooIsDBF() As Boolean
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  14 Feb 2007
        ' Description:	Get and set FieldValue
        '=============================================
        Get
            Return mvbooIsDBF
        End Get
        Set(ByVal value As Boolean)
            mvbooIsDBF = value
        End Set
    End Property


    Public Property gpintNo() As Integer
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  14 Feb 2007
        ' Description:	Get and set FieldValue
        '=============================================
        Get
            Return mvintNo
        End Get
        Set(ByVal value As Integer)
            mvintNo = value
        End Set
    End Property
    Public Property gpstrUserId() As String
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  14 Feb 2007
        ' Description:	Get and set UserId 
        '=============================================
        Get
            Return mvstrUserId
        End Get
        Set(ByVal value As String)
            mvstrUserId = value
        End Set
    End Property
    Public Property gpstrCurrentTime() As String
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  14 Feb 2007
        ' Description:	Get and set Current Time
        '=============================================
        Get
            Return mvstrCurrentTime
        End Get
        Set(ByVal value As String)
            mvstrCurrentTime = value
        End Set
    End Property
    Public Property gpstrId() As String
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  14 Feb 2007
        ' Description:	Get and set Id
        '=============================================
        Get
            Return mvstrId
        End Get
        Set(ByVal value As String)
            mvstrId = value
        End Set
    End Property
    Public Property gpNo() As Integer
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  14 Feb 2007
        ' Description:	Get and set Id
        '=============================================
        Get
            Return gmNo
        End Get
        Set(ByVal value As Integer)
            gmNo = value
        End Set
    End Property
    Public Property gpRowNo() As Integer
        ' =============================================
        ' Author:		Nancy Ng
        ' Create date:  14 Feb 2007
        ' Description:	Get and set Id
        '=============================================
        Get
            Return mvRowNo
        End Get
        Set(ByVal value As Integer)
            mvRowNo = value
        End Set
    End Property

    Public Function FormatDecimalwString(ByVal value As String, Optional ByVal numFormat As String = "###,###,##0.0000") As Decimal
        Dim nVal As New Decimal

        If value = "" Then value = "0"
        nVal = 0
        nVal = Convert.ToDecimal(value).ToString(numFormat)

        Return nVal

    End Function

    Public Function FormatIntegerString(ByVal value As String, Optional ByVal numFormat As String = "###,###,##0") As Decimal
        Dim nVal As New Decimal

        If value = "" Then value = "0"
        nVal = 0
        nVal = Convert.ToDecimal(value).ToString(numFormat)

        Return nVal

    End Function

    'Added by Ashwin
    Public Function CreateQRCode(ByVal rawText As String) As String

        Dim QrPath As String
        QrPath = Server.MapPath("../QRImg") + "/QRcode_" + System.DateTime.Now.ToFileTime() + ".png"
        Dim settings As New BarcodeSettings()
        settings.Type = BarCodeType.QRCode
        settings.Unit = System.Drawing.GraphicsUnit.Pixel
        settings.ShowText = False
        settings.ResolutionType = ResolutionType.UseDpi
        settings.X = 5.0F
        settings.Data = rawText
        Dim generator As New BarCodeGenerator(settings)
        Dim image As System.Drawing.Image = generator.GenerateImage()
        image.Save(QrPath)
        Return QrPath
    End Function

    'Added by Ashwin
End Class
