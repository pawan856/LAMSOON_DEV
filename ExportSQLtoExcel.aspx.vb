
Imports System.Data
Imports System.IO
Imports ExcelLibrary.SpreadSheet
Imports System.Data.SqlClient
Imports System.Configuration

Partial Class ExportSQLtoExcel
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private UiFun As New UIfunc
    Dim sqlString As String
    Dim nDataSource As DataTable
    Protected Const FUN_CODE As String = "RPT_EXPTOSQL"


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils(FUN_CODE, Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If
        If Not IsPostBack Then
            UiFun.load_ComboBox(ddlSelectForQuery, "select queryVal from QueryToExport order by queryVal", "queryVal", "queryVal", , Session("gSelectLabel"))
            ddlSelectForQuery.SelectedIndex = 0
        End If
        'GridTableData.Visible = False


        'UiFun.load_dropdown(ddlSelectForQuery, "select queryVal from QueryToExport ", "queryVal", "queryVal", , Session("gSelectLabel"))
        ar.hideForm(Me)
    End Sub

    Protected Sub ddlSelectForQuery_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlSelectForQuery.SelectedIndexChanged
        Dim newListItem As ListItem
        Dim gConn = gDB.getConnection()
        Dim nDataSource As New DataTable
        hdnParameter.Value = ""
        lblDescription.Text = ""
        lblCount.InnerText = ""
        GridTableData.Visible = False

        If ddlSelectForQuery.SelectedValue <> "" Then
            sqlString = "select querySelect,parameter,Description,editable from QueryToExport where queryVal='" & ddlSelectForQuery.SelectedValue & "'"
            nDataSource = gDB.getDataTable(sqlString, gConn)
            If nDataSource.Rows.Count > 0 Then
                Dim itmArray As String()
                lblDescription.Text = nDataSource.Rows(0).Item("Description").ToString
                txtSQL.InnerText = nDataSource.Rows(0).Item("querySelect").ToString
                'Dim editableBool As Boolean = nDataSource.Rows(0).Item("editable").ToString
                If nDataSource.Rows(0).Item("editable").ToString = "False" Then
                    txtSQL.Disabled = True
                Else
                    txtSQL.Disabled = False
                End If
                If (nDataSource.Rows(0).Item("parameter").ToString <> "") Then
                    hdnParameter.Value = nDataSource.Rows(0).Item("parameter").ToString
                    'Dim i As Integer = 0
                    'For Each txt As String In Split(nDataSource.Rows(0).Item("parameter").ToString, ",")
                    'i = i + 1
                    'Dim myTextBox As New TextBox
                    'Dim myLable As New Label

                    'If i Mod 3 = 0 Then
                    '    myLable.Text = "<br/><br/>  " & txt & ": "
                    'Else
                    '    myLable.Text = "  " & txt & ": "
                    'End If
                    'myLable.Style.Add("align", "left")
                    'myLable.Style.Add("font-Size", "20px")

                    'myTextBox.ID = txt
                    'myTextBox.CssClass = "TestTxt"
                    'myTextBox.Style.Add("font-Size", "20px")

                    'Me.rDropdown.Controls.Add(myLable)
                    'Me.rDropdown.Controls.Add(myTextBox)

                    Dim sb As StringBuilder = New StringBuilder()
                    sb.Append("<table id='tblparam'>")

                    For Each row As String In Split(hdnParameter.Value.ToString, "|")
                        sb.Append("<tr>")
                        For Each column As String In Split(row.ToString, ",")
                            sb.Append("<td style='background-color: #B8DBFD;border: 1px solid #ccc'>" & column & "</td><td><textarea id='" & column & "'  name='" & column & "'></textarea></td>")
                        Next
                        sb.Append("</tr>")
                    Next
                    sb.Append("</table>")

                    ltTable.Text = sb.ToString
                    'Next
                Else
                    ltTable.Text = ""
                End If

            End If
        Else
            txtSQL.InnerText = ""
        End If
    End Sub
    Protected Sub btnDownloadExcel_Click(sender As Object, e As EventArgs)
        Dim gConn = gDB.getConnection()
        Dim nDataSource As DataSet

        Try
            Dim Param As String = ""

            'For Stock Balance after Lot Allocation REPORT '
            sqlString = txtSQL.InnerText.ToString()
            If hdnParameter.Value <> "" Then
                Param = hdnParameter.Value.Replace("|", ",")
                For Each hdnVal As String In Split(Param.ToString, ",")
                    Dim txt As String = "'" & Request.Form(hdnVal) & "'"
                    If txt Is "" OrElse txt Is Nothing Then
                        txt = "''"
                    End If
                    sqlString = sqlString.Replace(hdnVal.Trim, txt.Trim)
                Next
            End If
            If sqlString.ToUpper().Contains("UPDATE ") OrElse sqlString.ToUpper().Contains("INSERT ") OrElse sqlString.ToUpper().Contains("DELETE ") OrElse sqlString.ToUpper().Contains("ALTER ") OrElse sqlString.ToUpper().Contains("TRUNCATE ") OrElse sqlString.ToUpper().Contains("CREATE ") OrElse sqlString.ToUpper().Contains("DROP ") Then
                UiFun.displayMsgNew(Me, "", "Only SELECT queries are allowed", Session("gLang"))
            Else
                If sqlString <> "" Then
                    nDataSource = gDB.getDataSet(sqlString, gConn)

                    WriteXLSFile(nDataSource)
                    Dim sb As StringBuilder = New StringBuilder()
                    sb.Append("<table id='tblparam'>")

                    For Each row As String In Split(hdnParameter.Value.ToString, "|")
                        sb.Append("<tr>")
                        For Each column As String In Split(row.ToString, ",")
                            If column <> "" Then
                                Dim txt As String = Request.Form(column)
                                sb.Append("<td style='background-color: #B8DBFD;border: 1px solid #ccc'>" & column & "</td><td><textarea id='" & column & "'  name='" & column & "'>" & txt.Trim & "</textarea></td>")
                            End If
                        Next
                        sb.Append("</tr>")
                    Next
                    sb.Append("</table>")

                    ltTable.Text = sb.ToString
                Else
                    UiFun.displayMsgNew(Me, "", "Please enter SQL query", Session("gLang"))
                End If
            End If

        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub


    Protected Sub btnSHowInTable_Click(sender As Object, e As EventArgs)
        Dim gConn = gDB.getConnection()
        Dim nDataSource As DataSet
        Dim str As StringBuilder = New StringBuilder()

        Dim Param As String = ""
        Try
            'For Stock Balance after Lot Allocation REPORT '
            sqlString = txtSQL.InnerText.ToString()

            If hdnParameter.Value <> "" Then
                Param = hdnParameter.Value.Replace("|", ",")
                For Each hdnVal As String In Split(Param.ToString, ",")
                    Dim txt As String = "'" & Request.Form(hdnVal) & "'"
                    If txt Is "" OrElse txt Is Nothing Then
                        txt = "''"
                    End If
                    sqlString = sqlString.Replace(hdnVal.Trim, txt.Trim)
                Next
            End If

            If sqlString.ToUpper().Contains("UPDATE ") OrElse sqlString.ToUpper().Contains("INSERT ") OrElse sqlString.ToUpper().Contains("DELETE ") OrElse sqlString.ToUpper().Contains("ALTER ") OrElse sqlString.ToUpper().Contains("TRUNCATE ") OrElse sqlString.ToUpper().Contains("CREATE ") OrElse sqlString.ToUpper().Contains("DROP ") Then
                UiFun.displayMsgNew(Me, "", "Only SELECT queries are allowed", Session("gLang"))
            Else
                If sqlString <> "" Then
                    nDataSource = gDB.getDataSet(sqlString, gConn)
                    lblCount.InnerText = "Total no. of rows :" & nDataSource.Tables(0).Rows.Count
                    GridTableData.Visible = True
                    GridTableData.DataSource = nDataSource
                    GridTableData.DataBind()


                    Dim sb As StringBuilder = New StringBuilder()
                    sb.Append("<table id='tblparam'>")

                    For Each row As String In Split(hdnParameter.Value.ToString, "|")
                        sb.Append("<tr>")
                        For Each column As String In Split(row.ToString, ",")
                            If column <> "" Then
                                Dim txt As String = Request.Form(column)
                                sb.Append("<td style='background-color: #B8DBFD;border: 1px solid #ccc'>" & column & "</td><td><textarea id='" & column & "'  name='" & column & "'>" & txt.Trim & "</textarea></td>")
                            End If
                        Next
                        sb.Append("</tr>")
                    Next
                    sb.Append("</table>")

                    ltTable.Text = sb.ToString
                Else
                    UiFun.displayMsgNew(Me, "", "Please enter SQL query", Session("gLang"))
                End If
            End If


        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub

    Public Function WriteXLSFile(ByVal pDataSet As DataSet) As Boolean
        Try
            'Create a workbook instance
            Dim workbook As Workbook = New Workbook()
            Dim worksheet As Worksheet
            Dim iRow As Integer = 0
            Dim iCol As Integer = 0
            Dim sTemp As String = String.Empty
            Dim dtTemp As DateTime
            Dim dTemp As Double = 0
            Dim iTemp As Integer = 0
            Dim count As Integer = 0
            Dim iTotalRows As Integer = 0
            Dim iSheetCount As Integer = 0

            'Read DataSet
            If Not pDataSet Is Nothing And pDataSet.Tables.Count > 0 Then

                'Traverse DataTable inside the DataSet
                For Each dt As DataTable In pDataSet.Tables

                    'Create a worksheet instance
                    iSheetCount = iSheetCount + 1

                    worksheet = New Worksheet("SQL Export" + iSheetCount.ToString())

                    'Write Table Header
                    iCol = 0
                    For Each dc As DataColumn In dt.Columns
                        worksheet.Cells(0, iCol) = New Cell(dc.ColumnName)
                        iCol = iCol + 1
                    Next

                    'Write Table Body
                    iRow = 1
                    For Each dr As DataRow In dt.Rows
                        iCol = 0
                        For Each dc As DataColumn In dt.Columns
                            sTemp = dr(dc.ColumnName).ToString()
                            Select Case dc.DataType
                                Case GetType(DateTime)
                                    DateTime.TryParse(sTemp, dtTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dtTemp, "MM/DD/YYYY hh:mm:ss")
                                Case GetType(Double)
                                    Double.TryParse(sTemp, dTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.0000")
                                Case GetType(Decimal)
                                    Decimal.TryParse(sTemp, dTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.0000")
                                Case Else
                                    worksheet.Cells(iRow, iCol) = New Cell(sTemp)
                            End Select
                            iCol = iCol + 1
                        Next
                        iRow = iRow + 1
                    Next

                    ''Attach worksheet to workbook
                    workbook.Worksheets.Add(worksheet)
                    iTotalRows = iTotalRows + iRow
                Next

            End If

            'Bug on Excel Library, min file size must be 7 Kb
            'thus we need to add empty row for safety
            If iTotalRows < 100 Then
                worksheet = New Worksheet("Sheet X")
                count = 1
                Do While count < 100
                    worksheet.Cells(count, 0) = New Cell(" ")
                    count = count + 1
                Loop
                workbook.Worksheets.Add(worksheet)
            End If

            'Export the Excel file.
            Response.Clear()
            Response.Buffer = True
            Response.Charset = ""
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Response.AddHeader("content-disposition", "attachment;filename=SQLExport.xls")
            Using MyMemoryStream As New MemoryStream()
                workbook.Save(MyMemoryStream)
                MyMemoryStream.WriteTo(Response.OutputStream)
                Response.Flush()
            End Using

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Protected Sub btnReset_Click(sender As Object, e As EventArgs)
        'Response.Redirect(HttpContext.Current.Request.Url.ToString(), True)
        Dim sb As StringBuilder = New StringBuilder()
        sb.Append("<table id='tblparam'>")

        For Each row As String In Split(hdnParameter.Value.ToString, "|")
            sb.Append("<tr>")
            For Each column As String In Split(row.ToString, ",")
                If column <> "" Then
                    sb.Append("<td style='background-color: #B8DBFD;border: 1px solid #ccc'>" & column & "</td><td><textarea id='" & column & "'  name='" & column & "'></textarea></td>")
                End If
            Next
            sb.Append("</tr>")
        Next
        sb.Append("</table>")
        ltTable.Text = sb.ToString
    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(control As Control)
        ' Verifies that the control is rendered  
    End Sub

End Class
