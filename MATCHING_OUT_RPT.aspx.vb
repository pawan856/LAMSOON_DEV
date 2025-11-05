
Imports System.Data
Imports System.IO
Imports ExcelLibrary.SpreadSheet
Imports System.Data.SqlClient
Imports System.Configuration

Partial Class MATCHING_OUT_RPT
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private UiFun As New UIfunc
    Dim sqlString As String
    Dim nDataSource As DataTable
    Protected Const FUN_CODE As String = "RPT_MATCHING_OUT"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils(FUN_CODE, Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If
        ar.hideForm(Me)
    End Sub

    Protected Sub btnDownloadExcel_Click(sender As Object, e As EventArgs) Handles btnDownloadExcel.Click
        Dim gConn = gDB.getConnection()
        Dim nDataSource As DataSet

        Try
            'For OUTSTANDING REPORT'
            sqlString = " exec sp_MachingOutRpt "

            nDataSource = gDB.getDataSet(sqlString, gConn)

            WriteXLSFile(nDataSource)

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
                    If (iSheetCount = 1) Then
                        worksheet = New Worksheet("IPR OUTSTANDING REPORT")
                    ElseIf (iSheetCount = 2) Then
                        worksheet = New Worksheet("NON PO REPORT")
                    ElseIf (iSheetCount = 3) Then
                        worksheet = New Worksheet("STOCK RETURN REPORT")
                    ElseIf (iSheetCount = 4) Then
                        worksheet = New Worksheet("STOCK ADJUSTMENT REPORT")
                    Else
                        worksheet = New Worksheet("STOCK BALANCE REPORT")
                    End If

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
                                    If Int32.TryParse(sTemp, iTemp) Then
                                        worksheet.Cells(iRow, iCol) = New Cell(Convert.ToInt32(iTemp), "0")
                                    Else
                                        worksheet.Cells(iRow, iCol) = New Cell(sTemp)
                                    End If
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
            Response.AddHeader("content-disposition", "attachment;filename=MatchingOutstandingReport.xls")
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

    Public Overrides Sub VerifyRenderingInServerForm(control As Control)
        ' Verifies that the control is rendered  
    End Sub

End Class
