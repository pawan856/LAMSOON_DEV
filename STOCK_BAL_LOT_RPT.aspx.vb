
Imports System.Data
Imports System.IO
Imports ExcelLibrary.SpreadSheet
Imports System.Data.SqlClient
Imports System.Configuration

Partial Class STOCK_BAL_LOT_RPT
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private UiFun As New UIfunc
    Dim sqlString As String
    Dim nDataSource As DataTable
    Protected Const FUN_CODE As String = "RPT_STOCK_BAL_LOT"

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
            'For Stock Balance after Lot Allocation REPORT '
            sqlString = " SELECT itm_loc_bal.STORER_CODE as STORER,	itm.ITM_CODE as 貨品內部編號,	itm.ITM_SKU_NO as 貨品編號,	itm.ITM_DESC AS 貨品名稱,itm_loc_bal.ILOC_WH as 子庫存,itm_loc_bal.ILOC_BATCH_NO as 批次 , " &
                        " itm.ITM_SHELF_LIFE aS 最少保質天數要求,	itm.ITM_UOM as 單位,	SUM(itm_loc_bal.ILOC_BAL_QTY) AS 庫存數量,SUM(rsvd.QTY) AS 已分佩數量,SUM(itm_loc_bal.ILOC_BAL_QTY)- SUM(rsvd.QTY) AS 分配餘數,  " &
                     " itm_loc_bal.ILOC_EXPIRY_DATE as 產品到期日 FROM WMS_ITEM_LOC_BAL itm_loc_bal 	LEFT JOIN WMS_WAVEPICK_RSVD rsvd  ON rsvd.wh_code = itm_loc_bal.ILOC_WH   AND rsvd.ITEM_CODE = itm_loc_bal.ITM_CODE AND itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO  " &
                     " AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE LEFT JOIN WMS_ITEM itm  ON itm.ITM_CODE = itm_loc_bal.ITM_CODE AND itm.STORER_CODE = itm_loc_bal.STORER_CODE WHERE  itm_loc_bal.ILOC_BAL_QTY > 0  " &
            " GROUP BY 	itm_loc_bal.STORER_CODE,itm_loc_bal.ILOC_WH,itm_loc_bal.ILOC_BATCH_NO,itm.ITM_CODE,	itm.ITM_DESC,rsvd.QTY,itm.ITM_SKU_NO,itm.ITM_SHELF_LIFE,itm.ITM_UOM,itm_loc_bal.ILOC_EXPIRY_DATE  " &
            " ORDER BY 	已分佩數量 DESC "

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

                    worksheet = New Worksheet("Stock Balance after Lot Allocation Report")

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
            Response.AddHeader("content-disposition", "attachment;filename=StockBalanceafterLotAllocationReport.xls")
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
