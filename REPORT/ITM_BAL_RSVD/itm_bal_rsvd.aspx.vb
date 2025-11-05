Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Globalization
Imports Microsoft.Reporting.WebForms.Internal.Soap.ReportingServices2005.Execution
Imports Microsoft.Reporting.WebForms
Imports ExcelLibrary.SpreadSheet
Imports System.Drawing

Partial Class REPORT_ITM_BAL_RSVD_itm_bal_rsvd
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private wmsFun As New WMSFunc
    Private UiFun As New UIfunc
    Private rptU As New ReportUtils
    Dim sqlString As String
    Dim nDataSource As DataTable
    Protected Const FUN_CODE As String = "RPT_CUS_SUPPORT"

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils(FUN_CODE, Session("usr_id"), Me)
        ar.hideForm(Me)

        Dim do_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        ViewState("STOERE_CODE") = ""
        ViewState("DO_CODE") = ""

        ViewState("DO_CODE") = Server.UrlDecode(Request("DO_CODE"))
        ViewState("STORER_CODE") = Session("str_code")

        If Not IsPostBack Then
            BindData()
        End If
    End Sub

    Private Sub BindData()
        Dim storer_code As String = ""
        storer_code = Session("str_code")

        sqlString = "SELECT itm_loc_bal.ILOC_WH AS WH_CODE,itm_loc_bal.ILOC_BATCH_NO LOT_NO,itm.ITM_CODE,itm.ITM_DESC,itm.ITM_SKU_NO,itm.ITM_SHELF_LIFE,SUM(itm_loc_bal.ILOC_BAL_QTY) AS Stock_Balance,itm_loc_bal.ILOC_EXPIRY_DATE as 'EXPIRY_DATE',itm.ITM_UOM as 'UOM', " &
                    " SUM(isnull(rsvd.QTY,0)) AS ALLOCATED,SUM(itm_loc_bal.ILOC_BAL_QTY)- SUM(isnull(rsvd.QTY,0)) AS Balance FROM WMS_ITEM_LOC_BAL itm_loc_bal " &
                    " LEFT JOIN WMS_WAVEPICK_RSVD rsvd ON rsvd.wh_code = itm_loc_bal.ILOC_WH AND rsvd.ITEM_CODE = itm_loc_bal.ITM_CODE " &
                    " AND itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE INNER JOIN WMS_ITEM itm ON itm.ITM_CODE = itm_loc_bal.ITM_CODE " &
                    " AND itm.STORER_CODE = itm_loc_bal.STORER_CODE LEFT JOIN WMS_DELV_ORDER DO ON DO.DO_CODE = rsvd.DO_CODE " &
                    " WHERE itm_loc_bal.STORER_CODE = '" & storer_code & "' AND itm_loc_bal.ILOC_BAL_QTY > 0 AND itm_loc_bal.ITM_CODE IN " &
                    " (SELECT COD_ITM_CODE FROM WMS_CUST_ORDER_D cod WHERE ALLOCATED = 'N' AND STORER_CODE = '" & storer_code & "' " &
                    " AND COD_WH_CODE = itm_loc_bal.ILOC_WH AND (SELECT COUNT(1) FROM WMS_DELV_ORDER_D WHERE DOD_CO_CODE LIKE '%' + cod.CO_CODE + '%' " &
                    " AND DOD_ITM_CODE = cod.COD_ITM_CODE)=0) GROUP BY itm_loc_bal.ILOC_WH,itm_loc_bal.ILOC_BATCH_NO,itm.ITM_CODE,itm.ITM_DESC,itm.ITM_SKU_NO, " &
                    " itm.ITM_SHELF_LIFE,itm_loc_bal.ILOC_EXPIRY_DATE,itm.ITM_UOM ORDER BY itm.ITM_SKU_NO DESC"
        nDataSource = gDB.getDataTable(sqlString)

        If nDataSource.Rows.Count > 0 Then
            Dim ITMBALRSVDReportModels As New List(Of ITMBALRSVDReportModel)
            For index As Integer = 0 To nDataSource.Rows.Count - 1
                Dim ITMBALRSVD As New ITMBALRSVDReportModel
                ITMBALRSVD.Stock_Balance = nDataSource.Rows(index)("Stock_Balance").ToString()
                'ITMBALRSVD.ILOC_EXPIRY_DATE = nDataSource.Rows(index)("EXPIRY_DATE").ToString()
                ITMBALRSVD.ITM_CODE = nDataSource.Rows(index)("ITM_CODE").ToString()
                ITMBALRSVD.ITM_DESC = nDataSource.Rows(index)("ITM_DESC").ToString()
                ITMBALRSVD.ITM_UOM = nDataSource.Rows(index)("UOM").ToString()
                ITMBALRSVD.ITM_SHELF_LIFE = nDataSource.Rows(index)("ITM_SHELF_LIFE").ToString()
                ITMBALRSVD.ITM_SKU_NO = nDataSource.Rows(index)("ITM_SKU_NO").ToString()
                ITMBALRSVD.lot_no = nDataSource.Rows(index)("lot_no").ToString()
                ITMBALRSVD.Allocated = nDataSource.Rows(index)("Allocated").ToString()
                ITMBALRSVD.WH_CODE = nDataSource.Rows(index)("WH_CODE").ToString()
                ITMBALRSVD.wh_loc = "" 'nDataSource.Rows(index)("wh_loc").ToString()
                ITMBALRSVD.Balance = nDataSource.Rows(index)("Balance").ToString()
                Dim dd As String = nDataSource.Rows(index)("EXPIRY_DATE").ToString()
                Dim ExpDate As DateTime = DateTime.Parse(dd, CultureInfo.InvariantCulture)
                Dim reformatted As String = ExpDate.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
                ITMBALRSVD.ILOC_EXPIRY_DATE = reformatted
                ITMBALRSVDReportModels.Add(ITMBALRSVD)
            Next

            Dim rootITMBALRSVD As List(Of RootITMBALRSVDReportModel) = ITMBALRSVDReportModels.GroupBy(Function(person) New With {Key person.ITM_SKU_NO, Key person.ITM_DESC}).
Select(Function(grp) New RootITMBALRSVDReportModel With {
.ITM_SKU_NO = grp.Key.ITM_SKU_NO,
.ITM_DESC = grp.Key.ITM_DESC,
.Total_Stock_Balance = grp.Sum(Function(item) Decimal.Parse(item.Stock_Balance)),
.Total_Allocated = grp.Sum(Function(item) Decimal.Parse(item.Allocated)),
.Total_Balance = grp.Sum(Function(item) Decimal.Parse(item.Balance)),
.ITMBALRSVD = grp.ToList()
}).ToList()

            rptCustomers.DataSource = rootITMBALRSVD
            rptCustomers.DataBind()
        End If

    End Sub

    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Response.Clear()
        Response.Buffer = True
        Response.AddHeader("content-disposition", "attachment;filename=ItemReservedBalanceExport.xls")
        Response.Charset = ""
        Response.ContentType = "application/vnd.ms-excel"
        Dim sw As New StringWriter()
        Dim hw As New HtmlTextWriter(sw)
        rptCustomers.RenderControl(hw)
        Response.Output.Write(sw.ToString())
        Response.Flush()
        Response.End()
    End Sub

    Protected Sub btnExportDataTable_Click(sender As Object, e As EventArgs) Handles btnExportDataTable.Click
        Dim storer_code As String = ""
        storer_code = Session("str_code")
        Dim nDataSource As DataSet
        sqlString = " SELECT itm.ITM_SKU_NO AS 貨品編號, itm.ITM_DESC AS 貨品名稱, itm_loc_bal.ILOC_WH AS 子庫存, itm_loc_bal.ILOC_BATCH_NO AS 批次, " &
"  itm.ITM_UOM AS 單位, SUM(itm_loc_bal.ILOC_BAL_QTY) AS 庫存數量, ( SELECT ISNULL( SUM(rsvd.QTY), 0 " &
"  ) FROM WMS_WAVEPICK_RSVD rsvd WHERE rsvd.wh_code = itm_loc_bal.ILOC_WH AND rsvd.ITEM_CODE = itm.ITM_CODE AND  " &
"  itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE " &
"  ) AS 已分佩數量, SUM( ISNULL(itm_loc_bal.ILOC_BAL_QTY, 0) )- ( SELECT ISNULL( SUM(rsvd.QTY), 0 ) FROM WMS_WAVEPICK_RSVD rsvd  " &
"  WHERE rsvd.wh_code = itm_loc_bal.ILOC_WH AND rsvd.ITEM_CODE = itm.ITM_CODE AND itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO  " &
"  AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE ) AS 分配餘數, FORMAT( itm_loc_bal.ILOC_EXPIRY_DATE, 'dd MMM yy' ) AS 產品到期日, " &
"  CAST( DATEDIFF( DAY, GETDATE(), ILOC_EXPIRY_DATE ) AS DECIMAL ) AS 剩余日數, " &
"  itm.ITM_SHELF_LIFE 保質天數, itm_loc_bal.STORER_CODE AS STORER, itm.ITM_CODE AS 貨品內部編號 FROM " &
"  WMS_ITEM_LOC_BAL itm_loc_bal LEFT JOIN WMS_ITEM itm ON itm.ITM_CODE = itm_loc_bal.ITM_CODE AND itm.STORER_CODE = itm_loc_bal.STORER_CODE  " &
"  LEFT JOIN WMS_WAREHOUSE WH ON itm_loc_bal.ILOC_WH = wh.WH_CODE WHERE itm_loc_bal.ILOC_BAL_QTY > = 0 AND wh.WH_CONT_PER1 = 'Y'  And itm_loc_bal.STORER_CODE = '" & storer_code & "' " &
"  GROUP BY itm_loc_bal.STORER_CODE, itm_loc_bal.ILOC_WH, itm_loc_bal.ILOC_BATCH_NO, itm.ITM_CODE, itm.ITM_DESC,  " &
"  itm.ITM_SKU_NO, itm.ITM_SHELF_LIFE, itm.ITM_UOM, itm_loc_bal.ILOC_EXPIRY_DATE, wh.WH_CONT_PER1 HAVING wh.WH_CONT_PER1 = 'Y'  " &
"  AND( SUM(itm_loc_bal.ILOC_BAL_QTY) > 0 OR ( SELECT ISNULL( SUM(rsvd.QTY), 0 ) FROM WMS_WAVEPICK_RSVD rsvd  " &
"  WHERE rsvd.wh_code = itm_loc_bal.ILOC_WH AND rsvd.ITEM_CODE = itm.ITM_CODE AND itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO  " &
"  AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE ) > 0 )  ORDER BY 貨品編號,子庫存 "

        sqlString = sqlString + "  SELECT itm.ITM_SKU_NO 貨品編號, itm.ITM_DESC 貨品名稱, itm_loc_bal.ILOC_WH AS 子庫存, itm_loc_bal.ILOC_BATCH_NO 批次, itm.ITM_SHELF_LIFE  保質天數, SUM(itm_loc_bal.ILOC_BAL_QTY)  " &
"  AS 庫存數量, SUM( ISNULL(rsvd.QTY, 0) ) AS 分配數量, SUM(itm_loc_bal.ILOC_BAL_QTY)- SUM(  ISNULL(rsvd.QTY, 0) ) AS 分配後數量, itm.ITM_UOM AS '單位', itm_loc_bal.ILOC_EXPIRY_DATE  " &
"  AS 產品到期日, CAST(DATEDIFF(DAY,GETDATE(), ILOC_EXPIRY_DATE  ) AS DECIMAL ) AS 剩余日數 FROM WMS_ITEM_LOC_BAL itm_loc_bal LEFT JOIN WMS_WAVEPICK_RSVD rsvd  " &
"  ON rsvd.wh_code = itm_loc_bal.ILOC_WH AND rsvd.ITEM_CODE = itm_loc_bal.ITM_CODE AND itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE " &
"   INNER JOIN WMS_ITEM itm ON itm.ITM_CODE = itm_loc_bal.ITM_CODE AND itm.STORER_CODE = itm_loc_bal.STORER_CODE LEFT JOIN WMS_DELV_ORDER DO ON DO.DO_CODE = rsvd.DO_CODE " &
"  WHERE  itm_loc_bal.STORER_CODE = '" & storer_code & "' AND itm_loc_bal.ILOC_BAL_QTY > 0 AND itm_loc_bal.ITM_CODE IN ( SELECT COD_ITM_CODE  FROM   WMS_CUST_ORDER_D cod " &
"    WHERE   ALLOCATED = 'N'   AND    STORER_CODE = '" & storer_code & "'   AND COD_WH_CODE = itm_loc_bal.ILOC_WH  AND ( SELECT COUNT(1)    FROM WMS_DELV_ORDER_D " &
"      WHERE DOD_CO_CODE LIKE '%' + cod.CO_CODE + '%' AND DOD_ITM_CODE = cod.COD_ITM_CODE   )=0 ) GROUP BY itm_loc_bal.ILOC_WH, itm_loc_bal.ILOC_BATCH_NO, itm.ITM_CODE, " &
"  itm_sku_no, itm.ITM_DESC, rsvd.QTY, itm.ITM_SKU_NO, itm.ITM_SHELF_LIFE, itm.ITM_UOM, itm_loc_bal.ILOC_EXPIRY_DATE ORDER BY itm.ITM_SKU_NO DESC  "

        nDataSource = gDB.getDataSet(sqlString)
        Try
            Dim excelFolder As String
            excelFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempFiles\" & "_ItemBalanceItems_" & System.DateTime.UtcNow.ToString("MM-dd-yyyy hh-mm-ss") & ".xls")
            Dim _Result As Boolean = WriteXLSFile(excelFolder, nDataSource)
            If _Result = True Then
                Response.ContentType = ContentType
                Response.AppendHeader("Content-Disposition", ("attachment; filename=" + Path.GetFileName(excelFolder)))
                Response.WriteFile(excelFolder)
                Response.End()
            End If
        Catch ex As Exception
            Response.Write(ex.Message)
            UiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
        End Try

    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(control As Control)
        ' Verifies that the control is rendered  
    End Sub

    Public Function WriteXLSFile(ByVal pFileName As String, ByVal pDataSet As DataSet) As Boolean
        Try
            'Create a workbook instance
            Dim workbook As Workbook = New Workbook()
            Dim worksheet As Worksheet
            Dim iRow As Integer = 0
            Dim iCol As Integer = 0
            Dim sTemp As String = String.Empty
            Dim dTemp As Double = 0
            Dim iTemp As Integer = 0
            Dim dtTemp As DateTime
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
                        worksheet = New Worksheet("Item Balance Report")
                    Else
                        worksheet = New Worksheet("Non allocated item balance")
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
                                    worksheet.Cells(iRow, iCol) = New Cell(dtTemp, "MM/DD/YYYY")
                                Case GetType(Double)
                                    Double.TryParse(sTemp, dTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.00")
                                Case GetType(Decimal)
                                    Decimal.TryParse(sTemp, dTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.00")
                                Case Else
                                    worksheet.Cells(iRow, iCol) = New Cell(sTemp)
                            End Select
                            iCol = iCol + 1
                        Next
                        iRow = iRow + 1
                    Next

                    'Attach worksheet to workbook
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

            workbook.Save(pFileName)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class
