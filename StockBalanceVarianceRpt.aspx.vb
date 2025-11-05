
Imports System.Data
Imports System.IO

Partial Class StockBalanceVarianceRpt
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private UiFun As New UIfunc
    Dim sqlString As String
    Dim nDataSource As DataTable
    Protected Const FUN_CODE As String = "RPT_BAL_VAR"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils(FUN_CODE, Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        ar.hideForm(Me)

        If Not IsPostBack Then
            BindData()
        End If

    End Sub

    Private Sub BindData()
        Try

            sqlString = "SELECT aa.ITM_CODE as 'WMS ITEM CODE',(case when isnull(aa.itm_code,'')='' then bb.ITEM_NUMBER else (select ITM_SKU_NO from WMS_ITEM where STORER_CODE=aa.STORER_CODE and ITM_CODE=aa.ITM_CODE) end) as ITM_SKU_NO, " &
            " (case when isnull(aa.itm_code,'')='' then (select ITM_NAME from WMS_ITEM where STORER_CODE=bb.IO_ID and ITM_CODE=bb.ITEM_ID) else (select ITM_NAME from WMS_ITEM where STORER_CODE=aa.STORER_CODE and ITM_CODE=aa.ITM_CODE) end) as ITM_NAME, " &
            " aa.ILOC_WH as 'WMS WAREHOUSE',aa.ILOC_BATCH_NO as 'WMS LOT #',aa.ILOC_LOC as 'WMS LOCATION',aa.ILOC_BAL_QTY as 'WMS STOCK BALANCE',bb.ITEM_ID AS 'EBS ITEM CODE',bb.SUBINVENTORY_CODE as 'EBS WAREHOUSE',bb.LOT_NUMBER as 'EBS LOT #', " &
            " PARSENAME(bb.LOCATOR, 4) + PARSENAME(bb.LOCATOR, 3) + PARSENAME(bb.LOCATOR, 2) + PARSENAME(bb.LOCATOR, 1) 'EBS LOCATION',bb.QTY as 'EBS QTY',aa.ILOC_BAL_QTY - bb.QTY as 'Different in QTY' FROM dbo.WMS_ITEM_LOC_BAL aa FULL OUTER JOIN EBS_WMS_STOCK_ONHAND bb ON " &
            " aa.ITM_CODE = bb.ITEM_ID AND aa.ILOC_WH = bb.SUBINVENTORY_CODE AND aa.ILOC_BATCH_NO = bb.LOT_NUMBER AND aa.STORER_CODE = bb.IO_ID AND aa.ILOC_LOC = PARSENAME(bb.LOCATOR, 4) + PARSENAME(bb.LOCATOR, 3) + PARSENAME(bb.LOCATOR, 2) + PARSENAME(bb.LOCATOR, 1) " &
            " where abs(aa.ILOC_BAL_QTY - bb.QTY) >= 0.01 or bb.QTY is NULL or aa.ILOC_BAL_QTY is Null"

            nDataSource = gDB.getDataTable(sqlString)

            If nDataSource IsNot Nothing AndAlso nDataSource.Rows.Count > 0 Then
                GridTableData.Visible = True
                btnExportExcel.Visible = True
                GridTableData.DataSource = nDataSource
                GridTableData.DataBind()
            Else
                GridTableData.Visible = False
                btnExportExcel.Visible = False
            End If

        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub

    Protected Sub btnExportExcel_Click(sender As Object, e As EventArgs) Handles btnExportExcel.Click
        Try
            Response.Clear()
            Response.Buffer = True
            Response.AddHeader("content-disposition", "attachment;filename=StockBalanceVarianceExport.xls")
            Response.Charset = ""
            Response.ContentType = "application/vnd.ms-excel"
            Using sw As New System.IO.StringWriter()
                Dim hw As New HtmlTextWriter(sw)
                GridTableData.RenderControl(hw)
                Response.Output.Write(sw.ToString())
                Response.Flush()
                Response.[End]()
            End Using
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub

End Class
