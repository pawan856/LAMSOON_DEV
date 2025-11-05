Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Globalization
Imports Microsoft.Reporting.WebForms.Internal.Soap.ReportingServices2005.Execution
Imports Microsoft.Reporting.WebForms
Imports System.Drawing
Imports Microsoft.ReportingServices.Rendering.ExcelRenderer

Partial Class REPORT_CUST_SUPPORT_cust_support
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private wmsFun As New WMSFunc
    Private UiFun As New UIfunc
    Dim sqlString As String
    Dim nDataSource As DataTable
    Protected Const FUN_CODE As String = "RPT_CUS_SUPPORT"

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils(FUN_CODE, Session("usr_id"), Me)

        If Session("usr_id") Is Nothing Or Session("usr_id") = "" Then
            Session.Remove("PAGE_SESSION_MENU_CODE")
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        ar.hideForm(Me)

        Dim imp_code As String = ""

        'Dim Paras(6) As Microsoft.Reporting.WebForms.ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        ViewState("STOERE_CODE") = ""
        ViewState("CO_CODE") = ""

        ViewState("CO_CODE") = Server.UrlDecode(Request("CO_CODE"))
        ViewState("STORER_CODE") = Session("str_code")

        If Not IsPostBack Then
            'BindData()

        End If
    End Sub

    Private Sub BindData()
        Dim storer_code As String = ""

        storer_code = Session("str_code")

        sqlString = "select t2.* from(select t.* from(SELECT (Select TOP(1) IO_CODE from EBS_WMS_COMPANY_MASTER where IO_ID=cod.STORER_CODE) IO_CODE, co.CO_INV_NO,itm.ITM_SKU_NO as SKU_NUMBER,'' FORMULA, cod.COD_ITM_DESC as SKU_NAME, cod.COD_Ticket_no CUST_PO_NUMBER,cod.COD_Carton_no, cod.COD_UOM,co.ROUTE_ID,convert(varchar(10),(co.CO_DATE),120) CO_DATE,cod.COD_Plant, co.CUS_NAME   as CUS_NAME, co.CUS_CODE,cod.CO_CODE As CUS_ORDER,co.CO_SENDER_REGION,  isnull((select top(1) t1.MIN_PROD_LIFE from  (Select ISNULL(MIN_PROD_LIFE,0)MIN_PROD_LIFE from WMS_CUSTOMER_RULE   where item_code = cod.COD_ITM_CODE And CUST_CODE = co.CUS_CODE union  Select ISNULL(MIN_PROD_LIFE,0)MIN_PROD_LIFE from WMS_CUSTOMER   Where  storer_code = cod.STORER_CODE And CUS_CODE = co.CUS_CODE) t1  order by  t1.MIN_PROD_LIFE desc),0) as DOD_MIN_PROD_DATE,  isnull((select top(1) t1.MIN_SELF_LIFE from  (Select ISNULL(MIN_SELF_LIFE,0)MIN_SELF_LIFE from WMS_CUSTOMER_RULE   where item_code = cod.COD_ITM_CODE And CUST_CODE = co.CUS_CODE union  Select ISNULL(MIN_SELF_LIFE,0)MIN_SELF_LIFE from WMS_CUSTOMER   Where  storer_code = cod.STORER_CODE And CUS_CODE = co.CUS_CODE) t1  order by  t1.MIN_SELF_LIFE desc),0) as DOD_MIN_SHELF_LIFE,  cod.STORER_CODE,cod.COD_WH_CODE,cod.COD_QTY,  (case when cod.ALLOCATED is Null and dod.DO_CODE is not NUll and dod.DOD_ITM_CODE=cod.COD_ITM_CODE then 'Y' when cod.ALLOCATED='N' and dod.DO_CODE   is not NUll and dod.DOD_ITM_CODE=cod.COD_ITM_CODE then 'Y' when cod.ALLOCATED='N' and dod.DO_CODE is not NUll and dod.DOD_ITM_CODE!=cod.COD_ITM_CODE then 'N'   when cod.ALLOCATED=NULL And dod.DO_CODE Is NUll then 'N' when cod.ALLOCATED='N' and dod.DO_CODE is NuLL then 'N' else 'None' end) as ALLOCATED,cod.COD_ITM_CODE,   cod.COD_PALLET_NO, dod.DO_CODE, cod.COD_PALLET_NO as DOD_LAST_LOT,   isnull((select top(1) t1.MAX_LOTS from  (Select ISNULL(MAX_LOTS,0)MAX_LOTS from WMS_CUSTOMER_RULE   where item_code=cod.COD_ITM_CODE and CUST_CODE=co.CUS_CODE union  Select ISNULL(MAX_LOTS,0)MIN_PROD_LIFE from WMS_CUSTOMER   Where STORER_CODE=cod.STORER_CODE and CUS_CODE=co.CUS_CODE) t1  order by  t1.MAX_LOTS desc),0) as DOD_MAX_LOT   FROM WMS_CUST_ORDER_D cod inner join WMS_CUST_ORDER co On co.CO_CODE=cod.CO_CODE   Left JOIN WMS_DELV_ORDER_D dod on cod.COD_ITM_CODE=dod.DOD_ITM_CODE AND cod.CO_CODE in (select items from dbo.split(dod.DOD_CO_CODE,','))   Left Join WMS_ITEM itm On cod.IMP_CODE = itm.IMP_CODE And cod.STORER_CODE = itm.STORER_CODE And cod.COD_ITM_CODE = itm.ITM_CODE And cod.COD_PACK_KEY = itm.PACK_KEY   Group BY itm.ITM_SKU_NO,cod.COD_ITM_DESC,cod.COD_Ticket_no,co.CUS_NAME,cod.COD_UOM,co.ROUTE_ID,co.CO_DATE,co.CUS_CODE,cod.CO_CODE,co.CO_SENDER_REGION, cod.COD_WH_CODE,cod.COD_QTY,cod.ALLOCATED,co.CUS_CODE,cod.COD_Carton_no,cod.COD_Plant,cod.COD_ITM_CODE,cod.COD_PALLET_NO,dod.DO_CODE,dod.DOD_ITM_CODE,   dod.DOD_LAST_LOT,co.CO_INV_NO,cod.STORER_CODE) as t where t.ALLOCATED!='None') as t2 where t2.STORER_CODE='" & storer_code & "' group by   t2.CO_INV_NO,t2.SKU_NUMBER,t2.CUST_PO_NUMBER,t2.FORMULA,t2.COD_Carton_no,t2.CO_DATE,t2.cod_uom,t2.ROUTE_ID,t2.CUS_CODE,t2.COD_Plant,t2.IO_CODE,t2.SKU_NAME,t2.CUS_NAME,t2.CUS_ORDER,t2.CO_SENDER_REGION,t2.DOD_MIN_PROD_DATE,t2.DOD_MIN_SHELF_LIFE,t2.COD_WH_CODE,   t2.COD_QTY,t2.ALLOCATED,t2.COD_ITM_CODE,t2.COD_PALLET_NO,t2.DO_CODE,t2.DOD_LAST_LOT,t2.DOD_MAX_LOT,t2.STORER_CODE order by t2.CUS_ORDER  "

        nDataSource = gDB.getDataTable(sqlString)

        If nDataSource.Rows.Count > 0 Then
            Dim CustSupportReportModels As New List(Of CustSupportReportModel)
            For index As Integer = 0 To nDataSource.Rows.Count - 1
                Dim CustSupport As New CustSupportReportModel
                CustSupport.SKU_NUMBER = nDataSource.Rows(index)("SKU_NUMBER").ToString()
                CustSupport.SKU_NAME = nDataSource.Rows(index)("SKU_NAME").ToString()
                'CustSupport.DOD_CO_CODE = nDataSource.Rows(index)("DOD_CO_CODE").ToString()
                CustSupport.CUS_ORDER = nDataSource.Rows(index)("CUS_ORDER").ToString()
                CustSupport.CUS_NAME = nDataSource.Rows(index)("CUS_NAME").ToString()
                CustSupport.CO_SENDER_REGION = nDataSource.Rows(index)("CO_SENDER_REGION").ToString()
                CustSupport.DOD_MIN_SHELF_LIFE = nDataSource.Rows(index)("DOD_MIN_SHELF_LIFE").ToString()
                CustSupport.DOD_MAX_LOT = nDataSource.Rows(index)("DOD_MAX_LOT").ToString()
                CustSupport.DOD_MIN_PROD_DATE = nDataSource.Rows(index)("DOD_MIN_PROD_DATE").ToString()
                CustSupport.COD_WH_CODE = nDataSource.Rows(index)("COD_WH_CODE").ToString()
                CustSupport.COD_PALLET_NO = nDataSource.Rows(index)("COD_PALLET_NO").ToString()
                CustSupport.COD_QTY = nDataSource.Rows(index)("COD_QTY").ToString()
                CustSupport.ALLOCATED = nDataSource.Rows(index)("ALLOCATED").ToString()
                CustSupport.CO_INV_NO = nDataSource.Rows(index)("CO_INV_NO").ToString()

                CustSupport.IO_CODE = nDataSource.Rows(index)("IO_CODE").ToString()
                CustSupport.CUST_PO_NUMBER = nDataSource.Rows(index)("CUST_PO_NUMBER").ToString()
                CustSupport.COD_Plant = nDataSource.Rows(index)("COD_Plant").ToString()
                CustSupport.CUS_CODE = nDataSource.Rows(index)("CUS_CODE").ToString()
                CustSupport.COD_Carton_no = nDataSource.Rows(index)("COD_Carton_no").ToString()
                CustSupport.Formula = nDataSource.Rows(index)("Formula").ToString()
                CustSupport.COD_UOM = nDataSource.Rows(index)("COD_UOM").ToString()
                CustSupport.CO_DATE = nDataSource.Rows(index)("CO_DATE").ToString()
                CustSupport.ROUTE_ID = nDataSource.Rows(index)("ROUTE_ID").ToString()
                CustSupportReportModels.Add(CustSupport)
            Next

            Dim rootCustSupports As List(Of RootCustSupportReportModel) = CustSupportReportModels.GroupBy(Function(person) New With {Key person.SKU_NUMBER, Key person.SKU_NAME}).
Select(Function(grp) New RootCustSupportReportModel With {
.SKU_NUMBER = grp.Key.SKU_NUMBER,
.SKU_NAME = grp.Key.SKU_NAME,
.Total_QTY = grp.Sum(Function(item) Decimal.Parse(item.COD_QTY)),
.CustSupports = grp.ToList()
}).ToList()

            rptCustomers.DataSource = rootCustSupports
            rptCustomers.DataBind()
        End If
    End Sub

    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Response.Clear()
        Response.Buffer = True
        Response.AddHeader("content-disposition", "attachment;filename=CustomeOrderAlloacationReport.xls")
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

        sqlString = " SELECT t2.* FROM (SELECT t.* FROM (SELECT (SELECT TOP(1) IO_CODE FROM EBS_WMS_COMPANY_MASTER WHERE IO_ID = cod.STORER_CODE) 倉庫編號, " &
                    " (CASE WHEN co.CO_SENDER_REGION = 'FALSE' AND ((cod.ALLOCATED IS NULL AND dod.DO_CODE IS NOT NULL AND dod.DOD_ITM_CODE = cod.COD_ITM_CODE)) THEN '' ELSE 'Y' END) AS 放單檢查, " &
                    " (CASE WHEN co.CO_SENDER_REGION = 'FALSE' THEN 'N' ELSE 'Y' END) AS 暫停信貸,(CASE WHEN CO_FTRACK_NO='ACTUAL' and  (dod.DO_CODE is NULL or cod.ALLOCATED = 'N') THEN 'A-NP' " &
                    " WHEN cod.ALLOCATED Is NULL And dod.DO_CODE Is NULL THEN 'Y' WHEN cod.ALLOCATED Is NULL And dod.DO_CODE Is Not NULL And dod.DOD_ITM_CODE = cod.COD_ITM_CODE THEN 'N' " &
                    " WHEN cod.ALLOCATED = 'N' And dod.DO_CODE Is Not NULL And dod.DOD_ITM_CODE = cod.COD_ITM_CODE THEN 'N' WHEN cod.ALLOCATED = 'N' AND dod.DO_CODE IS NOT NULL " &
                    " And dod.DOD_ITM_CODE ! = cod.COD_ITM_CODE THEN 'Y' WHEN cod.ALLOCATED = 'N' AND dod.DO_CODE IS NULL THEN 'Y' ELSE 'N' END) AS 分貨失敗,cod.COD_WH_CODE AS 子庫存, " &
                    " (SELECT TOP(1) DELIVERY_ID FROM EBS_WMS_TRANS_ITX_SO_HEADER WHERE SO_HEADER_ID = cod.COD_REF_NO ORDER BY TRANSACTION_ID DESC) AS 發票編號,	co.CO_INV_NO AS 訂單編號, " &
                    " cod.COD_Plant AS 訂單行號, " &
                    " cod.COD_Ticket_no 客戶訂單號, " &
                    " itm.ITM_SKU_NO AS 貨品編號, " &
                    " cod.COD_ITM_DESC AS 貨品名稱, " &
                    " co.CUS_CODE AS 客戶編號, " &
                    " co.CUS_NAME AS 客戶名稱, " &
                    " cod.COD_Carton_no AS 銷售員, " &
                    " cod.COD_QTY AS 訂購數量, " &
                    " cod.COD_UOM AS 單位, " &
                    " (SELECT ISNULL(SUM(ISNULL(itm_loc_bal.ILOC_BAL_QTY, 0)) - (SELECT ISNULL(SUM(rsvd.QTY),0) FROM WMS_WAVEPICK_RSVD rsvd WHERE " &
                    " rsvd.wh_code = cod.COD_WH_CODE AND rsvd.ITEM_CODE = cod.COD_ITM_CODE AND rsvd.STORER_CODE = cod.STORER_CODE),0) FROM WMS_ITEM_LOC_BAL itm_loc_bal " &
                    " WHERE itm_loc_bal.ILOC_BAL_QTY > = 0 AND itm_loc_bal.itm_code = cod.COD_ITM_CODE AND itm_loc_bal.ILOC_WH = cod.COD_WH_CODE AND itm_loc_bal.Storer_code = cod.STORER_CODE) AS 分配餘數, " &
                    " CONVERT(VARCHAR(10),(co.CO_DATE),120) 計劃發貨日期,	ISNULL((SELECT TOP(1) t1.MIN_PROD_LIFE FROM (SELECT ISNULL(MIN_PROD_LIFE, 0) MIN_PROD_LIFE FROM WMS_CUSTOMER_RULE " &
                    " WHERE ITEM_CODE = cod.COD_ITM_CODE AND CUST_CODE = co.CUS_CODE UNION SELECT ISNULL(MIN_PROD_LIFE, 0) MIN_PROD_LIFE FROM WMS_CUSTOMER " &
                    " WHERE STORER_CODE = cod.STORER_CODE AND CUS_CODE = co.CUS_CODE) t1 ORDER BY t1.MIN_PROD_LIFE DESC),0) AS 客戶後熟天數要求, " &
                    " ISNULL((SELECT TOP(1) t1.MIN_SELF_LIFE FROM (SELECT ISNULL(MIN_SELF_LIFE, 0) MIN_SELF_LIFE FROM WMS_CUSTOMER_RULE WHERE ITEM_CODE = cod.COD_ITM_CODE " &
                    " AND CUST_CODE = co.CUS_CODE UNION SELECT ISNULL(MIN_SELF_LIFE, 0) MIN_SELF_LIFE FROM WMS_CUSTOMER WHERE STORER_CODE = cod.STORER_CODE AND CUS_CODE = co.CUS_CODE) t1 " &
                    " ORDER BY t1.MIN_SELF_LIFE ASC),0) AS 客戶最少保質天數要求,cod.COD_PALLET_NO AS 客戶最後接受批次,ISNULL((SELECT TOP(1) t1.MAX_LOTS FROM (SELECT ISNULL(MAX_LOTS, 0) MAX_LOTS " &
                    " FROM WMS_CUSTOMER_RULE WHERE ITEM_CODE = cod.COD_ITM_CODE AND CUST_CODE = co.CUS_CODE UNION SELECT ISNULL(MAX_LOTS, 0) MIN_PROD_LIFE FROM WMS_CUSTOMER " &
                    " WHERE STORER_CODE = cod.STORER_CODE AND CUS_CODE = co.CUS_CODE) t1 ORDER BY t1.MAX_LOTS DESC),0) AS 客戶可接受最多批次,co.ROUTE_ID AS 車線,cod.STORER_CODE AS WMS倉庫編號, " &
                    " cod.COD_ITM_CODE AS 貨品內部編號,itm.ITM_SHELF_LIFE AS 保質天數,cod.CO_CODE AS WMS訂單編號,CO_STATUS AS WMS訂單編號狀態,dod.DO_CODE AS WMS送貨單編號,co.CO_SENDER AS EBS用家, " &
                    " co.CO_EDI_SIR_NO AS 批號,cod.ALLOCATED AS ALLOCATED,CO_FTRACK_NO as TYPE,co.EBS_UPDATE_STATUS AS 已上傳EBS FROM WMS_CUST_ORDER_D cod INNER JOIN WMS_CUST_ORDER co ON co.CO_CODE = cod.CO_CODE " &
                    " LEFT JOIN WMS_DELV_ORDER_D dod ON cod.COD_ITM_CODE = dod.DOD_ITM_CODE AND cod.CO_CODE IN (SELECT items FROM dbo.split(dod.DOD_CO_CODE, ',')) LEFT JOIN WMS_ITEM itm ON cod.IMP_CODE = itm.IMP_CODE " &
                    " AND cod.STORER_CODE = itm.STORER_CODE AND cod.COD_ITM_CODE = itm.ITM_CODE AND cod.COD_PACK_KEY = itm.PACK_KEY WHERE CO_STATUS IN ('PICKED', 'PARTIAL') OR (CO_STATUS = 'NEW' AND (cod.ALLOCATED = 'N' " &
                    " Or cod.ALLOCATED Is NULL)) GROUP BY itm.ITM_SKU_NO,itm.ITM_SHELF_LIFE,cod.COD_ITM_DESC,cod.COD_Ticket_no,co.CUS_NAME,cod.COD_UOM,co.ROUTE_ID,co.CO_DATE,co.CUS_CODE,cod.CO_CODE,co.CO_SENDER_REGION, " &
                    " cod.COD_WH_CODE,cod.COD_QTY,cod.ALLOCATED,co.CUS_CODE,cod.COD_Carton_no,cod.COD_Plant,cod.COD_ITM_CODE,cod.COD_PALLET_NO,dod.DO_CODE,dod.DOD_ITM_CODE,dod.DOD_LAST_LOT,co.CO_INV_NO,COD.COD_REF_NO, " &
                    " cod.STORER_CODE,cod.ALLOCATED,CO_SENDER,CO_EDI_SIR_NO,CO_STATUS,CO_FTRACK_NO,CO.EBS_UPDATE_STATUS) AS t WHERE t.已上傳EBS <> 'Y') AS t2 WHERE t2.WMS倉庫編號 = '" & storer_code & "' " &
                    " GROUP BY t2.訂單編號,t2.貨品編號,t2.客戶訂單號,t2.放單檢查,t2.銷售員,t2.計劃發貨日期,t2.單位,t2.車線,t2.客戶編號,t2.訂單行號,t2.倉庫編號,t2.貨品名稱,t2.保質天數,t2.客戶名稱,t2.WMS訂單編號, " &
                    " t2.暫停信貸,t2.客戶後熟天數要求,	t2.客戶最少保質天數要求,t2.子庫存,	t2.訂購數量,	t2.分貨失敗,	t2.貨品內部編號,	t2.WMS送貨單編號,t2.客戶最後接受批次,t2.客戶可接受最多批次,	t2.發票編號,	t2.分配餘數, " &
                    " t2.EBS用家,t2.Allocated,t2.type,t2.批號,t2.WMS倉庫編號,	t2.WMS訂單編號狀態,t2.已上傳EBS ORDER BY t2.WMS訂單編號 "



        nDataSource = gDB.getDataTable(sqlString)
        GridTableData.DataSource = nDataSource
        GridTableData.DataBind()

        Try
            Response.Clear()
            Response.Buffer = True
            Response.AddHeader("content-disposition", "attachment;filename=SalesAdminStorageReportRaw.xls")
            Response.Charset = ""
            Response.ContentType = "application/vnd.ms-excel"
            GridTableData.Visible = True
            Using sw As New System.IO.StringWriter()
                Dim hw As New HtmlTextWriter(sw)
                GridTableData.RenderControl(hw)
                Response.Output.Write(sw.ToString())
                Response.Flush()
                Response.[End]()
            End Using
            nDataSource.Clear()
            GridTableData.DataSource = nDataSource
            GridTableData.DataSource = Nothing
            GridTableData.Visible = False
        Catch ex As Exception
            nDataSource.Clear()
            GridTableData.DataSource = nDataSource
            GridTableData.DataSource = Nothing
            GridTableData.Visible = False
            Response.Write(ex.Message)
            UiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
        End Try

    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(control As Control)
        ' Verifies that the control is rendered  
    End Sub

End Class
