Imports System.IO
Imports System.Net.Mail
Imports ExcelLibrary
Imports ExcelLibrary.SpreadSheet
Imports System.Data
Partial Class SALES_ITM_BALReport
    Inherits System.Web.UI.Page
    Private moduleAction As String = ""
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim storercode As String = STORER_CODE.SelectedValue
        Session("str_code") = storercode

        If Not IsPostBack Then
            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
        End If
    End Sub

    Private Function validateImp() As Boolean
        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty, please select Storer!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If
        Return True
    End Function

    Protected Sub btnSales_Report_click(sender As Object, e As EventArgs) Handles btnSales_Report.Click
        If validateImp() Then
            Dim url As String = "~/../REPORT/CUST_SUPPORT/cust_support.aspx"
            Dim s As String = "window.open('" & url + "', 'CUST_SUPPORT', 'width=800,height=800,left=100,top=100,resizable=yes');"
            ClientScript.RegisterStartupScript(Me.GetType(), "script", s, True)
        End If
        'STORER_CODE.SelectedValue = ""
        'Response.Redirect("REPORT/CUST_SUPPORT/cust_support.aspx")
    End Sub

    Public Sub btnITM_BAL_Report_click(sender As Object, e As EventArgs) Handles btnITM_BAL_Report.Click
        If validateImp() Then
            Dim url As String = "~/../REPORT/ITM_BAL_RSVD/itm_bal_rsvd.aspx"
            Dim s As String = "window.open('" & url + "', 'ITM_BAL_RSVD', 'width=800,height=800,left=100,top=100,resizable=yes');"
            ClientScript.RegisterStartupScript(Me.GetType(), "script", s, True)
        End If
        'STORER_CODE.SelectedValue = ""
        'Response.Redirect("REPORT/ITM_BAL_RSVD/itm_bal_rsvd.aspx")
    End Sub

    Public Sub btnSendEmail_Click(sender As Object, e As EventArgs) Handles btnSendEmail.Click
        If validateImp() Then
            Dim sqlString As String
            Dim nDataSource As DataSet
            '            sqlString = " SELECT t2.* FROM (SELECT t.* FROM (SELECT (SELECT TOP(1) IO_CODE FROM EBS_WMS_COMPANY_MASTER WHERE IO_ID = cod.STORER_CODE) 倉庫編號, " &
            '                    " (CASE WHEN co.CO_SENDER_REGION = 'FALSE' AND ((cod.ALLOCATED IS NULL AND dod.DO_CODE IS NOT NULL AND dod.DOD_ITM_CODE = cod.COD_ITM_CODE)) THEN '' ELSE 'Y' END) AS 放單檢查, " &
            '                    " (CASE WHEN co.CO_SENDER_REGION = 'FALSE' THEN 'N' ELSE 'Y' END) AS 暫停信貸,(CASE WHEN CO_FTRACK_NO='ACTUAL' and  (dod.DO_CODE is NULL or cod.ALLOCATED = 'N') THEN 'A-NP' " &
            '                    " WHEN cod.ALLOCATED Is NULL And dod.DO_CODE Is NULL THEN 'Y' WHEN cod.ALLOCATED Is NULL And dod.DO_CODE Is Not NULL And dod.DOD_ITM_CODE = cod.COD_ITM_CODE THEN 'N' " &
            '                    " WHEN cod.ALLOCATED = 'N' And dod.DO_CODE Is Not NULL And dod.DOD_ITM_CODE = cod.COD_ITM_CODE THEN 'N' WHEN cod.ALLOCATED = 'N' AND dod.DO_CODE IS NOT NULL " &
            '                    " And dod.DOD_ITM_CODE ! = cod.COD_ITM_CODE THEN 'Y' WHEN cod.ALLOCATED = 'N' AND dod.DO_CODE IS NULL THEN 'Y' ELSE 'N' END) AS 分貨失敗,cod.COD_WH_CODE AS 子庫存, " &
            '                    " (SELECT TOP(1) DELIVERY_ID FROM EBS_WMS_TRANS_ITX_SO_HEADER WHERE SO_HEADER_ID = cod.COD_REF_NO ORDER BY TRANSACTION_ID DESC) AS 發票編號,	co.CO_INV_NO AS 訂單編號, " &
            '                    " cod.COD_Plant AS 訂單行號, " &
            '                    " cod.COD_Ticket_no 客戶訂單號, " &
            '                    " itm.ITM_SKU_NO AS 貨品編號, " &
            '                    " cod.COD_ITM_DESC AS 貨品名稱, " &
            '                    " co.CUS_CODE AS 客戶編號, " &
            '                    " co.CUS_NAME AS 客戶名稱, " &
            '                    " cod.COD_Carton_no AS 銷售員, " &
            '                    " cod.COD_QTY AS 訂購數量, " &
            '                    " cod.COD_UOM AS 單位, " &
            '                    " (SELECT ISNULL(SUM(ISNULL(itm_loc_bal.ILOC_BAL_QTY, 0)) - (SELECT ISNULL(SUM(rsvd.QTY),0) FROM WMS_WAVEPICK_RSVD rsvd WHERE " &
            '                    " rsvd.wh_code = cod.COD_WH_CODE AND rsvd.ITEM_CODE = cod.COD_ITM_CODE AND rsvd.STORER_CODE = cod.STORER_CODE),0) FROM WMS_ITEM_LOC_BAL itm_loc_bal " &
            '                    " WHERE itm_loc_bal.ILOC_BAL_QTY > = 0 AND itm_loc_bal.itm_code = cod.COD_ITM_CODE AND itm_loc_bal.ILOC_WH = cod.COD_WH_CODE AND itm_loc_bal.Storer_code = cod.STORER_CODE) AS 分配餘數, " &
            '                    " CONVERT(VARCHAR(10),(co.CO_DATE),120) 計劃發貨日期,	ISNULL((SELECT TOP(1) t1.MIN_PROD_LIFE FROM (SELECT ISNULL(MIN_PROD_LIFE, 0) MIN_PROD_LIFE FROM WMS_CUSTOMER_RULE " &
            '                    " WHERE ITEM_CODE = cod.COD_ITM_CODE AND CUST_CODE = co.CUS_CODE UNION SELECT ISNULL(MIN_PROD_LIFE, 0) MIN_PROD_LIFE FROM WMS_CUSTOMER " &
            '                    " WHERE STORER_CODE = cod.STORER_CODE AND CUS_CODE = co.CUS_CODE) t1 ORDER BY t1.MIN_PROD_LIFE DESC),0) AS 客戶後熟天數要求, " &
            '                    " ISNULL((SELECT TOP(1) t1.MIN_SELF_LIFE FROM (SELECT ISNULL(MIN_SELF_LIFE, 0) MIN_SELF_LIFE FROM WMS_CUSTOMER_RULE WHERE ITEM_CODE = cod.COD_ITM_CODE " &
            '                    " AND CUST_CODE = co.CUS_CODE UNION SELECT ISNULL(MIN_SELF_LIFE, 0) MIN_SELF_LIFE FROM WMS_CUSTOMER WHERE STORER_CODE = cod.STORER_CODE AND CUS_CODE = co.CUS_CODE) t1 " &
            '                    " ORDER BY t1.MIN_SELF_LIFE ASC),0) AS 客戶最少保質天數要求,cod.COD_PALLET_NO AS 客戶最後接受批次,ISNULL((SELECT TOP(1) t1.MAX_LOTS FROM (SELECT ISNULL(MAX_LOTS, 0) MAX_LOTS " &
            '                    " FROM WMS_CUSTOMER_RULE WHERE ITEM_CODE = cod.COD_ITM_CODE AND CUST_CODE = co.CUS_CODE UNION SELECT ISNULL(MAX_LOTS, 0) MIN_PROD_LIFE FROM WMS_CUSTOMER " &
            '                    " WHERE STORER_CODE = cod.STORER_CODE AND CUS_CODE = co.CUS_CODE) t1 ORDER BY t1.MAX_LOTS DESC),0) AS 客戶可接受最多批次,co.ROUTE_ID AS 車線,cod.STORER_CODE AS WMS倉庫編號, " &
            '                    " cod.COD_ITM_CODE AS 貨品內部編號,itm.ITM_SHELF_LIFE AS 保質天數,cod.CO_CODE AS WMS訂單編號,CO_STATUS AS WMS訂單編號狀態,dod.DO_CODE AS WMS送貨單編號,co.CO_SENDER AS EBS用家, " &
            '                    " co.CO_EDI_SIR_NO AS 批號,cod.ALLOCATED AS ALLOCATED,CO_FTRACK_NO as TYPE,co.EBS_UPDATE_STATUS AS 已上傳EBS FROM WMS_CUST_ORDER_D cod INNER JOIN WMS_CUST_ORDER co ON co.CO_CODE = cod.CO_CODE " &
            '                    " LEFT JOIN WMS_DELV_ORDER_D dod ON cod.COD_ITM_CODE = dod.DOD_ITM_CODE AND cod.CO_CODE IN (SELECT items FROM dbo.split(dod.DOD_CO_CODE, ',')) LEFT JOIN WMS_ITEM itm ON cod.IMP_CODE = itm.IMP_CODE " &
            '                    " AND cod.STORER_CODE = itm.STORER_CODE AND cod.COD_ITM_CODE = itm.ITM_CODE AND cod.COD_PACK_KEY = itm.PACK_KEY WHERE CO_STATUS IN ('PICKED', 'PARTIAL') OR (CO_STATUS = 'NEW' AND (cod.ALLOCATED = 'N' " &
            '                    " Or cod.ALLOCATED Is NULL)) GROUP BY itm.ITM_SKU_NO,itm.ITM_SHELF_LIFE,cod.COD_ITM_DESC,cod.COD_Ticket_no,co.CUS_NAME,cod.COD_UOM,co.ROUTE_ID,co.CO_DATE,co.CUS_CODE,cod.CO_CODE,co.CO_SENDER_REGION, " &
            '                    " cod.COD_WH_CODE,cod.COD_QTY,cod.ALLOCATED,co.CUS_CODE,cod.COD_Carton_no,cod.COD_Plant,cod.COD_ITM_CODE,cod.COD_PALLET_NO,dod.DO_CODE,dod.DOD_ITM_CODE,dod.DOD_LAST_LOT,co.CO_INV_NO,COD.COD_REF_NO, " &
            '                    " cod.STORER_CODE,cod.ALLOCATED,CO_SENDER,CO_EDI_SIR_NO,CO_STATUS,CO_FTRACK_NO,CO.EBS_UPDATE_STATUS) AS t WHERE t.已上傳EBS <> 'Y') AS t2 WHERE t2.WMS倉庫編號 = '" & STORER_CODE.SelectedValue & "' " &
            '                    " GROUP BY t2.訂單編號,t2.貨品編號,t2.客戶訂單號,t2.放單檢查,t2.銷售員,t2.計劃發貨日期,t2.單位,t2.車線,t2.客戶編號,t2.訂單行號,t2.倉庫編號,t2.貨品名稱,t2.保質天數,t2.客戶名稱,t2.WMS訂單編號, " &
            '                    " t2.暫停信貸,t2.客戶後熟天數要求,	t2.客戶最少保質天數要求,t2.子庫存,	t2.訂購數量,	t2.分貨失敗,	t2.貨品內部編號,	t2.WMS送貨單編號,t2.客戶最後接受批次,t2.客戶可接受最多批次,	t2.發票編號,	t2.分配餘數, " &
            '                    " t2.EBS用家,t2.Allocated,t2.type,t2.批號,t2.WMS倉庫編號,	t2.WMS訂單編號狀態,t2.已上傳EBS ORDER BY t2.WMS訂單編號 "


            '            sqlString = sqlString + " SELECT itm.ITM_SKU_NO AS 貨品編號, itm.ITM_DESC AS 貨品名稱, itm_loc_bal.ILOC_WH AS 子庫存, itm_loc_bal.ILOC_BATCH_NO AS 批次, " &
            '"  itm.ITM_UOM AS 單位, SUM(itm_loc_bal.ILOC_BAL_QTY) AS 庫存數量, ( SELECT ISNULL( SUM(rsvd.QTY), 0 " &
            '"  ) FROM WMS_WAVEPICK_RSVD rsvd WHERE rsvd.wh_code = itm_loc_bal.ILOC_WH AND rsvd.ITEM_CODE = itm.ITM_CODE AND  " &
            '"  itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE " &
            '"  ) AS 已分佩數量, SUM( ISNULL(itm_loc_bal.ILOC_BAL_QTY, 0) )- ( SELECT ISNULL( SUM(rsvd.QTY), 0 ) FROM WMS_WAVEPICK_RSVD rsvd  " &
            '"  WHERE rsvd.wh_code = itm_loc_bal.ILOC_WH AND rsvd.ITEM_CODE = itm.ITM_CODE AND itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO  " &
            '"  AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE ) AS 分配餘數, FORMAT( itm_loc_bal.ILOC_EXPIRY_DATE, 'dd MMM yy' ) AS 產品到期日, " &
            '"  CAST( DATEDIFF( DAY, GETDATE(), ILOC_EXPIRY_DATE ) AS DECIMAL ) AS 剩余日數, " &
            '"  itm.ITM_SHELF_LIFE 保質天數, itm_loc_bal.STORER_CODE AS STORER, itm.ITM_CODE AS 貨品內部編號 FROM " &
            '"  WMS_ITEM_LOC_BAL itm_loc_bal LEFT JOIN WMS_ITEM itm ON itm.ITM_CODE = itm_loc_bal.ITM_CODE AND itm.STORER_CODE = itm_loc_bal.STORER_CODE  " &
            '"  LEFT JOIN WMS_WAREHOUSE WH ON itm_loc_bal.ILOC_WH = wh.WH_CODE WHERE itm_loc_bal.ILOC_BAL_QTY > = 0 AND wh.WH_CONT_PER1 = 'Y'  And itm_loc_bal.STORER_CODE = '" & STORER_CODE.SelectedValue & "' " &
            '"  GROUP BY itm_loc_bal.STORER_CODE, itm_loc_bal.ILOC_WH, itm_loc_bal.ILOC_BATCH_NO, itm.ITM_CODE, itm.ITM_DESC,  " &
            '"  itm.ITM_SKU_NO, itm.ITM_SHELF_LIFE, itm.ITM_UOM, itm_loc_bal.ILOC_EXPIRY_DATE, wh.WH_CONT_PER1 HAVING wh.WH_CONT_PER1 = 'Y'  " &
            '"  AND( SUM(itm_loc_bal.ILOC_BAL_QTY) > 0 OR ( SELECT ISNULL( SUM(rsvd.QTY), 0 ) FROM WMS_WAVEPICK_RSVD rsvd  " &
            '"  WHERE rsvd.wh_code = itm_loc_bal.ILOC_WH AND rsvd.ITEM_CODE = itm.ITM_CODE AND itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO  " &
            '"  AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE ) > 0 )  ORDER BY 貨品編號,子庫存 "

            '            sqlString = sqlString + "  SELECT itm.ITM_SKU_NO 貨品編號, itm.ITM_DESC 貨品名稱, itm_loc_bal.ILOC_WH AS 子庫存, itm_loc_bal.ILOC_BATCH_NO 批次, itm.ITM_SHELF_LIFE  保質天數, SUM(itm_loc_bal.ILOC_BAL_QTY)  " &
            '"  AS 庫存數量, SUM( ISNULL(rsvd.QTY, 0) ) AS 分配數量, SUM(itm_loc_bal.ILOC_BAL_QTY)- SUM(  ISNULL(rsvd.QTY, 0) ) AS 分配後數量, itm.ITM_UOM AS '單位', itm_loc_bal.ILOC_EXPIRY_DATE  " &
            '"  AS 產品到期日, CAST(DATEDIFF(DAY,GETDATE(), ILOC_EXPIRY_DATE  ) AS DECIMAL ) AS 剩余日數 FROM WMS_ITEM_LOC_BAL itm_loc_bal LEFT JOIN WMS_WAVEPICK_RSVD rsvd  " &
            '"  ON rsvd.wh_code = itm_loc_bal.ILOC_WH AND rsvd.ITEM_CODE = itm_loc_bal.ITM_CODE AND itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE " &
            '"   INNER JOIN WMS_ITEM itm ON itm.ITM_CODE = itm_loc_bal.ITM_CODE AND itm.STORER_CODE = itm_loc_bal.STORER_CODE LEFT JOIN WMS_DELV_ORDER DO ON DO.DO_CODE = rsvd.DO_CODE " &
            '"  WHERE  itm_loc_bal.STORER_CODE = '" & STORER_CODE.SelectedValue & "' AND itm_loc_bal.ILOC_BAL_QTY > 0 AND itm_loc_bal.ITM_CODE IN ( SELECT COD_ITM_CODE  FROM   WMS_CUST_ORDER_D cod " &
            '"    WHERE   ALLOCATED = 'N'   AND    STORER_CODE = '" & STORER_CODE.SelectedValue & "'   AND COD_WH_CODE = itm_loc_bal.ILOC_WH  AND ( SELECT COUNT(1)    FROM WMS_DELV_ORDER_D " &
            '"      WHERE DOD_CO_CODE LIKE '%' + cod.CO_CODE + '%' AND DOD_ITM_CODE = cod.COD_ITM_CODE   )=0 ) GROUP BY itm_loc_bal.ILOC_WH, itm_loc_bal.ILOC_BATCH_NO, itm.ITM_CODE, " &
            '"  itm_sku_no, itm.ITM_DESC, rsvd.QTY, itm.ITM_SKU_NO, itm.ITM_SHELF_LIFE, itm.ITM_UOM, itm_loc_bal.ILOC_EXPIRY_DATE ORDER BY itm.ITM_SKU_NO DESC "

            '            sqlString = sqlString + " SELECT itm.ITM_SKU_NO 貨品編號, itm.ITM_DESC 貨品名稱, itm_loc_bal.ILOC_WH AS 子庫存, itm_loc_bal.ILOC_BATCH_NO 批次, " &
            '"  itm.ITM_SHELF_LIFE 保質天數, AVG(itm_loc_bal.ILOC_BAL_QTY) AS 庫存數量, SUM( ISNULL(rsvd.QTY, 0) ) AS 分配數量, " &
            '"  AVG(itm_loc_bal.ILOC_BAL_QTY)- SUM( ISNULL(rsvd.QTY, 0) ) AS 分配後數量, itm.ITM_UOM AS '單位', itm_loc_bal.ILOC_EXPIRY_DATE AS 產品到期日, " &
            '"  CAST( DATEDIFF( DAY, GETDATE(), ILOC_EXPIRY_DATE ) AS DECIMAL ) AS 剩余日數, itm.ITM_CODE FROM " &
            '"  WMS_ITEM_LOC_BAL itm_loc_bal LEFT JOIN WMS_WAVEPICK_RSVD rsvd ON rsvd.wh_code = itm_loc_bal.ILOC_WH " &
            '"  AND rsvd.ITEM_CODE = itm_loc_bal.ITM_CODE AND itm_loc_bal.ILOC_BATCH_NO = rsvd.LOT_NO AND itm_loc_bal.STORER_CODE = rsvd.STORER_CODE " &
            '"  INNER JOIN WMS_ITEM itm ON itm.ITM_CODE = itm_loc_bal.ITM_CODE AND itm.STORER_CODE = itm_loc_bal.STORER_CODE " &
            '"  LEFT JOIN WMS_DELV_ORDER DO ON DO.DO_CODE = rsvd.DO_CODE WHERE itm_loc_bal.STORER_CODE = '" & STORER_CODE.SelectedValue & "' " &
            '"  AND  itm_loc_bal.ILOC_BAL_QTY > 0 AND itm_loc_bal.ITM_CODE IN ( SELECT DISTINCT COD_ITM_CODE FROM WMS_CUST_ORDER_D cod " &
            '"  INNER JOIN WMS_CUST_ORDER co ON co.CO_CODE = cod.CO_CODE LEFT JOIN WMS_DELV_ORDER_D dod ON cod.COD_ITM_CODE = dod.DOD_ITM_CODE " &
            '"  AND cod.CO_CODE IN ( SELECT items FROM dbo.split(dod.DOD_CO_CODE, ',') ) WHERE EBS_UPDATE_STATUS = 'N' " &
            '"  AND COD_WH_CODE = itm_loc_bal.ILOC_WH AND ( cod.ALLOCATED = 'N' OR CO_STATUS = 'NEW' OR ( SELECT COUNT(1) FROM " &
            '"  WMS_DELV_ORDER_D WHERE DOD_CO_CODE LIKE '%' + cod.CO_CODE + '%' AND DOD_ITM_CODE = cod.COD_ITM_CODE )= 0 " &
            '"  or dod.do_code is NULL ) )GROUP BY itm_loc_bal.ILOC_WH, itm_loc_bal.ILOC_BATCH_NO, itm.ITM_CODE, ITM_SKU_NO, itm.ITM_DESC, " &
            '"  itm.ITM_SKU_NO, itm.ITM_SHELF_LIFE, itm.ITM_UOM, itm_loc_bal.ILOC_EXPIRY_DATE ORDER BY 貨品編號, 子庫存  "

            nDataSource = gDB.getDataSet("exec sp_StorageBalReport " & STORER_CODE.SelectedValue)
            'nDataSource = gDB.getDataSet(sqlString)

            Dim excelFolder As String
            excelFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempFiles\" & "_SalesAdminReportItems_" & System.DateTime.UtcNow.ToString("MM-dd-yyyy hh-mm-ss") & ".xls")
            Dim _Result As Boolean = WriteXLSFile(excelFolder, nDataSource)
            If _Result = True Then
                If (STORER_CODE.SelectedValue = "104") Then
                    SendMailToUser(excelFolder, "flourorderlot@lamsoon.com", "201")
                ElseIf (STORER_CODE.SelectedValue = "105") Then
                    SendMailToUser(excelFolder, "oilorderlot@lamsoon.com", "301")
                ElseIf (STORER_CODE.SelectedValue = "106") Then
                    SendMailToUser(excelFolder, "homecareorderlot@lamsoon.com", "401")
                ElseIf (STORER_CODE.SelectedValue = "362") Then
                    SendMailToUser(excelFolder, "flourorderlot@lamsoon.com", "204")
                End If
            End If

        End If
        'STORER_CODE.SelectedValue = ""
        'Response.Redirect("REPORT/ITM_BAL_RSVD/itm_bal_rsvd.aspx")
    End Sub
    Private Sub SendMailToUser(excelFolder As String, ToEmail As String, IO_CODE As String)
        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            'Smtp_Server.UseDefaultCredentials = True
            'Smtp_Server.Host = "hksmtp.lamsoon.com"
            'e_mail = New MailMessage()
            'e_mail.From = New MailAddress("hkwms@lamsoon.com")
            'e_mail.To.Add(ToEmail)
            'e_mail.Subject = "Sales Admin Shortage report from WMS for IO Code '" & IO_CODE & "' on " & System.DateTime.Now.ToString("dd-MMM-yyyy")
            'e_mail.IsBodyHtml = False
            'e_mail.Body = "Kindly find the report attachment."
            'e_mail.Attachments.Add(New Attachment(excelFolder))
            'Smtp_Server.Send(e_mail)

            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("augursmail@gmail.com", "Augurs@0009")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress("augursmail@gmail.com")
            e_mail.To.Add("madhvendra009@gmail.com")
            e_mail.To.Add("robertwong68@gmail.com")
            e_mail.Subject = "Sales Admin Shortage report from WMS for IO Code '" & IO_CODE & "' on " & System.DateTime.Now.ToString("dd-MMM-yyyy")
            e_mail.IsBodyHtml = False
            e_mail.Body = "Kindly find the report attachment."
            e_mail.Attachments.Add(New Attachment(excelFolder))
            Smtp_Server.Send(e_mail)
            lblMsg.Text = "Email Sent Successfully!"

        Catch error_t As Exception
            lblMsg.Text = error_t.ToString
        End Try
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
                        worksheet = New Worksheet("Sales Admin Shortage Report")
                    ElseIf (iSheetCount = 2) Then
                        worksheet = New Worksheet("Item Balance Report")
                    ElseIf (iSheetCount = 3) Then
                        worksheet = New Worksheet("Non allocated Item Balance Report")
                    Else
                        worksheet = New Worksheet("NEW NonAllocated report")
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
                                    worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.00")
                                Case GetType(Decimal)
                                    Decimal.TryParse(sTemp, dTemp)
                                    worksheet.Cells(iRow, iCol) = New Cell(dTemp, "#,##0.00")
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
