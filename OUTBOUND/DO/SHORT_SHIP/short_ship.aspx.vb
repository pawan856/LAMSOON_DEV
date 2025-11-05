Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Globalization
Imports Microsoft.Reporting.WebForms.Internal.Soap.ReportingServices2005.Execution
Imports Microsoft.Reporting.WebForms
Imports ExcelLibrary.SpreadSheet

Partial Class OUTBOUND_DO_SHORT_SHIP_short_ship
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private wmsFun As New WMSFunc
    Private UiFun As New UIfunc


    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        ar.hideForm(Me)

        'Dim sqlString As String
        'Dim nDataSource As DataTable
        Dim do_code As String = ""
        Dim storer_code As String = ""
        Dim imp_code As String = ""
        Dim tempSQL As String = ""
        Dim tempSQL2 As String = ""

        'Dim Paras(7) As Microsoft.Reporting.WebForms.ReportParameter

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")
        ViewState("STOERE_CODE") = ""
        ViewState("DO_CODE") = ""

        ViewState("DO_CODE") = Server.UrlDecode(Request("DO_CODE"))
        ViewState("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))

        If Not IsPostBack Then

            BindData()

            'storer_code = ViewState("STORER_CODE")
            'do_code = ViewState("DO_CODE")

            'sqlString = "SELECT WMS_DELV_ORDER_D.DOD_SEQ, WMS_ITEM.ITM_SKU_NO,WMS_DELV_ORDER_D.DOD_EXPIRY_DATE, WMS_DELV_ORDER_D.DOD_QTY, WMS_DELV_ORDER_D.DOD_ITM_DESC, WMS_DELV_ORDER_D.DOD_SS_QTY,WMS_CUST_ORDER.CUS_NAME" &
            '        " FROM WMS_DELV_ORDER_D INNER JOIN " &
            '                " WMS_ITEM On WMS_DELV_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE And WMS_DELV_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE And " &
            '                " WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE And WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY INNER JOIN " &
            '                " WMS_CUST_ORDER On WMS_DELV_ORDER_D.IMP_CODE=WMS_CUST_ORDER.IMP_CODE And " &
            '                " WMS_DELV_ORDER_D.STORER_CODE= WMS_CUST_ORDER.STORER_CODE And " &
            '                " WMS_DELV_ORDER_D.DOD_CO_CODE=WMS_CUST_ORDER.CO_CODE" &
            '                " Where WMS_DELV_ORDER_D.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "' AND WMS_DELV_ORDER_D.STORER_CODE='" & gU.dbEncode(storer_code) & "' " &
            '                " And WMS_DELV_ORDER_D.DO_CODE ='" & gU.dbEncode(do_code) & "' order by WMS_DELV_ORDER_D.DOD_ITM_DESC "
            'nDataSource = gDB.getDataTable(sqlString)

            'If nDataSource.Rows.Count > 0 Then
            '    Dim tempPath As String = gU.getConfig("SYSP_TEMP_DIR") & "\DO\" & do_code

            '    'Dim seqNo As String = ""
            '    Dim sku_no As String = ""
            '    Dim lot_no As String = ""
            '    Dim qty As Integer = 0
            '    Dim ss_qty As Integer = 0
            '    Dim itm_desc As String = ""
            '    Dim cus_name As String = ""
            '    Dim route As String = ""
            '    Dim dd As String = ""


            '    Dim dtl As DataTable

            '    sqlString = "SELECT WMS_DELV_ORDER.ROUTE_ID,WMS_DELV_ORDER.DO_DATE from WMS_DELV_ORDER"
            '    dtl = gDB.getDataTable(sqlString)


            '    route = dtl.Rows(0)("ROUTE_ID").ToString
            '    dd = dtl.Rows(0)("DO_DATE").ToString

            '    Dim ddate As DateTime = DateTime.Parse(dd, CultureInfo.InvariantCulture)
            '    Dim reformatted As String = ddate.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture)
            '    dd = reformatted

            '    'seqNo = nDataSource.Rows(0)("DOD_SEQ").ToString
            '    sku_no = nDataSource.Rows(0)("ITM_SKU_NO").ToString
            '    qty = nDataSource.Rows(0)("DOD_QTY").ToString
            '    itm_desc = nDataSource.Rows(0)("DOD_ITM_DESC").ToString
            '    ss_qty = nDataSource.Rows(0)("DOD_QTY").ToString
            '    cus_name = nDataSource.Rows(0)("CUS_NAME").ToString

            '    lot_no = nDataSource.Rows(0)("DOD_EXPIRY_DATE").ToString
            '    If lot_no = "" Then
            '        lot_no = ""
            '    Else
            '        Dim ddatetime As DateTime = DateTime.Parse(lot_no, CultureInfo.InvariantCulture)
            '        Dim reformattedDate As String = ddatetime.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
            '        lot_no = reformattedDate
            '    End If

            '    Paras(0) = New Microsoft.Reporting.WebForms.ReportParameter("ITM_SKU_NO", sku_no)
            '    Paras(1) = New Microsoft.Reporting.WebForms.ReportParameter("DOD_EXPIRY_DATE", lot_no)
            '    Paras(2) = New Microsoft.Reporting.WebForms.ReportParameter("DOD_QTY", qty)
            '    Paras(3) = New Microsoft.Reporting.WebForms.ReportParameter("DOD_ITM_DESC", itm_desc)
            '    Paras(4) = New Microsoft.Reporting.WebForms.ReportParameter("DOD_SS_QTY", ss_qty)
            '    Paras(5) = New Microsoft.Reporting.WebForms.ReportParameter("CUS_NAME", cus_name)
            '    Paras(6) = New Microsoft.Reporting.WebForms.ReportParameter("ROUTE_ID", route)
            '    Paras(7) = New Microsoft.Reporting.WebForms.ReportParameter("DO_DATE", dd)

            '    'reportSource(nDataSource, Paras)
            'Else

            '    Response.Write("No Check List record has been Found.")
            '    Response.End()
            'End If
        End If
    End Sub

    Private Sub BindData()
        Dim sqlString As String
        Dim nDataSource As DataTable
        Dim do_codes As String = ""
        Dim storer_code As String = ""

        storer_code = ViewState("STORER_CODE")
        do_codes = "0000000"

        Dim do_code_list As New List(Of String)
        do_code_list = CType(Session("do_code_list"), List(Of String))
        For Each item As String In do_code_list
            If do_codes = "" Then
                do_codes = "'" + item + "'"
            Else
                do_codes = do_codes + ",'" + item + "'"
            End If
        Next

        sqlString = "Select DO_DATE,Convert(nvarchar(20),dHdr.DO_DATE,105)DO_Display_DATE,st.STO_NAME_CH,st.STO_ADDR1,dhdr.ROUTE_ID,WMS_DELV_ORDER_D.DOD_CO_CODE,WMS_DELV_ORDER_D.DOD_UOM,WMS_DO_PICKLIST_D.PLD_WH,WMS_DO_PICKLIST_D.PLD_ITEM_NO,WMS_DO_PICKLIST_D.PLD_LOC,WMS_DO_PICKLIST_D.PLD_SEQ,WMS_DO_PICKLIST_D.PLD_BATCH_NO,WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_ITEM.ITM_SKU_NO,WMS_DO_PICKLIST_D.PLD_EXPIRY_DATE, WMS_DO_PICKLIST_D.PLD_ITEM_QTY, WMS_DELV_ORDER_D.DOD_ITM_DESC, WMS_DO_PICKLIST_D.PLD_SS_QTY, WMS_DELV_ORDER_D.DOD_CUS_NAME,(CASE WHEN Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_MIN_SHELF_LIFE) = '0' or Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_MIN_SHELF_LIFE)= '' THEN '' ELSE (N'最低保質期 : '+Convert(nvarchar(20), WMS_DELV_ORDER_D.DOD_MIN_SHELF_LIFE)+',') END + CASE WHEN Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_LAST_LOT) = '0' or Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_LAST_LOT)= '' THEN '' ELSE (N'最大手數 : '+Convert(nvarchar(20), WMS_DELV_ORDER_D.DOD_LAST_LOT)+',') END + CASE WHEN Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_MIN_PROD_DATE) = '0' or Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_MIN_PROD_DATE)= '' THEN '' ELSE (N'最低生產日期 : '+Convert(nvarchar(20), WMS_DELV_ORDER_D.DOD_MIN_PROD_DATE)+',') END + CASE WHEN Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_PALLET_NO) = '0' or Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_PALLET_NO)= '' THEN '' ELSE (N'不超過很多 : '+Convert(nvarchar(20), WMS_DELV_ORDER_D.DOD_PALLET_NO)) END) AS DOD_RULES " &
        " From WMS_DELV_ORDER_D INNER Join WMS_DELV_ORDER dHdr On dHdr.DO_CODE= WMS_DELV_ORDER_D.DO_CODE Inner Join WMS_STORER st ON dhdr.STORER_CODE=st.STORER_CODE And st.IMP_CODE=dHdr.IMP_CODE INNER Join " &
        " WMS_ITEM On WMS_DELV_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE And WMS_DELV_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE And " &
        " WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE And WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY INNER Join WMS_DO_PICKLIST_D On WMS_DELV_ORDER_D.DO_CODE=WMS_DO_PICKLIST_D.DO_CODE and WMS_DELV_ORDER_D.DOD_SEQ=WMS_DO_PICKLIST_D.DOD_SEQ AND WMS_DELV_ORDER_D.DOD_ITM_CODE=WMS_DO_PICKLIST_D.PLD_ITEM_NO Where WMS_DO_PICKLIST_D.DO_CODE In (" & do_codes & ") order by WMS_DO_PICKLIST_D.PLD_SS_QTY desc "

        nDataSource = gDB.getDataTable(sqlString)

        Dim delivery_Dt_lst As New List(Of String)
        Dim Route_lst As New List(Of String)
        If nDataSource.Rows.Count > 0 Then
            For Each row As DataRow In nDataSource.Rows
                delivery_Dt_lst.Add(row("DO_DATE"))
                Route_lst.Add(row("ROUTE_ID"))
            Next

            Dim ShortShipReportModels As New List(Of ShortShipReportModel)
            For index As Integer = 0 To nDataSource.Rows.Count - 1
                Dim ShortShip As New ShortShipReportModel
                ShortShip.STO_NAME_CH = nDataSource.Rows(index)("STO_NAME_CH").ToString()
                ShortShip.STO_ADDR1 = nDataSource.Rows(index)("STO_ADDR1").ToString()
                ShortShip.DO_DATE = Convert.ToDateTime(nDataSource.Rows(index)("DO_DATE"))
                ShortShip.DO_Display_DATE = nDataSource.Rows(index)("DO_Display_DATE").ToString.Trim
                ShortShip.ROUTE_ID = nDataSource.Rows(index)("ROUTE_ID").ToString()
                ShortShip.PLD_EXPIRY_DATE = nDataSource.Rows(index)("PLD_EXPIRY_DATE").ToString()
                ShortShip.PLD_WH = nDataSource.Rows(index)("PLD_WH").ToString()
                ShortShip.DOD_UOM = nDataSource.Rows(index)("DOD_UOM").ToString()
                ShortShip.PLD_LOC = nDataSource.Rows(index)("PLD_LOC").ToString()
                ShortShip.ITM_SKU_NO = nDataSource.Rows(index)("ITM_SKU_NO").ToString()
                ShortShip.PLD_ITEM_QTY = nDataSource.Rows(index)("PLD_ITEM_QTY").ToString()
                ShortShip.PLD_ITEM_NO = nDataSource.Rows(index)("PLD_ITEM_NO").ToString()
                ShortShip.DOD_ITM_DESC = nDataSource.Rows(index)("DOD_ITM_DESC").ToString()
                ShortShip.PLD_SS_QTY = nDataSource.Rows(index)("PLD_SS_QTY").ToString()
                ShortShip.DOD_CUS_NAME = nDataSource.Rows(index)("DOD_CUS_NAME").ToString()
                ShortShip.DOD_RULES = nDataSource.Rows(index)("DOD_RULES").ToString()
                ShortShip.PLD_SEQ = nDataSource.Rows(index)("PLD_SEQ").ToString()
                ShortShip.PLD_BATCH_NO = nDataSource.Rows(index)("PLD_BATCH_NO").ToString()
                ShortShip.DOD_CO_CODE = nDataSource.Rows(index)("DOD_CO_CODE").ToString()
                ShortShipReportModels.Add(ShortShip)
            Next

            ShortShipReportModels.ForEach(Function(obj)
                                              Dim query = "'" + String.Join("','", obj.DOD_CO_CODE.Split(",")) + "'"
                                              query = "select '('+CONCAT(CO.CO_INV_NO,',', CO.CUS_NAME,',',cod.COD_QTY)+')' as CUST_INFO" &
                                                       " from WMS_CUST_ORDER_D cod inner join WMS_CUST_ORDER co ON co.co_code=cod.co_code " &
                                                       " where co.co_code in (" + query + ") and cod.COD_ITM_CODE='" + obj.PLD_ITEM_NO + "' "
                                              nDataSource = gDB.getDataTable(query)
                                              Dim customerInfo As New List(Of String)
                                              If nDataSource.Rows.Count > 0 Then
                                                  For Each row As DataRow In nDataSource.Rows
                                                      customerInfo.Add(row("CUST_INFO"))
                                                  Next
                                                  obj.CUST_INFO = String.Join("<br/>", customerInfo)
                                              End If
                                              Return True
                                          End Function)


            Dim rootShortShips As List(Of RootShortShipReportModel) = ShortShipReportModels.GroupBy(Function(person) New With {Key person.ROUTE_ID, Key person.DO_DATE}).
                Select(Function(grp) New RootShortShipReportModel With {
                .DO_DATE = grp.Key.DO_DATE.ToString,
                .DO_Display_DATE = grp.FirstOrDefault().DO_Display_DATE,
                .ROUTE_ID = grp.Key.ROUTE_ID,
                .ShortShips = grp.ToList(),
                 .STO_ADDR1 = grp.FirstOrDefault().STO_ADDR1,
                .STO_NAME_CH = grp.FirstOrDefault().STO_NAME_CH
              }).OrderBy(Function(g) g.DO_DATE).ThenBy(Function(p) p.ROUTE_ID).ToList()

            rptCustomers.DataSource = rootShortShips
            rptCustomers.DataBind()
        End If

    End Sub

    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Response.Clear()
        Response.Buffer = True
        Response.AddHeader("content-disposition", "attachment;filename=ShortShipListReport.xls")
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
        Dim sqlString As String
        Dim nDataSource As DataSet
        Dim do_codes As String = ""
        Dim storer_code As String = ""

        Try
            storer_code = ViewState("STORER_CODE")
            do_codes = "000000"
            Dim do_code_list As New List(Of String)
            do_code_list = CType(Session("do_code_list"), List(Of String))
            For Each item As String In do_code_list
                If do_codes = "" Then
                    do_codes = "'" + item + "'"
                Else
                    do_codes = do_codes + ",'" + item + "'"
                End If

            Next

            sqlString = "Select DO_DATE,Convert(nvarchar(20),dHdr.DO_DATE,105)DO_Display_DATE,st.STO_NAME_CH,st.STO_ADDR1,dhdr.ROUTE_ID,WMS_DELV_ORDER_D.DOD_CO_CODE,WMS_DELV_ORDER_D.DOD_UOM,WMS_DO_PICKLIST_D.PLD_WH,WMS_DO_PICKLIST_D.PLD_ITEM_NO,WMS_DO_PICKLIST_D.PLD_LOC,WMS_DO_PICKLIST_D.PLD_SEQ,WMS_DO_PICKLIST_D.PLD_BATCH_NO,WMS_DO_PICKLIST_D.PLD_ITEM_NO, WMS_ITEM.ITM_SKU_NO,WMS_DO_PICKLIST_D.PLD_EXPIRY_DATE, WMS_DO_PICKLIST_D.PLD_ITEM_QTY, WMS_DELV_ORDER_D.DOD_ITM_DESC, WMS_DO_PICKLIST_D.PLD_SS_QTY, WMS_DELV_ORDER_D.DOD_CUS_NAME,(CASE WHEN Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_MIN_SHELF_LIFE) = '0' or Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_MIN_SHELF_LIFE)= '' THEN '' ELSE (N'最低保質期 : '+Convert(nvarchar(20), WMS_DELV_ORDER_D.DOD_MIN_SHELF_LIFE)+',') END + CASE WHEN Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_LAST_LOT) = '0' or Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_LAST_LOT)= '' THEN '' ELSE (N'最大手數 : '+Convert(nvarchar(20), WMS_DELV_ORDER_D.DOD_LAST_LOT)+',') END + CASE WHEN Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_MIN_PROD_DATE) = '0' or Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_MIN_PROD_DATE)= '' THEN '' ELSE (N'最低生產日期 : '+Convert(nvarchar(20), WMS_DELV_ORDER_D.DOD_MIN_PROD_DATE)+',') END + CASE WHEN Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_PALLET_NO) = '0' or Convert(nvarchar(20),WMS_DELV_ORDER_D.DOD_PALLET_NO)= '' THEN '' ELSE (N'不超過很多 : '+Convert(nvarchar(20), WMS_DELV_ORDER_D.DOD_PALLET_NO)) END) AS DOD_RULES " &
        " From WMS_DELV_ORDER_D INNER Join WMS_DELV_ORDER dHdr On dHdr.DO_CODE= WMS_DELV_ORDER_D.DO_CODE Inner Join WMS_STORER st ON dhdr.STORER_CODE=st.STORER_CODE And st.IMP_CODE=dHdr.IMP_CODE INNER Join " &
        " WMS_ITEM On WMS_DELV_ORDER_D.IMP_CODE = WMS_ITEM.IMP_CODE And WMS_DELV_ORDER_D.STORER_CODE = WMS_ITEM.STORER_CODE And " &
        " WMS_DELV_ORDER_D.DOD_ITM_CODE = WMS_ITEM.ITM_CODE And WMS_DELV_ORDER_D.DOD_PACK_KEY = WMS_ITEM.PACK_KEY INNER Join WMS_DO_PICKLIST_D On WMS_DELV_ORDER_D.DO_CODE=WMS_DO_PICKLIST_D.DO_CODE and WMS_DELV_ORDER_D.DOD_SEQ=WMS_DO_PICKLIST_D.DOD_SEQ AND WMS_DELV_ORDER_D.DOD_ITM_CODE=WMS_DO_PICKLIST_D.PLD_ITEM_NO Where WMS_DO_PICKLIST_D.DO_CODE In (" & do_codes & ") order by WMS_DO_PICKLIST_D.PLD_SS_QTY desc "

            nDataSource = gDB.getDataSet(sqlString)

            WriteXLSFile(nDataSource)

        Catch ex As Exception
            Response.Write(ex.Message)
            UiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
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

                    worksheet = New Worksheet("SHORT SHIP REPORT")

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
            Response.AddHeader("content-disposition", "attachment;filename=ShortShipReport.xls")
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
