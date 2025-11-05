Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Linq
Partial Class OUTBOUND_DO_Wavepick
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private DDFORMAT As String = gU.getConfig("DDFORMATNo")
    Private lstError As New List(Of String)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        If Not IsPostBack Then
            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
        End If
    End Sub

    Function getLastLot(ByVal ITEM_CODE As String, ByVal STORER_CODE As String, ByVal CUS_CODE As String) As String

        Dim LAST_LOT As String = ""
        Dim SQLString As String = "Select iSnuLL(DEST_LOT_NUMBER,'')LAST_LOT from WMS_EBS_TRANS_ITX_SO so Where so.ITEM_ID = '" + ITEM_CODE + "' and so.ACCOUNT_NUMBER='" + CUS_CODE + "' and so.IO_ID = '" + STORER_CODE + "'"
        Dim dtItem As DataTable = gDB.getDataTable(SQLString)
        If dtItem IsNot Nothing AndAlso dtItem.Rows.Count > 0 Then
            LAST_LOT = dtItem.Rows(0)("LAST_LOT")
        End If
        Return LAST_LOT
    End Function

    'For CO details (ALLOCATED) update 
    Private Function updateCO_Details(ByVal ITEM_CODE As String, ByVal STORER_CODE As String, ByVal CUS_CODE As String, ByVal DO_CODE As String, ByVal ROUTE_ID As String, ByRef conn As SqlConnection, ByRef transaction As SqlTransaction) As String
        Dim UpdateSql As String
        Dim SQLString As String
        Dim cocodes As String = ""

        SQLString = "Select DOD_CO_CODE from WMS_DELV_ORDER_D where DO_CODE='" & DO_CODE & "'"
        Dim dtcoCodes = gDB.getDataTable(SQLString, conn, transaction)
        If (dtcoCodes.Rows.Count = 0) Then
            UpdateSql = "Update cod set cod.ALLOCATED='N' FROM WMS_CUST_ORDER_D cod Inner JOIN WMS_CUST_ORDER co ON co.CO_CODE=cod.CO_CODE where cod.COD_ITM_CODE = '" + ITEM_CODE + "' and cod.STORER_CODE = '" + STORER_CODE + "' and co.CUS_CODE='" + CUS_CODE + "' and co.CO_STATUS='NEW' and co.ROUTE_ID='" + ROUTE_ID + "'"
            gDB.amendData(UpdateSql, conn, transaction)
        End If
        For Each row As DataRow In dtcoCodes.Rows
            Dim co_Codes = row("DOD_CO_CODE").ToString().Split(",")
            For Each cocd As String In co_Codes
                If cocodes = "" Then
                    cocodes = "" + cocd + ""
                Else
                    cocodes = cocodes + "," + cocd + ""
                End If
            Next
        Next

        UpdateSql = "Update cod set cod.ALLOCATED='N' FROM WMS_CUST_ORDER_D cod Inner JOIN WMS_CUST_ORDER co ON co.CO_CODE=cod.CO_CODE where cod.CO_CODE in (select items from  dbo.split('" & cocodes & "',',')) and cod.COD_ITM_CODE = '" + ITEM_CODE + "' and cod.STORER_CODE = '" + STORER_CODE + "' and co.CUS_CODE='" + CUS_CODE + "' and co.ROUTE_ID='" + ROUTE_ID + "'"
        gDB.amendData(UpdateSql, conn, transaction)
        Return True
    End Function

    Private Function custRuleChk(ByRef conn As SqlConnection, ByRef transaction As SqlTransaction, ByVal RULE_COUNT As Int16, ByVal WH_CODE As String, ByVal DO_CODE As String, ByVal IMP_CODE As String, ByVal STORER_CODE As String, ByVal ITEM_CODE As String, ByVal QTY As Double, ByVal ROUTE_ID As String, ByVal DO_DATE As DateTime, ByVal Optional CUS_CODE As String = "", ByVal Optional LAST_LOT As String = "") As List(Of Dictionary(Of String, Object))
        Dim SQLString As String
        Dim dtRule As DataTable
        Dim dtStock As DataTable
        Dim returnList As New List(Of Dictionary(Of String, Object))
        Dim IsFound = 0
        Dim AllotedLotCount = 0
        Try


            SQLString = "Select a.ITM_CODE,a.ITM_SKU_NO,a.ITM_DESC,b.STORER_CODE,b.STO_NAME from WMS_ITEM a Inner Join WMS_STORER b On a.IMP_CODE=b.IMP_CODE and a.STORER_CODE=b.STORER_CODE Where a.IMP_CODE='" + IMP_CODE + "' and a.STORER_CODE='" + STORER_CODE + "' and ITM_CODE='" + ITEM_CODE + "' And PACK_KEY='1'"
            Dim dtItem As DataTable = gDB.getDataTable(SQLString, conn, transaction)

            SQLString = "Select distinct a.CUS_NAME,a.CUS_CODE from WMS_CUSTOMER a Where a.IMP_CODE='" + IMP_CODE + "' and CUS_CODE='" + CUS_CODE.Split(",")(0) + "' " 'and a.STORER_CODE='" + STORER_CODE + "' 
            Dim dtCust As DataTable = gDB.getDataTable(SQLString, conn, transaction)

            If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                SQLString = "Select a.*,IsNull(b.ITM_SHELF_LIFE,0)ITM_SHELF_LIFE from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1 order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC "
            Else
                SQLString = "Select a.*,IsNull(b.ITM_SHELF_LIFE,0)ITM_SHELF_LIFE from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1 order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
            End If
            dtStock = gDB.getDataTable(SQLString, conn, transaction)

            If dtStock IsNot Nothing And dtStock.Rows.Count > 0 Then
                If CUS_CODE <> "" Then
                    CUS_CODE = CUS_CODE.Split(",")(0)
                    SQLString = "Select ISNULL(MPL_FLAG,0)MPL_FLAG,ISNULL(MIN_PROD_LIFE,0)MIN_PROD_LIFE,ISNULL(MSL_FLAG,0)MSL_FLAG,ISNULL(MIN_SELF_LIFE,0)MIN_SELF_LIFE,ISNULL(MB_FLAG,0)MB_FLAG,ISNULL(MAX_BATCHES,0)MAX_BATCHES,ISNULL(LOTS_CANNOT_BE_EARLIER,0)LOTS_CANNOT_BE_EARLIER,ISNULL(MAX_LOTS,0)MAX_LOTS from WMS_CUSTOMER_RULE where item_code='" & gU.dbEncode(ITEM_CODE) & "' and CUST_CODE='" & gU.dbEncode(CUS_CODE) & "' "
                    dtRule = gDB.getDataTable(SQLString, conn, transaction)
                    If dtRule Is Nothing Or dtRule.Rows.Count <= 0 Then
                        SQLString = "Select ISNULL(MPL_FLAG,0)MPL_FLAG,ISNULL(MIN_PROD_LIFE,0)MIN_PROD_LIFE,ISNULL(MSL_FLAG,0)MSL_FLAG,ISNULL(MIN_SELF_LIFE,0)MIN_SELF_LIFE,ISNULL(MB_FLAG,0)MB_FLAG,ISNULL(MAX_BATCHES,0)MAX_BATCHES,ISNULL(LOTS_CANNOT_BE_EARLIER,0)LOTS_CANNOT_BE_EARLIER,ISNULL(MAX_LOTS,0)MAX_LOTS from WMS_CUSTOMER Where  IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and CUS_CODE='" & gU.dbEncode(CUS_CODE) & "'"
                        dtRule = gDB.getDataTable(SQLString, conn, transaction)
                    End If

                    If dtRule.Rows.Count > 0 AndAlso RULE_COUNT > 0 Then
                        Dim dtToday As DateTime = DO_DATE
                        'System.DateTime.Now
                        Dim PrevLot As DateTime = Nothing
                        returnList.Clear()
                        IsFound = 0
                        Dim MAX_LOTS = Convert.ToInt32(dtRule.Rows(0)("MAX_LOTS").ToString())
                        For Each rowStk In dtStock.Rows


                            Dim dtExpiry As DateTime = Convert.ToDateTime(rowStk("ILOC_EXPIRY_DATE").ToString)
                            Dim strBatch As String = rowStk("ILOC_BATCH_NO").ToString
                            Dim dtLot As DateTime = Nothing
                            If strBatch IsNot Nothing AndAlso strBatch.Length >= 8 Then
                                dtLot = Convert.ToDateTime(strBatch.Substring(0, 4) + "-" + strBatch.Substring(4, 2) + "-" + strBatch.Substring(6, 2))
                            End If
                            SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(rowStk("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(rowStk("ILOC_BATCH_NO")) & "'  and WH_CODE='" & WH_CODE & "' and WH_LOC='" & gU.dbEncode(rowStk("ILOC_LOC")) & "'"
                            'and ROUTE_ID='" & gU.dbEncode(ROUTE_ID) & "' and DO_DATE='" & gU.dbEncode(DO_DATE) & "'
                            Dim dtRSVD As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                            If dtRSVD IsNot Nothing AndAlso dtRSVD.Rows.Count > 0 Then
                                rowStk("ILOC_BAL_QTY") = Convert.ToDouble(rowStk("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVD.Rows(0)("QTY").ToString)
                            End If

                            If dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If dtLot <> Nothing Then
                                        If LAST_LOT.Length >= 8 Then
                                            Dim Last_LOT_DT = Convert.ToDateTime(LAST_LOT.Substring(0, 4) + "-" + LAST_LOT.Substring(4, 2) + "-" + LAST_LOT.Substring(6, 2))
                                            If (dtToday - dtLot).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_PROD_LIFE")) Then
                                                If dtExpiry.Subtract(dtToday).Days > Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                                    If dtLot >= Last_LOT_DT Then
                                                        If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE)-1 >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0  and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                                        Else
                                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE)-1 >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                                        End If
                                                        Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                                        For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                            SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                                            Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                                            If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                                dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                                            End If
                                                        Next
                                                        dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                                        Dim tmpList = New List(Of Dictionary(Of String, Object))
                                                        Dim AvQty As Double = 0
                                                        If MAX_LOTS > 0 Then
                                                            For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                                If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)
                                                                    Exit For
                                                                Else
                                                                    For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                                        If MAX_LOTS = 1 Then
                                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)
                                                                                Exit For
                                                                            End If

                                                                        ElseIf MAX_LOTS = 2 Then

                                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)

                                                                                tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)
                                                                                Exit For
                                                                            End If

                                                                        ElseIf MAX_LOTS = 3 Then
                                                                            If dtTmp.Rows.Count >= j + 1 Then
                                                                                If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                                    tmpList.Add(tmpDict)

                                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                                    tmpList.Add(tmpDict)

                                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                                    tmpList.Add(tmpDict)
                                                                                    Exit For
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next
                                                                End If
                                                                If tmpList.Count > 0 Then
                                                                    Exit For
                                                                End If
                                                            Next
                                                            If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                                For Each item As Dictionary(Of String, Object) In tmpList
                                                                    Dim dictParam As New Dictionary(Of String, Object)()
                                                                    dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                                    dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                                    dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                                    dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                                    dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                                    dictParam.Add("ILOC_BAL_QTY", 0)
                                                                    dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    Dim balQty As Double = 0.00
                                                                    If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                                        balQty = QTY

                                                                    Else
                                                                        balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                                    End If
                                                                    dictParam.Add("DOD_PK_QTY", balQty)

                                                                    QTY = QTY - balQty
                                                                    Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                                    returnList.Add(dictParam)
                                                                    IsFound = 1
                                                                    AllotedLotCount = AllotedLotCount + 1
                                                                Next
                                                            End If
                                                        Else
                                                            For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                                AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))

                                                            Next
                                                            If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                                Dim dictParam As New Dictionary(Of String, Object)()
                                                                dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                                dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                                dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                                dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                                dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                                dictParam.Add("ILOC_BAL_QTY", 0)
                                                                dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                                dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                Dim balQty As Double = 0.00
                                                                If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                                    balQty = QTY
                                                                    dictParam.Add("DOD_PK_QTY", QTY)
                                                                Else
                                                                    balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                                    dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                                End If
                                                                QTY = QTY - balQty
                                                                Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                                returnList.Add(dictParam)
                                                                IsFound = 1
                                                                AllotedLotCount = AllotedLotCount + 1
                                                                'Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If dtLot <> Nothing Then
                                        If LAST_LOT.Length >= 8 Then
                                            Dim Last_LOT_DT = Convert.ToDateTime(LAST_LOT.Substring(0, 4) + "-" + LAST_LOT.Substring(4, 2) + "-" + LAST_LOT.Substring(6, 2))
                                            If (dtToday - dtLot).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_PROD_LIFE")) Then
                                                If dtLot >= Last_LOT_DT Then
                                                    If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                                        SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1 order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                                    Else
                                                        SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "'  and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1 order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                                    End If
                                                    Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                                    For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                        SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                                        Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                                        If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                            dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                                        End If
                                                    Next
                                                    dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                                    Dim tmpList = New List(Of Dictionary(Of String, Object))
                                                    Dim AvQty As Double = 0
                                                    If MAX_LOTS > 0 Then
                                                        For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                            If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            Else
                                                                For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                                    If MAX_LOTS = 1 Then
                                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)
                                                                            Exit For
                                                                        End If

                                                                    ElseIf MAX_LOTS = 2 Then

                                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)

                                                                            tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)
                                                                            Exit For
                                                                        End If

                                                                    ElseIf MAX_LOTS = 3 Then
                                                                        If dtTmp.Rows.Count >= j + 1 Then
                                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)

                                                                                tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)

                                                                                tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next
                                                            End If
                                                            If tmpList.Count > 0 Then
                                                                Exit For
                                                            End If
                                                        Next
                                                        If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                            For Each item As Dictionary(Of String, Object) In tmpList
                                                                Dim dictParam As New Dictionary(Of String, Object)()
                                                                dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                                dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                                dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                                dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                                dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                                dictParam.Add("ILOC_BAL_QTY", 0)
                                                                dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                                dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                Dim balQty As Double = 0.00
                                                                If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                                    balQty = QTY

                                                                Else
                                                                    balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                                End If
                                                                dictParam.Add("DOD_PK_QTY", balQty)

                                                                QTY = QTY - balQty
                                                                Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                                returnList.Add(dictParam)
                                                                IsFound = 1
                                                                AllotedLotCount = AllotedLotCount + 1
                                                            Next
                                                        End If
                                                    Else
                                                        For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                            AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))

                                                        Next
                                                        If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                            Dim dictParam As New Dictionary(Of String, Object)()
                                                            dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                            dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                            dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                            dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                            dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                            dictParam.Add("ILOC_BAL_QTY", 0)
                                                            dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                            dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            Dim balQty As Double = 0.00
                                                            If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                                balQty = QTY
                                                                dictParam.Add("DOD_PK_QTY", QTY)
                                                            Else
                                                                balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                                dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                            End If
                                                            QTY = QTY - balQty
                                                            Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                            returnList.Add(dictParam)
                                                            IsFound = 1
                                                            AllotedLotCount = AllotedLotCount + 1
                                                            'Exit For
                                                        End If
                                                    End If

                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If dtLot <> Nothing Then
                                        If (dtToday - dtLot).Days >= Convert.ToInt16(dtRule.Rows(0)("MIN_PROD_LIFE")) Then
                                            If dtExpiry.Subtract(dtToday).Days > Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                                If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                                    SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE)-1 >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                                Else
                                                    SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE)-1 >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and a.ITM_CODE ='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1 order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                                End If
                                                Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                                For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                    SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                                    Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                                    If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                        dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                                    End If
                                                Next
                                                dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                                Dim tmpList = New List(Of Dictionary(Of String, Object))
                                                Dim AvQty As Double = 0
                                                If MAX_LOTS > 0 Then
                                                    For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                        If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                            tmpList.Add(tmpDict)
                                                            Exit For
                                                        Else
                                                            For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                                If MAX_LOTS = 1 Then
                                                                    If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)
                                                                        Exit For
                                                                    End If

                                                                ElseIf MAX_LOTS = 2 Then

                                                                    If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)

                                                                        tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)
                                                                        Exit For
                                                                    End If

                                                                ElseIf MAX_LOTS = 3 Then
                                                                    If dtTmp.Rows.Count >= j + 1 Then
                                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)

                                                                            tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)

                                                                            tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)
                                                                            Exit For
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                        End If
                                                        If tmpList.Count > 0 Then
                                                            Exit For
                                                        End If
                                                    Next
                                                    If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                        For Each item As Dictionary(Of String, Object) In tmpList
                                                            Dim dictParam As New Dictionary(Of String, Object)()
                                                            dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                            dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                            dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                            dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                            dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                            dictParam.Add("ILOC_BAL_QTY", 0)
                                                            dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                            dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            Dim balQty As Double = 0.00
                                                            If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                                balQty = QTY

                                                            Else
                                                                balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                            End If
                                                            dictParam.Add("DOD_PK_QTY", balQty)

                                                            QTY = QTY - balQty
                                                            Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                            returnList.Add(dictParam)
                                                            IsFound = 1
                                                            AllotedLotCount = AllotedLotCount + 1
                                                        Next
                                                    End If
                                                Else
                                                    For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                        AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))

                                                    Next
                                                    If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                        Dim dictParam As New Dictionary(Of String, Object)()
                                                        dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                        dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                        dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                        dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                        dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                        dictParam.Add("ILOC_BAL_QTY", 0)
                                                        dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                        dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        Dim balQty As Double = 0.00
                                                        If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                            balQty = QTY
                                                            dictParam.Add("DOD_PK_QTY", QTY)
                                                        Else
                                                            balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                            dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                        End If
                                                        QTY = QTY - balQty
                                                        Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                        returnList.Add(dictParam)
                                                        IsFound = 1
                                                        AllotedLotCount = AllotedLotCount + 1
                                                        'Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" Then
                                If LAST_LOT IsNot Nothing AndAlso LAST_LOT <> "" Then
                                    If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                        If LAST_LOT.Length >= 8 Then
                                            Dim Last_LOT_DT = Convert.ToDateTime(LAST_LOT.Substring(0, 4) + "-" + LAST_LOT.Substring(4, 2) + "-" + LAST_LOT.Substring(6, 2))
                                            If dtExpiry.Subtract(dtToday).Days > Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                                If dtLot >= Last_LOT_DT Then
                                                    If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                                        SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE and DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE)-1 >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                                    Else
                                                        SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE and DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE)-1 >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                                    End If
                                                    Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                                    For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                        SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                                        Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                                        If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                            dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                                        End If
                                                    Next
                                                    dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                                    Dim tmpList = New List(Of Dictionary(Of String, Object))
                                                    Dim AvQty As Double = 0
                                                    If MAX_LOTS > 0 Then
                                                        For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                            If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            Else
                                                                For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                                    If MAX_LOTS = 1 Then
                                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)
                                                                            Exit For
                                                                        End If

                                                                    ElseIf MAX_LOTS = 2 Then

                                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)

                                                                            tmpDict = New Dictionary(Of String, Object)
                                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                                            tmpList.Add(tmpDict)
                                                                            Exit For
                                                                        End If

                                                                    ElseIf MAX_LOTS = 3 Then
                                                                        If dtTmp.Rows.Count >= j + 1 Then
                                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)

                                                                                tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)

                                                                                tmpDict = New Dictionary(Of String, Object)
                                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                                tmpList.Add(tmpDict)
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next
                                                            End If
                                                            If tmpList.Count > 0 Then
                                                                Exit For
                                                            End If
                                                        Next
                                                        If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                            For Each item As Dictionary(Of String, Object) In tmpList
                                                                Dim dictParam As New Dictionary(Of String, Object)()
                                                                dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                                dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                                dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                                dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                                dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                                dictParam.Add("ILOC_BAL_QTY", 0)
                                                                dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                                dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                Dim balQty As Double = 0.00
                                                                If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                                    balQty = QTY

                                                                Else
                                                                    balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                                End If
                                                                dictParam.Add("DOD_PK_QTY", balQty)

                                                                QTY = QTY - balQty
                                                                Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                                returnList.Add(dictParam)
                                                                IsFound = 1
                                                                AllotedLotCount = AllotedLotCount + 1
                                                            Next
                                                        End If
                                                    Else
                                                        For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                            AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))

                                                        Next
                                                        If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                            Dim dictParam As New Dictionary(Of String, Object)()
                                                            dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                            dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                            dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                            dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                            dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                            dictParam.Add("ILOC_BAL_QTY", 0)
                                                            dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                            dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            Dim balQty As Double = 0.00
                                                            If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                                balQty = QTY
                                                                dictParam.Add("DOD_PK_QTY", QTY)
                                                            Else
                                                                balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                                dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                            End If
                                                            QTY = QTY - balQty
                                                            Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                            returnList.Add(dictParam)
                                                            IsFound = 1
                                                            AllotedLotCount = AllotedLotCount + 1
                                                            'Exit For
                                                        End If
                                                    End If
                                                    'PrevLot = dtLot
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If LAST_LOT.Length >= 8 Then
                                        Dim Last_LOT_DT = Convert.ToDateTime(LAST_LOT.Substring(0, 4) + "-" + LAST_LOT.Substring(4, 2) + "-" + LAST_LOT.Substring(6, 2))
                                        If dtLot >= Last_LOT_DT Then
                                            If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                                SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                            Else
                                                SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE and Convert(datetime,SUBSTRING(a.ILOC_BATCH_NO,0,5)+'-'+SUBSTRING( a.ILOC_BATCH_NO,5,2)+'-'+SUBSTRING( a.ILOC_BATCH_NO,7,2)) >= '" + Last_LOT_DT + "' and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                            End If
                                            Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                            For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                                Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                                If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                    dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                                End If
                                            Next
                                            dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                            Dim tmpList = New List(Of Dictionary(Of String, Object))
                                            Dim AvQty As Double = 0
                                            If MAX_LOTS > 0 Then
                                                For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                    If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                        tmpList.Add(tmpDict)
                                                        Exit For
                                                    Else
                                                        For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                            If MAX_LOTS = 1 Then
                                                                If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)
                                                                    Exit For
                                                                End If

                                                            ElseIf MAX_LOTS = 2 Then

                                                                If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)

                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)
                                                                    Exit For
                                                                End If

                                                            ElseIf MAX_LOTS = 3 Then
                                                                If dtTmp.Rows.Count >= j + 1 Then
                                                                    If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)

                                                                        tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)

                                                                        tmpDict = New Dictionary(Of String, Object)
                                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                                        tmpList.Add(tmpDict)
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                    If tmpList.Count > 0 Then
                                                        Exit For
                                                    End If
                                                Next
                                                If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                    For Each item As Dictionary(Of String, Object) In tmpList
                                                        Dim dictParam As New Dictionary(Of String, Object)()
                                                        dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                        dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                        dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                        dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                        dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                        dictParam.Add("ILOC_BAL_QTY", 0)
                                                        dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                        dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        Dim balQty As Double = 0.00
                                                        If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                            balQty = QTY

                                                        Else
                                                            balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                        End If
                                                        dictParam.Add("DOD_PK_QTY", balQty)

                                                        QTY = QTY - balQty
                                                        Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                        returnList.Add(dictParam)
                                                        IsFound = 1
                                                        AllotedLotCount = AllotedLotCount + 1
                                                    Next
                                                End If
                                            Else
                                                For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                    AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))
                                                Next
                                                If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                    Dim dictParam As New Dictionary(Of String, Object)()
                                                    dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                    dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                    dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                    dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                    dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                    dictParam.Add("ILOC_BAL_QTY", 0)
                                                    dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                    dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                    dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                    dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                    Dim balQty As Double = 0.00
                                                    If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                        balQty = QTY
                                                        dictParam.Add("DOD_PK_QTY", QTY)
                                                    Else
                                                        balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                        dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                    End If
                                                    QTY = QTY - balQty
                                                    Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                    returnList.Add(dictParam)
                                                    IsFound = 1
                                                    AllotedLotCount = AllotedLotCount + 1
                                                    'Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("MSL_FLAG") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If dtExpiry.Subtract(dtToday).Days > Convert.ToInt16(dtRule.Rows(0)("MIN_SELF_LIFE")) Then
                                        'Dim StkbalQty = Double.Parse(dtRSVD.Rows(0)("QTY"))
                                        If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE)-1 >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                        Else
                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtToday + "', a.ILOC_EXPIRY_DATE)-1 >= " + dtRule.Rows(0)("MIN_SELF_LIFE") + " and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                        End If
                                        Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                        For i As Integer = 0 To dtTmp.Rows.Count - 1
                                            SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                            Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                            If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                            End If
                                        Next
                                        dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                        Dim tmpList = New List(Of Dictionary(Of String, Object))
                                        Dim AvQty As Double = 0
                                        If MAX_LOTS > 0 Then
                                            For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                    tmpList.Add(tmpDict)
                                                    Exit For
                                                Else
                                                    For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                        If MAX_LOTS = 1 Then
                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            End If

                                                        ElseIf MAX_LOTS = 2 Then

                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)

                                                                tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            End If

                                                        ElseIf MAX_LOTS = 3 Then
                                                            If dtTmp.Rows.Count >= j + 1 Then
                                                                If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)

                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)

                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                                If tmpList.Count > 0 Then
                                                    Exit For
                                                End If
                                            Next
                                            If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                For Each item As Dictionary(Of String, Object) In tmpList
                                                    Dim dictParam As New Dictionary(Of String, Object)()
                                                    dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                    dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                    dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                    dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                    dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                    dictParam.Add("ILOC_BAL_QTY", 0)
                                                    dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                    dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                    dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                    dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                    Dim balQty As Double = 0.00
                                                    If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                        balQty = QTY

                                                    Else
                                                        balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                    End If
                                                    dictParam.Add("DOD_PK_QTY", balQty)

                                                    QTY = QTY - balQty
                                                    Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                    returnList.Add(dictParam)
                                                    IsFound = 1
                                                    AllotedLotCount = AllotedLotCount + 1
                                                Next
                                            End If
                                        Else
                                            For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))
                                            Next
                                            If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                Dim dictParam As New Dictionary(Of String, Object)()
                                                dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                dictParam.Add("ILOC_BAL_QTY", 0)
                                                dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                Dim balQty As Double = 0.00
                                                If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                    balQty = QTY
                                                    dictParam.Add("DOD_PK_QTY", QTY)
                                                Else
                                                    balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                    dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                End If
                                                QTY = QTY - balQty
                                                Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                returnList.Add(dictParam)
                                                IsFound = 1
                                                AllotedLotCount = AllotedLotCount + 1
                                                'Exit For
                                            End If
                                        End If
                                        'PrevLot = dtLot
                                    End If
                                End If
                            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" Then
                                If AllotedLotCount < MAX_LOTS OrElse MAX_LOTS <= 0 Then
                                    If (dtToday - dtExpiry).Days <= Convert.ToInt16(dtRule.Rows(0)("MIN_PROD_LIFE")) Then
                                        If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                        Else
                                            SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where DATEDIFF(day, '" + dtLot + "', '" + dtToday + "') >= " + dtRule.Rows(0)("MIN_PROD_LIFE") + " and a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1 order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                        End If
                                        Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                        For i As Integer = 0 To dtTmp.Rows.Count - 1
                                            SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                            Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                            If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                                dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                            End If
                                        Next
                                        dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                        Dim tmpList = New List(Of Dictionary(Of String, Object))
                                        Dim AvQty As Double = 0
                                        If MAX_LOTS > 0 Then
                                            For i As Integer = 0 To dtTmp.Rows.Count - 1

                                                If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                    tmpList.Add(tmpDict)
                                                    Exit For
                                                Else
                                                    For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                        If MAX_LOTS = 1 Then
                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            End If

                                                        ElseIf MAX_LOTS = 2 Then

                                                            If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                                Dim tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)

                                                                tmpDict = New Dictionary(Of String, Object)
                                                                tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                tmpDict.Add("DOD_PK_QTY", 0)
                                                                tmpList.Add(tmpDict)
                                                                Exit For
                                                            End If

                                                        ElseIf MAX_LOTS = 3 Then
                                                            If dtTmp.Rows.Count >= j + 1 Then
                                                                If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                                    Dim tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)

                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)

                                                                    tmpDict = New Dictionary(Of String, Object)
                                                                    tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                                    tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                                    tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                                    tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                                    tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                                    tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                                    tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                                    tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                                    tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                                    tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                                    tmpDict.Add("DOD_PK_QTY", 0)
                                                                    tmpList.Add(tmpDict)
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                                If tmpList.Count > 0 Then
                                                    Exit For
                                                End If
                                            Next
                                            If QTY > 0 AndAlso tmpList.Count > 0 Then
                                                For Each item As Dictionary(Of String, Object) In tmpList
                                                    Dim dictParam As New Dictionary(Of String, Object)()
                                                    dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                                    dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                                    dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                                    dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                                    dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                                    dictParam.Add("ILOC_BAL_QTY", 0)
                                                    dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                    dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                    dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                    dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                    Dim balQty As Double = 0.00
                                                    If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                        balQty = QTY

                                                    Else
                                                        balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                                    End If
                                                    dictParam.Add("DOD_PK_QTY", balQty)

                                                    QTY = QTY - balQty
                                                    Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                    returnList.Add(dictParam)
                                                    IsFound = 1
                                                    AllotedLotCount = AllotedLotCount + 1
                                                Next
                                            End If
                                        Else
                                            For i As Integer = 0 To dtTmp.Rows.Count - 1
                                                AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))
                                            Next
                                            If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                                Dim dictParam As New Dictionary(Of String, Object)()
                                                dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                                dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                                dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                                dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                                dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                                dictParam.Add("ILOC_BAL_QTY", 0)
                                                dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                                dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                Dim balQty As Double = 0.00
                                                If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                                    balQty = QTY
                                                    dictParam.Add("DOD_PK_QTY", QTY)
                                                Else
                                                    balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                                    dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                                End If
                                                QTY = QTY - balQty
                                                Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                                returnList.Add(dictParam)
                                                IsFound = 1
                                                AllotedLotCount = AllotedLotCount + 1
                                                'Exit For
                                            End If
                                        End If

                                    End If
                                End If
                            Else
                                'Dim StkbalQty = Double.Parse(dtRSVD.Rows(0)("QTY"))
                                If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                                    SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1  order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                Else
                                    SQLString = "Select a.* from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103)+1 order by a.ILOC_EXPIRY_DATE,a.ILOC_BATCH_NO,a.ILOC_LOC"
                                End If
                                Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                                For i As Integer = 0 To dtTmp.Rows.Count - 1
                                    SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtTmp.Rows(i)("ILOC_LOC")) & "'"
                                    Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                                    If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                        dtTmp.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtTmp.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                                    End If
                                Next
                                dtTmp = dtTmp.Select("ILOC_BAL_QTY > 0", "ILOC_EXPIRY_DATE").CopyToDataTable()
                                Dim tmpList = New List(Of Dictionary(Of String, Object))
                                Dim AvQty As Double = 0
                                If MAX_LOTS > 0 Then
                                    For i As Integer = 0 To dtTmp.Rows.Count - 1

                                        If QTY > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) > 0 AndAlso Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                            Dim tmpDict = New Dictionary(Of String, Object)
                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                            tmpDict.Add("DOD_PK_QTY", 0)
                                            tmpList.Add(tmpDict)
                                            Exit For
                                        Else
                                            For j As Integer = i + 1 To dtTmp.Rows.Count - 1
                                                If MAX_LOTS = 1 Then
                                                    If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) >= QTY Then
                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                        tmpList.Add(tmpDict)
                                                        Exit For
                                                    End If

                                                ElseIf MAX_LOTS = 2 Then

                                                    If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) >= QTY Then

                                                        Dim tmpDict = New Dictionary(Of String, Object)
                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                        tmpList.Add(tmpDict)

                                                        tmpDict = New Dictionary(Of String, Object)
                                                        tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                        tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                        tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                        tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                        tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                        tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                        tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                        tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                        tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                        tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                        tmpDict.Add("DOD_PK_QTY", 0)
                                                        tmpList.Add(tmpDict)
                                                        Exit For
                                                    End If

                                                ElseIf MAX_LOTS = 3 Then
                                                    If dtTmp.Rows.Count >= j + 1 Then
                                                        If Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")) + Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")) >= QTY Then

                                                            Dim tmpDict = New Dictionary(Of String, Object)
                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(i)("ILOC_EXPIRY_DATE"))
                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(i)("ITM_CODE"))
                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(i)("ILOC_BATCH_NO"))
                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(i)("ILOC_LOC"))
                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(i)("ILOC_WH"))
                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY")))
                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                            tmpList.Add(tmpDict)

                                                            tmpDict = New Dictionary(Of String, Object)
                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j)("ILOC_EXPIRY_DATE"))
                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j)("ITM_CODE"))
                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j)("ILOC_BATCH_NO"))
                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j)("ILOC_LOC"))
                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j)("ILOC_WH"))
                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j)("ILOC_BAL_QTY")))
                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                            tmpList.Add(tmpDict)

                                                            tmpDict = New Dictionary(Of String, Object)
                                                            tmpDict.Add("ILOC_EXPIRY_DATE", dtTmp.Rows(j + 1)("ILOC_EXPIRY_DATE"))
                                                            tmpDict.Add("ITM_CODE", dtTmp.Rows(j + 1)("ITM_CODE"))
                                                            tmpDict.Add("ILOC_BATCH_NO", dtTmp.Rows(j + 1)("ILOC_BATCH_NO"))
                                                            tmpDict.Add("ILOC_LOC", dtTmp.Rows(j + 1)("ILOC_LOC"))
                                                            tmpDict.Add("ILOC_WH", dtTmp.Rows(j + 1)("ILOC_WH"))
                                                            tmpDict.Add("ILOC_BAL_QTY", Double.Parse(dtTmp.Rows(j + 1)("ILOC_BAL_QTY")))
                                                            tmpDict.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                                            tmpDict.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                                            tmpDict.Add("DOD_LAST_LOT", LAST_LOT)
                                                            tmpDict.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                                            tmpDict.Add("DOD_PK_QTY", 0)
                                                            tmpList.Add(tmpDict)
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                        If tmpList.Count > 0 Then
                                            Exit For
                                        End If
                                    Next
                                    If QTY > 0 AndAlso tmpList.Count > 0 Then
                                        For Each item As Dictionary(Of String, Object) In tmpList
                                            Dim dictParam As New Dictionary(Of String, Object)()
                                            dictParam.Add("ILOC_EXPIRY_DATE", item("ILOC_EXPIRY_DATE"))
                                            dictParam.Add("ITM_CODE", item("ITM_CODE"))
                                            dictParam.Add("ILOC_BATCH_NO", item("ILOC_BATCH_NO"))
                                            dictParam.Add("ILOC_LOC", item("ILOC_LOC"))
                                            dictParam.Add("ILOC_WH", item("ILOC_WH"))
                                            dictParam.Add("ILOC_BAL_QTY", 0)
                                            dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                            dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                            dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                            dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                            Dim balQty As Double = 0.00
                                            If Double.Parse(item("ILOC_BAL_QTY")) >= (QTY) Then
                                                balQty = QTY

                                            Else
                                                balQty = Double.Parse(item("ILOC_BAL_QTY"))


                                            End If
                                            dictParam.Add("DOD_PK_QTY", balQty)

                                            QTY = QTY - balQty
                                            Dim result = SaveReserve(conn, transaction, item("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, item("ITM_CODE"), item("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                            returnList.Add(dictParam)
                                            IsFound = 1
                                            AllotedLotCount = AllotedLotCount + 1
                                        Next
                                    End If
                                Else
                                    For i As Integer = 0 To dtTmp.Rows.Count - 1
                                        AvQty = AvQty + Double.Parse(dtTmp.Rows(i)("ILOC_BAL_QTY"))
                                    Next
                                    If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) > 0 AndAlso AvQty >= QTY Then

                                        Dim dictParam As New Dictionary(Of String, Object)()
                                        dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                        dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                        dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                        dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                        dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                        dictParam.Add("ILOC_BAL_QTY", 0)
                                        dictParam.Add("DOD_MIN_PROD_DATE", dtRule.Rows(0)("MIN_PROD_LIFE"))
                                        dictParam.Add("DOD_MIN_SHELF_LIFE", dtRule.Rows(0)("MIN_SELF_LIFE"))
                                        dictParam.Add("DOD_LAST_LOT", LAST_LOT)
                                        dictParam.Add("DOD_MAX_LOT", dtRule.Rows(0)("MAX_LOTS"))
                                        Dim balQty As Double = 0.00
                                        If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                            balQty = QTY
                                            dictParam.Add("DOD_PK_QTY", QTY)
                                        Else
                                            balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                            dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                        End If
                                        QTY = QTY - balQty
                                        Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                        returnList.Add(dictParam)
                                        IsFound = 1
                                        AllotedLotCount = AllotedLotCount + 1
                                        'Exit For
                                    End If
                                End If

                            End If
                        Next
                        If IsFound = 0 Then
                            lstError.Add("Storer : " + dtItem.Rows(0)("STO_NAME") + ",ITEM : " + dtItem.Rows(0)("ITM_SKU_NO") + "{" + dtItem.Rows(0)("ITM_DESC") + "/" + ITEM_CODE + "}" + ", Qty: " + QTY.ToString + " Route: " + ROUTE_ID.ToString + " ,Customer: {" + CUS_CODE.ToString + "/" + dtCust.Rows(0)("CUS_NAME") + "} , CO STATUS : NEW is not available in stock!")
                            updateCO_Details(ITEM_CODE, STORER_CODE, CUS_CODE, DO_CODE, ROUTE_ID, conn, transaction)
                        End If
                    Else
                        IsFound = 0
                        Dim AvlQty As Double = 0.0000
                        For i As Integer = 0 To dtStock.Rows.Count - 1
                            SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtStock.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtStock.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtStock.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtStock.Rows(i)("ILOC_LOC")) & "'"
                            Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                            If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                                dtStock.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtStock.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                            End If
                            AvlQty = AvlQty + Convert.ToDouble(dtStock.Rows(i)("ILOC_BAL_QTY"))
                        Next
                        For Each rowStk In dtStock.Rows
                            'SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(rowStk("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(rowStk("ILOC_BATCH_NO")) & "'  and WH_CODE='" & WH_CODE & "' and WH_LOC='" & gU.dbEncode(rowStk("ILOC_LOC")) & "'"
                            ''and ROUTE_ID='" & gU.dbEncode(ROUTE_ID) & "' and DO_DATE='" & gU.dbEncode(DO_DATE) & "'
                            'Dim dtRSVD As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                            'If dtRSVD IsNot Nothing AndAlso dtRSVD.Rows.Count > 0 Then
                            '    rowStk("ILOC_BAL_QTY") = Convert.ToDouble(rowStk("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVD.Rows(0)("QTY").ToString)
                            'End If
                            'If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                            '    SQLString = "Select IsNUll(Sum(a.ILOC_BAL_QTY),0)AvQty from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103) "
                            'Else
                            '    SQLString = "Select IsNUll(Sum(a.ILOC_BAL_QTY),0)AvQty from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0  and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103) "
                            'End If
                            'Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                            If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) >= 0 AndAlso AvlQty >= QTY Then 'AndAlso (Convert.ToDouble(dtTmp.Rows(0)("AvQty").ToString) - Convert.ToDouble(dtRSVD.Rows(0)("QTY").ToString)) >= QTY Then
                                Dim dictParam As New Dictionary(Of String, Object)()
                                dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                                dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                                dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                                dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                                dictParam.Add("ILOC_BAL_QTY", 0)
                                dictParam.Add("DOD_MIN_PROD_DATE", "0")
                                dictParam.Add("DOD_MIN_SHELF_LIFE", "0")
                                dictParam.Add("DOD_LAST_LOT", "")
                                dictParam.Add("DOD_MAX_LOT", "0")
                                Dim balQty As Double = 0.00

                                If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                    balQty = QTY
                                    dictParam.Add("DOD_PK_QTY", QTY)
                                Else
                                    balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                    dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                                End If
                                QTY = QTY - balQty
                                Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                                returnList.Add(dictParam)
                                IsFound = 1


                                'Exit For

                            End If
                        Next
                        If IsFound = 0 Then
                            lstError.Add("Storer : " + dtItem.Rows(0)("STO_NAME") + ",ITEM : " + dtItem.Rows(0)("ITM_SKU_NO") + "{" + dtItem.Rows(0)("ITM_DESC") + "/" + ITEM_CODE + "}" + ", Qty: " + QTY.ToString + " Route: " + ROUTE_ID.ToString + " ,Customer: {" + CUS_CODE.ToString + "/" + dtCust.Rows(0)("CUS_NAME") + "} , CO STATUS : NEW is not available in stock!")
                            updateCO_Details(ITEM_CODE, STORER_CODE, CUS_CODE, DO_CODE, ROUTE_ID, conn, transaction)
                            'lstError.Add(ITEM_CODE + " Qty: " + QTY.ToString + " For Route: " + ROUTE_ID.ToString + " ,Customer: " + CUS_CODE.ToString + " Not available in stock!")
                        End If
                    End If
                Else
                    IsFound = 0
                    Dim AvlQty As Double = 0.0000
                    For i As Integer = 0 To dtStock.Rows.Count - 1
                        SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(dtStock.Rows(i)("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(dtStock.Rows(i)("ILOC_BATCH_NO")) & "'  and WH_CODE='" & gU.dbEncode(dtStock.Rows(i)("ILOC_WH")) & "' and WH_LOC='" & gU.dbEncode(dtStock.Rows(i)("ILOC_LOC")) & "'"
                        Dim dtRSVDTmp As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                        If dtRSVDTmp IsNot Nothing AndAlso dtRSVDTmp.Rows.Count > 0 Then
                            dtStock.Rows(i)("ILOC_BAL_QTY") = Convert.ToDouble(dtStock.Rows(i)("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVDTmp.Rows(0)("QTY").ToString)
                        End If
                        AvlQty = AvlQty + Convert.ToDouble(dtStock.Rows(i)("ILOC_BAL_QTY"))
                    Next
                    For Each rowStk In dtStock.Rows
                        'SQLString = "Select ISNULL(Sum(QTY),0)QTY from WMS_WAVEPICK_RSVD  Where ITEM_CODE='" & gU.dbEncode(rowStk("ITM_CODE")) & "' and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and LOT_NO='" & gU.dbEncode(rowStk("ILOC_BATCH_NO")) & "' and ROUTE_ID='" & gU.dbEncode(ROUTE_ID) & "' and DO_DATE='" & gU.dbEncode(DO_DATE) & "' and WH_CODE='" & WH_CODE & "' and WH_LOC='" & gU.dbEncode(rowStk("ILOC_LOC")) & "'"
                        'Dim dtRSVD As DataTable = gDB.getDataTable(SQLString, conn, transaction)
                        'If dtRSVD IsNot Nothing AndAlso dtRSVD.Rows.Count > 0 Then
                        '    rowStk("ILOC_BAL_QTY") = Convert.ToDouble(rowStk("ILOC_BAL_QTY").ToString) - Convert.ToDouble(dtRSVD.Rows(0)("QTY").ToString)
                        'End If

                        'If WH_CODE IsNot Nothing AndAlso WH_CODE <> "" Then
                        '    SQLString = "Select IsNUll(Sum(a.ILOC_BAL_QTY),0)AvQty from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and ILOC_WH='" & WH_CODE & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103) "
                        'Else
                        '    SQLString = "Select IsNUll(Sum(a.ILOC_BAL_QTY),0)AvQty from WMS_ITEM_LOC_BAL a inner join WMS_ITEM b on a.ITM_CODE=b.ITM_CODE and b.IMP_CODE=a.IMP_CODE and b.STORER_CODE=a.STORER_CODE   Where a.ITM_CODE='" & gU.dbEncode(ITEM_CODE) & "' and a.IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and a.STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and a.ILOC_BAL_QTY > 0 and CONVERT(datetime,a.ILOC_EXPIRY_DATE,103)>=CONVERT(datetime,'" & DO_DATE.ToString(gU.getConfig("DDFORMAT2")) & "',103) "
                        'End If
                        'Dim dtTmp = gDB.getDataTable(SQLString, conn, transaction)
                        If QTY > 0 AndAlso Double.Parse(rowStk("ILOC_BAL_QTY")) >= 0 AndAlso AvlQty >= QTY Then ' (Convert.ToDouble(dtTmp.Rows(0)("AvQty").ToString) - Convert.ToDouble(dtRSVD.Rows(0)("QTY").ToString)) >= QTY Then
                            Dim dictParam As New Dictionary(Of String, Object)()
                            dictParam.Add("ILOC_EXPIRY_DATE", rowStk("ILOC_EXPIRY_DATE"))
                            dictParam.Add("ITM_CODE", rowStk("ITM_CODE"))
                            dictParam.Add("ILOC_BATCH_NO", rowStk("ILOC_BATCH_NO"))
                            dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                            dictParam.Add("ILOC_WH", rowStk("ILOC_WH"))
                            dictParam.Add("ILOC_BAL_QTY", 0)
                            dictParam.Add("DOD_MIN_PROD_DATE", "0")
                            dictParam.Add("DOD_MIN_SHELF_LIFE", "0")
                            dictParam.Add("DOD_LAST_LOT", "")
                            dictParam.Add("DOD_MAX_LOT", "0")
                            Dim balQty As Double = 0.00

                            If Double.Parse(rowStk("ILOC_BAL_QTY")) >= (QTY) Then
                                balQty = QTY
                                dictParam.Add("DOD_PK_QTY", QTY)
                            Else
                                balQty = Double.Parse(rowStk("ILOC_BAL_QTY"))
                                dictParam.Add("DOD_PK_QTY", Double.Parse(rowStk("ILOC_BAL_QTY")))

                            End If
                            QTY = QTY - balQty
                            Dim result = SaveReserve(conn, transaction, rowStk("ILOC_LOC"), WH_CODE, IMP_CODE, STORER_CODE, CUS_CODE, rowStk("ITM_CODE"), rowStk("ILOC_BATCH_NO"), balQty, ROUTE_ID, DO_DATE, DO_CODE)
                            returnList.Add(dictParam)
                            IsFound = 1


                            'Exit For

                        End If
                    Next
                    If IsFound = 0 Then
                        lstError.Add("Storer : " + dtItem.Rows(0)("STO_NAME") + ",ITEM : " + dtItem.Rows(0)("ITM_SKU_NO") + "{" + dtItem.Rows(0)("ITM_DESC") + "/" + ITEM_CODE + "}" + ", Qty: " + QTY.ToString + " Route: " + ROUTE_ID.ToString + " ,Customer: {" + CUS_CODE.ToString + "/" + dtCust.Rows(0)("CUS_NAME") + "} , CO STATUS : NEW is not available in stock!")
                        updateCO_Details(ITEM_CODE, STORER_CODE, CUS_CODE, DO_CODE, ROUTE_ID, conn, transaction)
                        'lstError.Add(ITEM_CODE + " Qty: " + QTY.ToString + " For Route: " + ROUTE_ID.ToString + " ,Customer: " + CUS_CODE.ToString + " Not available in stock!")
                    End If
                End If

            Else
                lstError.Add("Storer : " + dtItem.Rows(0)("STO_NAME") + ",ITEM : " + dtItem.Rows(0)("ITM_SKU_NO") + "{" + dtItem.Rows(0)("ITM_DESC") + "/" + ITEM_CODE + "}" + ", Qty: " + QTY.ToString + " Route: " + ROUTE_ID.ToString + " ,Customer: {" + CUS_CODE.ToString + "/" + dtCust.Rows(0)("CUS_NAME") + "} , CO STATUS : NEW is not available in stock!")
                updateCO_Details(ITEM_CODE, STORER_CODE, CUS_CODE, DO_CODE, ROUTE_ID, conn, transaction)
                'lstError.Add(ITEM_CODE + " Qty:  " + QTY.ToString + " For Route: " + ROUTE_ID.ToString + " ,Customer: " + CUS_CODE.ToString + " not available in stock!")
                'Dim dictParam As New Dictionary(Of String, Object)()
                'dictParam.Add("ILOC_EXPIRY_DATE", "")
                'dictParam.Add("ITM_CODE", ITEM_CODE)
                'dictParam.Add("ILOC_BATCH_NO", "")
                'dictParam.Add("ILOC_LOC", "")
                'dictParam.Add("ILOC_WH", "")
                'dictParam.Add("ILOC_BAL_QTY", 0)
                'dictParam.Add("DOD_MIN_PROD_DATE", "0")
                'dictParam.Add("DOD_MIN_SHELF_LIFE", "0")
                'dictParam.Add("DOD_LAST_LOT", "")
                'dictParam.Add("DOD_MAX_LOT", "")
                'dictParam.Add("DOD_PK_QTY", "0")
                'SaveReserve(IMP_CODE, STORER_CODE, CUS_CODE, ITEM_CODE, "", QTY, ROUTE_ID, DO_DATE, DO_CODE)
                'returnList.Add(dictParam)
            End If
        Catch ex As Exception
            'transaction.Rollback()
            lstError.Add("Error : " + ex.Message)
        End Try
        Return returnList
    End Function

    Private Function IsRule(ByRef conn As SqlConnection, ByRef transaction As SqlTransaction, ByVal IMP_CODE As String, ByVal STORER_CODE As String, ByVal CUS_CODE As String, ByVal ITEM_CODE As String) As DataRow
        Dim SQLString As String
        Dim dtRule As DataTable = Nothing
        Dim hasRule As Int16 = 0
        SQLString = "Select [ITEM_CODE] ,[CUST_CODE] ,isnull(MPL_FLAG,0)MPL_FLAG ,[MIN_PROD_LIFE] ,isnull(MSL_FLAG,0) MSL_FLAG ,[MIN_SELF_LIFE] ,isnull(MB_FLAG,0)MB_FLAG ,isnull(MAX_BATCHES,0)MAX_BATCHES ,isnull([LOTS_CANNOT_BE_EARLIER],0) [LOTS_CANNOT_BE_EARLIER] ,[MAX_LOTS], [Series] from WMS_CUSTOMER_RULE where (MPL_FLAG=1 Or MSL_FLAG=1 Or MB_FLAG=1 or LOTS_CANNOT_BE_EARLIER =1) and item_code='" & gU.dbEncode(ITEM_CODE) & "' and CUST_CODE='" & gU.dbEncode(CUS_CODE) & "' "
        dtRule = gDB.getDataTable(SQLString, conn, transaction)
        If dtRule Is Nothing Or dtRule.Rows.Count <= 0 Then
            SQLString = "Select isnull(MPL_FLAG,0)MPL_FLAG,MIN_PROD_LIFE,isnull(MSL_FLAG,0) MSL_FLAG,MIN_SELF_LIFE,isnull(MB_FLAG,0)MB_FLAG,isnull(MAX_BATCHES,0)MAX_BATCHES,LOTS_CANNOT_BE_EARLIER from WMS_CUSTOMER Where (MPL_FLAG=1 Or MSL_FLAG=1 Or MB_FLAG=1 or LOTS_CANNOT_BE_EARLIER =1) and IMP_CODE='" & gU.dbEncode(IMP_CODE) & "' and STORER_CODE='" & gU.dbEncode(STORER_CODE) & "' and CUS_CODE='" & gU.dbEncode(CUS_CODE) & "'"
            dtRule = gDB.getDataTable(SQLString, conn, transaction)
        End If
        dtRule.Columns.Add("RuleCount")
        If dtRule IsNot Nothing AndAlso dtRule.Rows.Count > 0 Then
            If dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 4

            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1" Then
                dtRule.Rows(0)("RuleCount") = 3

            ElseIf (dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1") OrElse (dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1") OrElse (dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1") Then
                dtRule.Rows(0)("RuleCount") = 2

            ElseIf (dtRule.Rows(0)("MB_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "1") OrElse (dtRule.Rows(0)("MB_FLAG") = "1" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1") OrElse (dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MB_FLAG") = "1") Then
                dtRule.Rows(0)("RuleCount") = 2

            ElseIf dtRule.Rows(0)("MSL_FLAG") = "1" AndAlso dtRule.Rows(0)("MPL_FLAG") = "0" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "0" AndAlso dtRule.Rows(0)("MB_FLAG") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            ElseIf dtRule.Rows(0)("MPL_FLAG") = "1" AndAlso dtRule.Rows(0)("MSL_FLAG") = "0" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "0" AndAlso dtRule.Rows(0)("MB_FLAG") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            ElseIf dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "1" AndAlso dtRule.Rows(0)("MPL_FLAG") = "0" AndAlso dtRule.Rows(0)("MSL_FLAG") = "0" AndAlso dtRule.Rows(0)("MB_FLAG") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            ElseIf dtRule.Rows(0)("MB_FLAG") = "1" AndAlso dtRule.Rows(0)("MPL_FLAG") = "0" AndAlso dtRule.Rows(0)("MSL_FLAG") = "0" AndAlso dtRule.Rows(0)("LOTS_CANNOT_BE_EARLIER") = "0" Then
                dtRule.Rows(0)("RuleCount") = 1
            End If
        Else
            Return Nothing
        End If
        Return dtRule.Rows(0)
    End Function

    Protected Sub btnPicking_Click(sender As Object, e As System.EventArgs) Handles btnPicking.Click
        Dim SQLString As String
        Dim updateSQL As String = ""
        Dim STORER_TYPE As String = ""
        'Dim STORER_TYPEDt As DataTable
        Dim nextNo As String
        Dim coHdrDATES As DataTable
        Dim coHdrGrp As DataTable
        Dim coDtl As DataTable
        Dim Qty As Double = 0
        Dim gConn = gDB.getConnection()
        Dim transaction = Nothing ' gConn.BeginTransaction
        Dim SeqNo As Int16 = 0
        Dim OrderCount = 0
        lstError.Clear()
        ltlMessage.Text = ""
        Dim ruleWareList = New List(Of String)
        ruleWareList.Add("FG")
        ruleWareList.Add("VIP")
        Try
            Dim DO_DATES = ""
            Dim ROUTE_IDS = ""
            For Each item As ListItem In lboxCODate.Items
                If item.Selected = True Then
                    If DO_DATES = "" Then
                        DO_DATES = "'" + item.Value + "'"
                    Else
                        DO_DATES = DO_DATES + ",'" + item.Value + "'"
                    End If

                End If
            Next
            For Each item As ListItem In lboxCORoute.Items
                If item.Selected = True Then
                    If ROUTE_IDS = "" Then
                        ROUTE_IDS = "'" + item.Value + "'"
                    Else
                        ROUTE_IDS = ROUTE_IDS + ",'" + item.Value + "'"
                    End If

                End If
            Next

            'SQLString = "select * from WMS_STORER where STO_REM='NEW_NEW' AND STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "'"
            'STORER_TYPEDt = gDB.getDataTable(SQLString, gConn, transaction)

            'If STORER_TYPEDt IsNot Nothing AndAlso STORER_TYPEDt.Rows.Count > 0 Then
            '    STORER_TYPE = "NEW_NEW"
            '    SQLString = "Select Distinct IMP_CODE,STORER_CODE,CO_DATE,ROUTE_ID,CO_INV_NO FROM WMS_CUST_ORDER SS Where ss.CO_STATUS='NEW' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "' "
            'Else
            SQLString = "Select Distinct IMP_CODE,STORER_CODE,CO_DATE,ROUTE_ID FROM WMS_CUST_ORDER SS Where ss.CO_STATUS='NEW' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "' "
            'End If
            If ROUTE_IDS <> "" AndAlso DO_DATES <> "" Then
                SQLString += " and (CO_DATE in ( " + DO_DATES + ") And ROUTE_ID in (" + ROUTE_IDS + "))"
                coHdrDATES = gDB.getDataTable(SQLString, gConn, transaction)
                'ElseIf ROUTE_IDS <> "" Then 
                '    SQLString += " and (ROUTE_ID in (" + ROUTE_IDS + "))"
                '    coHdrDATES = gDB.getDataTable(SQLString, gConn, transaction)
                'ElseIf DO_DATES <> "" Then
                '    SQLString += " and (CO_DATE in (" + DO_DATES + "))"
                '    coHdrDATES = gDB.getDataTable(SQLString, gConn, transaction)
            Else
                SQLString = ""
                coHdrDATES = Nothing
            End If

            If coHdrDATES IsNot Nothing AndAlso coHdrDATES.Rows.Count > 0 Then
                For Each dtRow As DataRow In coHdrDATES.Rows
                    Dim STORER_CODE = dtRow("STORER_CODE")
                    Dim IMP_CODE = dtRow("IMP_CODE")
                    'If STORER_TYPE = "NEW_NEW" Then
                    '    SQLString = "SELECT Distinct SS.ROUTE_ID,SS.CO_INV_NO,STUFF((SELECT ', ' + US.CO_CODE FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' and US.CO_INV_NO = '" & dtRow("CO_INV_NO") & "' and us.STORER_CODE=ss.STORER_CODE and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "' FOR XML PATH('')), 1, 1, '') [CO_CODES],STUFF((SELECT ', ' + US.CUS_CODE FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' and US.CO_INV_NO = '" & dtRow("CO_INV_NO") & "' and us.STORER_CODE=ss.STORER_CODE and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "' FOR XML PATH('')), 1, 1, '') [CUS_CODES],STUFF((SELECT ', ' + US.CUS_NAME FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' and US.CO_INV_NO = '" & dtRow("CO_INV_NO") & "' and us.STORER_CODE=ss.STORER_CODE and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "' FOR XML PATH('')), 1, 1, '') [CUS_NAMES] FROM WMS_CUST_ORDER SS Where ss.ROUTE_ID = '" & gU.dbEncode(dtRow("ROUTE_ID")) & "' and ss.CO_Date ='" & gU.dbEncode(dtRow("CO_DATE")) & "' and ss.CO_STATUS='NEW' and ss.CO_INV_NO = '" & dtRow("CO_INV_NO") & "' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "'"
                    'Else
                    SQLString = "SELECT Distinct SS.ROUTE_ID,STUFF((SELECT ', ' + US.CO_CODE FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' and us.STORER_CODE=ss.STORER_CODE and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "' FOR XML PATH('')), 1, 1, '') [CO_CODES],STUFF((SELECT ', ' + US.CUS_CODE FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' and us.STORER_CODE=ss.STORER_CODE and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "' FOR XML PATH('')), 1, 1, '') [CUS_CODES],STUFF((SELECT ', ' + US.CUS_NAME FROM WMS_CUST_ORDER US WHERE US.ROUTE_ID = SS.ROUTE_ID and US.CO_DATE=SS.CO_DATE and US.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' and us.STORER_CODE=ss.STORER_CODE and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "' FOR XML PATH('')), 1, 1, '') [CUS_NAMES] FROM WMS_CUST_ORDER SS Where ss.ROUTE_ID = '" & gU.dbEncode(dtRow("ROUTE_ID")) & "' and ss.CO_Date ='" & gU.dbEncode(dtRow("CO_DATE")) & "' and ss.CO_STATUS='NEW' and ss.STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "' and ss.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "'"
                    'End If
                    coHdrGrp = gDB.getDataTable(SQLString, gConn, transaction)
                    If coHdrGrp IsNot Nothing AndAlso coHdrGrp.Rows.Count > 0 Then
                        coHdrGrp.DefaultView.Sort = "ROUTE_ID asc"
                        coHdrGrp = coHdrGrp.DefaultView.ToTable()

                        'If STORER_TYPE = "NEW_NEW" Then
                        '    coHdrGrp.DefaultView.Sort = "CO_INV_NO asc" 'Sort with SO number
                        '    coHdrGrp = coHdrGrp.DefaultView.ToTable()
                        'End If

                        For Each dtGrpRow As DataRow In coHdrGrp.Rows
                            nextNo = DB.getDocNo("Do", gConn, transaction)
                            SQLString = "INSERT INTO WMS_DELV_ORDER(IMP_CODE ,STORER_CODE ,DO_CODE,ROUTE_ID ,DO_STATUS ,DO_ISSUED_BY ,DO_CO_CODE ,DO_CUS_REF_NO ,DO_DATE ,CUS_CODE ,CUS_NAME ,DO_POSTED_DATE ,DO_POSTED_BY ,DO_TYPE ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
                            SQLString += "VALUES("
                            SQLString += "'" & gU.dbEncode(IMP_CODE.Trim) & "',"
                            SQLString += "'" & gU.dbEncode(STORER_CODE.ToString.Trim) & "',"
                            SQLString += "'" & nextNo & "',"
                            SQLString += "'" & gU.dbEncode(dtGrpRow("ROUTE_ID").ToString.Trim) & "',"
                            SQLString += "'DRAFT',"
                            SQLString += "'" & Session("usr_id") & "',"
                            SQLString += "'" & gU.dbEncode(dtGrpRow("CO_CODES").ToString.Trim) & "',"
                            SQLString += "'NA',"
                            SQLString += "'" & dtRow("CO_DATE").ToString.Trim & "',"
                            SQLString += "'" & gU.dbEncode(dtGrpRow("CUS_CODES").ToString.Trim) & "',"
                            SQLString += "N'" & gU.dbEncode(dtGrpRow("CUS_NAMES").ToString.Trim) & "'," '
                            SQLString += "GETDATE(),"
                            SQLString += "'" & Session("usr_id") & "',"
                            SQLString += "'NA',"
                            SQLString += "'" & Session("usr_id") & "',"
                            SQLString += "GETDATE(),"
                            SQLString += "GETDATE(),"
                            SQLString += "'" & Session("usr_id") & "'"
                            SQLString += ")"
                            gDB.amendData(SQLString, gConn, transaction)


                            SQLString = "Select isnull(COD_PCS_UOM,0) COD_PCS_UOM,isnull(COD_TOTPCS,0) COD_TOTPCS,isnull(COD_TOT_WGT,0) COD_TOT_WGT, " &
" isnull(COD_TOT_CBM,0) COD_TOT_CBM,(case when LEN(COD_PALLET_NO)<8 then '19000101' else COD_PALLET_NO end) COD_PALLET_NO,a.*,b.ROUTE_ID,IsNUll(a.COD_WH_CODE,'') as COD_WH,IsNull(c.WH_TYPE,'')WH_TYPE,iSnuLL(b.CUS_CODE,'')CUS_CODE,iSnULL(b.CUS_NAME,'')CUS_NAME from WMS_CUST_ORDER_D a Inner join WMS_CUST_ORDER b On a.CO_CODE = b.CO_CODE and a.STORER_CODE = b.STORER_CODE Inner join WMS_WAREHOUSE c On a.COD_WH_CODE = c.WH_CODE and a.IMP_CODE=c.IMP_CODE " &
                            "WHERE  " &
                            "b.STORER_CODE = '" & gU.dbEncode(STORER_CODE) & "'" &
                            "and b.IMP_CODE = '" & gU.dbEncode(IMP_CODE) & "' and a.CO_CODE in (" & gU.dbEncode(dtGrpRow("CO_CODES").ToString.Trim) & ")  Order by b.ROUTE_ID"
                            coDtl = gDB.getDataTable(SQLString, gConn, transaction)
                            If coDtl IsNot Nothing AndAlso coDtl.Rows.Count > 0 Then


                                coDtl.Columns.Add("IsRule")
                                coDtl.Columns.Add("MIN_SELF_LIFE")
                                coDtl.Columns.Add("MIN_PROD_LIFE")
                                coDtl.Columns.Add("MAX_BATCHES")
                                coDtl.Columns.Add("MPL_FLAG")
                                coDtl.Columns.Add("MB_FLAG")
                                coDtl.Columns.Add("MSL_FLAG")
                                coDtl.Columns.Add("LOTS_CANNOT_BE_EARLIER")

                                'Dim dtView As New DataView(codt)
                                'Dim dtTemp As DataTable = dtView.ToTable(True, "CUS_CODE", "COD_ITM_CODE")
                                Dim dtWPickNORULE As New DataTable
                                Dim dtWPick1RULE As New DataTable
                                Dim dtWPickRULE As New DataTable
                                Dim grpQTY As Double = 0
                                Dim grpCus As String = ""
                                Dim custCodes As New List(Of String)
                                Dim custNames As New List(Of String)
                                Dim coCodes As New List(Of String)
                                Dim prevItem As String = coDtl.Rows(0)("COD_ITM_CODE").ToString()
                                Dim RowList As New List(Of DataRow)
                                dtWPickNORULE = coDtl.Clone()
                                dtWPick1RULE = coDtl.Clone()
                                dtWPickRULE = coDtl.Clone()
                                dtWPickNORULE.Rows.Clear()
                                dtWPick1RULE.Rows.Clear()
                                dtWPickRULE.Rows.Clear()

                                For Each row As DataRow In coDtl.Rows
                                    'Dim LAST_LOT As String = getLastLot(row("COD_ITM_CODE").ToString(), STORER_CODE, row("CUS_CODE").ToString())
                                    'row("COD_PALLET_NO") = LAST_LOT

                                    'Dim query As String = "Update WMS_CUST_ORDER_D Set COD_PALLET_NO='" + LAST_LOT + "' Where IMP_CODE='" + row("IMP_CODE").ToString() + "' and  STORER_CODE='" + row("STORER_CODE").ToString() + "'  and CO_CODE='" + row("CO_CODE").ToString() + "' and COD_SEQ='" + row("COD_SEQ").ToString() + "'"
                                    'gDB.amendData(query)

                                    Dim RuleRow = IsRule(gConn, transaction, IMP_CODE, STORER_CODE, row("CUS_CODE").ToString(), row("COD_ITM_CODE").ToString())
                                    If RuleRow IsNot Nothing AndAlso ruleWareList.Contains(row("WH_TYPE").ToString()) = True Then
                                        row("IsRule") = RuleRow("RuleCount")
                                        row("MIN_SELF_LIFE") = RuleRow("MIN_SELF_LIFE")
                                        row("MIN_PROD_LIFE") = RuleRow("MIN_PROD_LIFE")
                                        row("MAX_BATCHES") = RuleRow("MAX_BATCHES")

                                        row("MPL_FLAG") = RuleRow("MPL_FLAG")
                                        row("MSL_FLAG") = RuleRow("MSL_FLAG")
                                        row("MB_FLAG") = RuleRow("MB_FLAG")
                                        row("LOTS_CANNOT_BE_EARLIER") = RuleRow("LOTS_CANNOT_BE_EARLIER")

                                    ElseIf RuleRow Is Nothing OrElse ruleWareList.Contains(row("WH_TYPE").ToString()) = False Then
                                        row("IsRule") = "0"
                                        row("MIN_SELF_LIFE") = "0"
                                        row("MIN_PROD_LIFE") = "0"
                                        row("MAX_BATCHES") = "0"
                                        row("MPL_FLAG") = "0"
                                        row("MSL_FLAG") = "0"
                                        row("MB_FLAG") = "0"
                                        row("LOTS_CANNOT_BE_EARLIER") = "0"
                                    End If
                                Next
                                'Dim tmpTable = coDtl.DefaultView.ToTable(True, "CUS_CODE", "COD_ITM_CODE")

                                'No Rule
                                Dim tmpTable = If(coDtl.Select("IsRule = 0 And (WH_TYPE ='FG' OR WH_TYPE ='VIP')").Length > 0, coDtl.Select("IsRule = 0 And (WH_TYPE ='FG' OR WH_TYPE ='VIP')").CopyToDataTable(), New DataTable())
                                If tmpTable.Rows.Count > 0 Then
                                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE", "COD_WH")
                                    For Each row As DataRow In tmpTableDistinct.Rows
                                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' AND COD_WH='" + row("COD_WH").ToString() + "' ")
                                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
                                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' AND COD_WH='" + row("COD_WH").ToString() + "' ")
                                        If tmpArra.Count > 0 Then
                                            For Each item As DataRow In tmpArra
                                                custCodes.Add(item("CUS_CODE"))
                                                custNames.Add(item("CUS_NAME"))
                                                coCodes.Add(item("CO_CODE").ToString().Trim)
                                            Next
                                            Dim rgRow = coDtl.NewRow()
                                            'rgRow.ItemArray = row.ItemArray
                                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
                                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
                                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
                                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
                                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()
                                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
                                            rgRow("COD_PCS_UOM") = If(tmpArra(0)("COD_PCS_UOM") IsNot DBNull.Value, tmpArra(0)("COD_PCS_UOM").ToString(), "0")
                                            rgRow("COD_TOTPCS") = If(tmpArra(0)("COD_TOTPCS") IsNot DBNull.Value, tmpArra(0)("COD_TOTPCS").ToString(), "0")
                                            rgRow("COD_TOT_WGT") = If(tmpArra(0)("COD_TOT_WGT") IsNot DBNull.Value, tmpArra(0)("COD_TOT_WGT").ToString(), "0")
                                            rgRow("COD_TOT_CBM") = If(tmpArra(0)("COD_TOT_CBM") IsNot DBNull.Value, tmpArra(0)("COD_TOT_CBM").ToString(), "0")
                                            rgRow("COD_WH") = tmpArra(0)("COD_WH").ToString()
                                            rgRow("COD_QTY") = grpQTY
                                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
                                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())
                                            rgRow("CUS_NAME") = String.Join(",", custNames.Distinct().ToList())
                                            rgRow("CO_CODE") = String.Join(",", coCodes.Distinct().ToList())
                                            dtWPickNORULE.Rows.Add(rgRow.ItemArray)
                                            RowList.Add(row)
                                        End If


                                    Next
                                End If
                                tmpTable = If(coDtl.Select("IsRule = 0 And WH_TYPE <> 'FG' AND WH_TYPE <> 'VIP'").Length > 0, coDtl.Select("IsRule = 0 And WH_TYPE <> 'FG' AND WH_TYPE <> 'VIP'").CopyToDataTable(), New DataTable())
                                If tmpTable.Rows.Count > 0 Then
                                    tmpTable.DefaultView.Sort = "COD_WH"
                                    tmpTable = tmpTable.DefaultView.ToTable()
                                    For Each row As DataRow In tmpTable.Rows
                                        Dim rgRow = coDtl.NewRow()
                                        rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
                                        rgRow("COD_PALLET_NO") = row("COD_PALLET_NO").ToString()
                                        rgRow("COD_CARTON_NO") = row("COD_CARTON_NO").ToString()
                                        rgRow("COD_PACK_KEY") = row("COD_PACK_KEY").ToString()
                                        rgRow("COD_ITM_DESC") = row("COD_ITM_DESC").ToString()
                                        rgRow("COD_UOM") = row("COD_UOM").ToString()
                                        rgRow("COD_PCS_UOM") = If(row("COD_PCS_UOM") IsNot DBNull.Value, row("COD_PCS_UOM").ToString(), "0")
                                        rgRow("COD_TOTPCS") = If(row("COD_TOTPCS") IsNot DBNull.Value, row("COD_TOTPCS").ToString(), "0")
                                        rgRow("COD_TOT_WGT") = If(row("COD_TOT_WGT") IsNot DBNull.Value, row("COD_TOT_WGT").ToString(), "0")
                                        rgRow("COD_TOT_CBM") = If(row("COD_TOT_CBM") IsNot DBNull.Value, row("COD_TOT_CBM").ToString(), "0")
                                        rgRow("COD_QTY") = row("COD_QTY").ToString()
                                        rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
                                        rgRow("CUS_CODE") = row("CUS_CODE").ToString()
                                        rgRow("CUS_NAME") = row("CUS_NAME").ToString()
                                        rgRow("CO_CODE") = row("CO_CODE").ToString().Trim
                                        rgRow("COD_WH") = row("COD_WH").ToString()
                                        dtWPickNORULE.Rows.Add(rgRow.ItemArray)
                                        RowList.Add(row)
                                    Next
                                End If
                                'No Rule
                                tmpTable = If(coDtl.Select("IsRule = 1 And MSL_FLAG = 1 AND MPL_FLAG = 0 AND LOTS_CANNOT_BE_EARLIER = 0").Length > 0, coDtl.Select("IsRule = 1 And MSL_FLAG = 1 AND MPL_FLAG = 0 AND LOTS_CANNOT_BE_EARLIER = 0").CopyToDataTable(), New DataTable())
                                If tmpTable.Rows.Count > 0 Then
                                    tmpTable.DefaultView.Sort = "MIN_SELF_LIFE"
                                    tmpTable = tmpTable.DefaultView.ToTable()
                                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE", "MIN_SELF_LIFE", "COD_WH")
                                    For Each row As DataRow In tmpTableDistinct.Rows

                                        custCodes.Clear()
                                        custNames.Clear()
                                        coCodes.Clear()
                                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And MIN_SELF_LIFE='" + row("MIN_SELF_LIFE").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "'")
                                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
                                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And MIN_SELF_LIFE='" + row("MIN_SELF_LIFE").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "'")
                                        If tmpArra.Count > 0 Then
                                            For Each item As DataRow In tmpArra
                                                custCodes.Add(item("CUS_CODE"))
                                                custNames.Add(item("CUS_NAME"))
                                                coCodes.Add(item("CO_CODE").ToString().Trim)
                                            Next
                                            Dim rgRow = coDtl.NewRow()
                                            'rgRow.ItemArray = row.ItemArray
                                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
                                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
                                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
                                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
                                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()

                                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
                                            rgRow("COD_PCS_UOM") = If(tmpArra(0)("COD_PCS_UOM") IsNot DBNull.Value, tmpArra(0)("COD_PCS_UOM").ToString(), "0")
                                            rgRow("COD_TOTPCS") = If(tmpArra(0)("COD_TOTPCS") IsNot DBNull.Value, tmpArra(0)("COD_TOTPCS").ToString(), "0")
                                            rgRow("COD_TOT_WGT") = If(tmpArra(0)("COD_TOT_WGT") IsNot DBNull.Value, tmpArra(0)("COD_TOT_WGT").ToString(), "0")
                                            rgRow("COD_TOT_CBM") = If(tmpArra(0)("COD_TOT_CBM") IsNot DBNull.Value, tmpArra(0)("COD_TOT_CBM").ToString(), "0")
                                            rgRow("COD_WH") = tmpArra(0)("COD_WH").ToString()
                                            rgRow("COD_QTY") = grpQTY
                                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
                                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())
                                            rgRow("CUS_NAME") = String.Join(",", custNames.Distinct().ToList())
                                            rgRow("CO_CODE") = String.Join(",", coCodes.Distinct().ToList())
                                            dtWPick1RULE.Rows.Add(rgRow.ItemArray)
                                            RowList.Add(row)

                                        End If


                                    Next

                                End If


                                tmpTable = If(coDtl.Select("IsRule = 1 And LOTS_CANNOT_BE_EARLIER = 1 And MSL_FLAG = 0 And MPL_FLAG = 0").Length > 0, coDtl.Select("IsRule = 1 And LOTS_CANNOT_BE_EARLIER = 1 And MSL_FLAG = 0 And MPL_FLAG = 0").CopyToDataTable(), New DataTable())

                                If tmpTable.Rows.Count > 0 Then
                                    tmpTable.DefaultView.Sort = "MIN_SELF_LIFE"
                                    tmpTable = tmpTable.DefaultView.ToTable()
                                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE", "COD_PALLET_NO", "COD_WH")
                                    For Each row As DataRow In tmpTableDistinct.Rows

                                        custCodes.Clear()
                                        custNames.Clear()
                                        coCodes.Clear()
                                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And COD_PALLET_NO='" + row("COD_PALLET_NO").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "'")
                                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
                                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And COD_PALLET_NO='" + row("COD_PALLET_NO").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "'")
                                        If tmpArra.Count > 0 Then
                                            For Each item As DataRow In tmpArra
                                                custCodes.Add(item("CUS_CODE"))
                                                custNames.Add(item("CUS_NAME"))
                                                coCodes.Add(item("CO_CODE").ToString().Trim)
                                            Next
                                            Dim rgRow = coDtl.NewRow()
                                            'rgRow.ItemArray = row.ItemArray
                                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
                                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
                                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
                                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
                                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()

                                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
                                            rgRow("COD_PCS_UOM") = If(tmpArra(0)("COD_PCS_UOM") IsNot DBNull.Value, tmpArra(0)("COD_PCS_UOM").ToString(), "0")
                                            rgRow("COD_TOTPCS") = If(tmpArra(0)("COD_TOTPCS") IsNot DBNull.Value, tmpArra(0)("COD_TOTPCS").ToString(), "0")
                                            rgRow("COD_TOT_WGT") = If(tmpArra(0)("COD_TOT_WGT") IsNot DBNull.Value, tmpArra(0)("COD_TOT_WGT").ToString(), "0")
                                            rgRow("COD_TOT_CBM") = If(tmpArra(0)("COD_TOT_CBM") IsNot DBNull.Value, tmpArra(0)("COD_TOT_CBM").ToString(), "0")
                                            rgRow("COD_WH") = tmpArra(0)("COD_WH").ToString()
                                            rgRow("COD_QTY") = grpQTY
                                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
                                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())
                                            rgRow("CUS_NAME") = String.Join(",", custNames.Distinct().ToList())
                                            rgRow("CO_CODE") = String.Join(",", coCodes.Distinct().ToList())
                                            dtWPick1RULE.Rows.Add(rgRow.ItemArray)
                                            RowList.Add(row)
                                        End If


                                    Next
                                End If
                                '
                                SeqNo = 1
                                Dim DOD_GRP_SEQ = 1
                                For Each rowD As DataRow In dtWPick1RULE.Rows
                                    Dim LOTNO As String = ""
                                    Dim LOTQTY As Double = 0
                                    Dim chkdList = custRuleChk(gConn, transaction, 1, rowD("COD_WH").ToString.Trim, nextNo, IMP_CODE, STORER_CODE, rowD("COD_ITM_CODE").ToString.Trim, rowD("COD_QTY"), rowD("ROUTE_ID").ToString.Trim, Convert.ToDateTime(dtRow("CO_DATE").ToString.Trim), rowD("CUS_CODE").ToString.Trim, rowD("COD_PALLET_NO").ToString.Trim)
                                    chkdList = chkdList.Where(Function(d) d("DOD_PK_QTY") <> "0").ToList()

                                    For Each itemDict In chkdList
                                        SQLString = ""
                                        SQLString = "INSERT INTO WMS_DELV_ORDER_D(IMP_CODE ,STORER_CODE ,DO_CODE,DOD_WH_CODE,DOD_LOC_WH,DOD_EXPIRY_DATE ,DOD_SEQ ,DOD_DISP_SEQ ,DOD_PALLET_NO ,DOD_CARTON_NO ,DOD_ITM_CODE ,DOD_PACK_KEY ,DOD_ITM_DESC ,DOD_PACK_NO ,DOD_PACK_TYPE ,DOD_QTY ,DOD_BATCH_NO,DOD_UOM ,DOD_PCS_UOM ,DOD_TOTPCS ,DOD_TOT_WGT ,DOD_TOT_CBM  ,DOD_DELI_QTY,DOD_PK_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT,DOD_CUS_CODE,DOD_CUS_NAME,DOD_CO_CODE,DOD_GRP_SEQ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
                                        SQLString += "VALUES("
                                        SQLString += "'" & gU.dbEncode(IMP_CODE.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(STORER_CODE.ToString.Trim) & "',"
                                        SQLString += "'" & nextNo & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_WH").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_LOC").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_EXPIRY_DATE").ToString.Trim) & "',"
                                        SQLString += "'" & (SeqNo) & "',"
                                        SQLString += "'" & (SeqNo) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PALLET_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_CARTON_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ITM_CODE").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PACK_KEY").ToString.Trim) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("COD_ITM_DESC").ToString.Trim) & "',"
                                        SQLString += "'0',"
                                        SQLString += "'0',"
                                        SQLString += "" & gU.dbEncode(itemDict("DOD_PK_QTY").ToString.Trim) & ","
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_BATCH_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_UOM").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PCS_UOM").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOTPCS").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_WGT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_CBM").ToString.Trim) & "',"
                                        SQLString += "0,"
                                        SQLString += "0,"
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MIN_PROD_DATE").ToString.Trim) & "',"
                                        SQLString += "'" & If(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim = "0", gU.dbEncode(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim), (Convert.ToInt32(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim)).ToString) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_LAST_LOT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MAX_LOT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("CUS_CODE").ToString.Trim) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("CUS_NAME").ToString.Trim()) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("CO_CODE").ToString.Trim()) & "',"
                                        SQLString += "N'" & gU.dbEncode(DOD_GRP_SEQ.ToString.Trim()) & "',"
                                        SQLString += "'" & Session("usr_id") & "',"
                                        SQLString += "GETDATE(),"
                                        SQLString += "GETDATE(),"
                                        SQLString += "'" & Session("usr_id") & "'"
                                        SQLString += ")"
                                        SeqNo = SeqNo + 1
                                        gDB.amendData(SQLString, gConn, transaction)
                                    Next
                                    DOD_GRP_SEQ = DOD_GRP_SEQ + 1
                                Next
                                tmpTable.Clear()
                                tmpTable = If(coDtl.Select("IsRule > 1").Length > 0, coDtl.Select("IsRule > 1").CopyToDataTable(), New DataTable())

                                If tmpTable.Rows.Count > 0 Then
                                    tmpTable.DefaultView.Sort = "MIN_SELF_LIFE"
                                    tmpTable = tmpTable.DefaultView.ToTable()
                                    'For Each row As DataRow In tmpTable.Rows
                                    Dim NewTempDt As DataTable
                                    NewTempDt = coDtl.Clone()
                                    NewTempDt.Rows.Clear()
                                    Dim tmpTableDistinct = tmpTable.DefaultView.ToTable(True, "COD_ITM_CODE", "COD_PALLET_NO", "COD_WH", "MIN_SELF_LIFE", "MIN_PROD_LIFE", "MAX_BATCHES", "MPL_FLAG", "MB_FLAG", "LOTS_CANNOT_BE_EARLIER")
                                    For Each row As DataRow In tmpTableDistinct.Rows

                                        custCodes.Clear()
                                        custNames.Clear()
                                        coCodes.Clear()
                                        Dim rowQty = tmpTable.Compute("Sum(COD_QTY)", "COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And COD_PALLET_NO='" + row("COD_PALLET_NO").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "' And MIN_SELF_LIFE ='" + row("MIN_SELF_LIFE").ToString() + "' And MAX_BATCHES ='" + row("MAX_BATCHES").ToString() + "'  And MPL_FLAG ='" + row("MPL_FLAG").ToString() + "' And MB_FLAG ='" + row("MB_FLAG").ToString() + "' And LOTS_CANNOT_BE_EARLIER ='" + row("LOTS_CANNOT_BE_EARLIER").ToString() + "' ")
                                        grpQTY = If(rowQty IsNot DBNull.Value, rowQty, 0)
                                        Dim tmpArra() As DataRow = tmpTable.Select("COD_ITM_CODE='" + row("COD_ITM_CODE").ToString() + "' And COD_PALLET_NO='" + row("COD_PALLET_NO").ToString() + "' And COD_WH='" + row("COD_WH").ToString() + "' And MIN_SELF_LIFE ='" + row("MIN_SELF_LIFE").ToString() + "' And MAX_BATCHES ='" + row("MAX_BATCHES").ToString() + "' And MPL_FLAG ='" + row("MPL_FLAG").ToString() + "' And MB_FLAG ='" + row("MB_FLAG").ToString() + "' And LOTS_CANNOT_BE_EARLIER ='" + row("LOTS_CANNOT_BE_EARLIER").ToString() + "'")
                                        If tmpArra.Count > 0 Then
                                            For Each item As DataRow In tmpArra
                                                custCodes.Add(item("CUS_CODE"))
                                                custNames.Add(item("CUS_NAME"))
                                                coCodes.Add(item("CO_CODE").ToString().Trim)
                                            Next
                                            Dim rgRow = coDtl.NewRow()
                                            'rgRow.ItemArray = row.ItemArray
                                            rgRow("COD_ITM_CODE") = row("COD_ITM_CODE").ToString()
                                            rgRow("COD_PALLET_NO") = tmpArra(0)("COD_PALLET_NO").ToString()
                                            rgRow("COD_CARTON_NO") = tmpArra(0)("COD_CARTON_NO").ToString()
                                            rgRow("COD_PACK_KEY") = tmpArra(0)("COD_PACK_KEY").ToString()
                                            rgRow("COD_ITM_DESC") = tmpArra(0)("COD_ITM_DESC").ToString()

                                            rgRow("COD_UOM") = tmpArra(0)("COD_UOM").ToString()
                                            rgRow("COD_PCS_UOM") = If(tmpArra(0)("COD_PCS_UOM") IsNot DBNull.Value, tmpArra(0)("COD_PCS_UOM").ToString(), "0")
                                            rgRow("COD_TOTPCS") = If(tmpArra(0)("COD_TOTPCS") IsNot DBNull.Value, tmpArra(0)("COD_TOTPCS").ToString(), "0")
                                            rgRow("COD_TOT_WGT") = If(tmpArra(0)("COD_TOT_WGT") IsNot DBNull.Value, tmpArra(0)("COD_TOT_WGT").ToString(), "0")
                                            rgRow("COD_TOT_CBM") = If(tmpArra(0)("COD_TOT_CBM") IsNot DBNull.Value, tmpArra(0)("COD_TOT_CBM").ToString(), "0")
                                            rgRow("COD_WH") = tmpArra(0)("COD_WH").ToString()
                                            rgRow("COD_QTY") = grpQTY
                                            rgRow("ROUTE_ID") = dtGrpRow("ROUTE_ID").ToString.Trim
                                            rgRow("CUS_CODE") = String.Join(",", custCodes.Distinct().ToList())
                                            rgRow("CUS_NAME") = String.Join(",", custNames.Distinct().ToList())
                                            rgRow("CO_CODE") = String.Join(",", coCodes.Distinct().ToList())
                                            NewTempDt.Rows.Add(rgRow.ItemArray)
                                            RowList.Add(row)
                                        End If
                                    Next

                                    For Each rowD As DataRow In NewTempDt.Rows
                                        Dim LOTNO As String = ""
                                        Dim LOTQTY As Double = 0

                                        Dim chkdList = custRuleChk(gConn, transaction, 1, rowD("COD_WH").ToString.Trim, nextNo, IMP_CODE, STORER_CODE, rowD("COD_ITM_CODE").ToString.Trim, rowD("COD_QTY"), rowD("ROUTE_ID").ToString.Trim, Convert.ToDateTime(dtRow("CO_DATE").ToString.Trim), rowD("CUS_CODE").ToString.Trim, rowD("COD_PALLET_NO").ToString.Trim)
                                        chkdList = chkdList.Where(Function(d) d("DOD_PK_QTY") <> "0").ToList()

                                        For Each itemDict In chkdList
                                            SQLString = ""

                                            SQLString = "INSERT INTO WMS_DELV_ORDER_D(IMP_CODE ,STORER_CODE ,DO_CODE,DOD_WH_CODE,DOD_LOC_WH,DOD_EXPIRY_DATE ,DOD_SEQ ,DOD_DISP_SEQ ,DOD_PALLET_NO ,DOD_CARTON_NO ,DOD_ITM_CODE ,DOD_PACK_KEY ,DOD_ITM_DESC ,DOD_PACK_NO ,DOD_PACK_TYPE ,DOD_QTY ,DOD_BATCH_NO,DOD_UOM ,DOD_PCS_UOM ,DOD_TOTPCS ,DOD_TOT_WGT ,DOD_TOT_CBM  ,DOD_DELI_QTY,DOD_PK_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT,DOD_CUS_CODE,DOD_CUS_NAME,DOD_CO_CODE,DOD_GRP_SEQ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
                                            SQLString += "VALUES("
                                            SQLString += "'" & gU.dbEncode(IMP_CODE.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(STORER_CODE.ToString.Trim) & "',"
                                            SQLString += "'" & nextNo & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("COD_WH").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(itemDict("ILOC_LOC").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(itemDict("ILOC_EXPIRY_DATE").ToString.Trim) & "',"
                                            SQLString += "'" & (SeqNo) & "',"
                                            SQLString += "'" & (SeqNo) & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("COD_PALLET_NO").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("COD_CARTON_NO").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(itemDict("ITM_CODE").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("COD_PACK_KEY").ToString.Trim) & "',"
                                            SQLString += "N'" & gU.dbEncode(rowD("COD_ITM_DESC").ToString.Trim) & "',"
                                            SQLString += "'0',"
                                            SQLString += "'0',"
                                            SQLString += "" & gU.dbEncode(itemDict("DOD_PK_QTY").ToString.Trim) & ","
                                            SQLString += "'" & gU.dbEncode(itemDict("ILOC_BATCH_NO").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("COD_UOM").ToString.Trim) & "',"
                                            SQLString += "" & gU.dbEncode(rowD("COD_PCS_UOM").ToString.Trim) & ","
                                            SQLString += "" & gU.dbEncode(rowD("COD_TOTPCS").ToString.Trim) & ","
                                            SQLString += "" & gU.dbEncode(rowD("COD_TOT_WGT").ToString.Trim) & ","
                                            SQLString += "" & gU.dbEncode(rowD("COD_TOT_CBM").ToString.Trim) & ","
                                            SQLString += "0,"
                                            SQLString += "0,"
                                            SQLString += "" & gU.dbEncode(itemDict("DOD_MIN_PROD_DATE").ToString.Trim) & ","
                                            SQLString += "'" & If(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim = "0", gU.dbEncode(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim), (Convert.ToInt32(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim)).ToString) & "',"
                                            SQLString += "'" & gU.dbEncode(itemDict("DOD_LAST_LOT").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(itemDict("DOD_MAX_LOT").ToString.Trim) & "',"
                                            SQLString += "'" & gU.dbEncode(rowD("CUS_CODE").ToString.Trim) & "',"
                                            SQLString += "N'" & gU.dbEncode(rowD("CUS_NAME").ToString.Trim()) & "',"
                                            SQLString += "N'" & gU.dbEncode(rowD("CO_CODE").ToString.Trim()) & "',"
                                            SQLString += "N'" & gU.dbEncode(DOD_GRP_SEQ.ToString.Trim()) & "',"
                                            SQLString += "'" & Session("usr_id") & "',"
                                            SQLString += "GETDATE(),"
                                            SQLString += "GETDATE(),"
                                            SQLString += "'" & Session("usr_id") & "'"  'dictParam.Add("ILOC_LOC", rowStk("ILOC_LOC"))
                                            SQLString += ")"
                                            SeqNo = SeqNo + 1
                                            gDB.amendData(SQLString, gConn, transaction)
                                        Next
                                        DOD_GRP_SEQ = DOD_GRP_SEQ + 1
                                    Next

                                    'Next


                                End If

                                For Each rowD As DataRow In dtWPickNORULE.Rows
                                    Dim LOTNO As String = ""
                                    Dim LOTQTY As Double = 0

                                    Dim chkdList = custRuleChk(gConn, transaction, 0, rowD("COD_WH").ToString.Trim, nextNo, IMP_CODE, STORER_CODE, rowD("COD_ITM_CODE").ToString.Trim, rowD("COD_QTY"), rowD("ROUTE_ID").ToString.Trim, Convert.ToDateTime(dtRow("CO_DATE").ToString.Trim), rowD("CUS_CODE").ToString.Trim, rowD("COD_PALLET_NO").ToString.Trim)
                                    chkdList = chkdList.Where(Function(d) d("DOD_PK_QTY") <> "0").ToList()


                                    For Each itemDict In chkdList
                                        SQLString = ""

                                        SQLString = "INSERT INTO WMS_DELV_ORDER_D(IMP_CODE ,STORER_CODE ,DO_CODE,DOD_WH_CODE,DOD_LOC_WH,DOD_EXPIRY_DATE ,DOD_SEQ ,DOD_DISP_SEQ ,DOD_PALLET_NO ,DOD_CARTON_NO ,DOD_ITM_CODE ,DOD_PACK_KEY ,DOD_ITM_DESC ,DOD_PACK_NO ,DOD_PACK_TYPE ,DOD_QTY ,DOD_BATCH_NO,DOD_UOM ,DOD_PCS_UOM ,DOD_TOTPCS ,DOD_TOT_WGT ,DOD_TOT_CBM  ,DOD_DELI_QTY,DOD_PK_QTY,DOD_MIN_PROD_DATE,DOD_MIN_SHELF_LIFE,DOD_LAST_LOT,DOD_MAX_LOT,DOD_CUS_CODE,DOD_CUS_NAME,DOD_CO_CODE,DOD_GRP_SEQ ,SYS_LUB ,SYS_LUD ,SYS_CD ,SYS_CB)"
                                        SQLString += "VALUES("
                                        SQLString += "'" & gU.dbEncode(IMP_CODE.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(STORER_CODE.ToString.Trim) & "',"
                                        SQLString += "'" & nextNo & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_WH").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_LOC").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_EXPIRY_DATE").ToString.Trim) & "',"
                                        SQLString += "'" & (SeqNo) & "',"
                                        SQLString += "'" & (SeqNo) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PALLET_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_CARTON_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("ITM_CODE").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PACK_KEY").ToString.Trim) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("COD_ITM_DESC").ToString.Trim) & "',"
                                        SQLString += "'0',"
                                        SQLString += "'0',"
                                        SQLString += "" & gU.dbEncode(itemDict("DOD_PK_QTY").ToString.Trim) & ","
                                        SQLString += "'" & gU.dbEncode(itemDict("ILOC_BATCH_NO").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_UOM").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_PCS_UOM").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOTPCS").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_WGT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("COD_TOT_CBM").ToString.Trim) & "',"
                                        SQLString += "0,"
                                        SQLString += "0,"
                                        SQLString += "" & gU.dbEncode(itemDict("DOD_MIN_PROD_DATE").ToString.Trim) & ","
                                        SQLString += "'" & If(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim = "0", gU.dbEncode(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim), (Convert.ToInt32(itemDict("DOD_MIN_SHELF_LIFE").ToString.Trim)).ToString) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_LAST_LOT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(itemDict("DOD_MAX_LOT").ToString.Trim) & "',"
                                        SQLString += "'" & gU.dbEncode(rowD("CUS_CODE").ToString.Trim) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("CUS_NAME").ToString.Trim()) & "',"
                                        SQLString += "N'" & gU.dbEncode(rowD("CO_CODE").ToString.Trim()) & "',"
                                        SQLString += "N'" & gU.dbEncode(DOD_GRP_SEQ.ToString.Trim()) & "',"
                                        SQLString += "'" & Session("usr_id") & "',"
                                        SQLString += "GETDATE(),"
                                        SQLString += "GETDATE(),"
                                        SQLString += "'" & Session("usr_id") & "'"
                                        SQLString += ")"
                                        SeqNo = SeqNo + 1
                                        gDB.amendData(SQLString, gConn, transaction)
                                    Next
                                    DOD_GRP_SEQ = DOD_GRP_SEQ + 1
                                Next

                                SQLString = "Select DOD_CUS_CODE,DOD_CUS_NAME,DOD_CO_CODE from WMS_DELV_ORDER_D " &
                                "WHERE  " &
                                "DO_CODE = '" & gU.dbEncode(nextNo) & "'"
                                Dim dtHdr = gDB.getDataTable(SQLString, gConn, transaction)
                                Dim lstName As New List(Of String)
                                Dim lstCode As New List(Of String)
                                Dim lstcoCode As New List(Of String)
                                For Each row As DataRow In dtHdr.Rows
                                    Dim Names = row("DOD_CUS_NAME").ToString().Split(",")
                                    Dim Codes = row("DOD_CUS_CODE").ToString().Split(",")
                                    Dim co_Codes = row("DOD_CO_CODE").ToString().Split(",")
                                    For Each nm As String In Names
                                        lstName.Add(nm)
                                    Next
                                    For Each cd As String In Codes
                                        lstCode.Add(cd)
                                    Next
                                    For Each cocd As String In co_Codes
                                        lstcoCode.Add(cocd.ToString().Trim)
                                    Next

                                Next
                                SQLString = "Update WMS_DELV_ORDER Set CUS_CODE='" + String.Join(",", lstCode.Distinct().ToList()) + "',CUS_NAME=N'" & gU.dbEncode(String.Join(", ", lstName.Distinct().ToList())) & "',DO_CO_CODE=N'" + String.Join(", ", lstcoCode.Distinct().ToList()) + "' WHERE DO_CODE = '" & gU.dbEncode(nextNo) & "'"
                                gDB.amendData(SQLString, gConn, transaction)
                                If lstcoCode IsNot Nothing AndAlso lstcoCode.Count > 0 Then
                                    updateSQL = "Update WMS_CUST_ORDER set CO_STATUS='PICKED'" &
                                        " Where CO_STATUS='NEW' and CO_CODE in (" + String.Join(",", lstcoCode.Distinct().ToList()) + ") "
                                    gDB.amendData(updateSQL, gConn, transaction)
                                End If

                                OrderCount = OrderCount + 1
                            End If
                        Next
                    End If

                Next

                'transaction.Commit()
                uiFun.displayMsg(Me, "", OrderCount & " orders has been picked successfully!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "No orders to pick !", Session("gLang"))
            End If
        Catch ex As Exception
            'transaction.Rollback()
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
            ltlMessage.Text += "<p style='color:red'>" + lstError.Count.ToString + " records skipped </p>"
            For Each msg In lstError
                ltlMessage.Text += "<p style='color:red'>" + msg + " </p>"
                SaveErrorLog(msg)
            Next
            BindDate()
            BindCODate()

        End Try
    End Sub

    'For unassign orders
    Protected Sub btnUnAssign_Click(sender As Object, e As System.EventArgs) Handles btnUnAssign.Click
        Dim updateSQL As String = ""
        Dim gconn As SqlConnection
        Dim transaction As SqlTransaction
        Dim orderCount = 0
        gconn = gDB.getConnection()
        transaction = gconn.BeginTransaction()
        ltlMessage.Text = ""
        Dim codtl As DataTable = Nothing
        Try
            Dim DO_DATES = ""
            Dim ROUTE_IDS = ""
            For Each item As ListItem In lboxDate.Items
                If item.Selected = True Then
                    If DO_DATES = "" Then
                        DO_DATES = "'" + item.Value + "'"
                    Else
                        DO_DATES = DO_DATES + ",'" + item.Value + "'"
                    End If

                End If
            Next

            For Each item As ListItem In lboxRoute.Items
                If item.Selected = True Then
                    If ROUTE_IDS = "" Then
                        ROUTE_IDS = "'" + item.Value + "'"
                    Else
                        ROUTE_IDS = ROUTE_IDS + ",'" + item.Value + "'"
                    End If

                End If
            Next

            Dim SQLstring = "SELECT imp_code,storer_code,DO_CODE FROM WMS_DELV_ORDER WHERE DO_STATUS != 'POSTED' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "'"
            'Dim SQLstring = "SELECT imp_code,storer_code,DO_CODE FROM WMS_DELV_ORDER WHERE DO_STATUS='DRAFT' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "'"
            If ROUTE_IDS <> "" AndAlso DO_DATES <> "" Then
                SQLstring += " and (DO_DATE in ( " + DO_DATES + ") And ROUTE_ID in (" + ROUTE_IDS + "))"
                codtl = gDB.getDataTable(SQLstring, gconn, transaction)
                'ElseIf ROUTE_IDS <> "" Then
                '    SQLstring += " and (ROUTE_ID in (" + ROUTE_IDS + "))"
                '    codtl = gDB.getDataTable(SQLstring, gconn, transaction)
                'ElseIf DO_DATES <> "" Then
                '    SQLstring += " and (DO_DATE in (" + DO_DATES + "))"
                '    codtl = gDB.getDataTable(SQLstring, gconn, transaction)
            Else
                SQLstring = ""
                codtl = Nothing
            End If

            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then
                For Each rowD As DataRow In codtl.Rows
                    updateSQL = "Update WMS_DELV_ORDER_D set DOD_ASN_BY=NULL, DOD_STATUS=NULL, DOD_ASN_DT=NULL Where DO_CODE = ('" + rowD("DO_CODE") + "') And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "'"
                    gDB.amendData(updateSQL, gconn, transaction)
                    orderCount = orderCount + 1
                Next
                transaction.Commit()
                uiFun.displayMsg(Me, "", "The order has been unassign!!", Session("gLang"))
                ltlMessage.Text += "<p style='color:green'>" + orderCount.ToString + " delivery order(s) has been unassigned! </p>"
                'Else
                '    uiFun.displayMsg(Me, "", "No records found to unassigned!", Session("gLang"))
            End If


        Catch ex As Exception
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If

        Finally
            If gconn IsNot Nothing Then
                If gconn.State = ConnectionState.Open Then
                    gconn.Close()
                    gconn.Dispose()
                End If
            End If
            BindDate()
            BindCODate()
        End Try
    End Sub

    Protected Sub btnRelease_Click(sender As Object, e As System.EventArgs) Handles btnRelease.Click
        Dim updateSQL As String = ""
        Dim gconn As SqlConnection
        Dim transaction As SqlTransaction
        Dim orderCount = 0
        gconn = gDB.getConnection()
        transaction = gconn.BeginTransaction()
        ltlMessage.Text = ""
        Dim codtl As DataTable = Nothing
        Try
            Dim DO_DATES = ""
            Dim ROUTE_IDS = ""
            For Each item As ListItem In lboxDate.Items
                If item.Selected = True Then
                    If DO_DATES = "" Then
                        DO_DATES = "'" + item.Value + "'"
                    Else
                        DO_DATES = DO_DATES + ",'" + item.Value + "'"
                    End If

                End If
            Next
            For Each item As ListItem In lboxRoute.Items
                If item.Selected = True Then
                    If ROUTE_IDS = "" Then
                        ROUTE_IDS = "'" + item.Value + "'"
                    Else
                        ROUTE_IDS = ROUTE_IDS + ",'" + item.Value + "'"
                    End If

                End If
            Next
            Dim SQLstring = "SELECT imp_code,storer_code,DO_CODE FROM WMS_DELV_ORDER WHERE DO_STATUS='DRAFT' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "'"
            If ROUTE_IDS <> "" AndAlso DO_DATES <> "" Then
                SQLstring += " and (DO_DATE in ( " + DO_DATES + ") And ROUTE_ID in (" + ROUTE_IDS + "))"
                codtl = gDB.getDataTable(SQLstring, gconn, transaction)
                'ElseIf ROUTE_IDS <> "" Then
                '    SQLstring += " and (ROUTE_ID in (" + ROUTE_IDS + "))"
                '    codtl = gDB.getDataTable(SQLstring, gconn, transaction)
                'ElseIf DO_DATES <> "" Then
                '    SQLstring += " and (DO_DATE in (" + DO_DATES + "))"
                '    codtl = gDB.getDataTable(SQLstring, gconn, transaction)
            Else
                SQLstring = ""
                codtl = Nothing
            End If

            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then
                For Each rowD As DataRow In codtl.Rows
                    updateSQL = "Update WMS_DELV_ORDER set DO_STATUS='RELEASED' Where DO_CODE='" + rowD("DO_CODE") + "' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "' and (select sum(DOD_QTY) from WMS_DELV_ORDER_D where DO_CODE='" + rowD("DO_CODE") + "' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "')>0"
                    gDB.amendData(updateSQL, gconn, transaction)
                    orderCount = orderCount + 1
                Next
                transaction.Commit()
                uiFun.displayMsg(Me, "", "The order has been released!!", Session("gLang"))
                ltlMessage.Text += "<p style='color:green'>" + orderCount.ToString + " delivery order(s) released! </p>"
            Else
                uiFun.displayMsg(Me, "", "No records found to release!", Session("gLang"))
            End If


        Catch ex As Exception
            If Not transaction Is Nothing Then
                transaction.Rollback()
                transaction = Nothing
            End If



        Finally
            If gconn IsNot Nothing Then
                If gconn.State = ConnectionState.Open Then
                    gconn.Close()
                    gconn.Dispose()
                End If
            End If
            BindDate()
            BindCODate()

        End Try
    End Sub

    Protected Sub btnPickingDelete_Click(sender As Object, e As System.EventArgs) Handles btnPickingDelete.Click
        Dim codt As DataTable
        Dim codtl As DataTable
        Dim gconn As SqlConnection
        Dim transaction As SqlTransaction
        gconn = gDB.getConnection()
        transaction = gconn.BeginTransaction()
        Dim orderCount = 0
        Dim SQLstring As String = ""
        Dim javaStr = ""
        Try
            If hfDODelete.Value.ToString() = "Y" Then
                DeleteDO()
            Else
                ltlMessage.Text = ""
                Dim DO_DATES = ""
                Dim ROUTE_IDS = ""
                For Each item As ListItem In lboxDate.Items
                    If item.Selected = True Then
                        If DO_DATES = "" Then
                            DO_DATES = "'" + item.Value + "'"
                        Else
                            DO_DATES = DO_DATES + ",'" + item.Value + "'"
                        End If

                    End If
                Next
                For Each item As ListItem In lboxRoute.Items
                    If item.Selected = True Then
                        If ROUTE_IDS = "" Then
                            ROUTE_IDS = "'" + item.Value + "'"
                        Else
                            ROUTE_IDS = ROUTE_IDS + ",'" + item.Value + "'"
                        End If

                    End If
                Next
                SQLstring = "SELECT imp_code,storer_code,DO_CODE,do_status,IsNUll(DO_CO_CODE,'')DO_CO_CODE FROM WMS_DELV_ORDER WHERE DO_CODE <> '' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "'"
                '" WHERE DO_STATUS ='DRAFT'"
                Dim listReleased As New List(Of String)
                If ROUTE_IDS <> "" AndAlso DO_DATES <> "" Then
                    SQLstring += " and (DO_DATE in ( " + DO_DATES + ") And ROUTE_ID in (" + ROUTE_IDS + "))"
                    codt = gDB.getDataTable(SQLstring)
                    'ElseIf ROUTE_IDS <> "" Then
                    '    SQLstring += " and (ROUTE_ID in (" + ROUTE_IDS + "))"
                    '    codt = gDB.getDataTable(SQLstring)
                    'ElseIf DO_DATES <> "" Then
                    '    SQLstring += " and (DO_DATE in (" + DO_DATES + "))"
                    '    codt = gDB.getDataTable(SQLstring)
                Else
                    SQLstring = ""
                    codt = Nothing
                End If

                If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                    For Each row In codt.Rows
                        If row("do_status").ToString().ToUpper() = "RELEASED" Then
                            listReleased.Add(row("DO_CODE").ToString())
                        ElseIf row("do_status").ToString().ToUpper() = "POSTED" Then
                            listReleased.Add(row("DO_CODE").ToString())
                        ElseIf row("do_status").ToString().ToUpper() = "PICKED" Then
                            listReleased.Add(row("DO_CODE").ToString())
                        End If
                    Next

                End If
                If listReleased.Count > 0 Then
                    javaStr = "var x=confirm(""DO: [" & String.Join(",", listReleased.Distinct().ToList()) & "] has been released to PDA, are you sure to delete this Route, PDA corresponding Goods Issue will be deleted also."");if (x==true){var pass=prompt('Enter password');if(pass=='Passw0rd'){deleteDo();}else{alert('Wrong passcode!');}}"
                    ScriptManager.RegisterStartupScript(Me, Me.GetType, "CONFIRM_WARN", javaStr, True)
                Else
                    DeleteDO()
                End If
            End If
        Catch ex As Exception
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))
        End Try
    End Sub

    Private Sub DeleteDO()
        Dim updateSQL As String = ""
        Dim SQLstring As String = ""
        Dim codt As DataTable
        Dim codtl As DataTable
        Dim gconn As SqlConnection
        Dim transaction As SqlTransaction
        gconn = gDB.getConnection()
        transaction = gconn.BeginTransaction()
        Dim orderCount = 0

        Try
            hfDODelete.Value = "N"
            ltlMessage.Text = ""
            Dim DO_DATES = ""
            Dim ROUTE_IDS = ""
            For Each item As ListItem In lboxDate.Items
                If item.Selected = True Then
                    If DO_DATES = "" Then
                        DO_DATES = "'" + item.Value + "'"
                    Else
                        DO_DATES = DO_DATES + ",'" + item.Value + "'"
                    End If

                End If
            Next
            For Each item As ListItem In lboxRoute.Items
                If item.Selected = True Then
                    If ROUTE_IDS = "" Then
                        ROUTE_IDS = "'" + item.Value + "'"
                    Else
                        ROUTE_IDS = ROUTE_IDS + ",'" + item.Value + "'"
                    End If

                End If
            Next
            SQLstring = "SELECT imp_code,storer_code,DO_CODE,IsNUll(DO_CO_CODE,'')DO_CO_CODE FROM WMS_DELV_ORDER WHERE DO_CODE <> '' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "'"
            '" WHERE DO_STATUS ='DRAFT'"

            If ROUTE_IDS <> "" AndAlso DO_DATES <> "" Then
                SQLstring += " and (DO_DATE in ( " + DO_DATES + ") And ROUTE_ID in (" + ROUTE_IDS + "))"
                codt = gDB.getDataTable(SQLstring, gconn, transaction)
                'ElseIf ROUTE_IDS <> "" Then
                '    SQLstring += " and (ROUTE_ID in (" + ROUTE_IDS + "))"
                '    codt = gDB.getDataTable(SQLstring, gconn, transaction)
                'ElseIf DO_DATES <> "" Then
                '    SQLstring += " and (DO_DATE in (" + DO_DATES + "))"
                '    codt = gDB.getDataTable(SQLstring, gconn, transaction)
            Else
                SQLstring = ""
                codt = Nothing
            End If

            If codt IsNot Nothing AndAlso codt.Rows.Count > 0 Then
                For Each rowD As DataRow In codt.Rows
                    updateSQL = "Delete from WMS_DELV_ORDER_D " &
                            "Where imp_code='" & gU.dbEncode(rowD("imp_code")) & "' and storer_code='" & gU.dbEncode(rowD("storer_code")) & "' and DO_CODE='" & (rowD("DO_CODE").ToString.Trim) & "'"
                    gDB.amendData(updateSQL, gconn, transaction)

                    updateSQL = "Delete from WMS_DELV_ORDER " &
                            " Where DO_CODE='" & (rowD("DO_CODE").ToString.Trim) & "' and imp_code='" & gU.dbEncode(rowD("imp_code")) & "' and storer_code='" & gU.dbEncode(rowD("storer_code")) & "'"
                    gDB.amendData(updateSQL, gconn, transaction)

                    updateSQL = "Delete from WMS_WAVEPICK_RSVD " &
                           " Where DO_CODE='" & (rowD("DO_CODE").ToString.Trim) & "' and imp_code='" & gU.dbEncode(rowD("imp_code")) & "' and storer_code='" & gU.dbEncode(rowD("storer_code")) & "'"
                    gDB.amendData(updateSQL, gconn, transaction)

                    updateSQL = "Delete From WMS_DELV_ORDER_TEMP " &
                           " Where DO_CODE='" & (rowD("DO_CODE").ToString.Trim) & "' "
                    gDB.amendData(updateSQL, gconn, transaction)

                    If rowD("DO_CO_CODE").ToString <> "" Then
                        If ROUTE_IDS <> "" Then
                            updateSQL = "Update WMS_CUST_ORDER set CO_STATUS='NEW'" & " where  CO_DATE in (" + DO_DATES + ") And ROUTE_ID in (" + ROUTE_IDS + ") And CO_STATUS='PICKED' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "'"
                            '" Where CO_CODE in (" & (rowD("DO_CO_CODE").ToString.Trim) & ") and imp_code='" & gU.dbEncode(rowD("imp_code").ToString.Trim) & "' and storer_code='" & gU.dbEncode(rowD("storer_code").ToString.Trim) & "'"
                        Else
                            updateSQL = "Update WMS_CUST_ORDER set CO_STATUS='NEW'" & " where  CO_DATE in (" + DO_DATES + ") And CO_STATUS='PICKED' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "'"
                        End If
                        gDB.amendData(updateSQL, gconn, transaction)
                    End If
                    orderCount = orderCount + 1
                Next

                transaction.Commit()
                ltlMessage.Text += "<p style='color:red'>" + orderCount.ToString + " delivery order(s) deleted! </p>"
                uiFun.displayMsg(Me, "", "The order has been deleted!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "No records found to delete!", Session("gLang"))
            End If

            BindDate()
            BindCODate()

        Catch ex As Exception
            transaction.Rollback()
            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))


        Finally
            If gconn IsNot Nothing Then
                If gconn.State = ConnectionState.Open Then
                    gconn.Close()
                    gconn.Dispose()
                End If
            End If

        End Try
    End Sub

    Private Function SaveReserve(ByRef conn As SqlConnection, ByRef transaction As SqlTransaction, ByVal WH_LOC As String, ByVal WH_CODE As String, ByVal IMP_CODE As String, ByVal STORER_CODE As String, ByVal CUS_CODE As String, ByVal ITEM_CODE As String, ByVal LOT_NO As String, ByVal QTY As Double, ByVal ROUTE_ID As String, ByVal DO_DATE As DateTime, ByVal DO_CODE As String) As Int16

        Dim sql_string As String = ""
        sql_string += "insert into WMS_WAVEPICK_RSVD (imp_code, storer_code, CUS_CODE,ITEM_CODE,LOT_NO,QTY,ROUTE_ID,DO_DATE,DO_CODE,WH_CODE,WH_LOC)"
        sql_string += "values ("
        sql_string += "'" & gU.dbEncode(IMP_CODE.Trim) & "',"
        sql_string += "'" & gU.dbEncode(STORER_CODE.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(CUS_CODE.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(ITEM_CODE.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(LOT_NO.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(QTY.ToString.Trim) & "', "
        sql_string += "'" & gU.dbEncode(ROUTE_ID.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(DO_DATE.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(DO_CODE.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(WH_CODE.ToString.Trim) & "',"
        sql_string += "'" & gU.dbEncode(WH_LOC.ToString.Trim) & "'"
        sql_string += ")"

        Dim RowCount = gDB.amendData(sql_string, conn, transaction)
        Return RowCount
    End Function

    Protected Sub btnPickingDeleteCO_Click(sender As Object, e As EventArgs) Handles btnPickingDeleteCO.Click
        Dim gconn As SqlConnection
        Dim transaction As SqlTransaction
        gconn = gDB.getConnection()
        transaction = gconn.BeginTransaction()
        Dim orderCount = 0
        Dim codtl As DataTable = Nothing

        Try
            ltlMessage.Text = ""
            Dim DO_DATES = ""
            Dim ROUTE_IDS = ""
            For Each item As ListItem In lboxCODate.Items
                If item.Selected = True Then
                    If DO_DATES = "" Then
                        DO_DATES = "'" + item.Value + "'"
                    Else
                        DO_DATES = DO_DATES + ",'" + item.Value + "'"
                    End If

                End If
            Next
            For Each item As ListItem In lboxCORoute.Items
                If item.Selected = True Then
                    If ROUTE_IDS = "" Then
                        ROUTE_IDS = "'" + item.Value + "'"
                    Else
                        ROUTE_IDS = ROUTE_IDS + ",'" + item.Value + "'"
                    End If

                End If
            Next


            Dim SQLstring = "SELECT imp_code,storer_code,CO_CODE FROM WMS_CUST_ORDER WHERE CO_STATUS ='NEW' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "'"
            If ROUTE_IDS <> "" AndAlso DO_DATES <> "" Then
                SQLstring += " and (CO_DATE in ( " + DO_DATES + ") And ROUTE_ID in (" + ROUTE_IDS + "))"
                codtl = gDB.getDataTable(SQLstring)
            ElseIf ROUTE_IDS <> "" Then
                SQLstring += " and (ROUTE_ID in (" + ROUTE_IDS + "))"
                codtl = gDB.getDataTable(SQLstring, gconn, transaction)
            ElseIf DO_DATES <> "" Then
                SQLstring += " and (CO_DATE in (" + DO_DATES + "))"
                codtl = gDB.getDataTable(SQLstring, gconn, transaction)
            Else
                codtl = gDB.getDataTable(SQLstring, gconn, transaction)
            End If

            If codtl IsNot Nothing AndAlso codtl.Rows.Count > 0 Then
                For Each rowD As DataRow In codtl.Rows
                    Dim updateSQL = "Delete from WMS_CUST_ORDER_D " &
                                    " Where imp_code='" & gU.dbEncode(rowD("imp_code").ToString.Trim) & "' and storer_code='" & gU.dbEncode(rowD("storer_code").ToString.Trim) & "' and CO_CODE='" & (rowD("CO_CODE").ToString.Trim) & "'"
                    gDB.amendData(updateSQL, gconn, transaction)

                    updateSQL = "Delete from WMS_CUST_ORDER " &
                                    " Where CO_CODE='" & (rowD("CO_CODE").ToString.Trim) & "' and imp_code='" & gU.dbEncode(rowD("imp_code").ToString.Trim) & "' and storer_code='" & gU.dbEncode(rowD("storer_code").ToString.Trim) & "'"
                    gDB.amendData(updateSQL, gconn, transaction)
                    orderCount = orderCount + 1

                Next
                transaction.Commit()
                uiFun.displayMsg(Me, "", "The order has been deleted!", Session("gLang"))
                ltlMessage.Text += "<p style='color:red'>" + orderCount.ToString + " customer order(s) deleted </p>"
            Else
                uiFun.displayMsg(Me, "", "No records found to delete!", Session("gLang"))
            End If
            BindDate()
            BindCODate()
        Catch ex As Exception

            uiFun.displayMsg(Me, "", ex.Message, Session("gLang"))

            transaction.Rollback()
        Finally
            If gconn IsNot Nothing Then
                If gconn.State = ConnectionState.Open Then
                    gconn.Close()
                    gconn.Dispose()
                End If
            End If
        End Try
    End Sub

    Private Function BindCODate() As Int16
        Dim SQLString As String
        Dim dtRule As DataTable = Nothing
        Dim hasRule As Int16 = 0
        SQLString = "SELECT Distinct Convert(nvarchar(20), CO_DATE,106)CO_DATE_TXT,CO_DATE FROM WMS_CUST_ORDER Where CO_STATUS='NEW' AND CO_FTRACK_NO='ACTUAL' And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "' Order by CO_DATE desc"
        'WHERE DO_STATUS ='DRAFT'
        dtRule = gDB.getDataTable(SQLString)
        lboxCODate.Items.Clear()
        lboxCORoute.Items.Clear()
        If dtRule IsNot Nothing Or dtRule.Rows.Count > 0 Then
            For Each drow In dtRule.Rows
                Dim litem As New ListItem()
                litem.Text = drow("CO_DATE_TXT")
                litem.Value = drow("CO_DATE_TXT")
                lboxCODate.Items.Add(litem)

            Next
        End If
        Return 0
    End Function

    Private Function BindCORoute() As Int16
        Dim SQLString As String
        Dim dtRule As DataTable = Nothing
        Dim hasRule As Int16 = 0
        lboxCORoute.Items.Clear()
        Dim DO_DATES = ""
        For Each item As ListItem In lboxCODate.Items
            If item.Selected = True Then
                If DO_DATES = "" Then
                    DO_DATES = "'" + item.Value + "'"
                Else
                    DO_DATES = DO_DATES + ",'" + item.Value + "'"
                End If

            End If
        Next
        If DO_DATES <> "" Then


            SQLString = "SELECT Distinct IsNull(ROUTE_ID,'')ROUTE_ID, (Case When (select count(1) from WMS_CUST_ORDER where CO_STATUS='NEW' AND CO_DATE in (" + DO_DATES + ")  And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "' and CO_FTRACK_NO='TRIAL' AND ROUTE_ID=co.ROUTE_ID)>0 Then IsNull(ROUTE_ID,'')+' (TRIAL)'  ELSE IsNull(ROUTE_ID,'') End) as ROUTE_Name FROM WMS_CUST_ORDER co Where CO_STATUS='NEW' AND CO_DATE in (" + DO_DATES + ")  And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "' Order by ROUTE_Name"
            'WHERE DO_STATUS ='DRAFT'
            dtRule = gDB.getDataTable(SQLString)

            If dtRule IsNot Nothing Or dtRule.Rows.Count > 0 Then
                For Each drow In dtRule.Rows
                    If drow("ROUTE_ID") IsNot Nothing AndAlso drow("ROUTE_ID").ToString <> "" Then
                        Dim litem As New ListItem()
                        litem.Text = drow("ROUTE_Name")
                        litem.Value = drow("ROUTE_ID")
                        lboxCORoute.Items.Add(litem)
                    End If
                Next
                For Each item As ListItem In lboxCORoute.Items
                    item.Selected = True
                Next
            End If
        End If
        Return 0
    End Function

    Private Function BindDate() As Int16
        Dim SQLString As String
        Dim dtRule As DataTable = Nothing
        Dim hasRule As Int16 = 0
        SQLString = "SELECT Distinct Convert(nvarchar(20), DO_DATE,106)DO_DATE_TXT,DO_DATE FROM WMS_DELV_ORDER Where DO_STATUS!='POSTED' AND STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "' Order by DO_DATE"
        'WHERE DO_STATUS ='DRAFT'
        dtRule = gDB.getDataTable(SQLString)
        lboxDate.Items.Clear()
        lboxRoute.Items.Clear()
        If dtRule IsNot Nothing Or dtRule.Rows.Count > 0 Then
            For Each drow In dtRule.Rows
                Dim litem As New ListItem()
                litem.Text = drow("DO_DATE_TXT")
                litem.Value = drow("DO_DATE_TXT")
                lboxDate.Items.Add(litem)

            Next
        End If
        Return 0
    End Function

    Private Function BindRoute() As Int16
        Dim SQLString As String
        Dim dtRule As DataTable = Nothing
        Dim hasRule As Int16 = 0
        lboxRoute.Items.Clear()
        Dim DO_DATES = ""
        For Each item As ListItem In lboxDate.Items
            If item.Selected = True Then
                If DO_DATES = "" Then
                    DO_DATES = "'" + item.Value + "'"
                Else
                    DO_DATES = DO_DATES + ",'" + item.Value + "'"
                End If

            End If
        Next
        If DO_DATES <> "" Then


            SQLString = "SELECT Distinct IsNull(ROUTE_ID,'')ROUTE_ID, (IsNull(ROUTE_ID,'')+ ' ('+DO_STATUS+')' + ISNULL((select Top(1) ' ('+CO_FTRACK_NO+')' from WMS_CUST_ORDER where CO_FTRACK_NO='TRIAL' and CO_CODE in  (select items from  dbo.split(DO.DO_CO_CODE,','))),''))ROUTE_TEXT FROM WMS_DELV_ORDER DO WHERE DO_DATE in (" + DO_DATES + ")  And STORER_CODE='" + STORER_CODE.SelectedValue.ToString() + "' Order By ROUTE_ID"
            'WHERE DO_STATUS ='DRAFT'
            dtRule = gDB.getDataTable(SQLString)

            If dtRule IsNot Nothing Or dtRule.Rows.Count > 0 Then
                For Each drow In dtRule.Rows
                    If drow("ROUTE_ID") IsNot Nothing AndAlso drow("ROUTE_ID").ToString <> "" Then
                        Dim litem As New ListItem()
                        litem.Text = drow("ROUTE_TEXT")
                        litem.Value = drow("ROUTE_ID")
                        lboxRoute.Items.Add(litem)
                    End If
                Next
                For Each item As ListItem In lboxRoute.Items
                    item.Selected = True
                Next
            End If


        End If
        Return 0
    End Function

    Private Function SaveErrorLog(ByVal Msg As String) As Int16
        Dim RowCount = 0
        Try
            Dim gConn As SqlConnection
            gConn = gDB.getConnection()

            'Dim transaction As SqlTransaction
            'transaction = gConn.BeginTransaction()

            Dim sql_string As String = ""
            sql_string += "insert into WMS_ERROR_LOG (ErrType, ErrMsg,ErrDate)"
            sql_string += "values("
            sql_string += "'WAVEPICKING',"
            sql_string += "N'" & gU.dbEncode(Msg.ToString.Trim) & "',"
            sql_string += "'" & DateTime.Now & "' "
            sql_string += ")"
            RowCount = gDB.amendData(sql_string, gConn, Nothing)

        Catch ex As Exception

        End Try
        Return RowCount
    End Function
    Protected Sub lboxDate_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lboxDate.SelectedIndexChanged
        BindRoute()
    End Sub

    Protected Sub lboxCODate_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lboxCODate.SelectedIndexChanged
        BindCORoute()
    End Sub

    Protected Sub btnSelectCODate_Click(sender As Object, e As EventArgs) Handles btnSelectCODate.Click

        Dim itm = lboxCODate.SelectedValue
        If itm IsNot Nothing AndAlso itm <> "" Then
            For Each item As ListItem In lboxCODate.Items
                item.Selected = False
            Next
            'btnSelectCODate.Text = "Select All"
        Else
            For Each item As ListItem In lboxCODate.Items
                item.Selected = True
            Next
            'btnSelectCODate.Text = "Deselect All"
        End If
        BindCORoute()

    End Sub

    Protected Sub btnSelectCORoute_Click(sender As Object, e As EventArgs) Handles btnSelectCORoute.Click
        Dim itm = lboxCORoute.SelectedValue
        If itm IsNot Nothing AndAlso itm <> "" Then
            For Each item As ListItem In lboxCORoute.Items
                item.Selected = False
            Next
            'btnSelectCODate.Text = "Select All"
        Else
            For Each item As ListItem In lboxCORoute.Items
                item.Selected = True
            Next
            'btnSelectCODate.Text = "Deselect All"
        End If
    End Sub

    Protected Sub btnSelectDODate_Click(sender As Object, e As EventArgs) Handles btnSelectDODate.Click
        Dim itm = lboxDate.SelectedValue
        If itm IsNot Nothing AndAlso itm <> "" Then
            For Each item As ListItem In lboxDate.Items
                item.Selected = False
            Next
            'btnSelectCODate.Text = "Select All"
        Else
            For Each item As ListItem In lboxDate.Items
                item.Selected = True
            Next
            'btnSelectCODate.Text = "Deselect All"
        End If
        BindRoute()
    End Sub

    Protected Sub btnSelectRoute_Click(sender As Object, e As EventArgs) Handles btnSelectRoute.Click
        Dim itm = lboxRoute.SelectedValue
        If itm IsNot Nothing AndAlso itm <> "" Then
            For Each item As ListItem In lboxRoute.Items
                item.Selected = False
            Next
            'btnSelectCODate.Text = "Select All"
        Else
            For Each item As ListItem In lboxRoute.Items
                item.Selected = True
            Next
            'btnSelectCODate.Text = "Deselect All"
        End If
    End Sub

    Protected Sub btnBackUpStock_Click(sender As Object, e As EventArgs) Handles btnBackUpStock.Click
        Dim SQLString = "Select IsNull(sum(ILOC_BAL_QTY),0)Qty from WMS_ITEM_LOC_BAL"
        Dim dtTemp = gDB.getDataTable(SQLString)
        If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
            If Convert.ToDouble(dtTemp.Rows(0)("Qty").ToString) > 0 Then
                SQLString = "Update WMS_ITEM_LOC_BAL Set ILOC_BAL_CBM = ILOC_BAL_QTY"
                gDB.amendData(SQLString)
                uiFun.displayMsg(Me, "", "Backup success.", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "No stock found", Session("gLang"))
            End If
        Else
            uiFun.displayMsg(Me, "", "No stock found", Session("gLang"))
        End If
    End Sub

    Protected Sub btnRestoreStock_Click(sender As Object, e As EventArgs) Handles btnRestoreStock.Click

        Dim SQLString = "Select IsNull(sum(ILOC_BAL_CBM),0)Qty from WMS_ITEM_LOC_BAL"
        Dim dtTemp = gDB.getDataTable(SQLString)
        If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
            If Convert.ToDouble(dtTemp.Rows(0)("Qty").ToString) > 0 Then
                SQLString = "Update WMS_ITEM_LOC_BAL Set  ILOC_BAL_QTY=ILOC_BAL_CBM "
                gDB.amendData(SQLString)

                SQLString = "Update WMS_ITEM_LOC_BAL Set  ILOC_BAL_CBM = 0"
                gDB.amendData(SQLString)
                uiFun.displayMsg(Me, "", "Restoration success.", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "No stock found", Session("gLang"))
            End If
        Else
            uiFun.displayMsg(Me, "", "No stock found", Session("gLang"))
        End If

    End Sub

    Protected Sub STORER_CODE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles STORER_CODE.SelectedIndexChanged
        BindDate()
        BindCODate()
    End Sub
End Class
