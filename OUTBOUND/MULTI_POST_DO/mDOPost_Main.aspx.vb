Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OUTBOUND_MULTI_DO_mDOPost_Main
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private dl As New DocLink
    Private st As New StockTrans
    Private moduleAction As String = ""
    Private DDFORMAT2 As String = "dd/MM/yyyy"
    Private DDFORMAT As String = gU.getConfig("DDFORMATNO")
    Private imp_code As String = ""

    Private dt As New DataTable
    Private exceptionEditList As List(Of String)

    Private Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 3600
        DDFORMAT2 = gU.getConfig("DDFORMAT2")
        'DDFORMAT = gU.getConfig("DDFORMAT")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_MDOPOST", Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            ViewState("dt") = Nothing

            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
        End If

        If Session("gLang") = "E" Then
            lheader.Text = "Batch Post DO:"
            lbl_STORER_CODE.Text = "Organizations"
            lbl_DO_EDI_SIR_NO.Text = "Customer Order No."
            lbl_SO_NO.Text = "SO No."
            lbl_DO_DATE.Text = "DO Date"
            lblGV.Text = "DO List"


            btnSearch.Text = "Search"
            btnReset.Text = "Reset"
            btnSelAll.Text = "Select All"
            btnUnSel.Text = "Unselect All"

            BtnGen.Text = "Post"
            cfm1.ConfirmText = "All selected DO will be posted, Confirm to proceed?"

            btnGen2.Text = "Post"
            cfm2.ConfirmText = "All selected DO will be posted, Confirm to proceed?"

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "批次出庫:"
            lbl_STORER_CODE.Text = "部門"
            lbl_DO_EDI_SIR_NO.Text = "訂單號碼"
            lbl_SO_NO.Text = "所以不行"
            lbl_DO_DATE.Text = "訂單日期"
            lblGV.Text = "DO List"

            lblGV.Text = "提貨單"
            btnSearch.Text = "搜尋"
            btnReset.Text = "重置"
            btnSelAll.Text = "全選"
            btnUnSel.Text = "全取消"

            BtnGen.Text = "出庫"
            cfm1.ConfirmText = "確定要將已選提貨單出庫?"

            btnGen2.Text = "出庫"
            cfm2.ConfirmText = "確定要將已選提貨單出庫?"

        End If

    End Sub

    Protected Sub BindGV()
        Dim tempStr As String = ""
        Dim selectSQL As String = ""
        Dim dt As DataTable

        If STORER_CODE.SelectedValue <> "" Then
            tempStr = " AND STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "'"
        End If

        Dim keyArr As String()
        Dim FrToArr As String()
        Dim tempApp As String = ""
        Dim tempApp2 As String = ""


        If Not String.IsNullOrWhiteSpace(DO_EDI_SIR_NO.Text) Then
            tempStr &= " AND ','+Replace(DO_CO_CODE,' ','')+',' like '%," & DO_EDI_SIR_NO.Text.Trim & ",%' "
        End If

        If SO_NUMBER.Text.Trim <> "" Then
            tempStr &= " And ','+Replace(SO_NUMBER,' ','')+',' like '%," & SO_NUMBER.Text.Trim & ",%' "
        End If

        If DO_DATE_FR.Text.Trim <> "" Then
            tempStr &= " AND  CONVERT(datetime,DO_DATE," & DDFORMAT & ") >= CONVERT(datetime,'" & DO_DATE_FR.Text.Trim & "'," & DDFORMAT & ") "
        End If

        If DO_DATE_TO.Text.Trim <> "" Then
            tempStr &= " AND CONVERT(datetime,DO_DATE," & DDFORMAT & ") < CONVERT(datetime,'" & DO_DATE_TO.Text.Trim & "'," & DDFORMAT & ") + 1 "
        End If


        If tempStr <> "" Then
            selectSQL = " select * from(SELECT DISTINCT WMS_DELV_ORDER.IMP_CODE, WMS_DELV_ORDER.STORER_CODE,WMS_DELV_ORDER.ROUTE_ID, WMS_DELV_ORDER.DO_CODE, " &
                    " WMS_DELV_ORDER.DO_CO_CODE, WMS_DELV_ORDER.DO_STATUS,CONVERT(varchar, WMS_DELV_ORDER.DO_DATE,103 ) as DO_DATE,WMS_DELV_ORDER.CUS_CODE,WMS_DELV_ORDER.CUS_NAME, " &
                    " (STUFF(( select ', ' + CO_INV_NO from WMS_CUST_ORDER where CO_CODE in (select items from dbo.Split(REPLACE((WMS_DELV_ORDER.DO_CO_CODE),' ',''),',')) FOR XML PATH('')  ) ,1,2,'')) SO_NUMBER , " &
                    " WMS_DELV_ORDER.DO_EDI_SIR_NO,wms_storer.sto_name, '' as checkYN FROM WMS_DELV_ORDER INNER JOIN  WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND " &
                    " WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE INNER JOIN  WMS_STORER ON WMS_DELV_ORDER.IMP_CODE = WMS_STORER.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_STORER.STORER_CODE WHERE DO_STATUS NOT in ('POSTED','CANCELLED','RELEASED')) as t where 1=1 " &
                    tempStr &
                    " ORDER BY t.DO_CODE DESC "
        Else
            selectSQL = " select * from(SELECT DISTINCT WMS_DELV_ORDER.IMP_CODE, WMS_DELV_ORDER.STORER_CODE,WMS_DELV_ORDER.ROUTE_ID, WMS_DELV_ORDER.DO_CODE, " &
                    " WMS_DELV_ORDER.DO_CO_CODE, WMS_DELV_ORDER.DO_STATUS,CONVERT(varchar, WMS_DELV_ORDER.DO_DATE,103 ) as DO_DATE,WMS_DELV_ORDER.CUS_CODE,WMS_DELV_ORDER.CUS_NAME, " &
                    " (STUFF(( select ', ' + CO_INV_NO from WMS_CUST_ORDER where CO_CODE in (select items from dbo.Split(REPLACE((WMS_DELV_ORDER.DO_CO_CODE),' ',''),',')) FOR XML PATH('')  ) ,1,2,'')) SO_NUMBER , " &
                    " WMS_DELV_ORDER.DO_EDI_SIR_NO,wms_storer.sto_name, '' as checkYN FROM WMS_DELV_ORDER INNER JOIN  WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND " &
                    " WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE INNER JOIN  WMS_STORER ON WMS_DELV_ORDER.IMP_CODE = WMS_STORER.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_STORER.STORER_CODE WHERE DO_STATUS NOT in ('POSTED','CANCELLED','RELEASED')) as t " &
                    " ORDER BY t.DO_CODE DESC "
        End If

        'selectSQL = " SELECT DISTINCT WMS_DELV_ORDER.IMP_CODE, WMS_DELV_ORDER.STORER_CODE,WMS_DELV_ORDER.ROUTE_ID, WMS_DELV_ORDER.DO_CODE, WMS_DELV_ORDER.DO_CO_CODE, WMS_DELV_ORDER.DO_STATUS,CONVERT(varchar, WMS_DELV_ORDER.DO_DATE," & DDFORMAT & " ) as DO_DATE, WMS_DELV_ORDER.CUS_CODE, " &
        '            " WMS_DELV_ORDER.CUS_NAME, (STUFF(( select ', ' + CO_INV_NO from WMS_CUST_ORDER where CO_CODE in (select items from dbo.Split(REPLACE((WMS_DELV_ORDER.DO_CO_CODE),' ',''),',')) FOR XML PATH('')  ) ,1,2,'')) SO_NUMBER , WMS_DELV_ORDER.DO_EDI_SIR_NO,wms_storer.sto_name, '' as checkYN, DO_STATUS " &
        '            " FROM WMS_DELV_ORDER INNER JOIN " &
        '            " WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND " &
        '            " WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE INNER JOIN " &
        '            " WMS_STORER ON WMS_DELV_ORDER.IMP_CODE = WMS_STORER.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_STORER.STORER_CODE " &
        '            " WHERE DO_STATUS NOT in ('POSTED','CANCELLED','RELEASED') " &
        '            tempStr &
        '            " ORDER BY DO_CODE DESC"

        dt = gDB.getDataTable(selectSQL)

        ViewState("dt") = dt

        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If

        GridView1.DataBind()
        changeLabel()
        GVPNL.Visible = True
    End Sub

    Protected Sub Page_LoadComplete(sender As Object, e As System.EventArgs) Handles Me.LoadComplete
        Dim sm As ScriptManager = ScriptManager.GetCurrent(Page)

        sm.RegisterAsyncPostBackControl(btnSearch)
        sm.RegisterAsyncPostBackControl(btnSelAll)
        sm.RegisterAsyncPostBackControl(btnUnSel)
        sm.RegisterAsyncPostBackControl(BtnGen)
        sm.RegisterAsyncPostBackControl(btnGen2)

    End Sub

    Protected Sub changeLabel()
        REM **********************
        REM Use for re-create the label to change the Langauge
        REM Modify Here

        Call cU.newChangeGVLabel(GridView1, "", "")
        Call cU.newChangeGVLabel(GridView1, "Organizations", "部門")
        Call cU.newChangeGVLabel(GridView1, "DO Code", "送貨指令")
        Call cU.newChangeGVLabel(GridView1, "Customer Order Number", "訂單號碼")
        Call cU.newChangeGVLabel(GridView1, "DO Date", "日期")
        Call cU.newChangeGVLabel(GridView1, "Route", "路線")
        Call cU.newChangeGVLabel(GridView1, "SO No.", "SO No.")
        Call cU.newChangeGVLabel(GridView1, "Customer Name", "客戶名稱")
        Call cU.newChangeGVLabel(GridView1, "Status", "狀態")

        REM **********************
    End Sub


    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                If DataBinder.Eval(e.Row.DataItem, "checkYN").ToString.Trim = "Y" Then
                    CType(e.Row.FindControl("checkYN"), CheckBox).Text = True
                End If

                CType(e.Row.FindControl("sto_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "sto_name").ToString.Trim
                CType(e.Row.FindControl("STORER_CODE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "STORER_CODE").ToString.Trim

                CType(e.Row.FindControl("DO_CODE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DO_CODE").ToString.Trim
                CType(e.Row.FindControl("DO_EDI_SIR_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DO_CO_CODE").ToString.Trim 'DataBinder.Eval(e.Row.DataItem, "DO_EDI_SIR_NO").ToString.Trim
                CType(e.Row.FindControl("DO_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DO_DATE").ToString.Trim
                CType(e.Row.FindControl("ROUTE_ID"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ROUTE_ID").ToString.Trim
                CType(e.Row.FindControl("SO_NUMBER"), Label).Text = DataBinder.Eval(e.Row.DataItem, "SO_NUMBER").ToString.Trim
                CType(e.Row.FindControl("CUS_NAME"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CUS_NAME").ToString.Trim
                CType(e.Row.FindControl("DO_STATUS"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DO_STATUS").ToString.Trim
        End Select
    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As System.EventArgs) Handles btnSearch.Click
        BindGV()
    End Sub

    Protected Sub BtnGen_Click(sender As Object, e As System.EventArgs) Handles BtnGen.Click
        MultiPost()
    End Sub

    Protected Sub btnGen2_Click(sender As Object, e As System.EventArgs) Handles btnGen2.Click
        MultiPost()
    End Sub

    Protected Function MultiPost() As Boolean
        Dim noError As Boolean = True
        Dim errMsg As New StringBuilder

        Dim checkedDO As String = ""

        If GridView1 IsNot Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If DirectCast(GridView1.Rows(i).FindControl("checkYN"), CheckBox).Checked Then
                    checkedDO = gU.appendToList(checkedDO, "'" & DirectCast(GridView1.Rows(i).FindControl("STORER_CODE"), HiddenField).Value & "_000_" & DirectCast(GridView1.Rows(i).FindControl("DO_CODE"), Label).Text & "'")
                End If
            Next

            If checkedDO <> "" Then
                Dim updateSql, selectSql As String
                Dim pl_dt As DataTable
                Dim itmKey As String
                Dim qtyDict, cbmDict, wgtDict As Dictionary(Of String, Double)
                Dim gConn As SqlConnection
                Dim transaction As SqlTransaction
                Dim lCOD_SEQ As String
                Dim paP As GlobalDBFunc.DBCmdPara

                Dim doDT, MainDT As DataTable

                Dim dtlQtyDict, plQtyDict As Dictionary(Of String, Double)
                Dim itmKey2 As String
                Dim keys As Dictionary(Of String, Double).KeyCollection
                Dim itmArray As String()


                Dim lDO_CODE As String = ""
                Dim lStorer_code As String = ""
                Dim safeToGo As Boolean = False

                paP = New GlobalDBFunc.DBCmdPara
                selectSql = "SELECT WMS_DELV_ORDER.STORER_CODE, WMS_DELV_ORDER.DO_CODE,WMS_DELV_ORDER.ROUTE_ID,WMS_DELV_ORDER.DO_CO_CODE, WMS_DELV_ORDER.DO_EDI_SIR_NO, WMS_DELV_ORDER.CUS_CODE " &
                            "FROM WMS_DELV_ORDER " &
                            "WHERE WMS_DELV_ORDER.IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND WMS_DELV_ORDER.DO_STATUS != 'POSTED' AND WMS_DELV_ORDER.DO_STATUS != 'INPROGRESS' AND WMS_DELV_ORDER.STORER_CODE + '_000_' + WMS_DELV_ORDER.DO_CODE IN (" & checkedDO & ") ORDER BY STORER_CODE, DO_CODE "

                MainDT = gDB.getDataTable(selectSql, , , , paP)

                Dim UpdateDOStatus As String = "update WMS_DELV_ORDER set DO_STATUS='INPROGRESS' WHERE WMS_DELV_ORDER.IMP_CODE='" & (Session("IMP_CODE")) & "' AND WMS_DELV_ORDER.DO_STATUS != 'POSTED' AND WMS_DELV_ORDER.DO_STATUS != 'INPROGRESS' AND WMS_DELV_ORDER.STORER_CODE + '_000_' + WMS_DELV_ORDER.DO_CODE IN (" & checkedDO & ")  "
                gDB.amendData(UpdateDOStatus)

                If MainDT.Rows.Count > 0 Then

                    For i = 0 To MainDT.Rows.Count - 1
                            lDO_CODE = MainDT.Rows(i).Item("DO_CODE").ToString.Trim
                        lStorer_code = MainDT.Rows(i).Item("STORER_CODE").ToString.Trim

                        gConn = gDB.getConnection()
                        transaction = gConn.BeginTransaction()

                        paP = New GlobalDBFunc.DBCmdPara
                        selectSql = "Select DO_STATUS from wms_delv_order WHERE IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND STORER_CODE=" & paP.AP(lStorer_code) & " AND DO_CODE=" & paP.AP(lDO_CODE)
                        Dim checkDO As DataTable = gDB.getDataTable(selectSql, gConn, transaction, , paP)
                        If checkDO.Rows(0).Item("DO_STATUS").ToString.Trim <> "POSTED" Then

                            paP = New GlobalDBFunc.DBCmdPara
                            selectSql = "SELECT d.DOD_SEQ, d.DOD_DISP_SEQ, d.DOD_PALLET_NO, d.DOD_CARTON_NO, d.DOD_PACK_NO, d.DOD_ITM_CODE, d.DOD_PACK_KEY, " &
                                    "d.DOD_ITM_DESC, d.DOD_PACK_TYPE, d.DOD_QTY, d.DOD_UOM, d.DOD_CUT_YN, d.DOD_QTY2, d.DOD_UOM2, d.DOD_PCS_UOM, d.DOD_TOTPCS, d.DOD_TOT_WGT, d.DOD_TOT_CBM, " &
                                    "d.DOD_REM, d.DOD_TICKET_NO, d.DOD_WH_CODE, d.DOD_VND_CODE, d.DOD_BATCH_NO, d.IMP_CODE, d.STORER_CODE, d.DO_CODE, i.ITM_SKU_NO, i.ITM_DESC, " &
                                    "Convert(varchar,d.DOD_EXPIRY_DATE, " & gU.getConfig("DDFORMATNo") & ") as DOD_EXPIRY_DATE, " &
                                    "Convert(varchar,d.DOD_MANU_DATE, " & gU.getConfig("DDFORMATNo") & ") as DOD_MANU_DATE, " &
                                    "d.DOD_TROLLEY_ID, i.ITM_SERIAL_NO_YN, i.ITM_TYPE " &
                                    "from WMS_DELV_ORDER_D d " &
                                    "Left outer join WMS_ITEM i " &
                                    "ON d.IMP_CODE = i.IMP_CODE " &
                                    "and d.STORER_CODE = i.STORER_CODE " &
                                    "and d.DOD_ITM_CODE = i.ITM_CODE " &
                                    "and d.DOD_PACK_KEY = i.PACK_KEY " &
                                    "WHERE d.IMP_CODE=" & paP.AP(Session("IMP_CODE")) & " AND d.STORER_CODE=" & paP.AP(lStorer_code) & " AND d.DO_CODE=" & paP.AP(lDO_CODE)

                            doDT = gDB.getDataTable(selectSql, , , , paP)

                            safeToGo = True





                            Try
                                qtyDict = New Dictionary(Of String, Double)
                                cbmDict = New Dictionary(Of String, Double)
                                wgtDict = New Dictionary(Of String, Double)


                                dtlQtyDict = New Dictionary(Of String, Double)
                                plQtyDict = New Dictionary(Of String, Double)

                                For Each rows As DataRow In doDT.Rows
                                    itmKey = gU.decodeNullOrEmpty(rows.Item("dod_itm_code").ToString.Trim, "") & "#_#" &
                                        gU.decodeNullOrEmpty(rows.Item("dod_pack_key").ToString.Trim, "") & "#_#" &
                                        gU.decodeNullOrEmpty(rows.Item("dod_pallet_no").ToString.Trim, "") & "#_#" &
                                        gU.decodeNullOrEmpty(rows.Item("dod_batch_no").ToString.Trim, "")

                                    If qtyDict.ContainsKey(itmKey) Then
                                        qtyDict.Item(itmKey) = qtyDict.Item(itmKey) + gU.decodeEmptyCdbl(rows.Item("dod_qty").ToString.Trim, 0)
                                        cbmDict.Item(itmKey) = cbmDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(rows.Item("dod_tot_cbm").ToString.Trim, 0), 14)
                                        wgtDict.Item(itmKey) = wgtDict.Item(itmKey) + Math.Round(gU.decodeEmptyCdbl(rows.Item("dod_tot_wgt").ToString.Trim, 0), 14)
                                    Else
                                        qtyDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("dod_qty").ToString.Trim, 0))
                                        cbmDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("dod_tot_cbm").ToString.Trim, 0))
                                        wgtDict.Add(itmKey, gU.decodeEmptyCdbl(rows.Item("dod_tot_wgt").ToString.Trim, 0))
                                    End If


                                    itmKey2 = gU.decodeNullOrEmpty(rows.Item("dod_itm_code").ToString.Trim, "") & "#_#" & gU.decodeNullOrEmpty(rows.Item("dod_pack_key").ToString.Trim, "")

                                    If gU.decodeNullOrEmpty(rows.Item("itm_serial_no_yn").ToString.Trim, "") = "Y" Then
                                        If dtlQtyDict.ContainsKey(itmKey2) Then
                                            dtlQtyDict.Item(itmKey2) = dtlQtyDict.Item(itmKey2) + CDbl(gU.decodeNullOrEmpty(rows.Item("dod_qty2").ToString.Trim, "0")) * CDbl(gU.decodeNullOrEmpty(rows.Item("dod_qty").ToString.Trim, "0"))
                                        Else
                                            dtlQtyDict.Add(itmKey2, CDbl(gU.decodeNullOrEmpty(rows.Item("dod_qty2").ToString.Trim, "0")) * CDbl(gU.decodeNullOrEmpty(rows.Item("dod_qty").ToString.Trim, "0")))
                                        End If
                                    Else
                                        If dtlQtyDict.ContainsKey(itmKey2) Then
                                            dtlQtyDict.Item(itmKey2) = dtlQtyDict.Item(itmKey2) + CDbl(gU.decodeNullOrEmpty(rows.Item("dod_qty").ToString.Trim, "0"))
                                        Else
                                            dtlQtyDict.Add(itmKey2, CDbl(gU.decodeNullOrEmpty(rows.Item("dod_qty").ToString.Trim, "0")))
                                        End If
                                    End If
                                Next


                                paP = New GlobalDBFunc.DBCmdPara
                                selectSql = "SELECT distinct d.PLD_SEQ, d.DO_CODE, d.PLD_PICKED_BY, d.PLD_ITEM_NO, i.ITM_SKU_NO, d.PLD_PACK_KEY, d.PLD_PALLET_NO, d.PLD_ITEM_QTY, l.ILOC_BAL_QTY, " &
                                            "d.PLD_WH, d.PLD_LOC, d.PLD_ORG_LOC, d.PLD_REMARK, d.PLD_FLOOR, d.PLD_AREA, d.PLD_RACK,d.pld_ss_qty, d.PLD_BIN, d.PLD_IS_LOAN, d.PLD_DO_QTY, " &
                                            "d.PLD_BATCH_NO, d.PLD_FOI_QTY, d.PLD_SERIAL_NO, d.PLD_QTY2, i.ITM_SERIAL_NO_YN, i.itm_type, " &
                                            "ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(pick_item.picked_qty, 0) as stock_qty, " &
                                            "Convert(int, PLD_SPLIT_FR) as PLD_SPLIT_FR_INT, Convert(int, PLD_SEQ) AS PLD_SEQ_INT, " &
                                            "Convert(varchar, d.pld_expiry_date, " & gU.getConfig("DDFORMATNo") & ") as pld_expiry_date, " &
                                            "Convert(varchar,d.pld_manu_date, " & gU.getConfig("DDFORMATNo") & ") as pld_manu_date, " &
                                            "Convert(varchar,ISNULL(d.pld_expiry_date, l.ILOC_EXPIRY_DATE), " & gU.getConfig("DDFORMATNo") & ") as ILOC_EXPIRY_DATE, " &
                                            "isnull(dod.dod_disp_seq, 9999) as dod_disp_seq, b.BN_CSMS_CODE, " &
                                            "PLD_TO_DRUM_ID, PLD_TO_DRUM_CABLE_LIST, PLD_TO_DRUM_ILOC_SEQ,'' as from_drum_id,  " &
                                            "0.0 as HOLD_QTY, 0.0 as TOTAL_BAL, 0.0 as AVAIL_QTY, 0.0 as TOTAL_ITEM_QTY, 0.0 as TOTAL_FOI_QTY, 'U' as mFlag " &
                                        "from WMS_DO_PICKLIST_D d " &
                                        "left outer join WMS_ITEM_LOC_BAL l " &
                                            "on d.IMP_CODE = l.IMP_CODE " &
                                            "and d.STORER_CODE = l.STORER_CODE " &
                                            "and d.PLD_ITEM_NO = l.ITM_CODE " &
                                            "and l.ILOC_WH = d.PLD_WH " &
                                            "and d.PLD_PACK_KEY = l.PACK_KEY " &
                                            "and ISNULL(d.PLD_PALLET_NO,'000') = ISNULL(l.ILOC_PALLET_NO,'000') " &
                                            "and d.PLD_LOC = l.ILOC_LOC " &
                                            "and ISNULL(d.PLD_BATCH_NO, '') = ISNULL(l.ILOC_BATCH_NO, '') " &
                                        "left outer join WMS_DATE_CODE DC " &
                                            "on l.ILOC_BATCH_NO = DC.DC_DATE_CODE " &
                                            "and d.IMP_CODE = DC.IMP_CODE " &
                                            "and d.STORER_CODE = DC.STORER_CODE " &
                                        "left outer join WMS_ITEM i " &
                                            "on d.IMP_CODE = i.IMP_CODE " &
                                            "and d.STORER_CODE = i.STORER_CODE " &
                                            "and d.PLD_ITEM_NO = i.ITM_CODE " &
                                            "and d.PLD_PACK_KEY = i.PACK_KEY " &
                                        "LEFT OUTER JOIN ( " &
                                                "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " &
                                                "from wms_do_picklist_d p, wms_delv_order d " &
                                                "where d.imp_code = p.imp_code " &
                                                "and d.storer_code = p.storer_code " &
                                                "and d.do_code = p.do_code " &
                                                "and d.do_status = 'PICKED' " &
                                                "and d.do_code <> '" & gU.dbEncode(lDO_CODE) & "' " &
                                                "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " &
                                            "ON d.IMP_CODE = PICK_ITEM.IMP_CODE " &
                                            "AND d.STORER_CODE = PICK_ITEM.STORER_CODE " &
                                            "AND d.PLD_ITEM_NO = PICK_ITEM.PLD_ITEM_NO " &
                                            "AND d.PLD_PACK_KEY = PICK_ITEM.PLD_PACK_KEY " &
                                            "AND ISNULL(d.PLD_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " &
                                            "AND d.PLD_LOC = PICK_ITEM.PLD_LOC " &
                                        "left outer join (" &
                                                "select imp_code, storer_code, do_code, dod_itm_code, dod_pack_key, min(dod_disp_seq) as dod_disp_seq " &
                                                "from WMS_DELV_ORDER_D " &
                                                "group by imp_code, storer_code, do_code, dod_itm_code, dod_pack_key) DOD " &
                                            "ON d.IMP_CODE = DOD.IMP_CODE " &
                                            "AND d.STORER_CODE = DOD.STORER_CODE " &
                                            "AND d.DO_CODE = DOD.DO_CODE " &
                                            "AND d.PLD_ITEM_NO = DOD.dod_itm_code " &
                                            "AND d.PLD_PACK_KEY = DOD.dod_pack_key " &
                                        "left outer join wms_wh_bin b " &
                                            "on d.PLD_LOC = b.LOC_KEY " &
                                        "where d.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                                        "and d.DO_CODE = '" & gU.dbEncode(lDO_CODE) & "' " &
                                        "and d.STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' "


                                pl_dt = gDB.getDataTable(selectSql, gConn, transaction, , paP)

                                If pl_dt.Rows.Count = 0 Then
                                    If Session("gLang") = "E" Then
                                        errMsg.AppendLine("Error Post in DO No:" & lDO_CODE & ", No Picklist is found!")
                                    Else
                                        errMsg.AppendLine("出庫錯誤 提貨號碼:" & lDO_CODE & ", 找不到執貨單!")
                                    End If
                                    transaction.Rollback()
                                    gConn.Close()
                                    gConn.Dispose()
                                    noError = False
                                    safeToGo = False

                                    Continue For
                                End If

                                Dim emptyLoc As Boolean = False
                                Dim qtyError As Boolean = False

                                For Each rows As DataRow In pl_dt.Rows
                                    If rows.Item("pld_loc").ToString.Trim = "" Then
                                        emptyLoc = True
                                    End If

                                    If gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString, 0) > gU.decodeEmptyCdbl(rows.Item("pld_foi_qty").ToString, 0) Then
                                        If Session("gLang") = "E" Then
                                            errMsg.AppendLine("Error Post in DO No:" & lDO_CODE & ",Picked qty is greater than FFI qty, " & rows.Item("pld_item_no").ToString.Trim & ", please change the qty in Picking List!")
                                        Else
                                            errMsg.AppendLine("出庫錯誤 提貨號碼:" & lDO_CODE & "物料號碼:" & rows.Item("pld_item_no").ToString.Trim & ", 請修正取貨單內的數量!")
                                        End If

                                        qtyError = True
                                    End If

                                    itmKey2 = rows.Item("pld_item_no").ToString.Trim & "#_#" & rows.Item("pld_pack_key").ToString.Trim

                                    If rows.Item("itm_serial_no_yn").ToString.Trim = "Y" Then
                                        If plQtyDict.ContainsKey(itmKey2) Then
                                            plQtyDict.Item(itmKey2) = plQtyDict.Item(itmKey2) + CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(rows.Item("pld_qty2"), 0), 0))
                                        Else
                                            plQtyDict.Add(itmKey2, CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(rows.Item("pld_qty2"), 0), 0)))
                                        End If
                                    Else
                                        If plQtyDict.ContainsKey(itmKey2) Then
                                            plQtyDict.Item(itmKey2) = plQtyDict.Item(itmKey2) + CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(rows.Item("pld_item_qty"), 0), 0))
                                        Else
                                            plQtyDict.Add(itmKey2, CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(rows.Item("pld_item_qty"), 0), 0)))
                                        End If
                                    End If


                                Next

                                If emptyLoc Then
                                    If Session("gLang") = "E" Then
                                        errMsg.AppendLine("Error Post in DO No:" & lDO_CODE & ", No stock has been pick from location!")
                                    Else
                                        errMsg.AppendLine("出庫錯誤 提貨號碼:" & lDO_CODE & ", 執貨未完成!")
                                    End If
                                    transaction.Rollback()
                                    gConn.Close()
                                    gConn.Dispose()

                                    UpdateDOStatus = "update WMS_DELV_ORDER set DO_STATUS='PICKED' WHERE DO_STATUS='INPROGRESS'"
                                    gDB.amendData(UpdateDOStatus)

                                    noError = False
                                    safeToGo = False

                                    Continue For
                                End If

                                If qtyError Then

                                    transaction.Rollback()
                                    gConn.Close()
                                    gConn.Dispose()
                                    noError = False
                                    safeToGo = False

                                    Continue For
                                End If


                                keys = dtlQtyDict.Keys

                                qtyError = False

                                For k = 0 To keys.Count - 1
                                    If plQtyDict.ContainsKey(keys(k)) Then
                                        If Math.Round(dtlQtyDict.Item(keys(k)), 4) < Math.Round(plQtyDict.Item(keys(k)), 4) Then
                                            itmArray = Split(keys(k), "#_#")
                                            If Session("gLang") = "E" Then
                                                errMsg.AppendLine("Error Post in DO No:" & lDO_CODE & "Item qty is less than picked qty, " & itmArray(0) & "-" & itmArray(1) & ", please change the picked qty in Picking List!")
                                            Else
                                                errMsg.AppendLine("出庫錯誤 提貨號碼:" & lDO_CODE & ",貨單數量少於實取貨量, " & itmArray(0) & "-" & itmArray(1) & ", 請修正取貨單內的數量!")
                                            End If

                                            qtyError = True
                                        End If
                                    Else
                                        itmArray = Split(keys(k), "#_#")
                                        If Session("gLang") = "E" Then
                                            errMsg.AppendLine("Error Post in DO No:" & lDO_CODE & "There is no Picking List for item, " & itmArray(0) & "-" & itmArray(1) & ", please check the Picking List!")
                                        Else
                                            errMsg.AppendLine("出庫錯誤 提貨號碼:" & lDO_CODE & ",在取貨單中找不到貨品, " & itmArray(0) & "-" & itmArray(1) & ", 請檢楂取貨單!")
                                        End If

                                        qtyError = True
                                    End If
                                Next

                                If qtyError Then
                                    transaction.Rollback()
                                    gConn.Close()
                                    gConn.Dispose()

                                    UpdateDOStatus = "update WMS_DELV_ORDER set DO_STATUS='PICKED' WHERE DO_STATUS='INPROGRESS'"
                                    gDB.amendData(UpdateDOStatus)

                                    noError = False
                                    safeToGo = False

                                    Continue For
                                End If

                                Dim MANU_DATE As String = ""
                                Dim EXP_DATE As String = ""

                                'Sum total foi qty and picked qty for check stock bal validation
                                Dim objPlt As PickListTable
                                objPlt = New PickListTable(MainDT.Rows(i).Item("DO_CO_CODE").ToString.Trim, lDO_CODE, lStorer_code, Session("IMP_CODE"))
                                objPlt.updateTotalQty(pl_dt)
                                objPlt.updateHoldBal(pl_dt)
                                objPlt = Nothing

                                'Check total bal (for hold stock, becoz hold stock hv no loc)
                                For Each rows As DataRow In pl_dt.Rows
                                    If CDbl(rows.Item("hold_qty")) > 0 Then
                                        If CDbl(rows.Item("total_item_qty")) > CDbl(rows.Item("avail_qty")) Then
                                            If Session("gLang") = "E" Then
                                                errMsg.AppendLine("Error Post in DO No:" & lDO_CODE & ", Item " & gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "") & " | Batch NO:" &
                                                              gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                                              "don't have enough total stock to delivery!")
                                            Else
                                                errMsg.AppendLine("出庫錯誤 提貨號碼:" & lDO_CODE & ", 物料" & gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "") & " | 批次號:" &
                                                              gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                                              "沒有足夠總貨存出貨!")
                                            End If
                                            transaction.Rollback()
                                            gConn.Close()
                                            gConn.Dispose()

                                            UpdateDOStatus = "update WMS_DELV_ORDER set DO_STATUS='PICKED' WHERE DO_STATUS='INPROGRESS'"
                                            gDB.amendData(UpdateDOStatus)

                                            noError = False
                                            safeToGo = False
                                        End If
                                    End If
                                Next

                                If Not safeToGo Then Continue For

                                Dim haveReDrumYN As String = "N"
                                Dim hasCable As Boolean = False
                                'Check have redrum
                                For Each rows As DataRow In pl_dt.Rows
                                    If rows.Item("PLD_TO_DRUM_ID").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim <> "" Then
                                        haveReDrumYN = "Y"
                                    End If

                                    If rows.Item("ITM_TYPE").ToString.Trim = "CABLE" Then
                                        hasCable = True
                                    End If
                                Next

                                'Check for each loc bal
                                For Each rows As DataRow In pl_dt.Rows
                                    Dim SrchStr As String = ""
                                    SrchStr += "select ILOC_BAL_QTY,ILOC_BATCH_NO,ILOC_LOC,ILOC_PALLET_NO, Convert(varchar, ILOC_EXPIRY_DATE, " & gU.getConfig("DDFORMATNo") & ") as ExpDt, Convert(varchar, ILOC_MANU_DATE, " & gU.getConfig("DDFORMATNo") & ") as ManuDt, Convert(datetime, ILOC_EXPIRY_DATE, " & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(datetime,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' "
                                    SrchStr += "AND STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' "
                                    SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                                    SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                                    SrchStr += "AND ILOC_LOC = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_loc").ToString.Trim, "")) & "' "
                                    SrchStr += "AND ILOC_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "' "
                                    'SrchStr += "AND ILOC_PALLET_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, ""), "000")) & "' "

                                    If rows.Item("pld_batch_no").ToString.Trim <> "" Then
                                        SrchStr += "AND ILOC_BATCH_NO = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "")) & "' "
                                    Else
                                        SrchStr += "AND ISNULL(ILOC_BATCH_NO,'') = '' "
                                    End If

                                    Dim qtydt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)


                                    MANU_DATE = qtydt.Rows(0).Item("ManuDt").ToString.Trim 'rows.Item("PLD_MANU_DATE").ToString.Trim
                                    EXP_DATE = qtydt.Rows(0).Item("ExpDt").ToString.Trim 'rows.Item("PLD_EXPIRY_DATE").ToString.Trim

                                    If qtydt.Rows.Count > 0 Then
                                        Dim stQTY As Double = gU.decodeEmptyCdbl(qtydt.Rows(0).Item("ILOC_BAL_QTY").ToString, 0)
                                        Dim plQTY As Double = gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)
                                        Dim ssQTY As Double = gU.decodeEmptyCdbl(rows.Item("pld_ss_qty").ToString.Trim, 0)

                                        If ssQTY > 0 Then
                                            SrchStr = "Select IsNUll(Max(Seq_No),0)+1 as SeqNo from CC_ITEM "
                                            Dim Tmpdt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)
                                            If Tmpdt IsNot Nothing AndAlso Tmpdt.Rows.Count > 0 Then
                                                Dim CC_SEQ_NO = Tmpdt.Rows(0)("SeqNo").ToString
                                                Dim insertSql As String = "insert into CC_ITEM (IMP_CODE,STORER_CODE,ITEM_CODE,SEQ_NO,STATUS,CDate,CByFk,LOT_NO,WH_CODE,WH_LOC)Values('WMS','" + gU.dbEncode(lStorer_code) + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "','" + CC_SEQ_NO.ToString + "','NEW','" + System.DateTime.Now.ToString + "','" + Session("usr_id").ToString + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "')"
                                                gDB.amendData(insertSql, gConn, transaction)
                                            End If
                                        End If

                                        If plQTY = 0 Then
                                            If Session("gLang") = "E" Then
                                                errMsg.AppendLine("Error Post in DO No:" & lDO_CODE & ", Item " & gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "") & " | Batch No:" &
                                                              gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                                              "No Picked Quantity!")
                                            Else
                                                errMsg.AppendLine("出庫錯誤 提貨號碼:" & lDO_CODE & ", 物料" & gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "") & " | 批次號:" &
                                                              gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                                              "實取貨量不能為零!")
                                            End If

                                            'transaction.Rollback()
                                            'gConn.Close()
                                            'gConn.Dispose()
                                            noError = False
                                            safeToGo = False
                                        Else
                                            If stQTY < plQTY Then
                                                Dim qtyTmpdt As DataTable = New DataTable()

                                                SrchStr = "select IsNUll(ILOC_BAL_QTY,0)ILOC_BAL_QTY,ILOC_WH,STORER_CODE,ITM_CODE,ILOC_AREA,ILOC_BATCH_NO,ILOC_PALLET_NO,ILOC_LOC,Convert(datetime,ILOC_EXPIRY_DATE," & DDFORMAT & ") as EXPIRY_DATE, ILOC_EXPIRY_DATE, Convert(datetime,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = 'WMS' "
                                                SrchStr += "AND STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' "
                                                SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                                                SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                                                SrchStr += "AND ILOC_BAL_QTY >= '" & (plQTY - stQTY) & "' "
                                                SrchStr += "AND ILOC_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "' AND ILOC_BAL_QTY > 0 "
                                                SrchStr += "AND ILOC_LOC != '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "' "
                                                SrchStr += "AND ILOC_BATCH_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' Order by ILOC_EXPIRY_DATE desc"
                                                qtyTmpdt = gDB.getDataTable(SrchStr, gConn, transaction)

                                                If qtyTmpdt Is Nothing OrElse qtyTmpdt.Rows.Count = 0 Then
                                                    SrchStr = "select IsNUll(ILOC_BAL_QTY,0)ILOC_BAL_QTY,ILOC_WH,STORER_CODE,ITM_CODE,ILOC_AREA,ILOC_BATCH_NO,ILOC_PALLET_NO,ILOC_LOC,Convert(datetime,ILOC_EXPIRY_DATE," & DDFORMAT & ") as EXPIRY_DATE, ILOC_EXPIRY_DATE, Convert(datetime,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = 'WMS' "
                                                    SrchStr += "AND STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' "
                                                    SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                                                    SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                                                    SrchStr += "AND ILOC_BAL_QTY >= '" & (plQTY - stQTY) & "' "
                                                    SrchStr += "AND ILOC_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "' AND ILOC_BAL_QTY > 0 "
                                                    SrchStr += "AND ILOC_LOC+ILOC_BATCH_NO != '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) + gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' Order by ILOC_EXPIRY_DATE desc"
                                                    qtyTmpdt = gDB.getDataTable(SrchStr, gConn, transaction)
                                                End If

                                                If qtyTmpdt IsNot Nothing AndAlso qtyTmpdt.Rows.Count > 0 AndAlso Convert.ToDouble(qtyTmpdt.Rows(0)("ILOC_BAL_QTY")) >= (plQTY - stQTY) Then
                                                    Dim stNewQTY As Double = 0
                                                    stNewQTY = Convert.ToDouble(qtyTmpdt.Rows(0)("ILOC_BAL_QTY"))

                                                    If (stNewQTY + stQTY) >= plQTY Then
                                                        'INSPOSTDATA(qtyTmpdt.Rows(0)("STORER_CODE").ToString.Trim, qtydt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtydt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_WH").ToString.Trim, qtyTmpdt.Rows(0)("ITM_CODE").ToString.Trim, stQTY + (plQTY - stQTY), stNewQTY - (plQTY - stQTY), stQTY, stNewQTY, 0, qtyTmpdt.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim, qtydt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, transaction, gConn)
                                                        INSPOSTDATA(qtyTmpdt.Rows(0)("STORER_CODE").ToString.Trim, qtydt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtydt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_WH").ToString.Trim, qtyTmpdt.Rows(0)("ITM_CODE").ToString.Trim, stQTY + (plQTY - stQTY), stNewQTY - (plQTY - stQTY), stQTY, stNewQTY, (plQTY - stQTY), qtydt.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim, qtydt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, transaction, gConn)
                                                    End If
                                                    SrchStr = "Select IsNUll(Max(Seq_No),0)+1 as SeqNo from CC_ITEM "
                                                    Dim Tmpdt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)
                                                    If Tmpdt IsNot Nothing AndAlso Tmpdt.Rows.Count > 0 Then
                                                        Dim CC_SEQ_NO = Tmpdt.Rows(0)("SeqNo").ToString
                                                        Dim insertSql As String = "insert into CC_ITEM (IMP_CODE,STORER_CODE,ITEM_CODE,SEQ_NO,STATUS,CDate,CByFk,LOT_NO,WH_CODE,WH_LOC)Values('WMS','" + gU.dbEncode(lStorer_code) + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "','" + CC_SEQ_NO.ToString + "','NEW','" + System.DateTime.Now.ToString + "','" + Session("usr_id").ToString + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "')"
                                                        gDB.amendData(insertSql, gConn, transaction)
                                                    End If
                                                Else
                                                    If Session("gLang") = "E" Then
                                                        errMsg.AppendLine("Error Post in DO No:" & lDO_CODE & ", Item " & gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "") & " | Batch No:" &
                                                              gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                                              "don't have enough stock to delivery!")
                                                    Else
                                                        errMsg.AppendLine("出庫錯誤 提貨號碼:" & lDO_CODE & ", 物料" & gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "") & " | 批次號:" &
                                                              gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                                              "沒有足夠貨存出貨!")
                                                    End If

                                                    'transaction.Rollback()
                                                    'gConn.Close()
                                                    'gConn.Dispose()
                                                End If

                                            End If
                                        End If


                                        If MANU_DATE = "" Then
                                            MANU_DATE = qtydt.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim
                                        End If
                                        If EXP_DATE = "" Then
                                            EXP_DATE = qtydt.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
                                        End If
                                    Else

                                        Dim plQTY As Double = gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)
                                        Dim qtyTmpdt As DataTable = New DataTable()

                                        SrchStr = "select IsNUll(ILOC_BAL_QTY,0)ILOC_BAL_QTY,ILOC_WH,STORER_CODE,ITM_CODE,ILOC_AREA,ILOC_BATCH_NO,ILOC_PALLET_NO,ILOC_LOC,Convert(datetime,ILOC_EXPIRY_DATE," & DDFORMAT & ") as EXPIRY_DATE, ILOC_EXPIRY_DATE, Convert(datetime,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = 'WMS' "
                                        SrchStr += "AND STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' "
                                        SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                                        SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                                        SrchStr += "AND ILOC_BAL_QTY >= '" & plQTY & "' "
                                        SrchStr += "AND ILOC_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "' AND ILOC_BAL_QTY > 0 "
                                        SrchStr += "AND ILOC_LOC != '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "' "
                                        SrchStr += "AND ILOC_BATCH_NO = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' Order by ILOC_EXPIRY_DATE desc"
                                        qtyTmpdt = gDB.getDataTable(SrchStr, gConn, transaction)

                                        If qtyTmpdt Is Nothing OrElse qtyTmpdt.Rows.Count = 0 Then
                                            SrchStr = "select IsNUll(ILOC_BAL_QTY,0)ILOC_BAL_QTY,ILOC_WH,STORER_CODE,ITM_CODE,ILOC_AREA,ILOC_BATCH_NO,ILOC_PALLET_NO,ILOC_LOC,Convert(datetime,ILOC_EXPIRY_DATE," & DDFORMAT & ") as EXPIRY_DATE, ILOC_EXPIRY_DATE, Convert(datetime,ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = 'WMS' "
                                            SrchStr += "AND STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' "
                                            SrchStr += "AND ITM_CODE = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "' "
                                            SrchStr += "AND PACK_KEY = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_pack_key").ToString.Trim, "")) & "' "
                                            SrchStr += "AND ILOC_BAL_QTY >= '" & plQTY & "' "
                                            SrchStr += "AND ILOC_WH = '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "' AND ILOC_BAL_QTY > 0 "
                                            SrchStr += "AND ILOC_LOC+ILOC_BATCH_NO != '" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) + gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "' Order by ILOC_EXPIRY_DATE desc"
                                            qtyTmpdt = gDB.getDataTable(SrchStr, gConn, transaction)
                                        End If

                                        If qtyTmpdt IsNot Nothing AndAlso qtyTmpdt.Rows.Count > 0 AndAlso Convert.ToDouble(qtyTmpdt.Rows(0)("ILOC_BAL_QTY")) > plQTY Then
                                            Dim stQTY = Convert.ToDouble(qtyTmpdt.Rows(0)("ILOC_BAL_QTY"))
                                            If stQTY < plQTY Then

                                                Dim stNewQTY As Double = 0
                                                stNewQTY = Convert.ToDouble(qtyTmpdt.Rows(0)("ILOC_BAL_QTY"))
                                                If stNewQTY >= plQTY Then
                                                    'INSPOSTDATA(qtyTmpdt.Rows(0)("STORER_CODE").ToString.Trim, qtydt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtydt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_WH").ToString.Trim, qtyTmpdt.Rows(0)("ITM_CODE").ToString.Trim, stQTY + (plQTY - stQTY), stNewQTY - (plQTY - stQTY), stQTY, stNewQTY, 0, qtyTmpdt.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim, qtydt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, transaction, gConn)
                                                    INSPOSTDATA(qtyTmpdt.Rows(0)("STORER_CODE").ToString.Trim, qtydt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_BATCH_NO").ToString.Trim, qtydt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_LOC").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_WH").ToString.Trim, qtyTmpdt.Rows(0)("ITM_CODE").ToString.Trim, stQTY + (plQTY - stQTY), stNewQTY - (plQTY - stQTY), stQTY, stNewQTY, (plQTY - stQTY), qtydt.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_EXPIRY_DATE").ToString.Trim, qtydt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, qtyTmpdt.Rows(0)("ILOC_PALLET_NO").ToString.Trim, transaction, gConn)
                                                End If

                                                SrchStr = "Select IsNUll(Max(Seq_No),0)+1 as SeqNo from CC_ITEM "
                                                Dim Tmpdt As DataTable = gDB.getDataTable(SrchStr, gConn, transaction)
                                                If Tmpdt IsNot Nothing AndAlso Tmpdt.Rows.Count > 0 Then
                                                    Dim CC_SEQ_NO = Tmpdt.Rows(0)("SeqNo").ToString
                                                    Dim insertSql As String = "insert into CC_ITEM (IMP_CODE,STORER_CODE,ITEM_CODE,SEQ_NO,STATUS,CDate,CByFk,LOT_NO,WH_CODE,WH_LOC)Values('WMS','" + gU.dbEncode(lStorer_code) + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_item_no").ToString.Trim, "")) & "','" + CC_SEQ_NO.ToString + "','NEW','" + System.DateTime.Now.ToString + "','" + Session("usr_id").ToString + "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_batch_no").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_wh").ToString.Trim, "")) & "','" & gU.dbEncode(gU.decodeNull(rows.Item("pld_loc").ToString.Trim, "")) & "')"
                                                    gDB.amendData(insertSql, gConn, transaction)
                                                End If
                                            End If
                                        Else
                                            If Session("gLang") = "E" Then
                                                errMsg.AppendLine("Error Post in DO No:" & lDO_CODE & ", Item " & gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "") & " | Batch No:" &
                                                          gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                                          "don't have enough stock to delivery!")
                                            Else
                                                errMsg.AppendLine("出庫錯誤 提貨號碼:" & lDO_CODE & ", 物料" & gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "") & " | 批次號:" &
                                                          gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "") & " " &
                                                          "沒有足夠貨存出貨!")
                                            End If

                                            'transaction.Rollback()
                                            'gConn.Close()
                                            'gConn.Dispose()
                                        End If

                                    End If

                                    If Not safeToGo Then Exit For

                                    Dim Remark As String = "DO_CODE: " + rows.Item("do_code").ToString.Trim + ", ITEM_CODE: " + rows.Item("pld_item_no").ToString.Trim + ", PICKED_QTY: " + rows.Item("pld_item_qty").ToString.Trim
                                    Dim insertLog As String = "insert into [dbo].[ActionLogs] values ('" + gU.dbEncode(Remark) + "','DO_POST','WMS_DELV_ORDER',GetDate(),'" + Session("usr_id").ToString + "')"
                                    gDB.amendData(insertLog, gConn, transaction)

                                    st = New StockTrans

                                    st.IO_SEQ = ""
                                    st.STORER_CODE = lStorer_code
                                    st.ITM_CODE = gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "")
                                    st.PACK_KEY = gU.decodeNullOrEmpty(rows.Item("pld_pack_key").ToString.Trim, "")
                                    st.PALLET_NO = gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")
                                    st.IO_CUST_CODE = MainDT.Rows(i).Item("CUS_CODE").ToString.Trim
                                    st.IO_WH = gU.decodeNullOrEmpty(rows.Item("pld_wh").ToString.Trim, "")
                                    st.IO_AREA = gU.decodeNullOrEmpty(rows.Item("pld_area").ToString.Trim, "")
                                    st.IO_LOC = gU.decodeNullOrEmpty(rows.Item("pld_loc").ToString.Trim, "")
                                    st.IO_DOC = "DO"
                                    st.IO_DOC_ID = lDO_CODE
                                    st.IO_QTY = gU.decodeNullOrEmpty(rows.Item("pld_item_qty").ToString.Trim, "")
                                    st.lO_BATCH_NO = gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "")
                                    st.IO_MANU_DATE = MANU_DATE
                                    st.IO_EXPIRY_DATE = EXP_DATE

                                    itmKey = gU.decodeNullOrEmpty(rows.Item("pld_item_no").ToString.Trim, "") & "#_#" &
                                         gU.decodeNullOrEmpty(rows.Item("pld_pack_key").ToString.Trim, "") & "#_#" &
                                         gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "") & "#_#" &
                                         gU.decodeNullOrEmpty(rows.Item("pld_batch_no").ToString.Trim, "")

                                    If qtyDict.ContainsKey(itmKey) Then
                                        Dim tmpCbm As Decimal = 0
                                        Dim tmpwgt As Decimal = 0
                                        Dim tmpqty As Decimal = 0

                                        tmpCbm = cbmDict.Item(itmKey)
                                        tmpwgt = wgtDict.Item(itmKey)
                                        tmpqty = qtyDict.Item(itmKey)

                                        If tmpqty = 0 Then tmpqty = 1

                                        st.IO_CBM = Math.Round(tmpCbm / tmpqty * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
                                        st.IO_KG = Math.Round(tmpwgt / tmpqty * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
                                        'st.IO_CBM = Math.Round(tmpCbm * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
                                        'st.IO_KG = Math.Round(tmpwgt * gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0), 2)
                                    Else
                                        st.IO_CBM = 0
                                        st.IO_KG = 0
                                    End If

                                    If rows.Item("pld_serial_no").ToString.Trim = "" Then
                                        updateSql = "update WMS_CUST_ORDER_D " &
                                                "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) + " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " &
                                                    "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " &
                                                    "sys_lud = Getdate() " &
                                                "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                "and STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' " &
                                                "and CO_CODE = '" & gU.dbEncode(MainDT.Rows(i).Item("DO_CO_CODE").ToString.Trim) & "' " &
                                                "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " &
                                                "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " &
                                                "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' "

                                        gDB.amendData(updateSql, gConn, transaction)

                                    Else
                                        selectSql = "select COD_SEQ as value " &
                                                "from WMS_CUST_ORDER_D " &
                                                "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                "and STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' " &
                                                "and CO_CODE = '" & gU.dbEncode(MainDT.Rows(i).Item("DO_CO_CODE").ToString.Trim) & "' " &
                                                "AND COD_ITM_CODE = '" & gU.dbEncode(rows.Item("PLD_ITEM_NO").ToString.Trim) & "' " &
                                                "AND COD_PACK_KEY = '" & gU.dbEncode(rows.Item("PLD_PACK_KEY").ToString.Trim) & "' " &
                                                "AND ISNULL(COD_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("PLD_PALLET_NO").ToString.Trim, "000")) & "' " &
                                                "AND ISNULL(COD_POST_QTY, 0) < ISNULL(COD_QTY, 0) " &
                                                "order by convert(int, COD_SEQ) "

                                        lCOD_SEQ = DB.getValueFromSQL(selectSql, gConn, transaction)

                                        updateSql = "update WMS_CUST_ORDER_D " &
                                                "set COD_POST_QTY = ISNULL(COD_POST_QTY,0) + " & DB.decodeDBNull(rows.Item("pld_item_qty"), 0) & ", " &
                                                    "sys_lub = '" & gU.dbEncode(Session("usr_id")) & "', " &
                                                    "sys_lud = Getdate() " &
                                                "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                "and STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' " &
                                                "and CO_CODE = '" & gU.dbEncode(MainDT.Rows(i).Item("DO_CO_CODE").ToString.Trim) & "' " &
                                                "and COD_SEQ = '" & gU.dbEncode(lCOD_SEQ) & "' "

                                        gDB.amendData(updateSql, gConn, transaction)

                                    End If

                                    'Update coh_rel_qty for Hold stock logic
                                    updateSql = "update wms_cust_order_hold " &
                                            "set COH_REL_QTY = case when ISNULL(wms_cust_order_hold.coh_in_stock_qty, 0) < ISNULL(COH_REL_QTY,0) + " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " then wms_cust_order_hold.coh_in_stock_qty else ISNULL(COH_REL_QTY,0) + " & gU.dbEncode(gU.decodeEmptyCdbl(rows.Item("pld_item_qty").ToString.Trim, 0)) & " end , " &
                                                "sys_lub = '" & Session("usr_id") & "', " &
                                                "sys_lud = Getdate() " &
                                            "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                            "and STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' " &
                                            "and CO_CODE = '" & gU.dbEncode(MainDT.Rows(i).Item("DO_CO_CODE").ToString.Trim) & "' " &
                                            "and wms_cust_order_hold.coh_status <> 'RELEASE' " &
                                            "and ISNULL(wms_cust_order_hold.coh_in_stock_qty, 0) > 0 " &
                                            "and ISNULL(wms_cust_order_hold.COH_PALLET_NO, '000') = '" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("pld_pallet_no").ToString.Trim, "000")) & "' " &
                                            "and ISNULL(wms_cust_order_hold.COH_BATCH_NO, '') = ISNULL('" & gU.dbEncode(rows.Item("pld_batch_no").ToString.Trim) & "', '') " &
                                            "AND wms_cust_order_hold.COH_ITM_CODE = '" & gU.dbEncode(rows.Item("pld_item_no").ToString.Trim) & "' " &
                                            "AND wms_cust_order_hold.COH_PACK_KEY = '" & gU.dbEncode(rows.Item("pld_pack_key").ToString.Trim) & "' " &
                                            "and exists (" &
                                                "select 1 from WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " &
                                                "where wms_cust_order_hold.imp_code = d.imp_code " &
                                                "and wms_cust_order_hold.storer_code = d.storer_code " &
                                                "and wms_cust_order_hold.co_code = d.co_code " &
                                                "and wms_cust_order_hold.cod_seq = d.cod_seq " &
                                                "and c.imp_code = d.imp_code " &
                                                "and c.storer_code = d.storer_code " &
                                                "and c.co_code = d.co_code " &
                                                "and c.CO_STATUS not in ('CLOSED', 'CANCELLED')) "

                                    gDB.amendData(updateSql, gConn, transaction)

                                    st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)

                                    st.UpdateStockTrans("OUT", gConn, transaction)

                                    st.UpdateStockBalTrans("OUT", gConn, transaction)

                                    If rows.Item("pld_serial_no").ToString.Trim <> "" Then
                                        st.UpdateStockSerialTrans("OUT", gConn, transaction)

                                        st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                                    End If

                                    REM Update Transcation to PickList
                                    updateSql = "update WMS_DO_PICKLIST_D " &
                                            "set PLD_TX_ID = '" & st.IO_SYS_SEQ & "', " &
                                            "sys_lub = '" & Session("usr_id") & "', " &
                                            "sys_lud = Getdate() " &
                                            "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                            "and STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' " &
                                            "and DO_CODE = '" & gU.dbEncode(lDO_CODE) & "' " &
                                            "and PLD_SEQ = '" & gU.dbEncode(rows.Item("PLD_SEQ").ToString.Trim) & "' "

                                    gDB.amendData(updateSql, gConn, transaction)

                                Next

                                If safeToGo Then
                                    For Each rows As DataRow In pl_dt.Rows
                                        If rows.Item("PLD_TO_DRUM_ID").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim <> "" AndAlso rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim <> "" Then
                                            st.cableReDrum(rows.Item("PLD_TO_DRUM_ID").ToString.Trim, rows.Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim, rows.Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim, gConn, transaction)
                                        End If
                                    Next


                                    updateSql = "update WMS_CUST_ORDER " &
                                            "set CO_STATUS = 'CLOSED', " &
                                                "sys_lub = '" & Session("usr_id") & "', " &
                                                "sys_lud = Getdate() " &
                                            "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                "and STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' " &
                                                "and CO_CODE = '" & gU.dbEncode(MainDT.Rows(i).Item("DO_CO_CODE").ToString.Trim) & "' " &
                                            "AND NOT EXISTS (" &
                                                "SELECT 1 FROM WMS_CUST_ORDER_D D " &
                                                "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                                                "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                                                "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                                                "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, '') " &
                                                "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

                                    gDB.amendData(updateSql, gConn, transaction)

                                    updateSql = "update WMS_CUST_ORDER " &
                                            "set CO_STATUS = 'PARTIAL', " &
                                            "sys_lub = '" & Session("usr_id") & "', " &
                                            "sys_lud = Getdate() " &
                                            "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                "and STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' " &
                                                "and CO_CODE = '" & gU.dbEncode(MainDT.Rows(i).Item("DO_CO_CODE").ToString.Trim) & "' " &
                                            "AND EXISTS (" &
                                                "SELECT 1 FROM WMS_CUST_ORDER_D D " &
                                                "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                                                "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                                                "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                                                "AND D.COD_POST_QTY > 0) " &
                                            "AND EXISTS (" &
                                                "SELECT 1 FROM WMS_CUST_ORDER_D D " &
                                                "WHERE D.IMP_CODE = WMS_CUST_ORDER.IMP_CODE " &
                                                "AND D.STORER_CODE = WMS_CUST_ORDER.STORER_CODE " &
                                                "AND D.CO_CODE = WMS_CUST_ORDER.CO_CODE " &
                                                "group by D.IMP_CODE, D.STORER_CODE, D.CO_CODE, D.COD_ITM_CODE, D.COD_PACK_KEY, ISNULL(D.COD_PALLET_NO, '000'), ISNULL(D.COD_BATCH_NO, '') " &
                                                "having max(ISNULL(D.COD_POST_QTY, 0)) < sum(ISNULL(D.COD_QTY, 0))) "

                                    gDB.amendData(updateSql, gConn, transaction)

                                    updateSql = "update WMS_DELV_ORDER " &
                                            "set DO_STATUS = 'POSTED', DO_SAMPLE_CHECKED ='1', " &
                                            "DO_POSTED_DATE = Getdate(), " &
                                            "DO_POSTED_BY = '" & Session("usr_id") & "', " &
                                            "sys_lub = '" & Session("usr_id") & "', " &
                                            "sys_lud = Getdate() " &
                                            "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                            "and STORER_CODE = '" & gU.dbEncode(lStorer_code) & "' " &
                                            "and DO_CODE = '" & gU.dbEncode(lDO_CODE) & "' "

                                    gDB.amendData(updateSql, gConn, transaction)

                                    updateSql = "Delete from WMS_WAVEPICK_RSVD " &
                           " Where DO_CODE='" & gU.dbEncode(lDO_CODE) & "' and imp_code='" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code='" & gU.dbEncode(lStorer_code) & "'"
                                    gDB.amendData(updateSql, gConn, transaction)

                                    transaction.Commit()
                                    transaction = Nothing


                                End If

                            Catch ex As Exception
                                If Not transaction Is Nothing Then
                                    transaction.Rollback()
                                    transaction = Nothing
                                End If

                                noError = False
                                errMsg.AppendLine("Error Post in DO No:" & lDO_CODE & ", Post failed! " & ex.Message)
                            Finally
                                If gConn IsNot Nothing Then
                                    If gConn.State = ConnectionState.Open Then
                                        gConn.Close()
                                        gConn.Dispose()
                                    End If
                                End If
                                UpdateDOStatus = "update WMS_DELV_ORDER set DO_STATUS='PICKED' WHERE DO_STATUS='INPROGRESS'"
                                gDB.amendData(UpdateDOStatus)
                            End Try

                        End If

                    Next

                Else
                    noError = False
                    errMsg.AppendLine("No DO item has been founded")
                End If


                If noError Then
                    errMsg.AppendLine("Check DO has been posted")
                    uiFun.displayMsgNew(udp1, "", "DO Successfully posted!", Session("gLang"))
                Else
                    uiFun.displayMsgNew(udp1, "", "Error Occured! Please check the result.", Session("gLang"))
                End If

                errText.Text = errMsg.ToString

            Else
                If Session("gLang") = "C" Then
                    uiFun.displayMsgNew(udp1, "", "請選取任何提貨單!", Session("gLang"))
                Else
                    uiFun.displayMsgNew(udp1, "", "Please Select a DO!", Session("gLang"))
                End If

            End If

        Else
            If Session("gLang") = "C" Then
                uiFun.displayMsgNew(udp1, "", "請選取任何提貨單!", Session("gLang"))
            Else
                uiFun.displayMsgNew(udp1, "", "Please Select a DO!", Session("gLang"))
            End If
        End If

        Call BindGV()
        ScriptManager.RegisterStartupScript(udp1, udp1.GetType, "SCROLL", "Sys.WebForms.PageRequestManager.getInstance()._scrollPosition = null;window.scrollTo(0, 0);", True)

        Return noError
    End Function

    Public Sub INSPOSTDATA(STORER_CODE As String, BATCH_NO1 As String, BATCH_NO2 As String, LOCATION1 As String, LOCATION2 As String, WH_CODE As String, ADD_ITM_CODE As String, ADD_REV_QTY1 As Double, ADD_REV_QTY2 As Double, ADD_ORG_QTY1 As Double, ADD_ORG_QTY2 As Double, ADD_VAR_QTY As Double, ADD_EXPIRY_DATE1 As DateTime, ADD_EXPIRY_DATE2 As DateTime, ADD_PALLET_NO1 As String, ADD_PALLET_NO2 As String, transaction As SqlTransaction, gConn As SqlConnection)
        Dim cmd As New SqlCommand
        Dim sbCmdText As New StringBuilder
        Dim SQLString As String
        Dim nextNo As String
        Dim updtSql As String
        'Dim gConn = gDB.getConnection()
        'Dim transaction As SqlTransaction
        'transaction = gConn.BeginTransaction()

        Try
            nextNo = DB.getDocNo("SADJ", gConn, transaction)
            SQLString = "Select * from WMS_STOCK_ADJUST " &
                      "WHERE IMP_CODE ='WMS' and STORER_CODE='" & STORER_CODE & "' and AD_CODE='" & nextNo & "'"
            Dim dtROCode As DataTable
            dtROCode = gDB.getDataTable(SQLString, gConn, transaction)
            If dtROCode Is Nothing Or dtROCode.Rows.Count <= 0 Then
                'INSERT
                SQLString = "Insert into WMS_STOCK_ADJUST(IMP_CODE,STORER_CODE,AD_CODE,AD_DATE,AD_TYPE,AD_STATUS,AD_WH,SYS_CD,SYS_LUD,SYS_CB,SYS_LUB) VALUES (@IMP_CODE, @STORER_CODE,@AD_CODE,@AD_DATE,@AD_TYPE,@AD_STATUS,@AD_WH,@SYS_CD,@SYS_LUD,@SYS_CB,@SYS_LUB)"
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                cmd.Parameters.AddWithValue("@STORER_CODE", STORER_CODE)
                cmd.Parameters.AddWithValue("@AD_CODE", nextNo)
                cmd.Parameters.AddWithValue("@AD_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@AD_TYPE", "ADJ")
                cmd.Parameters.AddWithValue("@AD_STATUS", "NEW")
                cmd.Parameters.AddWithValue("@AD_WH", WH_CODE)
                cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                cmd.Parameters.AddWithValue("@SYS_LUB", DBNull.Value)
                cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@SYS_LUD", DBNull.Value)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()

                'for details'
                SQLString = "Insert into WMS_STOCK_ADJUST_D(IMP_CODE,STORER_CODE,AD_CODE,AD_SEQ,ADD_PACK_KEY,ADD_ORG_QTY,ADD_REV_QTY,ADD_EXPIRY_DATE,ADD_ITM_CODE,ADD_VAR_QTY,ADD_BATCH_NO,ADD_LOC,ADD_ORG_QTY2,ADD_REV_QTY2,ADD_VAR_QTY2,SYS_CD,SYS_LUD,SYS_CB,SYS_LUB,ADD_PALLET_NO) VALUES (@IMP_CODE,@STORER_CODE,@AD_CODE,@AD_SEQ,@ADD_PACK_KEY,@ADD_ORG_QTY,@ADD_REV_QTY,@ADD_EXPIRY_DATE,@ADD_ITM_CODE,@ADD_VAR_QTY,@ADD_BATCH_NO,@ADD_LOC,@ADD_ORG_QTY2,@ADD_REV_QTY2,@ADD_VAR_QTY2,@SYS_CD,@SYS_LUD,@SYS_CB,@SYS_LUB,@ADD_PALLET_NO)"
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                cmd.Parameters.AddWithValue("@STORER_CODE", STORER_CODE)
                cmd.Parameters.AddWithValue("@AD_CODE", nextNo)
                cmd.Parameters.AddWithValue("@AD_SEQ", "1")
                cmd.Parameters.AddWithValue("@ADD_PACK_KEY", "1")
                cmd.Parameters.AddWithValue("@ADD_ITM_CODE", ADD_ITM_CODE)
                cmd.Parameters.AddWithValue("@ADD_BATCH_NO", BATCH_NO1)
                cmd.Parameters.AddWithValue("@ADD_LOC", LOCATION1)
                cmd.Parameters.AddWithValue("@ADD_ORG_QTY", ADD_ORG_QTY1)
                cmd.Parameters.AddWithValue("@ADD_REV_QTY", ADD_REV_QTY1)
                cmd.Parameters.AddWithValue("@ADD_VAR_QTY", (ADD_REV_QTY1 - ADD_ORG_QTY1))
                cmd.Parameters.AddWithValue("@ADD_ORG_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_REV_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_VAR_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_EXPIRY_DATE", ADD_EXPIRY_DATE1)
                cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                cmd.Parameters.AddWithValue("@SYS_LUB", DBNull.Value)
                cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@SYS_LUD", DBNull.Value)
                cmd.Parameters.AddWithValue("@ADD_PALLET_NO", ADD_PALLET_NO1)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()

                SQLString = "Insert into WMS_STOCK_ADJUST_D(IMP_CODE,STORER_CODE,AD_CODE,AD_SEQ,ADD_PACK_KEY,ADD_ORG_QTY,ADD_REV_QTY,ADD_EXPIRY_DATE,ADD_ITM_CODE,ADD_VAR_QTY,ADD_BATCH_NO,ADD_LOC,ADD_ORG_QTY2,ADD_REV_QTY2,ADD_VAR_QTY2,SYS_CD,SYS_LUD,SYS_CB,SYS_LUB,ADD_PALLET_NO) VALUES (@IMP_CODE,@STORER_CODE,@AD_CODE,@AD_SEQ,@ADD_PACK_KEY,@ADD_ORG_QTY,@ADD_REV_QTY,@ADD_EXPIRY_DATE,@ADD_ITM_CODE,@ADD_VAR_QTY,@ADD_BATCH_NO,@ADD_LOC,@ADD_ORG_QTY2,@ADD_REV_QTY2,@ADD_VAR_QTY2,@SYS_CD,@SYS_LUD,@SYS_CB,@SYS_LUB,@ADD_PALLET_NO)"
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@IMP_CODE", "WMS")
                cmd.Parameters.AddWithValue("@STORER_CODE", STORER_CODE)
                cmd.Parameters.AddWithValue("@AD_CODE", nextNo)
                cmd.Parameters.AddWithValue("@AD_SEQ", "2")
                cmd.Parameters.AddWithValue("@ADD_PACK_KEY", "1")
                cmd.Parameters.AddWithValue("@ADD_ITM_CODE", ADD_ITM_CODE)
                cmd.Parameters.AddWithValue("@ADD_BATCH_NO", BATCH_NO2)
                cmd.Parameters.AddWithValue("@ADD_LOC", LOCATION2)
                cmd.Parameters.AddWithValue("@ADD_ORG_QTY", ADD_ORG_QTY2)
                cmd.Parameters.AddWithValue("@ADD_REV_QTY", ADD_REV_QTY2)
                cmd.Parameters.AddWithValue("@ADD_VAR_QTY", (ADD_REV_QTY2 - ADD_ORG_QTY2))
                cmd.Parameters.AddWithValue("@ADD_ORG_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_REV_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_VAR_QTY2", 0)
                cmd.Parameters.AddWithValue("@ADD_EXPIRY_DATE", ADD_EXPIRY_DATE2)
                cmd.Parameters.AddWithValue("@SYS_CB", Session("usr_id"))
                cmd.Parameters.AddWithValue("@SYS_LUB", DBNull.Value)
                cmd.Parameters.AddWithValue("@SYS_CD", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@SYS_LUD", DBNull.Value)
                cmd.Parameters.AddWithValue("@ADD_PALLET_NO", ADD_PALLET_NO2)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()
            End If

            SQLString = "SELECT [IMP_CODE],[STORER_CODE],[AD_CODE],[AD_SEQ],[ADD_ITM_CODE],[ADD_PACK_KEY] " &
",[ADD_LOC],[ADD_ORG_QTY],[ADD_REV_QTY],[ADD_VAR_QTY],[ADD_REM],[ADD_PALLET_NO] " &
",[SYS_LUB],[SYS_LUD],[SYS_CD],[SYS_CB],[ADD_BATCH_NO] " &
",[ADD_VND_CODE],convert(varchar(10),[ADD_EXPIRY_DATE],103) ADD_EXPIRY_DATE,convert(varchar(10),[ADD_MANU_DATE],103)ADD_MANU_DATE,[ADD_ORG_QTY2],[ADD_REV_QTY2] " &
",[ADD_VAR_QTY2],[ADD_SERIAL_NO],[DRUM_ID],[DRUM_LEVEL]FROM[dbo].[WMS_STOCK_ADJUST_D] " &
                      "WHERE IMP_CODE ='WMS' and STORER_CODE='" & STORER_CODE & "' and AD_CODE='" & nextNo & "'"
            Dim dtRODCode As DataTable
            dtRODCode = gDB.getDataTable(SQLString, gConn, transaction)

            If dtRODCode IsNot Nothing AndAlso dtRODCode.Rows.Count > 0 Then
                For Each rows As DataRow In dtRODCode.Rows
                    'FOR POSTING'
                    st.STORER_CODE = STORER_CODE
                    st.ITM_CODE = gU.decodeNull(rows.Item("add_itm_code").ToString.Trim, "")
                    st.PACK_KEY = gU.decodeNull(rows.Item("add_pack_key").ToString.Trim, "")
                    st.IO_CUST_CODE = ""
                    st.IO_AREA = ""
                    st.IO_DOC = "SADJ"
                    st.IO_DOC_ID = nextNo
                    st.IO_CBM = 0
                    st.IO_KG = 0
                    st.IO_WH = WH_CODE
                    st.PALLET_NO = gU.decodeNull(rows.Item("add_pallet_no").ToString.Trim, "")
                    st.IO_LOC = gU.decodeNull(rows.Item("add_loc").ToString.Trim, "")
                    st.lO_BATCH_NO = gU.decodeNull(rows.Item("add_batch_no").ToString.Trim, "")
                    st.IO_EXPIRY_DATE = gU.decodeNull(rows.Item("add_expiry_date").ToString.Trim, "")
                    st.IO_MANU_DATE = gU.decodeNull(rows.Item("add_manu_date").ToString.Trim, "")

                    Dim txQty As Double = 0

                    If gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") >
                        gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0")

                        st.IO_QTY = txQty
                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                        st.UpdateStockTrans("IN", gConn, transaction)
                        st.UpdateStockBalTrans("IN", gConn, transaction)

                        If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                            st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                            st.IOS_QTY2 = 1
                            st.UpdateStockSerialTrans("IN", gConn, transaction)
                            st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                        End If

                    ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") <
                        gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0")

                        st.IO_QTY = txQty
                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                        st.UpdateStockTrans("OUT", gConn, transaction)
                        st.UpdateStockBalTrans("OUT", gConn, transaction)

                        If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                            st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                            st.IOS_QTY2 = 1
                            st.UpdateStockSerialTrans("OUT", gConn, transaction)
                            st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                        End If

                    ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") >
                        gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0")

                        Dim drumDt As DataTable
                        SQLString = "SELECT ILBS_DRUM_ID, ILBS_DRUM_LEVEL FROM WMS_ITEM_LOC_BAL_S WHERE " &
                             "IMP_CODE = 'WMS' " &
                                    "AND STORER_CODE = '" & STORER_CODE & "' " &
                                    "AND ITM_CODE = '" & gU.decodeNull(ADD_ITM_CODE, "") & "' " &
                                    "AND PACK_KEY = '1' " &
                                    "AND ILBS_SERIAL_NO = '" & gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") & "' AND ILBS_QTY2 > 0 "
                        REM **********************

                        drumDt = gDB.getDataTable(SQLString, gConn, transaction)

                        If drumDt.Rows.Count > 0 Then
                            st.IOS_DRUM_ID = drumDt.Rows(0).Item("ILBS_DRUM_ID").ToString
                            st.IOS_DRUM_LEVEL = drumDt.Rows(0).Item("ILBS_DRUM_LEVEL")
                        End If

                        st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                        st.IOS_QTY2 = txQty
                        st.IO_QTY = 1
                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                        st.UpdateStockTrans("IN", gConn, transaction)
                        st.UpdateStockBalTrans("IN", gConn, transaction)
                        If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                            st.IOS_SL = "Y"
                            st.UpdateStockSerialTrans("IN", gConn, transaction)
                            st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                        End If

                    ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0") <
                        gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") Then

                        txQty = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty2").ToString.Trim, ""), "0") - gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty2").ToString.Trim, ""), "0")
                        st.IOS_SERIAL_NO = gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "")
                        st.IOS_QTY2 = txQty
                        st.IO_QTY = 1
                        Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, ""), gConn, transaction)
                        st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                        st.UpdateStockTrans("OUT", gConn, transaction)
                        st.UpdateStockBalTrans("OUT", gConn, transaction)
                        If gU.decodeNull(rows.Item("ADD_SERIAL_NO").ToString.Trim, "") <> "" Then
                            st.IOS_SL = "Y"
                            st.UpdateStockSerialTrans("OUT", gConn, transaction)
                            st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                        End If

                    ElseIf gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_rev_qty").ToString.Trim, ""), "0") =
                        gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("add_org_qty").ToString.Trim, ""), "0") Then

                        'If transaction IsNot Nothing Then
                        '    transaction.Rollback()
                        'End If
                        Exit Sub
                    End If
                Next

            End If

            updtSql = "update wms_stock_adjust " &
                            "set ad_status = 'POSTED' " &
                            "where imp_code = 'WMS' " &
                            "and storer_code = '" & STORER_CODE & "' " &
                            "and ad_code = '" & nextNo & "' "

            gDB.amendData(updtSql, gConn, transaction)

            'For STORER 
            Dim IO_CODE As String
            SQLString = "select Top(1) IO_CODE from EBS_WMS_COMPANY_MASTER where IO_ID = '" & STORER_CODE & "' "
            Dim dtIOCode As DataTable
            dtIOCode = gDB.getDataTable(SQLString, gConn, transaction)
            If dtIOCode IsNot Nothing AndAlso dtIOCode.Rows.Count > 0 Then
                IO_CODE = dtIOCode.Rows(0)("IO_CODE").ToString.Trim
            Else
                IO_CODE = ""
            End If

            'For ITEM NUMBER
            Dim ITEM_NUMBER As String
            Dim ITEM_UOM As String
            SQLString = "select Top(1) ITEM_NUMBER,PRIMARY_UOM_CODE from EBS_WMS_ITEM_MASTER where IO_ID = '" & STORER_CODE & "' and ITEM_ID='" & ADD_ITM_CODE & "'  "
            Dim dtITEMNUMBER As DataTable
            dtITEMNUMBER = gDB.getDataTable(SQLString, gConn, transaction)
            If dtITEMNUMBER IsNot Nothing AndAlso dtITEMNUMBER.Rows.Count > 0 Then
                ITEM_NUMBER = dtITEMNUMBER.Rows(0)("ITEM_NUMBER").ToString.Trim
                ITEM_UOM = dtITEMNUMBER.Rows(0)("PRIMARY_UOM_CODE").ToString.Trim
            Else
                ITEM_NUMBER = ""
                ITEM_UOM = ""
            End If

            'FOR SEQ NO.'
            Dim dtADJJSEQ As New DataTable
            SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT"
            dtADJJSEQ = gDB.getDataTable(SQLString, gConn, transaction)

            'FOR BATCH_NO'
            Dim dtADJBatchNo As New DataTable
            SQLString = "Select REPLACE('WMSLOT'+Convert(nvarchar(20),Convert(bigint,IsNUll('1'+IsNUll( Max(Substring(BATCH_NO,7,LEN(BATCH_NO))),'000000000000'),0))+1),'WMSLOT1','WMSLOT') as BATCH_NO From WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT"
            dtADJBatchNo = gDB.getDataTable(SQLString, gConn, transaction)

            'for LOCATION_ID'
            SQLString = " Select a.IO_ID,a.DESCRIPTION as LOCDESC, a.INVENTORY_LOCATION_ID as LOCID,a.SUBINVENTORY_CODE as WHCODE from EBS_WMS_WAREHOUSE_LOCATION a Where (a.SUBINVENTORY_CODE+ REPLACE(a.LOCATOR,'.',''))='" + WH_CODE + LOCATION2 + "' and IO_ID='" + STORER_CODE + "'"
            Dim dtLOC As New DataTable
            dtLOC = gDB.getDataTable(SQLString, gConn, transaction)
            Dim LocationId As String = "0"
            If dtLOC IsNot Nothing AndAlso dtLOC.Rows.Count > 0 Then
                LocationId = dtLOC.Rows(0)("LOCID").ToString()
            End If

            'for LOCATOR'
            Dim LOC As String
            LOC = LOCATION2
            Dim location As String = LOC.Substring(0, 2) + "." + LOC.Substring(2, 2) + "." + LOC.Substring(4, 2) + "." + LOC.Substring(6, 2)

            'for TRANSACTION_ID as use seq_no'
            Dim dtSeq As New DataTable
            SQLString = "Select IsNUll(Max(TRANSACTION_ID),0)+1 as TRANSACTION_ID from WMS_EBS_TRANS_ITX_ACTION"
            dtSeq = gDB.getDataTable(SQLString, gConn, transaction)

            'CompanyId
            Dim CompanyId As Integer
            Dim dtCompanyData As New DataTable
            SQLString = "select Top(1) Isnull(COMPANY_ID,0) COMPANY_ID from EBS_WMS_COMPANY_MASTER where IO_ID=" & STORER_CODE & ""
            dtCompanyData = gDB.getDataTable(SQLString, gConn, transaction)
            If (dtCompanyData IsNot Nothing AndAlso dtCompanyData.Rows.Count > 0) Then
                CompanyId = dtCompanyData.Rows(0)("COMPANY_ID")
            Else
                CompanyId = 0
            End If

            'for WMS_EBS_TRANS_ITX_ACTION'
            'INSERT
            SQLString = "Insert into WMS_EBS_TRANS_ITX_ACTION(SOURCE,ACTION,TRANSACTION_ID,IO_ID,DOC_TYPE,STATUS,BATCH_NO,COMPANY_ID,CREATION_DATE,LAST_UPDATE_DATE) VALUES (@SOURCE,@ACTION,@TRANSACTION_ID,@IO_ID,@DOC_TYPE,@STATUS,@BATCH_NO,@COMPANY_ID,@CREATION_DATE,@LAST_UPDATE_DATE)"
            cmd = New SqlCommand(SQLString, gConn, transaction)
            cmd.Parameters.AddWithValue("@SOURCE", "WMS")
            cmd.Parameters.AddWithValue("@ACTION", "NEW")
            cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
            cmd.Parameters.AddWithValue("@IO_ID", STORER_CODE)
            cmd.Parameters.AddWithValue("@DOC_TYPE", "LOT_ADJ")
            cmd.Parameters.AddWithValue("@COMPANY_ID", CompanyId)
            cmd.Parameters.AddWithValue("@STATUS", "NEW")
            cmd.Parameters.AddWithValue("@BATCH_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
            cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
            cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
            cmd.CommandType = System.Data.CommandType.Text
            cmd.ExecuteScalar()

            SQLString = "Select * from WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT " &
                      "WHERE TRANSACTION_ID ='" + dtSeq.Rows(0)("TRANSACTION_ID").ToString + "'"
            Dim dtSTKADJCode As DataTable
            dtSTKADJCode = gDB.getDataTable(SQLString, gConn, transaction)
            If dtSTKADJCode Is Nothing Or dtSTKADJCode.Rows.Count <= 0 Then
                'for Issue'
                'Create a WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT record 
                SQLString = "Insert into WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,TRANSACTION_TYPE,DOC_TYPE,IO_CODE,LOCATOR,LOCATION_ID,LOT_NUMBER,ITEM_NUMBER,SUBINVENTORY_CODE,QUANTITY,UOM_CODE,LAST_UPDATE_DATE,LAST_UPDATE_BY,CREATION_DATE,CREATION_BY) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@TRANSACTION_TYPE,@DOC_TYPE,@IO_CODE,@LOCATOR,@LOCATION_ID,@LOT_NUMBER,@ITEM_NUMBER,@SUBINVENTORY_CODE,@QUANTITY,@UOM_CODE,@LAST_UPDATE_DATE,@LAST_UPDATE_BY,@CREATION_DATE,@CREATION_BY)"
                'UOM_CODE,@UOM_CODE
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@SEQ_NO", dtADJJSEQ.Rows(0)("SEQ_NO").ToString.Trim)
                cmd.Parameters.AddWithValue("@BATCH_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                cmd.Parameters.AddWithValue("@TRANSACTION_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@TRANSACTION_TYPE", "I")
                cmd.Parameters.AddWithValue("@DOC_TYPE", "LOT_ADJ")
                cmd.Parameters.AddWithValue("@IO_CODE", IO_CODE)
                cmd.Parameters.AddWithValue("@SUBINVENTORY_CODE", WH_CODE)
                cmd.Parameters.AddWithValue("@LOCATOR", location)
                cmd.Parameters.AddWithValue("@LOCATION_ID", LocationId)
                cmd.Parameters.AddWithValue("@LOT_NUMBER", BATCH_NO2)
                cmd.Parameters.AddWithValue("@ITEM_NUMBER", ITEM_NUMBER)
                cmd.Parameters.AddWithValue("@QUANTITY", ADD_VAR_QTY)
                cmd.Parameters.AddWithValue("@UOM_CODE", ITEM_UOM)
                cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_id"))
                cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@LAST_UPDATE_BY", DBNull.Value)
                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()

                'for Receive'
                'FOR SEQ NO.'
                Dim dtADJSEQ As New DataTable
                SQLString = "Select IsNUll(Max(SEQ_NO),0)+1 as SEQ_NO from WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT"
                dtADJSEQ = gDB.getDataTable(SQLString, gConn, transaction)

                'for LOCATOR'
                Dim RCVLOC As String
                RCVLOC = LOCATION1
                Dim RCVlocation As String = RCVLOC.Substring(0, 2) + "." + RCVLOC.Substring(2, 2) + "." + RCVLOC.Substring(4, 2) + "." + RCVLOC.Substring(6, 2)

                'for LOCATION_ID'
                SQLString = " Select a.IO_ID,a.DESCRIPTION as LOCDESC, a.INVENTORY_LOCATION_ID as LOCID,a.SUBINVENTORY_CODE as WHCODE from EBS_WMS_WAREHOUSE_LOCATION a Where (a.SUBINVENTORY_CODE+ REPLACE(a.LOCATOR,'.',''))='" + WH_CODE + LOCATION1 + "' and IO_ID='" + STORER_CODE + "'"
                Dim dtRCVLOC As New DataTable
                dtRCVLOC = gDB.getDataTable(SQLString, gConn, transaction)
                Dim RCVLocationId As String = "0"
                If dtRCVLOC IsNot Nothing AndAlso dtRCVLOC.Rows.Count > 0 Then
                    RCVLocationId = dtRCVLOC.Rows(0)("LOCID").ToString()
                End If

                'Create a WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT record 
                SQLString = "Insert into WMS_EBS_TRANS_ITX_STOCK_ADJUSTMENT(SEQ_NO,BATCH_NO,TRANSACTION_ID,TRANSACTION_DATE,TRANSACTION_TYPE,DOC_TYPE,IO_CODE,LOCATOR,LOCATION_ID,LOT_NUMBER,ITEM_NUMBER,SUBINVENTORY_CODE,QUANTITY,UOM_CODE,LAST_UPDATE_DATE,LAST_UPDATE_BY,CREATION_DATE,CREATION_BY,EXPIRY_DATE) VALUES (@SEQ_NO,@BATCH_NO,@TRANSACTION_ID,@TRANSACTION_DATE,@TRANSACTION_TYPE,@DOC_TYPE,@IO_CODE,@LOCATOR,@LOCATION_ID,@LOT_NUMBER,@ITEM_NUMBER,@SUBINVENTORY_CODE,@QUANTITY,@UOM_CODE,@LAST_UPDATE_DATE,@LAST_UPDATE_BY,@CREATION_DATE,@CREATION_BY,@EXPIRY_DATE)"
                'UOM_CODE,@UOM_CODE
                cmd = New SqlCommand(SQLString, gConn, transaction)
                cmd.Parameters.AddWithValue("@SEQ_NO", dtADJSEQ.Rows(0)("SEQ_NO").ToString.Trim)
                cmd.Parameters.AddWithValue("@BATCH_NO", dtADJBatchNo.Rows(0)("BATCH_NO").ToString.Trim)
                cmd.Parameters.AddWithValue("@TRANSACTION_ID", dtSeq.Rows(0)("TRANSACTION_ID").ToString.Trim)
                cmd.Parameters.AddWithValue("@TRANSACTION_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@TRANSACTION_TYPE", "R")
                cmd.Parameters.AddWithValue("@DOC_TYPE", "LOT_ADJ")
                cmd.Parameters.AddWithValue("@IO_CODE", IO_CODE)
                cmd.Parameters.AddWithValue("@SUBINVENTORY_CODE", WH_CODE)
                cmd.Parameters.AddWithValue("@LOCATOR", RCVlocation)
                cmd.Parameters.AddWithValue("@LOCATION_ID", RCVLocationId)
                cmd.Parameters.AddWithValue("@LOT_NUMBER", BATCH_NO1)
                cmd.Parameters.AddWithValue("@ITEM_NUMBER", ITEM_NUMBER)
                cmd.Parameters.AddWithValue("@QUANTITY", ADD_VAR_QTY)
                cmd.Parameters.AddWithValue("@UOM_CODE", ITEM_UOM)
                cmd.Parameters.AddWithValue("@CREATION_BY", Session("usr_id"))
                cmd.Parameters.AddWithValue("@CREATION_DATE", System.DateTime.Now)
                cmd.Parameters.AddWithValue("@LAST_UPDATE_BY", DBNull.Value)
                cmd.Parameters.AddWithValue("@LAST_UPDATE_DATE", DBNull.Value)
                cmd.Parameters.AddWithValue("@EXPIRY_DATE", ADD_EXPIRY_DATE1)
                cmd.CommandType = System.Data.CommandType.Text
                cmd.ExecuteScalar()
            End If


        Catch ex As Exception
            Throw ex

        End Try

    End Sub

End Class

