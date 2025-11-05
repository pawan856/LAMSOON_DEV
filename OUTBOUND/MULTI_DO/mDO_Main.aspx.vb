Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OUTBOUND_MULTI_DO_mDO_Main
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private dl As New DocLink
    Private moduleAction As String = ""
    Private DDFORMAT2 As String = "dd/MM/yyyy"
    Private DDFORMAT As String = gU.getConfig("DDFORMATNO")
    Private imp_code As String = ""

    Private dt As New DataTable
    Private exceptionEditList As List(Of String)

    Private Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        DDFORMAT2 = gU.getConfig("DDFORMAT2")
        'DDFORMAT = gU.getConfig("DDFORMAT")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_MDO", Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If Not Session("IMP_CODE") Is Nothing Then imp_code = Session("IMP_CODE")

        If Not IsPostBack Then
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            ViewState("dt") = Nothing

            uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
        End If

        If Session("gLang") = "E" Then
            lheader.Text = "Batch Create DO:"
            lbl_STORER_CODE.Text = "Storer Code"
            lbl_CO_EDI_SIR_NO.Text = "Customer Order No."
            lbl_CO_DATE.Text = "CO Date"
            lblGV.Text = "CO List"


            btnSearch.Text = "Search"
            btnReset.Text = "Reset"
            btnSelAll.Text = "Select All"
            btnUnSel.Text = "Unselect All"

            BtnGen.Text = "Generate DO"
            cfm1.ConfirmText = "New DOs will be generated from selected CO, Confirm to proceed?"

            btnGen2.Text = "Generate DO"
            cfm2.ConfirmText = "New DOs will be generated from selected CO, Confirm to proceed?"

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "批次產生送貨指令"
            lbl_STORER_CODE.Text = "貨主"
            lbl_CO_EDI_SIR_NO.Text = "訂單號碼"
            lbl_CO_DATE.Text = "訂單日期"

            lblGV.Text = "客戶訂單"
            btnSearch.Text = "搜尋"
            btnReset.Text = "重置"
            btnSelAll.Text = "全選"
            btnUnSel.Text = "全取消"

            BtnGen.Text = "產生送貨指令"
            cfm1.ConfirmText = "將會產生送貨單,確認?"

            btnGen2.Text = "產生送貨指令"
            cfm2.ConfirmText = "將會產生送貨單,確認?"
        End If

    End Sub

    Protected Sub BindGV()
        Dim tempStr As String = ""
        Dim selectSQL As String = ""
        Dim dt As DataTable

        If STORER_CODE.SelectedValue <> "" Then
            tempStr = " AND WMS_CUST_ORDER.STORER_CODE='" & gU.dbEncode(STORER_CODE.SelectedValue) & "'"
        End If

        Dim keyArr As String()
        Dim FrToArr As String()
        Dim tempApp As String = ""
        Dim tempApp2 As String = ""


        If Not String.IsNullOrWhiteSpace(CO_EDI_SIR_NO.Text) Then
            If InStr(CO_EDI_SIR_NO.Text.Trim, "-") > 0 Then
                If InStr(CO_EDI_SIR_NO.Text.Trim, ",") > 0 Then
                    keyArr = gU.listToArray(CO_EDI_SIR_NO.Text.Trim)

                    For x = 0 To keyArr.Length - 1
                        If InStr(keyArr(x), "-") > 0 Then
                            FrToArr = keyArr(x).Split("-")

                            tempApp2 &= "(WMS_CUST_ORDER.CO_EDI_SIR_NO >='" & FrToArr(0).Trim & "' AND WMS_CUST_ORDER.CO_EDI_SIR_NO <='" & FrToArr(1).Trim & "') OR "

                        Else
                            tempApp = gU.appendToList(tempApp, "'" & keyArr(x).Trim & "'")
                        End If
                    Next

                    tempApp2 = Left(tempApp2, tempApp2.Length - 3)

                    If tempApp <> "" Then
                        tempStr &= " AND (WMS_CUST_ORDER.CO_EDI_SIR_NO IN (" & tempApp & ") OR (" & tempApp2 & ")) "
                    Else
                        tempStr &= " AND (" & tempApp2 & ")"
                    End If

                Else
                    FrToArr = CO_EDI_SIR_NO.Text.Trim.Split("-")

                    tempStr &= " AND (WMS_CUST_ORDER.CO_EDI_SIR_NO >='" & FrToArr(0).Trim & "' AND WMS_CUST_ORDER.CO_EDI_SIR_NO <='" & FrToArr(1).Trim & "') "

                End If

            ElseIf InStr(CO_EDI_SIR_NO.Text.Trim, ",") > 0 Then
                keyArr = gU.listToArray(CO_EDI_SIR_NO.Text.Trim)

                For x = 0 To keyArr.Length - 1
                    tempApp = gU.appendToList(tempApp, "'" & keyArr(x) & "'")
                Next

                tempStr &= " AND (WMS_CUST_ORDER.CO_EDI_SIR_NO IN (" & tempApp & ")) "

            Else
                tempStr &= " AND WMS_CUST_ORDER.CO_EDI_SIR_NO='" & CO_EDI_SIR_NO.Text.Trim & "' "
            End If

        End If

        If CO_DATE_FR.Text.Trim <> "" Then
            tempStr &= " AND CO_DATE >= CONVERT(datetime,'" & CO_DATE_FR.Text.Trim & "'," & DDFORMAT & ") "
        End If

        If CO_DATE_TO.Text.Trim <> "" Then
            tempStr &= " AND CO_DATE < CONVERT(datetime,'" & CO_DATE_TO.Text.Trim & "'," & DDFORMAT & ") + 1 "
        End If


        selectSQL = " SELECT DISTINCT WMS_CUST_ORDER.IMP_CODE, WMS_CUST_ORDER.STORER_CODE, WMS_CUST_ORDER.CO_CODE, WMS_CUST_ORDER.CO_STATUS,CONVERT(varchar, WMS_CUST_ORDER.CO_DATE," & DDFORMAT & " ) as CO_DATE, WMS_CUST_ORDER.CUS_CODE, " & _
                    " WMS_CUST_ORDER.CUS_NAME, WMS_CUST_ORDER.CO_EDI_SIR_NO,wms_storer.sto_name, '' as checkYN " & _
                    " FROM WMS_CUST_ORDER LEFT OUTER JOIN " & _
                    " WMS_CUST_ORDER_D ON WMS_CUST_ORDER.IMP_CODE = WMS_CUST_ORDER_D.IMP_CODE AND WMS_CUST_ORDER.STORER_CODE = WMS_CUST_ORDER_D.STORER_CODE AND " & _
                    " WMS_CUST_ORDER.CO_CODE = WMS_CUST_ORDER_D.CO_CODE INNER JOIN " & _
                    " WMS_STORER ON WMS_CUST_ORDER.IMP_CODE = WMS_STORER.IMP_CODE AND WMS_CUST_ORDER.STORER_CODE = WMS_STORER.STORER_CODE " & _
                    " WHERE CO_STATUS NOT in ('CLOSED','CANCELLED') AND (NOT EXISTS (Select 1 from wms_delv_order where WMS_CUST_ORDER.imp_code = WMS_DELV_ORDER.IMP_CODE AND WMS_CUST_ORDER.STORER_code = WMS_DELV_ORDER.STORER_CODE " & _
                    " and WMS_CUST_ORDER.CO_EDI_SIR_NO = WMS_DELV_ORDER.DO_EDI_SIR_NO AND DO_STATUS NOT in ('CANCELLED')) OR isNULL(CO_EDI_SIR_NO,'') = '') " & _
                    tempStr & _
                    " ORDER BY CO_CODE DESC"

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
        Call cU.newChangeGVLabel(GridView1, "Storer", "貨主")
        Call cU.newChangeGVLabel(GridView1, "CO Code", "客戶指令")
        Call cU.newChangeGVLabel(GridView1, "Customer Order Number", "訂單號碼")
        Call cU.newChangeGVLabel(GridView1, "CO Date", "日期")
        Call cU.newChangeGVLabel(GridView1, "Customer No.", "客戶號碼")
        Call cU.newChangeGVLabel(GridView1, "Customer Name", "客戶名稱")

        REM **********************
    End Sub


    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
         Select e.Row.RowType
            Case DataControlRowType.DataRow
                If DataBinder.Eval(e.Row.DataItem, "checkYN").ToString.Trim = "Y" Then
                    CType(e.Row.FindControl("checkYN"), CheckBox).Text = True
                End If

                CType(e.Row.FindControl("sto_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "sto_name").ToString.Trim
                CType(e.Row.FindControl("STORER_CODE"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "STORER_CODE").ToString.Trim

                CType(e.Row.FindControl("CO_CODE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CO_CODE").ToString.Trim
                CType(e.Row.FindControl("CO_EDI_SIR_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CO_EDI_SIR_NO").ToString.Trim
                CType(e.Row.FindControl("CO_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CO_DATE").ToString.Trim
                CType(e.Row.FindControl("CUS_CODE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CUS_CODE").ToString.Trim
                CType(e.Row.FindControl("CUS_NAME"), Label).Text = DataBinder.Eval(e.Row.DataItem, "CUS_NAME").ToString.Trim

        End Select



    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As System.EventArgs) Handles btnSearch.Click
        BindGV()
    End Sub

    Protected Sub MulitAddCOtoDO()
        Dim checkedCO As String = ""
        Dim selectSQL As String = ""
        Dim tempDT As DataTable
        Dim updateSQL As String = ""
        Dim successFlag As Boolean = False
        Dim PL_WH_CODE As String = ""

        Dim gConn As SqlConnection

        If GridView1 IsNot Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If DirectCast(GridView1.Rows(i).FindControl("checkYN"), CheckBox).Checked Then
                    checkedCO = gU.appendToList(checkedCO, "'" & DirectCast(GridView1.Rows(i).FindControl("STORER_CODE"), HiddenField).Value & "_000_" & DirectCast(GridView1.Rows(i).FindControl("CO_CODE"), Label).Text & "'")
                End If
            Next

            If checkedCO <> "" Then
                gConn = gDB.getConnection()

                Dim transaction As SqlTransaction
                transaction = gConn.BeginTransaction
                ' Start a local transaction
                Try
                    selectSQL = "SELECT d.imp_code, d.storer_code, d.CO_CODE, d.COD_SEQ, d.COD_QTY, m.CO_REM, m.CO_SHIP_MODE, Convert(varchar,m.CO_TARGET_DELDATE," & DDFORMAT & ") as CO_TARGET_DELDATE, " & _
                                "d.COD_DISP_SEQ, d.COD_PALLET_NO, d.COD_ITM_CODE, d.COD_PACK_KEY, d.COD_ITM_DESC, d.COD_PACKING, " & _
                                "m.CO_TRACK_NO, d.COD_UOM, d.COD_PCS_UOM, d.COD_TOT_WGT, d.COD_TOT_CBM, d.COD_REM, d.COD_VND_CODE, d.COD_BATCH_NO, " & _
                                "m.CUS_CODE, m.CUS_NAME, m.CO_ADDR1, m.CO_ADDR2, m.CO_ADDR3, m.CO_AREA_DEL, m.CO_REGION_DEL, " & _
                                "m.CO_COUNTRY_DEL, m.CO_CUS_CONT, m.CO_CUS_CONT_TEL, m.CO_TRACK_NO,m.CO_FTRACK_NO, m.CO_INV_NO, m.CO_CUS_REF_NO, i.ITM_SKU_NO, i.ITM_DESC, " & _
                                "v.AITM_QTY_PER_CTN, v.AITM_VOL, v.CARTON_CBM, d.COD_TICKET_NO,CO_WH_CODE, " & _
                                "d.COD_UOM2, d.COD_QTY2, d.COD_WH_CODE, " & _
                                "m.CO_PROVINCE, m.CO_CITY, " & _
                                "m.CO_SENDER, m.CO_SENDER_COUNTRY, m.CO_SENDER_PROVINCE, m.CO_SENDER_REGION, m.CO_SENDER_ADDR, m.CO_SENDER_TEL, " & _
                                "i.ITM_TYPE, i.ITM_SERIAL_NO_YN, " & _
                                "Convert(varchar,d.COD_EXPIRY_DATE, " & DDFORMAT & ") as COD_EXPIRY_DATE, " & _
                                "Convert(varchar,d.COD_MANU_DATE, " & DDFORMAT & ") as COD_MANU_DATE, m.CO_EDI_SIR_NO " & _
                             "from WMS_CUST_ORDER m " & _
                             "inner join WMS_CUST_ORDER_D d " & _
                                 "ON d.CO_CODE = m.CO_CODE " & _
                                 "and d.STORER_CODE = m.STORER_CODE " & _
                                 "and d.IMP_CODE = m.IMP_CODE " & _
                             "left outer join WMS_ITEM i " & _
                                 "on d.IMP_CODE = i.IMP_CODE " & _
                                 "and d.STORER_CODE = i.STORER_CODE " & _
                                 "and d.COD_ITM_CODE = i.ITM_CODE " & _
                                 "and d.COD_PACK_KEY = i.PACK_KEY " & _
                             "left outer join V_ALT_VEND_ITEM v " & _
                                 "on d.IMP_CODE = v.IMP_CODE " & _
                                 "and d.STORER_CODE = v.STORER_CODE " & _
                                 "and d.COD_ITM_CODE = v.ITM_CODE " & _
                                 "and d.COD_PACK_KEY = v.PACK_KEY " & _
                             "where d.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                             "and d.STORER_CODE + '_000_' + d.CO_CODE IN (" & checkedCO & ") " & _
                             "Order by d.CO_CODE, d.COD_SEQ"

                    tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                    If tempDT.Rows.Count > 0 Then
                        Dim lCO_CODE As String = ""
                        Dim nextDO As String = ""
                        Dim totWgt, totCBM, totPCs As Double
                        Dim doSeq As Integer = 1
                        Dim PLselectSQL As String = ""
                        Dim tempPLDT As DataTable

                        For i = 0 To tempDT.Rows.Count - 1
                            If lCO_CODE <> tempDT.Rows(i).Item("CO_CODE").ToString.Trim Then
                                If nextDO <> "" Then
                                    PLselectSQL = "SELECT DO_WH_CODE, DOD_SEQ FROM WMS_DELV_ORDER " & _
                                                    " INNER JOIN " & _
                                                    " WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND  " & _
                                                    " WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE WHERE WMS_DELV_ORDER.DO_CODE = '" & gU.dbEncode(nextDO) & "' order by DOD_SEQ"
                                    tempPLDT = gDB.getDataTable(PLselectSQL, gConn, transaction)
                                    For PLi = 0 To tempPLDT.Rows.Count - 1
                                        updateSQL = " INSERT INTO WMS_DO_PICKLIST_D " & _
                                                    " (IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_ITEM_QTY, PLD_WH, PLD_LOC, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN,  " & _
                                                    " SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY, PLD_EXPIRY_DATE, PLD_MANU_DATE,   " & _
                                                    " PLD_ORG_LOC, PLD_IS_LOAN, DOD_SEQ) " & _
                                                    " (SELECT WMS_DELV_ORDER.IMP_CODE, WMS_DELV_ORDER.STORER_CODE, WMS_DELV_ORDER.DO_CODE, DOD_SEQ, '" & Session("usr_id") & "',WMS_DELV_ORDER_D.DOD_ITM_CODE,  " & _
                                                    " WMS_DELV_ORDER_D.DOD_PACK_KEY, WMS_DELV_ORDER_D.DOD_PALLET_NO, WMS_DELV_ORDER_D.DOD_QTY,   " & _
                                                    " case when WMS_WH_BIN.WH_CODE is null then v.ILOC_WH else WMS_WH_BIN.WH_CODE end, case when WMS_WH_BIN.LOC_KEY is null then v.ILOC_LOC else WMS_WH_BIN.LOC_KEY end, case when WMS_WH_BIN.FL_NUM is null then v.ILOC_FLOOR else WMS_WH_BIN.FL_NUM end, case when WMS_WH_BIN.AR_CODE is null then v.ILOC_AREA else WMS_WH_BIN.AR_CODE end, case when WMS_WH_BIN.RK_CODE is null then v.ILOC_RACK else WMS_WH_BIN.RK_CODE end, case when WMS_WH_BIN.BN_CODE is null then v.ILOC_BIN else WMS_WH_BIN.BN_CODE end, " & _
                                                    " '" & Session("usr_id") & "', getdate(), getdate(), '" & Session("usr_id") & "', WMS_DELV_ORDER_D.DOD_QTY,  v.ILOC_BATCH_NO, " & _
                                                    " MAX(Case when v.ILOC_BAL_QTY > WMS_DELV_ORDER_D.DOD_QTY then WMS_DELV_ORDER_D.DOD_QTY else v.ILOC_BAL_QTY end) as FOI_QTY, WMS_DELV_ORDER_D.DOD_EXPIRY_DATE, WMS_DELV_ORDER_D.DOD_MANU_DATE, " & _
                                                    " case when WMS_WH_BIN.LOC_KEY is null then v.ILOC_LOC else WMS_WH_BIN.LOC_KEY end, 'N', DOD_SEQ " & _
                                                    " FROM WMS_DELV_ORDER " & _
                                                    " INNER JOIN " & _
                                                    " WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND  " & _
                                                    " WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE " & _
                                                    " LEFT JOIN " & _
                                                    " WMS_ITEM_WH ON WMS_ITEM_WH.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_ITEM_WH.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND  " & _
                                                    " WMS_ITEM_WH.ITM_CODE = WMS_DELV_ORDER_D.DOD_ITM_CODE AND WMS_ITEM_WH.PACK_KEY = WMS_DELV_ORDER_D.DOD_PACK_KEY AND  " & _
                                                    " WMS_ITEM_WH.WH_CODE = WMS_DELV_ORDER_D.DOD_WH_CODE " & _
                                                    " LEFT JOIN WMS_WH_BIN ON WMS_WH_BIN.LOC_KEY = WMS_ITEM_WH.IW_PICK_LOC " & _
                                                    " LEFT JOIN " & _
                                                    " (SELECT * " & _
                                                    " FROM   (SELECT *, ROW_NUMBER() " & _
                                                    " OVER(PARTITION BY IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILOC_PALLET_NO ORDER BY IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILOC_PALLET_NO) AS Seq " & _
                                                    " FROM WMS_ITEM_LOC_BAl WHERE ILOC_BAL_QTY > 0"
                                        If tempPLDT.Rows(PLi).Item("DOD_SEQ").ToString.Trim <> "" Then
                                            updateSQL = updateSQL & " AND ILOC_WH = '" & gU.dbEncode(tempPLDT.Rows(PLi).Item("DO_WH_CODE").ToString.Trim) & "' "
                                        End If
                                        updateSQL = updateSQL & ") a " & _
                                                    " WHERE  Seq = 1) v " & _
                                                    " ON  WMS_DELV_ORDER_D.IMP_CODE = v.IMP_CODE AND WMS_DELV_ORDER_D.STORER_CODE = v.STORER_CODE AND  " & _
                                                    " WMS_DELV_ORDER_D.DOD_ITM_CODE = v.ITM_CODE AND WMS_DELV_ORDER_D.DOD_PACK_KEY = v.PACK_KEY " & _
                                                    " WHERE WMS_DELV_ORDER.IMP_CODE = '" & Session("IMP_CODE") & "' AND WMS_DELV_ORDER.STORER_CODE='" & gU.dbEncode(tempDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "'" & _
                                                    " AND WMS_DELV_ORDER.DO_CODE='" & gU.dbEncode(nextDO) & "' " & _
                                                    " AND WMS_DELV_ORDER_D.DOD_SEQ='" & gU.dbEncode(tempPLDT.Rows(PLi).Item("DOD_SEQ").ToString.Trim) & "' " & _
                                                    " group by WMS_DELV_ORDER.IMP_CODE,WMS_DELV_ORDER.STORER_CODE,WMS_DELV_ORDER.DO_CODE,WMS_DELV_ORDER_D.DOD_SEQ,WMS_DELV_ORDER_D.DOD_ITM_CODE,WMS_DELV_ORDER_D.DOD_PACK_KEY,WMS_DELV_ORDER_D.DOD_PALLET_NO,WMS_DELV_ORDER_D.DOD_QTY,case when WMS_WH_BIN.WH_CODE is null then v.ILOC_WH else WMS_WH_BIN.WH_CODE end, case when WMS_WH_BIN.LOC_KEY is null then v.ILOC_LOC else WMS_WH_BIN.LOC_KEY end, case when WMS_WH_BIN.FL_NUM is null then v.ILOC_FLOOR else WMS_WH_BIN.FL_NUM end, case when WMS_WH_BIN.AR_CODE is null then v.ILOC_AREA else WMS_WH_BIN.AR_CODE end, case when WMS_WH_BIN.RK_CODE is null then v.ILOC_RACK else WMS_WH_BIN.RK_CODE end, case when WMS_WH_BIN.BN_CODE is null then v.ILOC_BIN else WMS_WH_BIN.BN_CODE end,WMS_DELV_ORDER_D.DOD_QTY,v.ILOC_BATCH_NO,WMS_DELV_ORDER_D.DOD_EXPIRY_DATE,WMS_DELV_ORDER_D.DOD_MANU_DATE " & _
                                                    ")"

                                        gDB.amendData(updateSQL, gConn, transaction)
                                    Next
                                End If

                                nextDO = DB.getDocNo("DO", gConn, transaction)

                                If tempDT.Rows(i).Item("STORER_CODE").ToString.Trim = "002" Or tempDT.Rows(i).Item("STORER_CODE").ToString.Trim = "007" Then
                                    If tempDT.Rows(i).Item("CO_COUNTRY_DEL").ToString.Trim = "CN" Then
                                        PL_WH_CODE = tempDT.Rows(i).Item("CO_COUNTRY_DEL").ToString.Trim
                                    Else
                                        PL_WH_CODE = "HKARIIXP"
                                    End If
                                Else
                                    PL_WH_CODE = tempDT.Rows(i).Item("CO_WH_CODE").ToString.Trim
                                End If

                                updateSQL = "insert into WMS_DELV_ORDER(IMP_CODE, STORER_CODE, DO_CODE, DO_STATUS, " & _
                                            "DO_CO_CODE, DO_CUS_REF_NO, DO_DATE, DO_TARGET_DELDATE, " & _
                                            "CUS_CODE, CUS_NAME, DO_ADDR1, DO_ADDR2, DO_ADDR3, DO_AREA_DEL, DO_REGION_DEL, " & _
                                            "DO_COUNTRY_DEL, DO_CUS_CONT, DO_CUS_CONT_TEL," & _
                                            "DO_REM, DO_track_no, DO_FTRACK_NO, " & _
                                            "DO_SHIP_MODE, DO_INV_NO, DO_EDI_SIR_NO,DO_WH_CODE," & _
                                            "DO_CONSIGNEE, DO_CONSIGNEE_ADDR1, DO_PROVINCE, DO_CITY, " & _
                                            "DO_SENDER, DO_SENDER_COUNTRY, DO_SENDER_PROVINCE, DO_SENDER_REGION, DO_SENDER_ADDR, DO_SENDER_TEL, " & _
                                             "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                            "values ( " & _
                                            "'" & Session("imp_code") & "', '" & gU.dbEncode(tempDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "', '" & gU.dbEncode(nextDO) & "', 'NEW', " & _
                                            "'" & gU.dbEncode(tempDT.Rows(i).Item("CO_CODE").ToString.Trim) & "', '" & gU.dbEncode(tempDT.Rows(i).Item("CO_CUS_REF_NO").ToString.Trim) & "', getdate()," & gU.convdbDate(tempDT.Rows(i).Item("CO_TARGET_DELDATE").ToString.Trim) & "," & _
                                            "'" & gU.dbEncode(tempDT.Rows(i).Item("CUS_CODE").ToString.Trim) & "', " & gU.convdbNVCData(tempDT.Rows(i).Item("CUS_NAME").ToString.Trim) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_ADDR1").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_ADDR2").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_ADDR3").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_AREA_DEL").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_REGION_DEL").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_COUNTRY_DEL").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_CUS_CONT").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_CUS_CONT_TEL").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_REM").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_TRACK_NO").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_FTRACK_NO").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_SHIP_MODE").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_INV_NO").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_EDI_SIR_NO").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(PL_WH_CODE)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_CUS_CONT").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_ADDR1").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_PROVINCE").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_CITY").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_SENDER").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_SENDER_COUNTRY").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_SENDER_PROVINCE").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_SENDER_REGION").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_SENDER_ADDR").ToString.Trim)) & ", " & _
                                            gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("CO_SENDER_TEL").ToString.Trim)) & ", " & _
                                            "'" & Session("usr_id") & "', Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                gDB.amendData(updateSQL, gConn, transaction)

                                REM **********************

                                REM Generate Document Link 
                                'dl.genDocLink("DO", DO_CODE.Text.Trim, "CO", DO_CO_CODE.text, STORER_CODE.SelectedValue, gConn, transaction)
                                dl.genDocLink("DO", nextDO, "CO", tempDT.Rows(i).Item("CO_CODE").ToString.Trim, tempDT.Rows(i).Item("STORER_CODE").ToString.Trim, gConn, transaction)

                                lCO_CODE = tempDT.Rows(i).Item("CO_CODE").ToString.Trim
                            End If

                            If gU.decodeEmptyCdbl(tempDT.Rows(i).Item("COD_QTY").ToString.Trim, 0) <> 0 Then
                                totWgt = gU.decodeEmptyCdbl(tempDT.Rows(i).Item("AITM_VOL").ToString.Trim, 0) * gU.decodeEmptyCdbl(tempDT.Rows(i).Item("COD_QTY").ToString.Trim, 0)
                                totCBM = gU.decodeEmptyCdbl(tempDT.Rows(i).Item("CARTON_CBM").ToString.Trim, 0) * gU.decodeEmptyCdbl(tempDT.Rows(i).Item("COD_QTY").ToString.Trim, 0)
                            Else
                                totWgt = gU.decodeEmptyCdbl(tempDT.Rows(i).Item("AITM_VOL").ToString.Trim, 0)
                                totCBM = gU.decodeEmptyCdbl(tempDT.Rows(i).Item("CARTON_CBM").ToString.Trim, 0)
                            End If

                            totPCs = gU.decodeEmptyCdbl(tempDT.Rows(i).Item("COD_PCS_UOM").ToString.Trim, 0) * gU.decodeEmptyCdbl(tempDT.Rows(i).Item("COD_QTY").ToString.Trim, 0)

                            doSeq = gU.decodeEmptyCInt(DB.getValueFromSQL("select MAX(CAST(DOD_SEQ AS int)) + 1 from WMS_DELV_ORDER_D " & _
                                    "where IMP_CODE = '" & gU.dbEncode(Session("imp_code")) & "' " & _
                                    "and STORER_CODE = '" & gU.dbEncode(tempDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "' " & _
                                    "and DO_CODE = '" & gU.dbEncode(nextDO) & "' ", gConn, transaction), 1)

                            updateSQL = "insert into WMS_DELV_ORDER_D " & _
                                        "(IMP_CODE, STORER_CODE, DO_CODE, DOD_SEQ, DOD_DISP_SEQ, DOD_PALLET_NO, DOD_ITM_CODE, DOD_PACK_KEY, " & _
                                        "DOD_ITM_DESC, DOD_PACK_TYPE, DOD_QTY, DOD_UOM, DOD_QTY2, DOD_UOM2, DOD_PCS_UOM, DOD_TOTPCS, DOD_TOT_WGT, DOD_TOT_CBM, DOD_REM, DOD_VND_CODE, " & _
                                        "DOD_BATCH_NO, DOD_EXPIRY_DATE, DOD_MANU_DATE,DOD_WH_CODE, " & _
                                        "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                        "values (" & _
                                        "'" & Session("imp_code") & "', '" & gU.dbEncode(tempDT.Rows(i).Item("STORER_CODE").ToString.Trim) & "', '" & gU.dbEncode(nextDO) & "', '" & gU.dbEncode(doSeq) & "', '" & gU.dbEncode(doSeq) & "'," & _
                                        gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(tempDT.Rows(i).Item("COD_PALLET_NO").ToString.Trim, "000"))) & "," & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("COD_ITM_CODE").ToString.Trim)) & "," & _
                                        gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("COD_PACK_KEY").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("COD_ITM_DESC").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("COD_PACKING").ToString.Trim)) & ", " & _
                                        gU.dbEncode(tempDT.Rows(i).Item("COD_QTY").ToString.Trim) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("COD_UOM").ToString.Trim)) & ", " & gU.decodeNullOrEmpty(gU.dbEncode(tempDT.Rows(i).Item("COD_QTY2").ToString.Trim), "NULL") & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("COD_UOM2").ToString.Trim)) & ", " & gU.dbEncode(tempDT.Rows(i).Item("COD_PCS_UOM").ToString.Trim) & ", " & _
                                        gU.dbEncode(totPCs) & "," & gU.dbEncode(totWgt) & "," & gU.dbEncode(totCBM) & "," & _
                                        gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("COD_REM").ToString.Trim)) & ", " & gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("COD_VND_CODE").ToString.Trim)) & ", " & _
                                        gU.convdbNVCData(gU.dbEncode(tempDT.Rows(i).Item("COD_BATCH_NO").ToString.Trim)) & ", " & gU.convdbDate(gU.dbEncode(tempDT.Rows(i).Item("COD_EXPIRY_DATE").ToString.Trim)) & ", " & gU.convdbDate(gU.dbEncode(tempDT.Rows(i).Item("COD_MANU_DATE").ToString.Trim)) & "," & _
                                        gU.convdbNVCData(gU.dbEncode(gU.decodeNullOrEmpty(tempDT.Rows(i).Item("COD_WH_CODE").ToString.Trim, "HK"))) & ", " & _
                                        "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                            gDB.amendData(updateSQL, gConn, transaction)


                        Next


                        If nextDO <> "" Then
                            PLselectSQL = "SELECT DO_WH_CODE, DOD_SEQ FROM WMS_DELV_ORDER " & _
                                                    " INNER JOIN " & _
                                                    " WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND  " & _
                                                    " WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE WHERE WMS_DELV_ORDER.DO_CODE = '" & gU.dbEncode(nextDO) & "' order by DOD_SEQ"
                            tempPLDT = gDB.getDataTable(PLselectSQL, gConn, transaction)
                            For PLi = 0 To tempPLDT.Rows.Count - 1
                                updateSQL = " INSERT INTO WMS_DO_PICKLIST_D " & _
                                            " (IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_ITEM_QTY, PLD_WH, PLD_LOC, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN,  " & _
                                            " SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY, PLD_EXPIRY_DATE, PLD_MANU_DATE,   " & _
                                            " PLD_ORG_LOC, PLD_IS_LOAN, DOD_SEQ) " & _
                                            " (SELECT WMS_DELV_ORDER.IMP_CODE, WMS_DELV_ORDER.STORER_CODE, WMS_DELV_ORDER.DO_CODE, DOD_SEQ, '" & Session("usr_id") & "',WMS_DELV_ORDER_D.DOD_ITM_CODE,  " & _
                                            " WMS_DELV_ORDER_D.DOD_PACK_KEY, WMS_DELV_ORDER_D.DOD_PALLET_NO, WMS_DELV_ORDER_D.DOD_QTY,   " & _
                                            " case when WMS_WH_BIN.WH_CODE is null then v.ILOC_WH else WMS_WH_BIN.WH_CODE end, case when WMS_WH_BIN.LOC_KEY is null then v.ILOC_LOC else WMS_WH_BIN.LOC_KEY end, case when WMS_WH_BIN.FL_NUM is null then v.ILOC_FLOOR else WMS_WH_BIN.FL_NUM end, case when WMS_WH_BIN.AR_CODE is null then v.ILOC_AREA else WMS_WH_BIN.AR_CODE end, case when WMS_WH_BIN.RK_CODE is null then v.ILOC_RACK else WMS_WH_BIN.RK_CODE end, case when WMS_WH_BIN.BN_CODE is null then v.ILOC_BIN else WMS_WH_BIN.BN_CODE end, " & _
                                            " '" & Session("usr_id") & "', getdate(), getdate(), '" & Session("usr_id") & "', WMS_DELV_ORDER_D.DOD_QTY,  v.ILOC_BATCH_NO, " & _
                                            " MAX(Case when v.ILOC_BAL_QTY > WMS_DELV_ORDER_D.DOD_QTY then WMS_DELV_ORDER_D.DOD_QTY else v.ILOC_BAL_QTY end) as FOI_QTY, WMS_DELV_ORDER_D.DOD_EXPIRY_DATE, WMS_DELV_ORDER_D.DOD_MANU_DATE, " & _
                                            " case when WMS_WH_BIN.LOC_KEY is null then v.ILOC_LOC else WMS_WH_BIN.LOC_KEY end, 'N', DOD_SEQ " & _
                                            " FROM WMS_DELV_ORDER " & _
                                            " INNER JOIN " & _
                                            " WMS_DELV_ORDER_D ON WMS_DELV_ORDER.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_DELV_ORDER.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND  " & _
                                            " WMS_DELV_ORDER.DO_CODE = WMS_DELV_ORDER_D.DO_CODE " & _
                                            " LEFT JOIN " & _
                                            " WMS_ITEM_WH ON WMS_ITEM_WH.IMP_CODE = WMS_DELV_ORDER_D.IMP_CODE AND WMS_ITEM_WH.STORER_CODE = WMS_DELV_ORDER_D.STORER_CODE AND  " & _
                                            " WMS_ITEM_WH.ITM_CODE = WMS_DELV_ORDER_D.DOD_ITM_CODE AND WMS_ITEM_WH.PACK_KEY = WMS_DELV_ORDER_D.DOD_PACK_KEY AND  " & _
                                            " WMS_ITEM_WH.WH_CODE = WMS_DELV_ORDER_D.DOD_WH_CODE " & _
                                            " LEFT JOIN WMS_WH_BIN ON WMS_WH_BIN.LOC_KEY = WMS_ITEM_WH.IW_PICK_LOC " & _
                                            " LEFT JOIN " & _
                                            " (SELECT * " & _
                                            " FROM   (SELECT *, ROW_NUMBER() " & _
                                            " OVER(PARTITION BY IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILOC_PALLET_NO ORDER BY IMP_CODE, STORER_CODE, ITM_CODE, PACK_KEY, ILOC_PALLET_NO) AS Seq " & _
                                            " FROM WMS_ITEM_LOC_BAl WHERE ILOC_BAL_QTY > 0"
                                If tempPLDT.Rows(PLi).Item("DOD_SEQ").ToString.Trim <> "" Then
                                    updateSQL = updateSQL & " AND ILOC_WH = '" & gU.dbEncode(tempPLDT.Rows(PLi).Item("DO_WH_CODE").ToString.Trim) & "' "
                                End If
                                updateSQL = updateSQL & ") a " & _
                                            " WHERE  Seq = 1) v " & _
                                            " ON  WMS_DELV_ORDER_D.IMP_CODE = v.IMP_CODE AND WMS_DELV_ORDER_D.STORER_CODE = v.STORER_CODE AND  " & _
                                            " WMS_DELV_ORDER_D.DOD_ITM_CODE = v.ITM_CODE AND WMS_DELV_ORDER_D.DOD_PACK_KEY = v.PACK_KEY " & _
                                            " WHERE WMS_DELV_ORDER.IMP_CODE = '" & Session("IMP_CODE") & "' AND WMS_DELV_ORDER.STORER_CODE='" & gU.dbEncode(tempDT.Rows(tempDT.Rows.Count - 1).Item("STORER_CODE").ToString.Trim) & "'" & _
                                            " AND WMS_DELV_ORDER.DO_CODE='" & gU.dbEncode(nextDO) & "' " & _
                                            " AND WMS_DELV_ORDER_D.DOD_SEQ='" & gU.dbEncode(tempPLDT.Rows(PLi).Item("DOD_SEQ").ToString.Trim) & "' " & _
                                            " group by WMS_DELV_ORDER.IMP_CODE,WMS_DELV_ORDER.STORER_CODE,WMS_DELV_ORDER.DO_CODE,WMS_DELV_ORDER_D.DOD_SEQ,WMS_DELV_ORDER_D.DOD_ITM_CODE,WMS_DELV_ORDER_D.DOD_PACK_KEY,WMS_DELV_ORDER_D.DOD_PALLET_NO,WMS_DELV_ORDER_D.DOD_QTY,case when WMS_WH_BIN.WH_CODE is null then v.ILOC_WH else WMS_WH_BIN.WH_CODE end, case when WMS_WH_BIN.LOC_KEY is null then v.ILOC_LOC else WMS_WH_BIN.LOC_KEY end, case when WMS_WH_BIN.FL_NUM is null then v.ILOC_FLOOR else WMS_WH_BIN.FL_NUM end, case when WMS_WH_BIN.AR_CODE is null then v.ILOC_AREA else WMS_WH_BIN.AR_CODE end, case when WMS_WH_BIN.RK_CODE is null then v.ILOC_RACK else WMS_WH_BIN.RK_CODE end, case when WMS_WH_BIN.BN_CODE is null then v.ILOC_BIN else WMS_WH_BIN.BN_CODE end,WMS_DELV_ORDER_D.DOD_QTY,v.ILOC_BATCH_NO,WMS_DELV_ORDER_D.DOD_EXPIRY_DATE,WMS_DELV_ORDER_D.DOD_MANU_DATE " & _
                                            ")"

                                gDB.amendData(updateSQL, gConn, transaction)
                            Next
                        End If



                        transaction.Commit()
                        successFlag = True

                        If Session("gLang") = "C" Then
                            uiFun.displayMsgNew(udp1, "", "成功產生送貨指令!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(udp1, "", "Do Successfully Generated!", Session("gLang"))
                        End If

                    Else
                        If Session("gLang") = "C" Then
                            uiFun.displayMsgNew(udp1, "", "找不到客戶指令!", Session("gLang"))
                        Else
                            uiFun.displayMsgNew(udp1, "", "Cannot Locate Any CO.", Session("gLang"))
                        End If

                    End If

                Catch ex As Exception
                    transaction.Rollback()
                    uiFun.displayMsgNew(udp1, "1008", "", Session("gLang"))
                Finally
                    If gConn IsNot Nothing Then
                        If gConn.State = ConnectionState.Open Then
                            gConn.Close()
                            gConn.Dispose()
                        End If
                    End If
                End Try


            Else
                If Session("gLang") = "C" Then
                    uiFun.displayMsgNew(udp1, "", "請選取任何客戶指令!", Session("gLang"))
                Else
                    uiFun.displayMsgNew(udp1, "", "Please Select a CO!", Session("gLang"))
                End If

            End If

        Else
            If Session("gLang") = "C" Then
                uiFun.displayMsgNew(udp1, "", "請選取任何客戶指令!", Session("gLang"))
            Else
                uiFun.displayMsgNew(udp1, "", "Please Select a CO!", Session("gLang"))
            End If
        End If

        If successFlag Then
            BindGV()
        End If

    End Sub

    Protected Sub BtnGen_Click(sender As Object, e As System.EventArgs) Handles BtnGen.Click
        MulitAddCOtoDO()
    End Sub

    Protected Sub btnGen2_Click(sender As Object, e As System.EventArgs) Handles btnGen2.Click
        MulitAddCOtoDO()
    End Sub
End Class

