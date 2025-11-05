Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OUTBOUND_DO_PickListLookup
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Private dt As New DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)
        If ar.sessionExpired = "Y" Then
            Exit Sub
        End If

        'If Not IsPostBack Then
        '    Session("pagemode") = Nothing
        '    Session("pagemode") = Request("mode")
        'End If

        If Session("pagemode") = "N" Then
            'CancelBtn.Visible = False
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Pick List"
            lbl_ITM_CODE.Text = "Item Code"
            lbl_ITM_SKU_NO.Text = "Stock No."
            lbl_PACK_KEY.Text = "Pack Key"
            lbl_REQ_QTY.Text = "Req. Qty"
            'saveBtn1.Text = "OK"
            'saveBtn2.Text = "OK"
        ElseIf Session("gLang") = "C" Then
            lheader.Text = "領料清單"
            lbl_ITM_CODE.Text = "物件號碼"
            lbl_ITM_SKU_NO.Text = "Stock No."
            lbl_PACK_KEY.Text = "封裝內碼"
            lbl_REQ_QTY.Text = "需求數量"
            'saveBtn1.Text = "确定"
            'saveBtn2.Text = "确定"
        End If
        'saveBtn1.Attributes.Add("onclick", "saveSelection();")
        'saveBtn1.Attributes.Add("onclick", "saveSelection();")
        REM **********************

        REM **********************
        REM Additional CSS
        'RT_TYPE.CssClass = "REQUIRED"
        REM **********************

        dt = Session("_M_OB_DO_TMP_pl_lkup_dt")
        If Not IsPostBack Then
            Call BindGV()
        End If

        ar.hideForm(Me, editMode)

        If Request("moduleAction") = "SAVEOK" Then
            save()
        End If
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here
                Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板編號", HorizontalAlign.Left)
                Call cU.changeGVLabel(oGridViewRow, e, "Qty", "領取數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "RCV Qty", "收貨數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Available Qty", "架上數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Bal Qty", "結餘數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "Batch No.", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Expiry Date", "到期日", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "WH", "倉庫", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Location", "位置", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Floor", "樓層", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Area", "地區", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Rack", "貨架", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Bin", "箱子", HorizontalAlign.Center)
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                REM **********************
                REM Modify Here
                'Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim
                CType(e.Row.FindControl("gra_pa_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRA_PA_QTY").ToString.Trim
                CType(e.Row.FindControl("iloc_pallet_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_PALLET_NO").ToString.Trim
                CType(e.Row.FindControl("stock_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "STOCK_QTY").ToString.Trim
                CType(e.Row.FindControl("iloc_bal_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_BAL_QTY").ToString.Trim
                CType(e.Row.FindControl("iloc_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_BATCH_NO").ToString.Trim
                CType(e.Row.FindControl("iloc_expiry_date"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_EXPIRY_DATE").ToString.Trim
                CType(e.Row.FindControl("iloc_wh"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_WH").ToString.Trim
                CType(e.Row.FindControl("iloc_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_LOC").ToString.Trim
                CType(e.Row.FindControl("iloc_floor"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_FLOOR").ToString.Trim
                CType(e.Row.FindControl("iloc_area"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_AREA").ToString.Trim
                CType(e.Row.FindControl("iloc_rack"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_RACK").ToString.Trim
                CType(e.Row.FindControl("iloc_bin"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_BIN").ToString.Trim

                Dim pl_dt As DataTable = Session("_M_OB_DO_TMP_pl_dt")
                Dim i As Integer

                For i = 0 To pl_dt.Rows.Count - 1
                    If pl_dt.Rows(i).Item("PLD_ITEM_NO").ToString = ITM_CODE.Text And _
                    pl_dt.Rows(i).Item("PLD_PACK_KEY").ToString = PACK_KEY.Text And _
                    pl_dt.Rows(i).Item("PLD_LOC").ToString = DataBinder.Eval(e.Row.DataItem, "ILOC_LOC").ToString.Trim And _
                    pl_dt.Rows(i).Item("PLD_BATCH_NO").ToString = DataBinder.Eval(e.Row.DataItem, "ILOC_BATCH_NO").ToString.Trim Then

                        'CType(e.Row.FindControl("pick_qty"), TextBox).Text = pl_dt.Rows(i).Item("PLD_ITEM_QTY").ToString
                        CType(e.Row.FindControl("foi_qty"), TextBox).Text = pl_dt.Rows(i).Item("PLD_FOI_QTY").ToString
                        Exit For
                    End If
                Next

                'If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                '    Call ar.hideGVRow(GridView1, e.Row)
                'End If
        End Select
    End Sub

    Private Function validateAll(Optional ByVal reqQty As Double = 0) As Boolean
        Dim i As Integer
        Dim javaStr As String
        Dim currPickQty As Double = 0

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("foi_qty"), TextBox).Text) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Invalid number, FFI Qty!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "無效的數字, 取貨數量!", Session("gLang"))
                    End If
                    Return False
                End If

                If Request("moduleAction") <> "SAVEOK" Then
                    If CType(GridView1.Rows(i).FindControl("foi_qty"), TextBox).Text.Trim <> "" Then
                        If CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("foi_qty"), TextBox).Text, "0")) > CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("STOCK_QTY"), Label).Text, "0")) Then
                            If Session("gLang") = "E" Then
                                javaStr = "if(confirm(""FFI qty is greater than stock qty, confirm to proceed?"")) saveok();"
                            Else
                                javaStr = "if(confirm(""取貨數量高於貨存, 確定輸入?"")) saveok();"
                            End If
                            If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                                Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
                            End If
                            Return False
                        End If
                    End If
                End If

                currPickQty += gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("foi_qty"), TextBox).Text, 0)
            Next
        End If

        If currPickQty > reqQty Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "FFI Qty is greater than Req. Qty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "取貨數量高於貨單要求數量!", Session("gLang"))
            End If
            Return False
        End If

        Return True
    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        If GridView1.Rows.Count > 0 Then
            Dim dtl_dt As DataTable = Session("dt")
            Dim reqQty As Double

            For i = 0 To dtl_dt.Rows.Count - 1
                If dtl_dt.Rows(i).Item("mFlag").ToString <> "D" Then
                    If dtl_dt.Rows(i).Item("DOD_ITM_CODE").ToString.Trim = ITM_CODE.Text And dtl_dt.Rows(i).Item("DOD_PACK_KEY").ToString.Trim = PACK_KEY.Text Then
                        reqQty = reqQty + CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(dtl_dt.Rows(i).Item("DOD_QTY"), 0), 0))
                    End If
                End If
            Next

            If validateAll(reqQty) Then
                If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                    Dim i, j As Integer
                    Dim pl_dt As DataTable = Session("_M_OB_DO_TMP_pl_dt")
                    Dim isPlUpdt As Boolean
                    Dim plSeq As Integer = CInt(Session("_M_OB_DO_TMP_pl_seq"))
                    Dim rows_count As Integer
                    Dim qtyLocDict As Dictionary(Of String, String) = New Dictionary(Of String, String)

                    For i = 0 To dt.Rows.Count - 1
                        If gU.decodeNullOrEmpty(DB.decodeDBNull(dt.Rows(i).Item("FOI_QTY"), 0), 0) > 0 Then
                            qtyLocDict.Add(dt.Rows(i).Item("ILOC_LOC").ToString & "#_#" & dt.Rows(i).Item("ILOC_BATCH_NO").ToString, "")
                            isPlUpdt = False
                            For j = 0 To pl_dt.Rows.Count - 1
                                If pl_dt.Rows(j).Item("mFlag").ToString <> "D" Then
                                    If pl_dt.Rows(j).Item("PLD_ITEM_NO").ToString = ITM_CODE.Text And _
                                        pl_dt.Rows(j).Item("PLD_PACK_KEY").ToString = PACK_KEY.Text And _
                                        pl_dt.Rows(j).Item("PLD_LOC").ToString = dt.Rows(i).Item("ILOC_LOC").ToString And _
                                        pl_dt.Rows(j).Item("PLD_BATCH_NO").ToString = dt.Rows(i).Item("ILOC_BATCH_NO").ToString Then

                                        'Update PL if matched location is found
                                        pl_dt.Rows(j).Item("PLD_ITEM_QTY") = dt.Rows(i).Item("FOI_QTY")
                                        pl_dt.Rows(j).Item("PLD_FOI_QTY") = dt.Rows(i).Item("FOI_QTY")
                                        pl_dt.Rows(j).Item("PLD_DO_QTY") = reqQty
                                        isPlUpdt = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If Not isPlUpdt Then
                                'Add a new row in PL if no matched location
                                pl_dt.Rows.Add()

                                plSeq = plSeq + 1

                                rows_count = pl_dt.Rows.Count

                                pl_dt.Rows(rows_count - 1).Item("PLD_SEQ") = CStr(plSeq)
                                pl_dt.Rows(rows_count - 1).Item("PLD_PICKED_BY") = Session("usr_id")
                                pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_NO") = ITM_CODE.Text
                                pl_dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = ITM_SKU_NO.Text
                                pl_dt.Rows(rows_count - 1).Item("PLD_PACK_KEY") = PACK_KEY.Text
                                pl_dt.Rows(rows_count - 1).Item("PLD_PALLET_NO") = dt.Rows(i).Item("ILOC_PALLET_NO")
                                pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_QTY") = dt.Rows(i).Item("FOI_QTY")
                                pl_dt.Rows(rows_count - 1).Item("PLD_FOI_QTY") = dt.Rows(i).Item("FOI_QTY")
                                'pl_dt.Rows(rows_count - 1).Item("PLD_ITEM_QTY") = 0

                                pl_dt.Rows(rows_count - 1).Item("STOCK_QTY") = dt.Rows(i).Item("STOCK_QTY")
                                pl_dt.Rows(rows_count - 1).Item("ILOC_BAL_QTY") = dt.Rows(i).Item("ILOC_BAL_QTY")


                                pl_dt.Rows(rows_count - 1).Item("HOLD_QTY") = 0
                                pl_dt.Rows(rows_count - 1).Item("TOTAL_BAL") = 0

                                'Will be updated by the updateTotalQty function
                                pl_dt.Rows(rows_count - 1).Item("TOTAL_ITEM_QTY") = 0
                                pl_dt.Rows(rows_count - 1).Item("TOTAL_FOI_QTY") = 0

                                pl_dt.Rows(rows_count - 1).Item("ILOC_EXPIRY_DATE") = dt.Rows(i).Item("ILOC_EXPIRY_DATE")

                                pl_dt.Rows(rows_count - 1).Item("PLD_EXPIRY_DATE") = dt.Rows(i).Item("ILOC_EXPIRY_DATE")
                                pl_dt.Rows(rows_count - 1).Item("PLD_MANU_DATE") = dt.Rows(i).Item("ILOC_MANU_DATE")

                                pl_dt.Rows(rows_count - 1).Item("PLD_WH") = dt.Rows(i).Item("ILOC_WH")
                                pl_dt.Rows(rows_count - 1).Item("PLD_LOC") = dt.Rows(i).Item("ILOC_LOC")
                                pl_dt.Rows(rows_count - 1).Item("PLD_ORG_LOC") = ORG_LOC.Value
                                pl_dt.Rows(rows_count - 1).Item("PLD_FLOOR") = dt.Rows(i).Item("ILOC_FLOOR")
                                pl_dt.Rows(rows_count - 1).Item("PLD_AREA") = dt.Rows(i).Item("ILOC_AREA")
                                pl_dt.Rows(rows_count - 1).Item("PLD_RACK") = dt.Rows(i).Item("ILOC_RACK")
                                pl_dt.Rows(rows_count - 1).Item("PLD_BIN") = dt.Rows(i).Item("ILOC_BIN")

                                pl_dt.Rows(rows_count - 1).Item("BN_CSMS_CODE") = dt.Rows(i).Item("BN_CSMS_CODE")
                                pl_dt.Rows(rows_count - 1).Item("DOD_DISP_SEQ") = DOD_DISP_SEQ.Value

                                pl_dt.Rows(rows_count - 1).Item("PLD_IS_LOAN") = "N"
                                pl_dt.Rows(rows_count - 1).Item("mFlag") = "N"

                                pl_dt.Rows(rows_count - 1).Item("PLD_BATCH_NO") = dt.Rows(i).Item("ILOC_BATCH_NO")
                                pl_dt.Rows(rows_count - 1).Item("PLD_DO_QTY") = reqQty
                            End If

                        End If
                    Next

                    Dim objPlt As PickListTable
                    objPlt = New PickListTable(CO_CODE.Value, DO_CODE.Value, STORER_CODE.Value, Session("IMP_CODE"))

                    objPlt.updateTotalQty(pl_dt)

                    objPlt.updateHoldBal(pl_dt)

                    'Set unpicked PL to zero qty
                    For j = 0 To pl_dt.Rows.Count - 1
                        If pl_dt.Rows(j).Item("PLD_ITEM_NO").ToString = ITM_CODE.Text And _
                            pl_dt.Rows(j).Item("PLD_PACK_KEY").ToString = PACK_KEY.Text Then

                            If Not qtyLocDict.ContainsKey(pl_dt.Rows(j).Item("PLD_LOC").ToString & "#_#" & _
                                                          pl_dt.Rows(j).Item("PLD_BATCH_NO").ToString) Then
                                pl_dt.Rows(j).Item("PLD_FOI_QTY") = 0
                                pl_dt.Rows(j).Item("PLD_ITEM_QTY") = 0
                                pl_dt.Rows(j).Item("STOCK_QTY") = 0
                                pl_dt.Rows(j).Item("ILOC_BAL_QTY") = 0
                                pl_dt.Rows(j).Item("ILOC_EXPIRY_DATE") = ""
                                pl_dt.Rows(j).Item("PLD_EXPIRY_DATE") = ""
                                pl_dt.Rows(j).Item("PLD_MANU_DATE") = ""
                                pl_dt.Rows(j).Item("PLD_WH") = ""
                                pl_dt.Rows(j).Item("PLD_LOC") = ""
                                pl_dt.Rows(j).Item("PLD_FLOOR") = ""
                                pl_dt.Rows(j).Item("PLD_AREA") = ""
                                pl_dt.Rows(j).Item("PLD_RACK") = ""
                                pl_dt.Rows(j).Item("PLD_BIN") = ""
                                pl_dt.Rows(j).Item("PLD_BATCH_NO") = ""

                                pl_dt.Rows(j).Item("BN_CSMS_CODE") = ""

                                pl_dt.Rows(j).Item("PLD_IS_LOAN") = "N"
                            End If
                        End If
                    Next

                    delZeroRow(pl_dt)

                    pl_dt.AcceptChanges()

                    Dim pi_dt As DataTable = Session("_M_OB_DO_TMP_pi_dt")

                    For k As Integer = 0 To pl_dt.Rows.Count - 1
                        For t As Integer = 0 To pi_dt.Rows.Count - 1
                            If pl_dt.Rows(k).Item("PLD_ITEM_NO").ToString = pi_dt.Rows(t).Item("pad_itm_code").ToString AndAlso _
                                        pl_dt.Rows(k).Item("PLD_PACK_KEY").ToString = pi_dt.Rows(t).Item("pad_pack_key").ToString Then

                                'pi_dt.Rows(t).Item("pad_qty") = pl_dt.Rows(k).Item("PLD_ITEM_QTY").ToString.Trim
                                pi_dt.Rows(t).Item("pad_qty") = pl_dt.Rows(k).Item("PLD_FOI_QTY").ToString.Trim
                                pi_dt.Rows(t).AcceptChanges()
                                Exit For
                            End If
                        Next
                    Next

                    Session("_M_OB_DO_TMP_pi_dt") = pi_dt

                    Session("_M_OB_DO_TMP_pl_seq") = CStr(plSeq)

                    'Session("_M_OB_DO_TMP_pl_dt") = dt
                    Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "close", "<script language=""JavaScript"">opener.refreshGV();window.close(""_self"");</script>")
                End If
            End If
        Else
            Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "close", "<script language=""JavaScript"">window.close(""_self"");</script>")
        End If
    End Sub

    Protected Sub BindGV()
        Dim selectSql As String
        Dim dtl_dt As DataTable = Session("dt")
        Dim reqQty As Double
        Dim i As Integer
        Dim ref_no As String

        IMP_CODE.Value = Server.UrlDecode(Request("IMP_CODE"))
        STORER_CODE.Value = Server.UrlDecode(Request("STORER_CODE"))
        ITM_CODE.Text = Request("ITM_CODE")
        ITM_SKU_NO.Text = Request("ITM_SKU_NO")
        PACK_KEY.Text = Server.UrlDecode(Request("PACK_KEY"))
        PALLET_NO.Value = Server.UrlDecode(Request("PALLET_NO"))
        ref_no = Server.UrlDecode(Request("REF_NO"))
        BATCH_NO.Value = Server.UrlDecode(Request("BATCH_NO"))

        ORG_LOC.Value = Server.UrlDecode(Request("ORG_LOC"))

        CO_CODE.Value = Server.UrlDecode(Request("CO_CODE"))
        DO_CODE.Value = Server.UrlDecode(Request("DO_CODE"))

        DOD_DISP_SEQ.Value = Server.UrlDecode(Request("DOD_DISP_SEQ"))

        reqQty = 0
        For i = 0 To dtl_dt.Rows.Count - 1
            If dtl_dt.Rows(i).Item("mFlag").ToString <> "D" Then
                If dtl_dt.Rows(i).Item("DOD_ITM_CODE").ToString.Trim = ITM_CODE.Text And dtl_dt.Rows(i).Item("DOD_PACK_KEY").ToString.Trim = PACK_KEY.Text Then
                    reqQty = reqQty + CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(dtl_dt.Rows(i).Item("DOD_QTY"), 0), 0))
                End If
            End If
        Next

        REQ_QTY.Text = CStr(reqQty)

        selectSql = "select l.ILOC_PALLET_NO, l.ILOC_BAL_QTY, l.ILOC_LOC, l.ILOC_WH, l.ILOC_FLOOR, l.ILOC_AREA, l.ILOC_RACK, l.ILOC_BIN, l.ILOC_BATCH_NO, r.GRA_PA_QTY, " & _
                        "ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(PICK_ITEM.picked_qty, 0) as STOCK_QTY, b.BN_CSMS_CODE, " & _
                        "convert(varchar, l.ILOC_EXPIRY_DATE, " & gU.getConfig("DDFORMATNo") & ") as ILOC_EXPIRY_DATE, " & _
                        "convert(varchar, l.ILOC_MANU_DATE, " & gU.getConfig("DDFORMATNo") & ") as ILOC_MANU_DATE " & _
                    "from WMS_ITEM_LOC_BAL l " & _
                    "inner join WMS_WH_BIN b " & _
                    "on l.ILOC_LOC = b.LOC_KEY " & _
                    "left outer join (" & _
                        "select p.GRA_BATCH_NO, p.GRA_LOC, sum(p.GRA_PA_QTY) as GRA_PA_QTY " & _
                        "from WMS_GOODSRCV g, WMS_GOODSRCV_PA p " & _
                        "where g.IMP_CODE = p.IMP_CODE " & _
                        "and g.STORER_CODE = p.STORER_CODE " & _
                        "and g.GR_CODE = p.GR_CODE " & _
                        "and g.GR_REF_NO = '" & gU.dbEncode(ref_no) & "' " & _
                        "and p.STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
                        "and p.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                        "and p.GRA_ITM_CODE = '" & gU.dbEncode(ITM_CODE.Text) & "' " & _
                        "and p.GRA_PACK_KEY = '" & gU.dbEncode(PACK_KEY.Text) & "' " & _
                        "and p.GRA_BATCH_NO = '" & gU.dbEncode(BATCH_NO.Value) & "' "

        If PALLET_NO.Value.Trim <> "" And PALLET_NO.Value.Trim <> "000" Then
            selectSql = selectSql & _
                    "and p.GRA_PALLET_NO = '" & gU.dbEncode(PALLET_NO.Value) & "' "
        End If

        selectSql = selectSql & _
                        "group by p.GRA_LOC, p.GRA_BATCH_NO) r " & _
                        "on l.ILOC_BATCH_NO = r.GRA_BATCH_NO " & _
                        "and l.ILOC_LOC = r.GRA_LOC " & _
                        "left outer join ( " & _
                            "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " & _
                            "from wms_do_picklist_d p, wms_delv_order d " & _
                            "where d.imp_code = p.imp_code " & _
                            "and d.storer_code = p.storer_code " & _
                            "and d.do_code = p.do_code " & _
                            "and d.do_status = 'PICKED' " & _
                            "and d.do_code <> '" & gU.dbEncode(DO_CODE.Value) & "' " & _
                            "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " & _
                        "on l.IMP_CODE = PICK_ITEM.IMP_CODE  " & _
                            "AND l.STORER_CODE = PICK_ITEM.STORER_CODE  " & _
                            "AND l.ITM_CODE = PICK_ITEM.PLD_ITEM_NO  " & _
                            "AND l.PACK_KEY = PICK_ITEM.PLD_PACK_KEY  " & _
                            "AND ISNULL(l.ILOC_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO  " & _
                            "AND ISNULL(l.ILOC_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO  " & _
                            "AND l.ILOC_LOC = PICK_ITEM.PLD_LOC  " & _
                    "where l.STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
                    "and l.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                    "and l.ITM_CODE = '" & gU.dbEncode(ITM_CODE.Text) & "' " & _
                    "and l.PACK_KEY = '" & gU.dbEncode(PACK_KEY.Text) & "' "
                  
                    

        'selectSql = "select l.ILOC_PALLET_NO, l.ILOC_BAL_QTY, l.ILOC_LOC, l.ILOC_WH, l.ILOC_FLOOR, l.ILOC_AREA, l.ILOC_RACK, l.ILOC_BIN " & _
        '            "from WMS_ITEM_LOC_BAL l " & _
        '            "where l.STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
        '            "and l.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
        '            "and l.ITM_CODE = '" & gU.dbEncode(ITM_CODE.Text) & "' " & _
        '            "and l.PACK_KEY = '" & gU.dbEncode(PACK_KEY.Text) & "' "

        If PALLET_NO.Value.Trim <> "" And PALLET_NO.Value.Trim <> "000" Then
            selectSql = selectSql & _
                    "and l.ILOC_PALLET_NO = '" & gU.dbEncode(PALLET_NO.Value) & "' "
        End If

        selectSql = selectSql & _
                    "and ISNULL(l.ILOC_BAL_QTY, 0) > 0 " & _
                    "order by l.ILOC_WH, l.ILOC_FLOOR, l.ILOC_AREA, l.ILOC_RACK, l.ILOC_BIN, STOCK_QTY desc "

        dt = gDB.getDataTable(selectSql)

        If dt.Rows.Count > 0 Then
            dt.Columns.Add("FOI_QTY", Type.GetType("System.Double"))
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
            If Session("gLang") = "E" Then
                GridView1.EmptyDataText = "No stock for requested item!"
            Else
                GridView1.EmptyDataText = "要求的物品沒有庫存!"
            End If
        End If

        Session("_M_OB_DO_TMP_pl_lkup_dt") = dt

        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub

    'This function should be same as DOMain
    Private Sub delZeroRow(ByRef dt As DataTable)
        Dim i As Integer
        Dim zeroItmDict As Dictionary(Of String, List(Of Integer))
        Dim qtyItmDict As Dictionary(Of String, String)
        Dim itmKey As String
        Dim tmpList, delList As List(Of Integer)
        Dim keys As Dictionary(Of String, List(Of Integer)).KeyCollection

        zeroItmDict = New Dictionary(Of String, List(Of Integer))
        qtyItmDict = New Dictionary(Of String, String)

        'Remove empty row, but keep one row at least for each item
        For i = 0 To dt.Rows.Count - 1
            If dt.Rows(i).Item("mFlag").ToString <> "D" Then
                itmKey = dt.Rows(i).Item("pld_item_no").ToString.Trim & "#_#" & _
                    dt.Rows(i).Item("pld_pack_key").ToString.Trim & "#_#" & _
                    gU.decodeNullOrEmpty(dt.Rows(i).Item("pld_pallet_no").ToString.Trim, "000")

                'If CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(dt.Rows(i).Item("pld_item_qty"), 0), 0)) > 0 Then
                If CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(dt.Rows(i).Item("pld_foi_qty"), 0), 0)) > 0 Then
                    If Not qtyItmDict.ContainsKey(itmKey) Then
                        qtyItmDict.Add(itmKey, "")
                    End If
                Else
                    If zeroItmDict.ContainsKey(itmKey) Then
                        tmpList = zeroItmDict.Item(itmKey)
                        tmpList.Add(i)
                    Else
                        tmpList = New List(Of Integer)
                        tmpList.Add(i)
                        zeroItmDict.Add(itmKey, tmpList)
                    End If
                End If
            End If
        Next

        keys = zeroItmDict.Keys
        delList = New List(Of Integer)

        For i = 0 To keys.Count - 1
            If qtyItmDict.ContainsKey(keys(i)) Then
                'One row contains qty, all zero row is not needed
                delList.AddRange(zeroItmDict.Item(keys(i)))
            Else
                'Keep first zero row if all row zero
                tmpList = zeroItmDict.Item(keys(i))
                tmpList.RemoveAt(0)
                If tmpList.Count > 0 Then
                    delList.AddRange(tmpList)
                End If
            End If
        Next

        delList.Sort()

        For i = delList.Count - 1 To 0 Step -1
            If dt.Rows(delList(i)).Item("mFlag") = "N" Then
                If CInt(dt.Rows(delList(i)).Item("PLD_SEQ")) = CInt(Session("_M_OB_DO_TMP_pl_seq")) Then
                    Session("_M_OB_DO_TMP_pl_seq") = CStr(CInt(Session("_M_OB_DO_TMP_pl_seq")) - 1)
                End If
                dt.Rows(delList(i)).Delete()
            Else
                dt.Rows(delList(i)).Item("mFlag") = "D"
            End If
        Next
        dt.AcceptChanges()
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub
End Class
