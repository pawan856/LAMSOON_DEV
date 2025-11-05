Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OUTBOUND_CO_HOLD
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As New AccessRightUtils
    Private ModuleAbb As String = "CO_MAIN"

    Private dt As New DataTable

    Private dtBatch As DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_CO", Session("usr_id"), Me)

        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If

        If Not IsPostBack Then
            'Session("pagemode") = Nothing
            'Session("pagemode") = Request("mode")

            'uiFun.load_dropdownBy_ColCode(Status, "WMS_COL_CODE.COLC_STATUS", Session("gSelectLabel"))
           
        Else
            'dt = ViewState("dt")

        End If


        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Hold Stock"
            lbl_CO_code.Text = "CO Code"
            lbl_cod_itm_code.Text = "Item Code"
            lbl_cod_itm_desc.Text = "Item Desc."
            'lbl_Bal_QTY.Text = "Available Item Balance Quantity"
            lbl_cod_batch_no.Text = "Item Batch No."

            btnSave1.Text = "OK"
            'btnSave1.Attributes("onclick") = "if(!confirm(""Are you sure to save this record?"")) return false;"

            btnSave2.Text = "OK"
            'btnSave2.Attributes("onclick") = "if(!confirm(""Are you sure to save this record?"")) return false;"


        ElseIf Session("gLang") = "C" Then
        
            'lheader.Text = "System Code Setup"
            'lbl_code.Text = "Code"
            'lbl_code_type.Text = "Code type"
            'lbl_colc_eng_value.Text = "Code Name"
            'lbl_display_seq.Text = "Display Seq."
            'lbl_status.Text = "Status"

            'saveBtn2.Text = "Save"
            'saveBtn2.Attributes("onclick") = "if(!confirm(""Are you sure to save this record?"")) return false;;"
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        'Status.CssClass = "REQUIRED"

        REM **********************

        If Session("pagemode") = "N" Then
            '.CssClass = "REQUIRED"
            '.CssClass = "REQUIRED"
        Else
            '.Enabled = False
        End If

        If Not IsPostBack Then
            'ViewState("dt") = Nothing
            ViewState("co_code") = ""
            ViewState("cod_seq") = ""
            ViewState("storer_code") = ""
            ViewState("qty_id") = ""
            ViewState("batch_no") = ""


            ViewState("co_code") = Server.UrlDecode(Request("co_code"))
            ViewState("cod_seq") = Server.UrlDecode(Request("cod_seq"))
            ViewState("storer_code") = Server.UrlDecode(Request("storer_code"))
            ViewState("qty_id") = Server.UrlDecode(Request("qty_id"))
            ViewState("batch_no") = Server.UrlDecode(Request("batch_no"))

            Call BindGV()
        Else

           
        End If

        ar.hideForm(Me)


    End Sub


    Private Function validateAll(Optional ByVal cFlag As String = "") As Boolean
        'Dim total_QTY As Integer = gU.decodeEmptyCInt(Bal_QTY.Text, 0)
        Dim temp_QTY As Double = 0

        Dim tempDT As New DataTable

        Dim co_code As String = ViewState("co_code")
        Dim cod_seq As String = ViewState("cod_seq")

        Dim i, j As Integer
        Dim availDt As DataTable

        availDt = Session("hold_avail_dt")

        tempDT = gU.getSessionTempData(ModuleAbb, "cloneDT", Nothing)

        If Not tempDT Is Nothing AndAlso tempDT.Rows.Count > 0 Then
            For i = 0 To tempDT.Rows.Count - 1
                If tempDT.Rows(i).Item("co_code").ToString.Trim = co_code AndAlso tempDT.Rows(i).Item("cod_seq").ToString.Trim = cod_seq AndAlso tempDT.Rows(i).Item("mFlag").ToString.Trim <> "D" Then
                    If CType(GridView1.Rows(i).FindControl("COH_STATUS"), DropDownList).SelectedValue = "HOLD" Then

                        If CType(GridView1.Rows(i).FindControl("COH_QTY"), TextBox).Text.Trim = "" Then
                            uiFun.displayMsg(Me, "", "Hold Qty. cannot be empty!", Session("gLang"))
                            Return False
                        Else

                            If Not gU.IsWholeNumber(CType(GridView1.Rows(i).FindControl("COH_QTY"), TextBox).Text.Trim) Then
                                uiFun.displayMsg(Me, "", "Hold Qty. must be postive integer!", Session("gLang"))
                                Return False
                            ElseIf CInt(CType(GridView1.Rows(i).FindControl("COH_QTY"), TextBox).Text.Trim) < 0 Then
                                uiFun.displayMsg(Me, "", "Hold Qty. must be larger than 0!", Session("gLang"))
                                Return False
                            Else

                                If CInt(CType(GridView1.Rows(i).FindControl("COH_QTY"), TextBox).Text.Trim) < gU.decodeEmptyCInt(CType(GridView1.Rows(i).FindControl("COH_REL_QTY"), Label).Text.Trim, 0) Then
                                    uiFun.displayMsg(Me, "", "Hold Qty. must be Larger than Picked Qty!", Session("gLang"))
                                    Return False
                                End If

                            End If
                        End If

                        If CType(GridView1.Rows(i).FindControl("COH_TYPE"), DropDownList).SelectedValue = "STOCK" Then
                            'temp_QTY += CType(GridView1.Rows(i).FindControl("COH_QTY"), TextBox).Text.Trim

                            For j = 0 To availDt.Rows.Count - 1
                                If CType(GridView1.Rows(i).FindControl("COH_BATCH_NO"), DropDownList).SelectedValue = availDt.Rows(j).Item("BATCH_NO").ToString.Trim Then
                                    If gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("COH_QTY"), TextBox).Text.Trim, 0) > gU.decodeEmptyCdbl(availDt.Rows(j).Item("AVAIL_QTY").ToString.Trim, 0) Then
                                        uiFun.displayMsg(Me, "", "Hold Stock Qty. cannot larger than available qty!", Session("gLang"))
                                        CType(GridView1.Rows(i).FindControl("COH_QTY"), TextBox).Focus()
                                        Return False
                                    End If
                                End If
                            Next

                        ElseIf CType(GridView1.Rows(i).FindControl("COH_TYPE"), DropDownList).SelectedValue = "RO" Then

                            If CType(GridView1.Rows(i).FindControl("RO_CODE"), TextBox).Text.Trim = "" Then
                                uiFun.displayMsg(Me, "", "Please Select a RO no. to hold RO!", Session("gLang"))
                                Return False
                            End If

                            If gU.decodeEmptyCInt(CType(GridView1.Rows(i).FindControl("ROD_QTY"), TextBox).Text.Trim, 0) < gU.decodeEmptyCInt(CType(GridView1.Rows(i).FindControl("COH_QTY"), TextBox).Text.Trim, 0) Then
                                uiFun.displayMsg(Me, "", "Hold RO Qty. cannot larger than Available RO QTY!", Session("gLang"))
                                Return False
                            End If
                        End If

                    End If

                End If
            Next


            'If temp_QTY > total_QTY Then
            '    uiFun.displayMsg(Me, "", "Hold Stock Qty. cannot larger than stock location balance!", Session("gLang"))
            '    Return False
            'End If

        End If


        'If Status.SelectedValue = "" Then
        '    If Session("gLang") = "E" Then
        '        uiFun.displayMsg(Me, "", lbl_status.Text & " cannot be empty!", Session("gLang"))
        '    Else
        '        uiFun.displayMsg(Me, "", lbl_status.Text & "不能空白!", Session("gLang"))
        '    End If
        '    Return False
        'End If

        'hold ro check (qty > in stock) = true
        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim sql_string As String = ""
        Dim up_sql As String = ""

        Dim c_code As String = ""
        Dim c_seq As String = ""

        Dim tempDT As New DataTable
        Dim tempCODDT As New DataTable
        Dim temp_HOLDQTY As Integer = 0

        If ViewState("co_code") <> "" Then

            c_code = ViewState("co_code")
            c_seq = ViewState("cod_seq")

        Else
            c_code = Server.UrlDecode(Request("co_code"))
            c_seq = Server.UrlDecode(Request("cod_seq"))

        End If

        If validateAll() Then
            tempDT = gU.getSessionTempData(ModuleAbb, "cloneDT", Nothing)

            If Not tempDT Is Nothing AndAlso tempDT.Rows.Count > 0 Then
                For i = 0 To tempDT.Rows.Count - 1
                    If tempDT.Rows(i).Item("cod_seq").ToString.Trim = c_seq Then
                        If tempDT.Rows(i).Item("mFlag").ToString.Trim <> "D" Then
                            tempDT.Rows(i).Item("COH_TYPE") = CType(GridView1.Rows(i).FindControl("COH_TYPE"), DropDownList).SelectedValue
                            tempDT.Rows(i).Item("COH_STATUS") = CType(GridView1.Rows(i).FindControl("COH_STATUS"), DropDownList).SelectedValue
                            tempDT.Rows(i).Item("RO_CODE") = CType(GridView1.Rows(i).FindControl("RO_CODE"), TextBox).Text.Trim
                            tempDT.Rows(i).Item("ROD_SEQ") = CType(GridView1.Rows(i).FindControl("ROD_SEQ"), HiddenField).Value
                            tempDT.Rows(i).Item("COH_QTY") = CType(GridView1.Rows(i).FindControl("COH_QTY"), TextBox).Text.Trim
                            tempDT.Rows(i).Item("COH_BATCH_NO") = CType(GridView1.Rows(i).FindControl("COH_BATCH_NO"), DropDownList).SelectedValue

                            If tempDT.Rows(i).Item("COH_STATUS").ToString.Trim = "HOLD" Then
                                temp_HOLDQTY += gU.decodeEmptyCInt(CType(GridView1.Rows(i).FindControl("COH_QTY"), TextBox).Text.Trim, 0)
                            End If
                        End If
                    End If
                Next

                tempDT.AcceptChanges()
                gU.setSessionTempData(ModuleAbb, "cloneDT", tempDT)

                Dim realDT As DataTable
                realDT = tempDT.Copy
                gU.setSessionTempData(ModuleAbb, "holdDT", tempDT)

                Dim qtyID As String = ViewState("qty_id")

                tempCODDT = gU.getSessionTempData(ModuleAbb, "codDT", Nothing)

                If Not tempCODDT Is Nothing AndAlso tempCODDT.Rows.Count > 0 Then
                    For i = 0 To tempCODDT.Rows.Count - 1
                        If tempCODDT.Rows(i).Item("cod_seq").ToString.Trim = c_seq Then
                            tempCODDT.Rows(i).Item("HOLD_QTY") = temp_HOLDQTY
                        End If
                    Next

                    tempCODDT.AcceptChanges()
                    gU.setSessionTempData(ModuleAbb, "codDT", tempCODDT)

                End If

                ClientScript.RegisterStartupScript(Me.GetType(), "updateHOLDQTY", "window.opener.document.getElementById('" & qtyID & "').value = '" & temp_HOLDQTY & "';", True)
                uiFun.displayMsg(Me, "", "Hold Record is saved temporarily!", Session("gLang"))
                ClientScript.RegisterStartupScript(Me.GetType(), "closeWin", "window.close();", True)
            End If

        End If
    End Sub

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim cat_sql As String = ""
        Dim cat_dt As New DataTable
        Dim dt As New DataTable
        Dim holdDT As New DataTable

        Dim c_code As String = ""
        Dim c_seq As String = ""
        Dim s_code As String = ""
        Dim b_no As String = ""

        Dim total_qty As Integer = 0
        Dim hold_qty As Integer = 0
        Dim avil_qty As Integer = 0

        Dim availDt As DataTable

        REM **********************
        REM Modify Here
        REM Primary Key Session

        If ViewState("co_code") <> "" Then

            c_code = ViewState("co_code")
            c_seq = ViewState("cod_seq")
            s_code = ViewState("storer_code")
            b_no = ViewState("batch_no")
        Else

            c_code = Server.UrlDecode(Request("co_code"))
            c_seq = Server.UrlDecode(Request("cod_seq"))
            s_code = Server.UrlDecode(Request("storer_code"))
            b_no = Server.UrlDecode(Request("batch_no"))

        End If

        REM **********************
        REM Modify Here
        REM Generate Data Table from Header

        SQLString = "select COD_ITM_CODE, COD_ITM_DESC, COD_BATCH_NO, COD_PACK_KEY, COD_PALLET_NO from wms_cust_order_d Where wms_cust_order_d.co_code = '" & gU.dbEncode(c_code) & "' " & _
                    " and wms_cust_order_d.storer_code = '" & gU.dbEncode(s_code) & "' " & _
                    " and wms_cust_order_d.imp_code = '" & gU.dbEncode(Session("imp_code")) & "'" & _
                    " and wms_cust_order_d.co_code = '" & gU.dbEncode(c_code) & "'" & _
                    " and wms_cust_order_d.cod_seq = '" & gU.dbEncode(c_seq) & "'"

        dt = gDB.getDataTable(SQLString)



        If dt.Rows.Count > 0 Then

            cod_itm_desc.Text = dt.Rows(0).Item("cod_itm_desc").ToString.Trim
            cod_itm_code.Text = dt.Rows(0).Item("cod_itm_code").ToString.Trim
            cod_batch_no.Text = b_no
            COD_PACK_KEY.Value = dt.Rows(0).Item("COD_PACK_KEY").ToString.Trim
            COD_PALLET_NO.Value = dt.Rows(0).Item("COD_PALLET_NO").ToString.Trim
            co_code.Text = c_code

            'SQLString = " SELECT SUM(wms_item_loc_bal.ILOC_BAL_QTY) as total_QTY FROM wms_item_loc_bal " & _
            '            " WHERE wms_item_loc_bal.IMP_CODE     = '" & gU.dbEncode(Session("imp_code")) & "' " & _
            '            " AND wms_item_loc_bal.STORER_CODE    = '" & gU.dbEncode(s_code) & "' " & _
            '            " AND wms_item_loc_bal.ITM_CODE       = '" & gU.dbEncode(dt.Rows(0).Item("cod_itm_code").ToString.Trim) & "' " & _
            '            " AND wms_item_loc_bal.PACK_KEY       = '" & gU.dbEncode(dt.Rows(0).Item("COD_PACK_KEY").ToString.Trim) & "' " & _
            '            " AND wms_item_loc_bal.ILOC_PALLET_NO = '" & gU.dbEncode(dt.Rows(0).Item("COD_PALLET_NO").ToString.Trim) & "' " & _
            '            " AND wms_item_loc_bal.ILOC_BATCH_NO  = '" & gU.dbEncode(dt.Rows(0).Item("COD_BATCH_NO").ToString.Trim) & "' "


            'SQLString = " Select (LOC_QTY - Hold_QTY) as total_qty from  " & _
            '            " (SELECT ISNULL(SUM(ILOC_BAL_QTY),0) as LOC_QTY " & _
            '            "  FROM wms_item_loc_bal where wms_item_loc_bal.IMP_CODE= '" & gU.dbEncode(Session("imp_code")) & "' AND wms_item_loc_bal.STORER_CODE='" & gU.dbEncode(s_code) & "' AND wms_item_loc_bal.ITM_CODE='" & gU.dbEncode(dt.Rows(0).Item("cod_itm_code").ToString.Trim) & "' " & _
            '            "  AND wms_item_loc_bal.PACK_KEY= '" & gU.dbEncode(dt.Rows(0).Item("COD_PACK_KEY").ToString.Trim) & "' AND wms_item_loc_bal.ILOC_PALLET_NO=ISNULL('" & gU.dbEncode(dt.Rows(0).Item("COD_PALLET_NO").ToString.Trim) & "','000') AND wms_item_loc_bal.ILOC_BATCH_NO='" & gU.dbEncode(b_no) & "') LOC,  " & _
            '            " (SELECT ISNULL(SUM(COH_QTY),0) as Hold_QTY " & _
            '            " FROM wms_cust_order, wms_cust_order_d, wms_cust_order_hold " & _
            '            "  WHERE wms_cust_order.IMP_CODE     = wms_cust_order_d.IMP_CODE " & _
            '            "  AND wms_cust_order.STORER_CODE = wms_cust_order_d.STORER_CODE " & _
            '            "  AND wms_cust_order.CO_CODE     = wms_cust_order_d.CO_CODE " & _
            '            "  AND wms_cust_order.CO_STATUS <> 'CANCELLED' " & _
            '            "  AND wms_cust_order_d.IMP_CODE     = wms_cust_order_hold.IMP_CODE " & _
            '            "  AND wms_cust_order_d.STORER_CODE = wms_cust_order_hold.STORER_CODE  " & _
            '            "  AND wms_cust_order_d.CO_CODE     = wms_cust_order_hold.CO_CODE  " & _
            '            "  AND wms_cust_order_d.COD_SEQ     = wms_cust_order_hold.COD_SEQ  " & _
            '            "  AND wms_cust_order_hold.COH_STATUS  = 'HOLD' AND wms_cust_order_hold.COH_TYPE  = 'STOCK' " & _
            '            "  AND wms_cust_order.CO_CODE <> '" & gU.dbEncode(c_code) & "' AND wms_cust_order_d.cod_seq <> '" & gU.dbEncode(c_seq) & "' " & _
            '            "  AND wms_cust_order_d.IMP_CODE= '" & gU.dbEncode(Session("imp_code")) & "' AND wms_cust_order_d.STORER_CODE='" & gU.dbEncode(s_code) & "' AND wms_cust_order_d.COD_ITM_CODE='" & gU.dbEncode(dt.Rows(0).Item("cod_itm_code").ToString.Trim) & "' " & _
            '            "  AND wms_cust_order_d.COD_PACK_KEY= '" & gU.dbEncode(dt.Rows(0).Item("COD_PACK_KEY").ToString.Trim) & "' AND ISNULL(wms_cust_order_d.COD_PALLET_NO,'000') = ISNULL('" & gU.dbEncode(dt.Rows(0).Item("COD_PALLET_NO").ToString.Trim) & "','000') AND wms_cust_order_d.COD_BATCH_NO='" & gU.dbEncode(b_no) & "') HOLD "

            'total_qty = gU.decodeEmptyCInt(DB.getValueFromSQL(SQLString), 0)

            SQLString = "select LOC.ILOC_BATCH_NO AS BATCH_NO, LOC.LOC_QTY - isnull(HOLD_STOCK.Hold_QTY, 0) AS AVAIL_QTY " & _
                        "FROM ( " & _
                            "select isnull(l.ILOC_BATCH_NO, '') as ILOC_BATCH_NO, sum(l.ILOC_BAL_QTY - isnull(PICK_ITEM.picked_qty, 0)) as LOC_QTY " & _
                            "from wms_item_loc_bal l " & _
                                "LEFT OUTER JOIN ( " & _
                                    "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " & _
                                    "from wms_do_picklist_d p, wms_delv_order d " & _
                                    "where d.imp_code = p.imp_code " & _
                                    "and d.storer_code = p.storer_code " & _
                                    "and d.do_code = p.do_code " & _
                                    "and d.do_status = 'PICKED' " & _
                                    "and d.do_co_code <> '" & gU.dbEncode(c_code) & "' " & _
                                    "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " & _
                                "ON  l.STORER_CODE = PICK_ITEM.STORER_CODE " & _
                                "AND l.ITM_CODE = PICK_ITEM.PLD_ITEM_NO " & _
                                "AND l.PACK_KEY = PICK_ITEM.PLD_PACK_KEY " & _
                                "AND ISNULL(l.ILOC_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " & _
                                "AND ISNULL(l.ILOC_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " & _
                                "AND l.ILOC_LOC = PICK_ITEM.PLD_LOC " & _
                            "where l.IMP_CODE = '" & gU.dbEncode(Session("imp_code")) & "' " & _
                            "AND l.STORER_CODE = '" & gU.dbEncode(s_code) & "' " & _
                            "AND l.ITM_CODE = '" & gU.dbEncode(dt.Rows(0).Item("cod_itm_code").ToString.Trim) & "' " & _
                            "AND l.PACK_KEY = '" & gU.dbEncode(dt.Rows(0).Item("COD_PACK_KEY").ToString.Trim) & "' " & _
                            "AND l.ILOC_PALLET_NO = ISNULL('" & gU.dbEncode(dt.Rows(0).Item("COD_PALLET_NO").ToString.Trim) & "', '000') "

            If cod_batch_no.Text.Trim <> "" Then
                SQLString &= "AND isnull(l.ILOC_BATCH_NO, '') = isnull('" & gU.dbEncode(cod_batch_no.Text.Trim) & "', '') "
            End If

            SQLString = SQLString & _
                            "and l.ILOC_BAL_QTY > 0 " & _
                            "group by ILOC_BATCH_NO) LOC " & _
                        "left outer join ( " & _
                            "select HOLD.COH_BATCH_NO, sum(hold.hold_qty - ISNULL(pick.picked_qty, 0)) as hold_qty " & _
                            "from (" & _
                                "select h.imp_code, h.storer_code, h.co_code, h.COH_ITM_CODE, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000') as COH_PALLET_NO, " & _
                                    "isnull(h.COH_BATCH_NO, '') as COH_BATCH_NO, SUM(ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0)) as Hold_QTY " & _
                                "from wms_cust_order c, wms_cust_order_d d, wms_cust_order_hold h " & _
                                "WHERE c.IMP_CODE = d.IMP_CODE " & _
                                "AND c.STORER_CODE = d.STORER_CODE " & _
                                "AND c.CO_CODE = d.CO_CODE " & _
                                "AND c.CO_STATUS <> 'CANCELLED' " & _
                                "AND d.IMP_CODE = h.IMP_CODE " & _
                                "AND d.STORER_CODE = h.STORER_CODE " & _
                                "AND d.CO_CODE = h.CO_CODE " & _
                                "AND d.COD_SEQ = h.COD_SEQ " & _
                                "AND h.COH_STATUS = 'HOLD' " & _
                                "AND h.COH_TYPE = 'STOCK' " & _
                                "and h.CO_CODE <> '" & gU.dbEncode(c_code) & "' " & _
                                "AND h.IMP_CODE = '" & gU.dbEncode(Session("imp_code")) & "' " & _
                                "AND h.STORER_CODE = '" & gU.dbEncode(s_code) & "' " & _
                                "AND h.COH_ITM_CODE = '" & gU.dbEncode(dt.Rows(0).Item("cod_itm_code").ToString.Trim) & "' " & _
                                "AND h.COH_PACK_KEY = '" & gU.dbEncode(dt.Rows(0).Item("COD_PACK_KEY").ToString.Trim) & "' " & _
                                "AND ISNULL(h.COH_PALLET_NO,'000') = ISNULL('" & gU.dbEncode(dt.Rows(0).Item("COD_PALLET_NO").ToString.Trim) & "', '000') "

            If cod_batch_no.Text.Trim <> "" Then
                SQLString &= "AND isnull(h.COH_BATCH_NO, '') = isnull('" & gU.dbEncode(cod_batch_no.Text.Trim) & "', '') "
            End If

            SQLString = SQLString & _
                                "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " & _
                                "group by h.imp_code, h.storer_code, h.co_code, h.COH_ITM_CODE, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000'), isnull(h.COH_BATCH_NO, '') " & _
                                ") HOLD " & _
                            "left outer join (" & _
                                "select p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, sum(p.pld_item_qty) as picked_qty " & _
                                "from wms_do_picklist_d p, wms_delv_order d " & _
                                "where d.imp_code = p.imp_code " & _
                                "and d.storer_code = p.storer_code " & _
                                "and d.do_code = p.do_code " & _
                                "and d.do_status = 'PICKED' " & _
                                "group by p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, '') " & _
                                ") pick " &
                            "on hold.imp_code = pick.imp_code " & _
                                "and hold.storer_code = pick.storer_code " & _
                                "and hold.co_code = pick.do_co_code " & _
                                "and hold.COH_ITM_CODE = pick.pld_item_no " & _
                                "and hold.COH_PACK_KEY = pick.pld_pack_key " & _
                                "and hold.COH_PALLET_NO = pick.pld_pallet_no " & _
                                "and hold.COH_BATCH_NO = pick.pld_batch_no " & _
                            "group by HOLD.COH_BATCH_NO) HOLD_STOCK " & _
                        "on LOC.ILOC_BATCH_NO = HOLD_STOCK.COH_BATCH_NO " & _
                        "WHERE LOC.LOC_QTY - isnull(HOLD_STOCK.Hold_QTY, 0) > 0 " & _
                        "order by 1, 2 "

            availDt = gDB.getDataTable(SQLString)

            gvAvailStock.DataSource = availDt
            gvAvailStock.DataBind()

            Session("hold_avail_dt") = availDt
        End If

        'Bal_QTY.Text = total_qty

        REM **********************

        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail

        holdDT = gU.getSessionTempData(ModuleAbb, "holdDT", Nothing)

        Dim holdROstr As String = ""

        For i = 0 To holdDT.Rows.Count - 1
            If holdDT.Rows(i).Item("COH_STATUS").ToString.Trim = "HOLD" AndAlso holdDT.Rows(i).Item("mFlag").ToString.Trim <> "D" Then
                If holdDT.Rows(i).Item("ro_code").ToString.Trim <> "" AndAlso holdDT.Rows(i).Item("co_code").ToString.Trim = c_code AndAlso holdDT.Rows(i).Item("cod_seq").ToString.Trim = c_seq Then
                    holdROstr &= "(" & holdDT.Rows(i).Item("ro_code").ToString.Trim & "||" & holdDT.Rows(i).Item("rod_seq").ToString.Trim & "),"
                End If
            End If
        Next

        gU.setSessionTempData(ModuleAbb, "holdROstr", holdROstr)

        Dim holdDTClone As DataTable

        holdDTClone = holdDT.Copy
        gU.setSessionTempData(ModuleAbb, "cloneDT", holdDTClone)

        GridView1.DataSource = holdDTClone
        GridView1.DataBind()

        REM **********************
    End Sub

    Protected Sub GridView1_DataBinding(sender As Object, e As System.EventArgs) Handles GridView1.DataBinding
        'Dim selectSql As String
        Dim availDt As DataTable

        'selectSql = "select dc_date_code as code, dc_date_code as name " & _
        '            "from wms_date_code " & _
        '            "where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' " & _
        '            "and storer_code = '" & gU.dbEncode(ViewState("storer_code")) & "' " & _
        '            "order by 1"

        'dtBatch = gDB.getDataTable(selectSql)

        availDt = Session("hold_avail_dt")

        dtBatch = availDt.Copy

        For i = 0 To dtBatch.Rows.Count - 1
            If dtBatch.Rows(i).Item("batch_no").ToString.Trim = "" Then
                dtBatch.Rows(i).Delete()
                dtBatch.AcceptChanges()
                Exit For
            End If
        Next

    End Sub

    Protected Sub GridView1_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Dim tempDT As New DataTable

        tempDT = gU.getSessionTempData(ModuleAbb, "cloneDT", Nothing)

        Select Case e.CommandName
            Case "DEL"

                tempDT.Rows(e.CommandArgument).Item("mFlag") = "D"
                tempDT.AcceptChanges()

                Dim SelectedRO As String = gU.getSessionTempData(ModuleAbb, "holdROstr", "")

                SelectedRO = Replace(SelectedRO, "(" & CType(GridView1.Rows(e.CommandArgument).FindControl("RO_CODE"), TextBox).Text.Trim & "||" & CType(GridView1.Rows(e.CommandArgument).FindControl("ROD_SEQ"), HiddenField).Value & "),", "")
                gU.setSessionTempData(ModuleAbb, "holdROstr", SelectedRO)

                gU.setSessionTempData(ModuleAbb, "cloneDT", tempDT)
                GridView1.DataSource = tempDT
                GridView1.DataBind()





        End Select
    End Sub


    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                uiFun.load_dropdown(CType(e.Row.FindControl("coh_batch_no"), DropDownList), dtBatch, "BATCH_NO", "BATCH_NO", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "coh_batch_no").ToString.Trim)

                CType(e.Row.FindControl("coh_type"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "coh_type").ToString.Trim
                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                    CType(e.Row.FindControl("coh_type"), DropDownList).Enabled = True
                    CType(e.Row.FindControl("coh_status"), DropDownList).Enabled = False
                Else
                    CType(e.Row.FindControl("coh_type"), DropDownList).Enabled = False
                    CType(e.Row.FindControl("coh_status"), DropDownList).Enabled = True
                End If
                CType(e.Row.FindControl("coh_status"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "coh_status").ToString.Trim
                CType(e.Row.FindControl("ori_status"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ori_status").ToString.Trim
                CType(e.Row.FindControl("sys_cd"), Label).Text = DataBinder.Eval(e.Row.DataItem, "sys_cd").ToString.Trim
                CType(e.Row.FindControl("COH_REL_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "COH_REL_DATE").ToString.Trim
                CType(e.Row.FindControl("ro_code"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ro_code").ToString.Trim
                CType(e.Row.FindControl("ro_code"), TextBox).Attributes.Add("readonly", "readonly")
                CType(e.Row.FindControl("rod_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "rod_seq").ToString.Trim

                CType(e.Row.FindControl("ROD_QTY"), TextBox).Attributes.Add("readonly", "readonly")

                Dim rod_qty As Integer = 0
                Dim holdRO_QTY As Integer = 0
                Dim getvalueSQL As String = ""
                Dim itemStr As String = cod_itm_code.Text.Trim & "|*|" & cod_batch_no.Text.Trim & "|*|" & COD_PACK_KEY.Value.Trim & "|*|" & COD_PALLET_NO.Value & "|*|" & ViewState("storer_code") & "|*|" & ViewState("co_code") & "|*|" & ViewState("cod_seq")
                If DataBinder.Eval(e.Row.DataItem, "coh_type").ToString.Trim = "RO" Then
                    getvalueSQL = "Select ROD_QTY from WMS_REPLENISH_D Where IMP_CODE='" & Session("IMP_CODE") & "' AND STORER_CODE='" & ViewState("storer_code") & "' AND RO_CODE='" & DataBinder.Eval(e.Row.DataItem, "ro_code").ToString.Trim & "' AND ROD_SEQ='" & DataBinder.Eval(e.Row.DataItem, "rod_seq").ToString.Trim & "' "

                    rod_qty = gU.decodeEmptyCInt(DB.getValueFromSQL(getvalueSQL), 0)

                    getvalueSQL = " Select ISNULL(SUM(COH_QTY),0) as hold_qty FROM wms_cust_order_hold, wms_cust_order " & _
                                  " WHERE wms_cust_order_hold.imp_code = wms_cust_order.imp_code " & _
                                  " AND wms_cust_order_hold.storer_code = wms_cust_order.storer_code " & _
                                  " AND wms_cust_order_hold.co_code = wms_cust_order.co_code " & _
                                  " AND wms_cust_order_hold.imp_code='" & Session("IMP_CODE") & "' AND wms_cust_order_hold.storer_code='" & ViewState("storer_code") & "' and coh_status='HOLD' and coh_type='RO' " & _
                                  " AND ro_code='" & DataBinder.Eval(e.Row.DataItem, "ro_code").ToString.Trim & "' and ROD_SEQ='" & DataBinder.Eval(e.Row.DataItem, "rod_seq").ToString.Trim & "' AND wms_cust_order.co_status <> 'CANCELLED' "

                    holdRO_QTY = gU.decodeEmptyCInt(DB.getValueFromSQL(getvalueSQL), 0)

                    CType(e.Row.FindControl("ROD_QTY"), TextBox).Text = rod_qty - holdRO_QTY

                    CType(e.Row.FindControl("ro_code"), TextBox).CssClass = "REQUIRED"
                    CType(e.Row.FindControl("Image_RO_LookUp"), Image).Attributes.Add("onclick", "javascript:ROLookUp('" & itemStr & "', '" & CType(e.Row.FindControl("ro_code"), TextBox).ClientID & "'," & _
                                                                                                 "'" & CType(e.Row.FindControl("rod_seq"), HiddenField).ClientID & "', '" & CType(e.Row.FindControl("ROD_QTY"), TextBox).ClientID & "');")

                Else

                    CType(e.Row.FindControl("ro_code"), TextBox).BackColor = Drawing.ColorTranslator.FromHtml("#F0F0F0")

                End If


                CType(e.Row.FindControl("COH_QTY"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "COH_QTY").ToString.Trim
                CType(e.Row.FindControl("COH_IN_STOCK_QTY"), Label).Text = DataBinder.Eval(e.Row.DataItem, "COH_IN_STOCK_QTY").ToString.Trim
                CType(e.Row.FindControl("COH_REL_QTY"), Label).Text = gU.decodeNullOrEmpty(DataBinder.Eval(e.Row.DataItem, "COH_REL_QTY").ToString.Trim, "0")

                CType(e.Row.FindControl("btnDEL"), Button).CommandArgument = e.Row.RowIndex

                If DataBinder.Eval(e.Row.DataItem, "cod_seq").ToString.Trim <> ViewState("cod_seq") Then e.Row.Visible = False
                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then e.Row.Visible = False

        End Select



    End Sub

    Protected Sub ddl_TypeChanged(sender As Object, e As EventArgs)

        Dim ddl As DropDownList = sender

        Dim grow As GridViewRow = ddl.NamingContainer
        Dim itemStr As String = cod_itm_code.Text.Trim & "|*|" & cod_batch_no.Text.Trim & "|*|" & COD_PACK_KEY.Value.Trim & "|*|" & COD_PALLET_NO.Value & "|*|" & ViewState("storer_code") & "|*|" & ViewState("co_code") & "|*|" & ViewState("cod_seq")


        If ddl.SelectedValue = "STOCK" Then
            CType(grow.FindControl("Image_RO_LookUp"), Image).Attributes.Remove("onclick")
            CType(grow.FindControl("ro_code"), TextBox).BackColor = Drawing.ColorTranslator.FromHtml("#F0F0F0")

        ElseIf ddl.SelectedValue = "RO" Then
            CType(grow.FindControl("Image_RO_LookUp"), Image).Attributes.Add("onclick", "javascript:ROLookUp('" & itemStr & "', '" & CType(grow.FindControl("ro_code"), TextBox).ClientID & "'," & _
                                                                             "'" & CType(grow.FindControl("rod_seq"), HiddenField).ClientID & "', '" & CType(grow.FindControl("ROD_QTY"), TextBox).ClientID & "');")
            CType(grow.FindControl("ro_code"), TextBox).BackColor = Drawing.ColorTranslator.FromHtml("#FFFA9C")
        End If


    End Sub

    Protected Sub btnSave1_Click(sender As Object, e As System.EventArgs) Handles btnSave1.Click
        Call save()
    End Sub

    Protected Sub btnSave2_Click(sender As Object, e As System.EventArgs) Handles btnSave2.Click
        Call save()
    End Sub

    Protected Sub btnAdd_Click(sender As Object, e As System.EventArgs) Handles btnAdd.Click
        Dim tempDT As New DataTable
        Dim newRow As DataRow

        tempDT = gU.getSessionTempData(ModuleAbb, "cloneDT", Nothing)
        If validateAll() Then
            If cU.gfBuildDataTableforGridView(tempDT, GridView1, True) Then

                newRow = tempDT.NewRow

                newRow.Item("mFlag") = "N"
                newRow.Item("IMP_CODE") = Session("IMP_CODE")
                newRow.Item("STORER_CODE") = ViewState("storer_code")
                newRow.Item("co_code") = ViewState("co_code")
                newRow.Item("COD_SEQ") = ViewState("cod_seq")
                newRow.Item("COH_TYPE") = "STOCK"
                newRow.Item("COH_STATUS") = "HOLD"
                newRow.Item("ori_status") = "HOLD"
                newRow.Item("COH_IN_STOCK_QTY") = 0
                newRow.Item("COH_QTY") = 0

                tempDT.Rows.Add(newRow)
                tempDT.AcceptChanges()

                gU.setSessionTempData(ModuleAbb, "cloneDT", tempDT)

                GridView1.DataSource = tempDT
                GridView1.DataBind()

            End If
        End If
    End Sub
End Class