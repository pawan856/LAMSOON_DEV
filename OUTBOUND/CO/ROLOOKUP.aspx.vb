Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient


Partial Class OUTBOUND_CO_ROLOOKUP
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private cU As New CommonUtils
    Private ar As New AccessRightUtils
    Private ModuleAbb As String = "CO_MAIN"

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
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
            ViewState("ITM_STR") = ""
            ViewState("RO_ID") = ""
            ViewState("SEQ_ID") = ""
            ViewState("QTY_ID") = ""

            ViewState("ITM_STR") = Server.UrlDecode(Request("ITM_STR"))
            ViewState("RO_ID") = Server.UrlDecode(Request("RO_ID"))
            ViewState("SEQ_ID") = Server.UrlDecode(Request("SEQ_ID"))
            ViewState("QTY_ID") = Server.UrlDecode(Request("QTY_ID"))
            ViewState("ori_RO") = Server.UrlDecode(Request("ori_RO"))

            Call BindGV()
        End If


    End Sub

    Public Sub BindGV()

        Dim itemStr As String = ""
        Dim itemPar() As String

        Dim dt As New DataTable

        itemStr = ViewState("ITM_STR")
        itemPar = Split(itemStr, "|*|")

        Dim SQLString As String = ""

        'itemStr = cod_itm_code.Text.Trim & "|*|" & cod_batch_no.Text.Trim & "|*|" & COD_PACK_KEY.Value.Trim & "|*|" & COD_PALLET_NO.Value


        item_code.Text = itemPar(0)
        batch_no.Text = itemPar(1)
        pack_key.Text = itemPar(2)
        pallet_no.Text = itemPar(3)
        storer_code.Value = itemPar(4)
        co_code.Value = itemPar(5)
        cod_seq.Value = itemPar(6)
        'SQLString = " SELECT wms_replenish.ro_code, wms_replenish_d.rod_seq, wms_storer.sto_name, wms_replenish_d.rod_qty," & _
        '            " to_char(rod_manu_date,'DD/MM/YYYY') as rod_manu_date, TO_CHAR(rod_expiry_date,'DD/MM/YYYY') as rod_expiry_date  " & _
        '            " FROM wms_replenish, wms_replenish_d, wms_storer  " & _
        '            " Where wms_replenish.IMP_CODE  = wms_replenish_d.IMP_CODE  " & _
        '            " AND wms_replenish.STORER_CODE = wms_replenish_d.STORER_CODE  " & _
        '            " AND wms_replenish.RO_CODE     = wms_replenish_d.RO_CODE " & _
        '            " AND wms_replenish.imp_code = wms_storer.imp_code " & _
        '            " AND wms_replenish.storer_code = wms_storer.storer_code " & _
        '            " AND wms_replenish_d.rod_qty > 0  AND wms_replenish.ro_status <> 'CLOSED' " & _
        '            " AND wms_replenish.imp_code='" & Session("IMP_CODE") & "' AND wms_replenish.STORER_CODE='" & gU.dbEncode(itemPar(4)) & "' " & _
        '            " AND wms_replenish_d.rod_itm_code='" & gU.dbEncode(itemPar(0)) & "' AND wms_replenish_d.rod_batch_no='" & gU.dbEncode(itemPar(1)) & "' " & _
        '            " AND wms_replenish_d.rod_pack_key='" & gU.dbEncode(itemPar(2)) & "' and ISNULL(wms_replenish_d.ROD_PALLET_NO,'000') =ISNULL('" & gU.dbEncode(itemPar(3)) & "','000') "

        SQLString = " SELECT wms_replenish.ro_code,  wms_replenish_d.rod_seq,  wms_storer.sto_name,  wms_replenish_d.rod_qty - ISNULL(gr.gr_qty,0)  as rod_qty, " & _
                    " Convert(varchar,rod_manu_date,103)   AS rod_manu_date,  Convert(varchar,rod_expiry_date,103) AS rod_expiry_date " & _
                    " FROM wms_replenish inner join wms_replenish_d " & _
                    " on  wms_replenish.IMP_CODE                 = wms_replenish_d.IMP_CODE AND wms_replenish.STORER_CODE = wms_replenish_d.STORER_CODE " & _
                    " AND wms_replenish.RO_CODE                    = wms_replenish_d.RO_CODE " & _
                    " inner join wms_storer " & _
                    " on wms_replenish.imp_code = wms_storer.imp_code AND wms_replenish.storer_code = wms_storer.storer_code " & _
                    " left outer join ( " & _
                    " SELECT ISNULL(Sum(grd_po_qty),0) as GR_QTY,wms_replenish.ro_code  FROM wms_replenish, wms_replenish_d, wms_goodsrcv_d, wms_goodsrcv " & _
                    " WHERE wms_replenish.IMP_CODE     = wms_replenish_d.IMP_CODE AND wms_replenish.STORER_CODE = wms_replenish_d.STORER_CODE " & _
                    " AND wms_replenish.RO_CODE     = wms_replenish_d.RO_CODE AND wms_goodsrcv_d.IMP_CODE     = wms_replenish.IMP_CODE " & _
                    " AND wms_goodsrcv_d.STORER_CODE = wms_replenish.STORER_CODE AND wms_replenish_d.ROD_ITM_CODE   = wms_goodsrcv_d.GRD_ITM_CODE " & _
                    " AND wms_replenish_d.ROD_PACK_KEY  = wms_goodsrcv_d.GRD_PACK_KEY AND wms_replenish_d.ROD_PALLET_NO = wms_goodsrcv_d.GRD_PALLET_NO " & _
                    " AND wms_replenish_d.ROD_BATCH_NO  = wms_goodsrcv_d.GRD_BATCH_NO AND wms_goodsrcv_d.IMP_CODE     = wms_goodsrcv.IMP_CODE " & _
                    " AND wms_goodsrcv_d.STORER_CODE = wms_goodsrcv.STORER_CODE AND wms_goodsrcv_d.GR_CODE     = wms_goodsrcv.GR_CODE " & _
                    " AND wms_goodsrcv.GR_DOC_NO = wms_replenish.ro_code " & _
                    " AND wms_goodsrcv.GR_DOC_TYPE ='RO' " & _
                    " AND wms_replenish.ro_status                 <> 'CLOSED' " & _
                    " AND wms_replenish.imp_code                   ='" & Session("IMP_CODE") & "' " & _
                    " AND wms_replenish.STORER_CODE                ='" & gU.dbEncode(itemPar(4)) & "' " & _
                    " AND wms_replenish_d.rod_itm_code             ='" & gU.dbEncode(itemPar(0)) & "' " & _
                    " AND wms_replenish_d.rod_batch_no             ='" & gU.dbEncode(itemPar(1)) & "' " & _
                    " AND wms_replenish_d.rod_pack_key             ='" & gU.dbEncode(itemPar(2)) & "' " & _
                    " AND ISNULL(wms_replenish_d.ROD_PALLET_NO,'000') =ISNULL('" & gU.dbEncode(itemPar(3)) & "','000')  " & _
                    " AND wms_goodsrcv.gr_status = 'POSTED'   group by wms_replenish.ro_code) gr " & _
                    " ON wms_replenish.ro_code = gr.ro_code " & _
                    " WHERE wms_replenish_d.rod_qty                  > 0 " & _
                    " AND wms_replenish.ro_status                 <> 'CLOSED' " & _
                    " AND wms_replenish.imp_code                   ='" & Session("IMP_CODE") & "' " & _
                    " AND wms_replenish.STORER_CODE                ='" & gU.dbEncode(itemPar(4)) & "' " & _
                    " AND wms_replenish_d.rod_itm_code             ='" & gU.dbEncode(itemPar(0)) & "' " & _
                    " AND wms_replenish_d.rod_batch_no             ='" & gU.dbEncode(itemPar(1)) & "' " & _
                    " AND wms_replenish_d.rod_pack_key             ='" & gU.dbEncode(itemPar(2)) & "' " & _
                    " AND ISNULL(wms_replenish_d.ROD_PALLET_NO,'000') =ISNULL('" & gU.dbEncode(itemPar(3)) & "','000') "


        dt = gDB.getDataTable(SQLString)

        If dt.Rows.Count > 0 Then
            storer_name.Text = dt.Rows(0).Item("sto_name").ToString.Trim
            GridView1.DataSource = dt
        Else
            storer_name.Text = itemPar(4)
            GridView1.DataSource = Nothing
        End If

        GridView1.DataBind()

    End Sub

    Protected Sub GridView1_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Dim ro_code As String = ""
        Dim ro_seq As String = ""
        Dim rod_qty As String = ""

        Dim ro_code_id As String = ViewState("RO_ID")
        Dim rod_seq_id As String = ViewState("SEQ_ID")
        Dim qty_id As String = ViewState("QTY_ID")

        Dim rindex As Integer
        Dim getvalueSQL As String = ""

        Dim holdRO_QTY As Integer = 0
        Dim selectedRO As String = gU.getSessionTempData(ModuleAbb, "holdROstr", "")

        Select Case e.CommandName
            Case "SELECT"

                rindex = e.CommandArgument

                ro_code = CType(GridView1.Rows(rindex).FindControl("ro_code"), Label).Text.Trim
                ro_seq = CType(GridView1.Rows(rindex).FindControl("rod_seq"), Label).Text.Trim
                rod_qty = CType(GridView1.Rows(rindex).FindControl("rod_qty"), Label).Text.Trim

                getvalueSQL = " Select ISNULL(SUM(COH_QTY),0) as hold_qty FROM wms_cust_order_hold, wms_cust_order " & _
                              " WHERE wms_cust_order_hold.imp_code = wms_cust_order.imp_code " & _
                              " AND wms_cust_order_hold.storer_code = wms_cust_order.storer_code " & _
                              " AND wms_cust_order_hold.co_code = wms_cust_order.co_code " & _
                              " AND wms_cust_order_hold.imp_code='" & Session("IMP_CODE") & "' AND wms_cust_order_hold.storer_code='" & gU.dbEncode(storer_code.Value) & "'" & _
                              " AND wms_cust_order_hold.co_code <> '" & gU.dbEncode(co_code.Value) & "' AND wms_cust_order_hold.cod_seq <> '" & gU.dbEncode(cod_seq.Value) & "' " & _
                              " and coh_status='HOLD' and coh_type='RO' " & _
                              " AND ro_code='" & gU.dbEncode(ro_code) & "' and ROD_SEQ='" & gU.dbEncode(ro_seq) & "' AND wms_cust_order.co_status not in ('CLOSED', 'CANCELLED') "

                holdRO_QTY = gU.decodeEmptyCInt(DB.getValueFromSQL(getvalueSQL), 0)

                selectedRO = Replace(selectedRO, ViewState("ori_RO"), "")
                gU.setSessionTempData(ModuleAbb, "holdROstr", selectedRO & "(" & ro_code & "||" & ro_seq & "),")

                ClientScript.RegisterStartupScript(Me.GetType(), "ROCode", "window.opener.document.getElementById('" & ro_code_id & "').value = '" & ro_code & "';", True)
                ClientScript.RegisterStartupScript(Me.GetType(), "ROSeq", "window.opener.document.getElementById('" & rod_seq_id & "').value = '" & ro_seq & "';", True)
                ClientScript.RegisterStartupScript(Me.GetType(), "ROQTY", "window.opener.document.getElementById('" & qty_id & "').value = '" & rod_qty - holdRO_QTY & "';", True)
                ClientScript.RegisterStartupScript(Me.GetType(), "CloseWin", "window.close();", True)

        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim holdROstr As String = gU.getSessionTempData(ModuleAbb, "holdROstr", "")
        Dim getValueSQL As String = ""
        holdROstr = Replace(holdROstr, ViewState("ori_RO"), "")

        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("ro_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ro_code").ToString.Trim
                CType(e.Row.FindControl("rod_seq"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_seq").ToString.Trim
                CType(e.Row.FindControl("rod_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "rod_qty").ToString.Trim
                CType(e.Row.FindControl("btnSel"), Button).CommandArgument = e.Row.RowIndex



                If gU.decodeEmptyCInt(DataBinder.Eval(e.Row.DataItem, "rod_qty").ToString.Trim, 0) <= 0 Then
                    e.Row.Visible = False
                End If

                If InStr(holdROstr, "(" & DataBinder.Eval(e.Row.DataItem, "ro_code").ToString.Trim & "||" & DataBinder.Eval(e.Row.DataItem, "rod_seq").ToString.Trim & ")") > 0 Then
                    e.Row.Visible = False
                End If

        End Select
    End Sub
End Class
