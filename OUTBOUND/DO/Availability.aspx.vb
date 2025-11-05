Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class Availability
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

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
            lheader.Text = "Quantity not Available"
            'saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            'saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[No. will be auto generated]"
            'End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "不可用检查"
            'saveBtn1.OnClientClick = "return confirm(""确定保存资料?"");"
            'saveBtn2.OnClientClick = "return confirm(""确定保存资料?"");"
            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[号码会自动产生]"
            'End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS
        'RT_TYPE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            'DO_CODE.CssClass = "REQUIRED"
            'STORER_CODE.CssClass = "REQUIRED"
        Else
            'DO_CODE.Enabled = False
            'STORER_CODE.Enabled = False
        End If

        'dt = Session("_M_OB_DO_TMP_pl_dt")
        If Not IsPostBack Then
            Call BindGV()
        End If

        'If DO_STATUS.Text = "CANCELLED" Then
        '    ar.sec_write = "N"
        '    CancelBtn.Visible = False
        'ElseIf DO_STATUS.Text = "POSTED" Then
        '    ar.sec_write = "N"
        '    btnPost.Visible = False
        'End If

        ar.hideForm(Me, editMode)
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here
                Call cU.changeGVLabel(oGridViewRow, e, "Item No.", "物件编号")
                Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封装内码")
                Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物件名称")

                Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板編號")
                Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "批次編號")

                Call cU.changeGVLabel(oGridViewRow, e, "Req. Qty", "选取数量")

                Call cU.changeGVLabel(oGridViewRow, e, "Avail. Qty", "可取数量")
                Call cU.changeGVLabel(oGridViewRow, e, "Hold Qty", "留貨數")
                Call cU.changeGVLabel(oGridViewRow, e, "Picked Qty", "已撿取數")
                Call cU.changeGVLabel(oGridViewRow, e, "Bal. Qty", "仓存数量")
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                REM **********************
                REM Modify Here

                CType(e.Row.FindControl("itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim
                CType(e.Row.FindControl("pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim
                CType(e.Row.FindControl("itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_name").ToString.Trim

                CType(e.Row.FindControl("pallet_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pallet_no").ToString.Trim
                CType(e.Row.FindControl("batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "batch_no").ToString.Trim

                CType(e.Row.FindControl("dod_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "dod_qty").ToString.Trim

                CType(e.Row.FindControl("avail_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "avail_qty").ToString.Trim
                CType(e.Row.FindControl("hold_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "hold_qty").ToString.Trim
                CType(e.Row.FindControl("itm_balance"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_balance").ToString.Trim

                CType(e.Row.FindControl("picked_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "picked_qty").ToString.Trim

                REM **********************
        End Select
    End Sub

    Protected Sub BindGV()
        Dim SQLString, selectSql As String
        Dim dt As DataTable

        'SQLString = "select i.itm_code, i.pack_key, i.itm_name, i.itm_balance, d.dod_qty " & _
        '            "from wms_item i, wms_delv_order_d d " & _
        '            "where d.imp_code = '" & gU.dbEncode(Server.UrlDecode(Request("imp_code"))) & "' " & _
        '            "and d.storer_code = '" & gU.dbEncode(Server.UrlDecode(Request("storer_code"))) & "' " & _
        '            "and d.do_code = '" & gU.dbEncode(Server.UrlDecode(Request("do_code"))) & "' " & _
        '            "and i.imp_code = d.imp_code " & _
        '            "and i.storer_code = d.storer_code " & _
        '            "and i.itm_code = d.dod_itm_code " & _
        '            "and i.pack_key = d.dod_pack_key " & _
        '            "and (i.itm_balance = 0 or ISNULL(i.itm_balance, 0) < ISNULL(d.dod_qty, 0)) " & _
        '            "order by i.itm_code, i.pack_key, i.itm_balance desc "


        SQLString = "select delv.imp_code, delv.storer_code, delv.dod_itm_code as itm_code, i.itm_name, delv.dod_pack_key as pack_key, delv.dod_pallet_no as pallet_no, delv.dod_batch_no as batch_no, delv.dod_qty, " &
                        "Case delv.dod_batch_no when null then HOLD_STOCK2.hold_qty else HOLD_STOCK.hold_qty end as hold_qty, " &
                        "ISNULL(Case when delv.dod_batch_no is null then bal2.bal_qty else bal.bal_qty end, 0) as itm_balance, " &
                        "Case when delv.dod_batch_no is null then pck2.picked_qty else pck.picked_qty end as picked_qty, " &
                        "case when " &
                            "ISNULL(case when delv.dod_batch_no  is null then bal2.bal_qty else bal.bal_qty end, 0) - " &
                                "ISNULL(case when delv.dod_batch_no is null then HOLD_STOCK2.hold_qty else HOLD_STOCK.hold_qty end, 0) - " &
                                "ISNULL(case when delv.dod_batch_no is null then pck2.picked_qty else pck.picked_qty end, 0) < 0 then 0 else " &
                            "ISNULL(case when delv.dod_batch_no is null then bal2.bal_qty else bal.bal_qty end, 0) - " &
                                "ISNULL(case when delv.dod_batch_no is null then HOLD_STOCK2.hold_qty else HOLD_STOCK.hold_qty end, 0) - " &
                                "ISNULL(case when delv.dod_batch_no is null then pck2.picked_qty else pck.picked_qty end, 0) end as avail_qty " &
                    "from ( " &
                        "select imp_code, storer_code, dod_itm_code, dod_pack_key, ISNULL(dod_pallet_no, '000') as dod_pallet_no, dod_batch_no, sum(dod_qty) as dod_qty " &
                        "from wms_delv_order_d d " &
                        "where imp_code = '" & gU.dbEncode(Server.UrlDecode(Request("imp_code"))) & "' " &
                        "and storer_code = '" & gU.dbEncode(Server.UrlDecode(Request("storer_code"))) & "' " &
                        "and DO_CODE = '" & gU.dbEncode(Server.UrlDecode(Request("do_code"))) & "' " &
                        "group by imp_code, storer_code, dod_itm_code, dod_pack_key, ISNULL(dod_pallet_no, '000'), dod_batch_no) delv inner join wms_item i " &
                        "on i.imp_code = delv.imp_code " &
                        "and i.storer_code = delv.storer_code " &
                        "and i.itm_code = delv.dod_itm_code " &
                        "and i.pack_key = delv.dod_pack_key " &
                        "left outer join ( " &
                            "select hold.imp_code, " &
                                "hold.storer_code, " &
                                "hold.COD_ITM_CODE, " &
                                "hold.COD_PACK_KEY, " &
                                "hold.COD_PALLET_NO, " &
                                "hold.COD_BATCH_NO, " &
                                "sum(hold.hold_qty - ISNULL(pick.picked_qty, 0)) as hold_qty " &
                            "from ( " &
                                "select h.imp_code, " &
                                    "h.storer_code, " &
                                    "h.co_code, " &
                                    "h.COH_ITM_CODE AS COD_ITM_CODE, " &
                                    "h.COH_PACK_KEY AS COD_PACK_KEY, " &
                                    "ISNULL(h.COH_PALLET_NO, '000') as COD_PALLET_NO, " &
                                    "h.COH_BATCH_NO AS COD_BATCH_NO, " &
                                    "sum(ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0)) as hold_qty " &
                                "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " &
                                "where h.imp_code = d.imp_code " &
                                "and h.storer_code = d.storer_code " &
                                "and h.co_code = d.co_code " &
                                "and h.cod_seq = d.cod_seq " &
                                "and c.imp_code = d.imp_code " &
                                "and c.storer_code = d.storer_code " &
                                "and c.co_code = d.co_code " &
                                "and h.coh_status <> 'RELEASE' " &
                                "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " &
                                "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " &
                                "and c.co_code <> '" & gU.dbEncode(Server.UrlDecode(Request("co_code"))) & "' " &
                                "group by h.imp_code, h.storer_code, h.co_code, h.COH_ITM_CODE, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000'), h.COH_BATCH_NO) hold " &
                            "left outer join " &
                                "( " &
                                "select p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, sum(p.pld_item_qty) as picked_qty " &
                                "from wms_do_picklist_d p, wms_delv_order d " &
                                "where d.imp_code = p.imp_code " &
                                "and d.storer_code = p.storer_code " &
                                "and d.do_code = p.do_code " &
                                "and d.do_status = 'PICKED' " &
                                "group by p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, '') " &
                                ") pick " &
                            "on hold.imp_code = pick.imp_code " &
                            "and hold.storer_code = pick.storer_code " &
                            "and hold.co_code = pick.do_co_code " &
                            "and hold.COD_ITM_CODE = pick.pld_item_no " &
                            "and hold.COD_PACK_KEY = pick.pld_pack_key " &
                            "and hold.COD_PALLET_NO = pick.pld_pallet_no " &
                            "and hold.COD_BATCH_NO = pick.pld_batch_no " &
                            "group by hold.imp_code, hold.storer_code, hold.COD_ITM_CODE, hold.COD_PACK_KEY, hold.COD_PALLET_NO, hold.COD_BATCH_NO) HOLD_STOCK " &
                        "on delv.imp_code = HOLD_STOCK.imp_code " &
                            "and delv.storer_code = HOLD_STOCK.storer_code " &
                            "and delv.dod_itm_code = HOLD_STOCK.cod_itm_code " &
                            "and delv.dod_pack_key = HOLD_STOCK.cod_pack_key " &
                            "and delv.dod_pallet_no = HOLD_STOCK.cod_pallet_no " &
                            "and delv.dod_batch_no = HOLD_STOCK.cod_batch_no " &
                        "left outer join ( " &
                            "select hold.imp_code, " &
                                "hold.storer_code, " &
                                "hold.COD_ITM_CODE, " &
                                "hold.COD_PACK_KEY, " &
                                "hold.COD_PALLET_NO, " &
                                "sum(hold.hold_qty - ISNULL(pick.picked_qty, 0)) as hold_qty " &
                            "from ( " &
                                "select h.imp_code, " &
                                    "h.storer_code, " &
                                    "h.co_code, " &
                                    "h.COH_ITM_CODE AS COD_ITM_CODE, " &
                                    "h.COH_PACK_KEY AS COD_PACK_KEY, " &
                                    "ISNULL(h.COH_PALLET_NO, '000') as COD_PALLET_NO, " &
                                    "sum(ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0)) as hold_qty " &
                                "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " &
                                "where h.imp_code = d.imp_code " &
                                "and h.storer_code = d.storer_code " &
                                "and h.co_code = d.co_code " &
                                "and h.cod_seq = d.cod_seq " &
                                "and c.imp_code = d.imp_code " &
                                "and c.storer_code = d.storer_code " &
                                "and c.co_code = d.co_code " &
                                "and h.coh_status <> 'RELEASE' " &
                                "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " &
                                "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " &
                                "and c.co_code <> '" & gU.dbEncode(Server.UrlDecode(Request("co_code"))) & "' " &
                                "group by h.imp_code, h.storer_code, h.co_code, h.COH_ITM_CODE, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000')) hold " &
                            "left outer join " &
                                "( " &
                                "select p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, sum(p.pld_item_qty) as picked_qty " &
                                "from wms_do_picklist_d p, wms_delv_order d " &
                                "where d.imp_code = p.imp_code " &
                                "and d.storer_code = p.storer_code " &
                                "and d.do_code = p.do_code " &
                                "and d.do_status = 'PICKED' " &
                                "group by p.imp_code, p.storer_code, d.do_co_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') " &
                                ") pick " &
                            "on hold.imp_code = pick.imp_code " &
                            "and hold.storer_code = pick.storer_code " &
                            "and hold.co_code = pick.do_co_code " &
                            "and hold.COD_ITM_CODE = pick.pld_item_no " &
                            "and hold.COD_PACK_KEY = pick.pld_pack_key " &
                            "and hold.COD_PALLET_NO = pick.pld_pallet_no " &
                            "group by hold.imp_code, hold.storer_code, hold.COD_ITM_CODE, hold.COD_PACK_KEY, hold.COD_PALLET_NO) HOLD_STOCK2 " &
                        "on delv.imp_code = HOLD_STOCK2.imp_code  " &
                            "and delv.storer_code = HOLD_STOCK2.storer_code  " &
                            "and delv.dod_itm_code = HOLD_STOCK2.cod_itm_code  " &
                            "and delv.dod_pack_key = HOLD_STOCK2.cod_pack_key  " &
                            "and delv.dod_pallet_no = HOLD_STOCK2.cod_pallet_no  " &
                        "Left outer join ( " &
                            "select l.imp_code, l.storer_code, l.ITM_CODE, l.PACK_KEY, ISNULL(l.ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, " &
                                "l.ILOC_BATCH_NO, sum(l.ILOC_BAL_QTY) as BAL_QTY " &
                            "from WMS_ITEM_LOC_BAL l, WMS_WH_AREA a " &
                            "where ISNULL(ILOC_BAL_QTY, 0) > 0 " &
                            "and l.IMP_CODE = a.IMP_CODE " &
                            "and l.ILOC_WH = a.WH_CODE " &
                            "and l.ILOC_FLOOR = a.FL_NUM " &
                            "and l.ILOC_AREA = a.AR_CODE " &
                            "and ISNULL(a.AR_DAMAGE_YN, '') <> 'Y' " &
                            "group by l.imp_code, l.storer_code, l.ITM_CODE, l.PACK_KEY, ISNULL(l.ILOC_PALLET_NO, '000'), l.ILOC_BATCH_NO) bal " &
                            "on delv.imp_code = bal.imp_code  " &
                            "and delv.storer_code = bal.storer_code  " &
                            "and delv.dod_itm_code = bal.ITM_CODE  " &
                            "and delv.dod_pack_key = bal.PACK_KEY  " &
                            "and delv.dod_pallet_no = bal.iloc_pallet_no  " &
                            "and delv.dod_batch_no = bal.iloc_batch_no  " &
                        "left outer join ( " &
                            "select l.imp_code, l.storer_code, l.ITM_CODE, l.PACK_KEY, ISNULL(l.ILOC_PALLET_NO, '000') AS ILOC_PALLET_NO, sum(l.ILOC_BAL_QTY) as BAL_QTY " &
                            "from WMS_ITEM_LOC_BAL l, WMS_WH_AREA a " &
                            "where ISNULL(l.ILOC_BAL_QTY, 0) > 0 " &
                            "and l.IMP_CODE = a.IMP_CODE " &
                            "and l.ILOC_WH = a.WH_CODE " &
                            "and l.ILOC_FLOOR = a.FL_NUM " &
                            "and l.ILOC_AREA = a.AR_CODE " &
                            "and ISNULL(a.AR_DAMAGE_YN, '') <> 'Y' " &
                            "group by l.imp_code, l.storer_code, l.ITM_CODE, l.PACK_KEY, ISNULL(l.ILOC_PALLET_NO, '000')) bal2 " &
                            "on delv.imp_code = bal2.imp_code  " &
                            "and delv.storer_code = bal2.storer_code  " &
                            "and delv.dod_itm_code = bal2.ITM_CODE  " &
                            "and delv.dod_pack_key = bal2.PACK_KEY  " &
                            "and delv.dod_pallet_no = bal2.iloc_pallet_no  " &
                        "left outer join (" &
                            "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, sum(p.pld_item_qty) as picked_qty " &
                            "from wms_do_picklist_d p, wms_delv_order d " &
                            "where d.imp_code = p.imp_code " &
                            "and d.storer_code = p.storer_code " &
                            "and d.do_code = p.do_code " &
                            "and d.do_status = 'PICKED' " &
                            "and d.do_code <> '" & gU.dbEncode(Server.UrlDecode(Request("do_code"))) & "' " &
                            "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, '')) pck " &
                            "on delv.imp_code = pck.imp_code  " &
                            "and delv.storer_code = pck.storer_code  " &
                            "and delv.dod_itm_code = pck.pld_item_no  " &
                            "and delv.dod_pack_key = pck.pld_pack_key  " &
                            "and delv.dod_pallet_no = pck.pld_pallet_no  " &
                            "and delv.dod_batch_no = pck.pld_batch_no  " &
                        "left outer join (" &
                            "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, sum(p.pld_item_qty) as picked_qty " &
                            "from wms_do_picklist_d p, wms_delv_order d " &
                            "where d.imp_code = p.imp_code " &
                            "and d.storer_code = p.storer_code " &
                            "and d.do_code = p.do_code " &
                            "and d.do_status = 'PICKED' " &
                            "and d.do_code <> '" & gU.dbEncode(Server.UrlDecode(Request("do_code"))) & "' " &
                            "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000')) pck2 " &
                            "on delv.imp_code = pck2.imp_code  " &
                            "and delv.storer_code = pck2.storer_code  " &
                            "and delv.dod_itm_code = pck2.pld_item_no  " &
                            "and delv.dod_pack_key = pck2.pld_pack_key  " &
                            "and delv.dod_pallet_no = pck2.pld_pallet_no  " &
                        "where ISNULL(Case delv.dod_batch_no when null then bal2.bal_qty else bal.bal_qty end, 0) - " &
                            "ISNULL(CaSE delv.dod_batch_no when null then HOLD_STOCK2.hold_qty else HOLD_STOCK.hold_qty end, 0) - " &
                            "ISNULL(Case delv.dod_batch_no when null then pck2.picked_qty else pck.picked_qty end, 0) " &
                            "< delv.dod_qty "

        '"group by h.imp_code, h.storer_code, h.co_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO) hold " & _



        '"( " & _
        '                "select h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000') as COD_PALLET_NO, d.COD_BATCH_NO, sum(h.coh_in_stock_qty - ISNULL(h.coh_rel_qty, 0)) as hold_qty " & _
        '                "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
        '                "where h.imp_code = d.imp_code " & _
        '                "and h.storer_code = d.storer_code " & _
        '                "and h.co_code = d.co_code " & _
        '                "and h.cod_seq = d.cod_seq " & _
        '                "and c.imp_code = d.imp_code " & _
        '                "and c.storer_code = d.storer_code " & _
        '                "and c.co_code = d.co_code " & _
        '                "and h.coh_status <> 'RELEASE' " & _
        '                "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
        '                "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " & _
        '                "and c.co_code <> '" & gU.dbEncode(Server.UrlDecode(Request("co_code"))) & "' " & _
        '                "group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000'), d.COD_BATCH_NO) hold, " & _
        '                "( " & _
        '                "select h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000') as COD_PALLET_NO, sum(h.coh_in_stock_qty - ISNULL(h.coh_rel_qty, 0)) as hold_qty " & _
        '                "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c " & _
        '                "where h.imp_code = d.imp_code " & _
        '                "and h.storer_code = d.storer_code " & _
        '                "and h.co_code = d.co_code " & _
        '                "and h.cod_seq = d.cod_seq " & _
        '                "and c.imp_code = d.imp_code " & _
        '                "and c.storer_code = d.storer_code " & _
        '                "and c.co_code = d.co_code " & _
        '                "and h.coh_status <> 'RELEASE' " & _
        '                "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " & _
        '                "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " & _
        '                "and c.co_code <> '" & gU.dbEncode(Server.UrlDecode(Request("co_code"))) & "' " & _
        '                "group by h.imp_code, h.storer_code, d.COD_ITM_CODE, d.COD_PACK_KEY, ISNULL(d.COD_PALLET_NO, '000')) hold2, " & _

        'Response.Write("SQLString = " & "<br>" & SQLString)

        dt = gDB.getDataTable(SQLString)
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
            GridView1.EmptyDataText = "All selected items are available"
        End If
        'Session("dt") = dt
        GridView1.DataBind()

        If dt.Rows.Count > 0 Then
            'tr_hold_title1.Visible = True
            'tr_hold_title2.Visible = True
            'tr_hold_title3.Visible = True
            'tr_hold_gv.Visible = True

            Dim itmKey, itmKeyList As String
            Dim i As Long
            Dim holdDt As DataTable

            itmKeyList = ""

            For i = 0 To dt.Rows.Count - 1

                If dt.Rows(i).Item("batch_no").ToString.Trim <> "" Then
                    itmKey = dt.Rows(i).Item("itm_code").ToString & "#_#" &
                            dt.Rows(i).Item("pack_key").ToString & "#_#" &
                            gU.decodeNullOrEmpty(dt.Rows(i).Item("pallet_no").ToString, "000") & "#_#" &
                            dt.Rows(i).Item("batch_no").ToString.Trim

                    itmKeyList = gU.appendToList(itmKeyList, itmKey)
                End If
            Next

            itmKeyList = gU.dbConvList(itmKeyList)

            selectSql = "select h.imp_code, h.storer_code, c.co_code, h.coh_itm_code as COD_ITM_CODE, i.itm_name, h.coh_pack_key as COD_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000') as COD_PALLET_NO, h.COH_BATCH_NO as COD_BATCH_NO, sum(h.coh_in_stock_qty - ISNULL(h.coh_rel_qty, 0)) as hold_qty " &
                        "from WMS_CUST_ORDER_HOLD h, WMS_CUST_ORDER_D d, WMS_CUST_ORDER c, wms_item i " &
                        "where h.imp_code = d.imp_code " &
                        "and h.storer_code = d.storer_code " &
                        "and h.co_code = d.co_code " &
                        "and h.cod_seq = d.cod_seq " &
                        "and c.imp_code = d.imp_code " &
                        "and c.storer_code = d.storer_code " &
                        "and c.co_code = d.co_code " &
                        "and h.coh_status <> 'RELEASE' " &
                        "and c.CO_STATUS not in ('CLOSED', 'CANCELLED') " &
                        "and ISNULL(h.coh_in_stock_qty, 0) - ISNULL(h.coh_rel_qty, 0) > 0 " &
                        "and i.imp_code = h.imp_code " &
                        "and i.storer_code = h.storer_code " &
                        "and i.itm_code = h.coh_itm_code " &
                        "and i.pack_key = h.coh_pack_key " &
                        "and h.IMP_CODE = '" & gU.dbEncode(Server.UrlDecode(Request("imp_code"))) & "' " &
                        "and h.STORER_CODE = '" & gU.dbEncode(Server.UrlDecode(Request("storer_code"))) & "' " &
                        "and c.co_code <> '" & gU.dbEncode(Server.UrlDecode(Request("co_code"))) & "' " &
                        "and h.COH_ITM_CODE + '#_#' + h.COH_PACK_KEY + '#_#' + ISNULL(h.COH_PALLET_NO, '000') + '#_#' + isnull(h.COH_BATCH_NO, '') in (" & itmKeyList & ") " &
                        "group by h.imp_code, h.storer_code, c.co_code, h.COH_ITM_CODE, i.itm_name, h.COH_PACK_KEY, ISNULL(h.COH_PALLET_NO, '000'), h.COH_BATCH_NO "

            holdDt = gDB.getDataTable(selectSql)

            'GridView2.DataSource = holdDt

            'GridView2.DataBind()
        End If
    End Sub

    'Protected Sub GridView2_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
    '    Select Case e.Row.RowType
    '        Case DataControlRowType.Header
    '            Dim oGridView As GridView = DirectCast(sender, GridView)
    '            Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

    '            REM **********************
    '            REM Use for re-create the label to change the Langauge
    '            REM Modify Here
    '            Call cU.changeGVLabel(oGridViewRow, e, "Item No.", "物件编号")
    '            Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封装内码")
    '            Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物件名称")

    '            Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板編號")
    '            Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "批次編號")

    '            Call cU.changeGVLabel(oGridViewRow, e, "CO Order Code", "客戶訂單號")

    '            Call cU.changeGVLabel(oGridViewRow, e, "Hold Qty", "留貨數")
    '            REM **********************

    '            oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
    '    End Select
    'End Sub

    'Protected Sub GridView2_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
    '    Select Case e.Row.RowType
    '        Case DataControlRowType.DataRow
    '            REM **********************
    '            REM Modify Here

    '            CType(e.Row.FindControl("cod_itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "cod_itm_code").ToString.Trim
    '            CType(e.Row.FindControl("cod_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "cod_pack_key").ToString.Trim
    '            CType(e.Row.FindControl("itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_name").ToString.Trim

    '            CType(e.Row.FindControl("cod_pallet_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "cod_pallet_no").ToString.Trim
    '            CType(e.Row.FindControl("cod_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "cod_batch_no").ToString.Trim

    '            CType(e.Row.FindControl("co_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "co_code").ToString.Trim

    '            CType(e.Row.FindControl("hold_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "hold_qty").ToString.Trim

    '            REM **********************
    '    End Select
    'End Sub

End Class
