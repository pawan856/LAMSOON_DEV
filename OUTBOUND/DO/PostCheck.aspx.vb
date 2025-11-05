Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OUTBOUND_DO_PostCheck
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils

    Private dt As New DataTable

    Private tabIndex As Long

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim moduleAction As String
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)
        If ar.sessionExpired = "Y" Then
            Exit Sub
        End If

        moduleAction = Request("moduleAction")

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Post Check"

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "發布檢查"

        End If
        REM **********************

        dt = Session("_M_OB_DO_TMP_pl_dt")

        If Not IsPostBack Or moduleAction = "RELOADPL" Then

            IMP_CODE.Value = Server.UrlDecode(Request("IMP_CODE"))
            STORER_CODE.Value = Server.UrlDecode(Request("STORER_CODE"))
            CO_CODE.Value = Server.UrlDecode(Request("CO_CODE"))
            DO_CODE.Value = Server.UrlDecode(Request("DO_CODE"))

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                'generatePickList()
            End If

            BindGV()
        End If

        ar.hideForm(Me, editMode)

    End Sub

    Protected Sub Page_LoadComplete(sender As Object, e As System.EventArgs) Handles Me.LoadComplete
        Dim sm As ScriptManager = ScriptManager.GetCurrent(Page)

        sm.RegisterAsyncPostBackControl(btnPost)
        sm.RegisterAsyncPostBackControl(btnPost2)

        'For i = 0 To GridView1.Rows.Count - 1
        '    sm.RegisterAsyncPostBackControl(CType(GridView1.Rows(i).FindControl("Image_Loc_LookUp"), ImageButton))
        'Next

        moduleAction.Value = ""
    End Sub



    Protected Sub GridView1_DataBinding(sender As Object, e As System.EventArgs) Handles GridView1.DataBinding
        tabIndex = 1
    End Sub


    Protected Sub GridView1_RowCreated(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here
                Call cU.changeGVLabel(oGridViewRow, e, "Seq No.", "編號", HorizontalAlign.Center)
                'Call cU.changeGVLabel(oGridViewRow, e, "Picked By", "領料者")
                'Call cU.changeGVLabel(oGridViewRow, e, "Item Code", "物件號碼")
                Call cU.changeGVLabel(oGridViewRow, e, "Stock No.", "Stock No.")
                'Call cU.changeGVLabel(oGridViewRow, e, "DO Qty", "貨單數量", HorizontalAlign.Right)
                'Call cU.changeGVLabel(oGridViewRow, e, "Location FFI Qty", "預取貨量", HorizontalAlign.Right)
                'Call cU.changeGVLabel(oGridViewRow, e, "Picked Qty", "實取貨量", HorizontalAlign.Right)
                'Call cU.changeGVLabel(oGridViewRow, e, "Picked Qty Check", "實取驗證", HorizontalAlign.Right)

                'Call cU.changeGVLabel(oGridViewRow, e, "WH", "倉庫", HorizontalAlign.Center)

                'Call cU.changeGVLabel(oGridViewRow, e, "Floor", "樓層", HorizontalAlign.Center)
                'Call cU.changeGVLabel(oGridViewRow, e, "Area", "地區", HorizontalAlign.Center)
                'Call cU.changeGVLabel(oGridViewRow, e, "Rack", "架子", HorizontalAlign.Center)
                'Call cU.changeGVLabel(oGridViewRow, e, "Bin", "箱子", HorizontalAlign.Center)

                'Call cU.changeGVLabel(oGridViewRow, e, "Location Available Qty", "架上數量", HorizontalAlign.Right)
                'Call cU.changeGVLabel(oGridViewRow, e, "Location Bal Qty", "現存數量", HorizontalAlign.Right)
                'Call cU.changeGVLabel(oGridViewRow, e, "Hold Qty<br>/ Total Bal", "留貨數<br>/ 總存數", HorizontalAlign.Right)
                'Call cU.changeGVLabel(oGridViewRow, e, "Expiry Date", "到期日", HorizontalAlign.Center)

                'Call cU.changeGVLabel(oGridViewRow, e, "Qty2", "數量2", HorizontalAlign.Right)
                'Call cU.changeGVLabel(oGridViewRow, e, "Qty2 Check", "數量2驗證", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Serial No.", "序號", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Serial No. Check", "序號驗證", HorizontalAlign.Right)
                'Call cU.changeGVLabel(oGridViewRow, e, "CSMS Code", "CSMS Code", HorizontalAlign.Center)
                'Call cU.changeGVLabel(oGridViewRow, e, "Location", "位置", HorizontalAlign.Center)
                'Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "批次編號")
                'Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內碼", HorizontalAlign.Center)
                'Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板編號", HorizontalAlign.Center)
                'Call cU.changeGVLabel(oGridViewRow, e, "Is Loan", "借貨", HorizontalAlign.Center)
                'Call cU.changeGVLabel(oGridViewRow, e, "Remarks", "備注", HorizontalAlign.Left)
                'Call cU.changeGVLabel(oGridViewRow, e, "", "")
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)

                'Case DataControlRowType.DataRow

        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                REM **********************
                REM Modify Here
                'Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("pld_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_SEQ").ToString.Trim
                CType(e.Row.FindControl("dod_disp_seq"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_DISP_SEQ").ToString.Trim
                'CType(e.Row.FindControl("pld_picked_by"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_PICKED_BY").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ITM_SKU_NO").ToString.Trim
                'CType(e.Row.FindControl("pld_foi_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_FOI_QTY").ToString.Trim
                'CType(e.Row.FindControl("pld_item_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_QTY").ToString.Trim
                'CType(e.Row.FindControl("pld_do_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pld_do_qty").ToString.Trim

                CType(e.Row.FindControl("pld_item_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_NO").ToString.Trim
                CType(e.Row.FindControl("pld_pack_key"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("pld_pallet_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_PALLET_NO").ToString.Trim
                CType(e.Row.FindControl("pld_batch_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_BATCH_NO").ToString.Trim
                CType(e.Row.FindControl("pld_loc"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_LOC").ToString.Trim
                CType(e.Row.FindControl("pld_item_qty"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_QTY").ToString.Trim
                CType(e.Row.FindControl("itm_serial_no_yn"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ITM_SERIAL_NO_YN").ToString.Trim
                CType(e.Row.FindControl("itm_type"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ITM_TYPE").ToString.Trim

                'CType(e.Row.FindControl("stock_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "stock_qty").ToString.Trim

                'If CDbl(gU.decodeNullOrEmpty(DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_QTY").ToString.Trim, "0")) > CDbl(gU.decodeNullOrEmpty(DataBinder.Eval(e.Row.DataItem, "stock_qty").ToString.Trim, "0")) Then
                '    CType(e.Row.FindControl("pld_item_qty"), TextBox).BackColor = Drawing.Color.Red
                'End If

                'CType(e.Row.FindControl("iloc_bal_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_BAL_QTY").ToString.Trim

                'CType(e.Row.FindControl("pld_qty2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_QTY2").ToString.Trim
                CType(e.Row.FindControl("pld_serial_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_SERIAL_NO").ToString.Trim

                'If DataBinder.Eval(e.Row.DataItem, "ITM_SERIAL_NO_YN").ToString.Trim <> "Y" Then
                '    CType(e.Row.FindControl("pld_qty2_chk"), TextBox).Style.Add("display", "none")
                '    CType(e.Row.FindControl("pld_serial_no_chk"), TextBox).Style.Add("display", "none")

                '    CType(e.Row.FindControl("pld_item_qty_chk"), TextBox).TabIndex = tabIndex
                '    tabIndex = tabIndex + 1
                'Else
                '    If DataBinder.Eval(e.Row.DataItem, "ITM_TYPE").ToString.Trim = "CABLE" Then
                '        CType(e.Row.FindControl("pld_item_qty_chk"), TextBox).Style.Add("display", "none")

                '        CType(e.Row.FindControl("pld_qty2_chk"), TextBox).TabIndex = tabIndex
                '        tabIndex = tabIndex + 1
                '        CType(e.Row.FindControl("pld_serial_no_chk"), TextBox).TabIndex = tabIndex
                '        tabIndex = tabIndex + 1
                '    Else
                '        CType(e.Row.FindControl("pld_qty2_chk"), TextBox).Style.Add("display", "none")

                '        CType(e.Row.FindControl("pld_item_qty_chk"), TextBox).TabIndex = tabIndex
                '        tabIndex = tabIndex + 1
                '        CType(e.Row.FindControl("pld_serial_no_chk"), TextBox).TabIndex = tabIndex
                '        tabIndex = tabIndex + 1
                '    End If
                'End If

                If DataBinder.Eval(e.Row.DataItem, "ITM_SERIAL_NO_YN").ToString.Trim = "Y" AndAlso DataBinder.Eval(e.Row.DataItem, "ITM_TYPE").ToString.Trim <> "CABLE" Then
                    CType(e.Row.FindControl("pld_serial_no_chk"), TextBox).TabIndex = tabIndex
                    tabIndex = tabIndex + 1
                Else
                    CType(e.Row.FindControl("pld_serial_no_chk"), TextBox).Style.Add("display", "none")
                End If

                'CType(e.Row.FindControl("pld_serial_no"), TextBox).Attributes.Add("readonly", "")

                'If CDbl(DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim) > 0 Then
                '    CType(e.Row.FindControl("hold_qty_desc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim & " / " & DataBinder.Eval(e.Row.DataItem, "TOTAL_BAL").ToString.Trim
                '    CType(e.Row.FindControl("avail_qty"), HiddenField).Value = CDbl(DataBinder.Eval(e.Row.DataItem, "TOTAL_BAL").ToString.Trim) - CDbl(DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim)
                'Else
                '    CType(e.Row.FindControl("hold_qty_desc"), Label).Text = ""
                '    CType(e.Row.FindControl("avail_qty"), HiddenField).Value = 0
                'End If
                'CType(e.Row.FindControl("hold_qty"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim
                'CType(e.Row.FindControl("total_bal"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "TOTAL_BAL").ToString.Trim


                'CType(e.Row.FindControl("dsp_pld_wh"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_WH").ToString.Trim

                'CType(e.Row.FindControl("iloc_expiry_date"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_EXPIRY_DATE").ToString.Trim

                'CType(e.Row.FindControl("pld_wh"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_WH").ToString.Trim

                'CType(e.Row.FindControl("dsp_pld_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_LOC").ToString.Trim
                'CType(e.Row.FindControl("pld_loc"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_LOC").ToString.Trim
                'CType(e.Row.FindControl("pld_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_LOC").ToString.Trim

                'CType(e.Row.FindControl("dsp_pld_floor"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_FLOOR").ToString.Trim
                'CType(e.Row.FindControl("pld_floor"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_FLOOR").ToString.Trim
                'CType(e.Row.FindControl("dsp_pld_area"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_AREA").ToString.Trim
                'CType(e.Row.FindControl("pld_area"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_AREA").ToString.Trim
                'CType(e.Row.FindControl("dsp_pld_rack"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_RACK").ToString.Trim
                'CType(e.Row.FindControl("pld_rack"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_RACK").ToString.Trim
                'CType(e.Row.FindControl("dsp_pld_bin"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_BIN").ToString.Trim
                'CType(e.Row.FindControl("pld_bin"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_BIN").ToString.Trim

                'CType(e.Row.FindControl("bn_csms_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "BN_CSMS_CODE").ToString.Trim

                'CType(e.Row.FindControl("pld_org_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_ORG_LOC").ToString.Trim
                'CType(e.Row.FindControl("pld_remark"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_REMARK").ToString.Trim

                'CType(e.Row.FindControl("pld_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pld_batch_no").ToString.Trim

                'If DataBinder.Eval(e.Row.DataItem, "PLD_IS_LOAN").ToString.Trim = "Y" Then
                '    CType(e.Row.FindControl("pld_is_loan"), CheckBox).Checked = True
                'End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    e.Row.Visible = False
                End If
        End Select
    End Sub

    Private Function validateAll(Optional ByVal flag As String = "") As Boolean
        Dim i As Integer

        If dt.Rows.Count <> GridView1.Rows.Count Then
            If Session("gLang") = "E" Then
                uiFun.displayMsgNew(updtPnlAlert, "", "Pick List has been changed. Please save the WIT first before Post", Session("gLang"))
            Else
                uiFun.displayMsgNew(updtPnlAlert, "", "取貨單有改動, 請先儲存WIT!", Session("gLang"))
            End If

            Return False
        End If

        For i = 0 To dt.Rows.Count - 1
            If dt.Rows(i).Item("pld_seq").ToString.Trim <> CType(GridView1.Rows(i).FindControl("pld_seq"), HiddenField).Value OrElse _
                dt.Rows(i).Item("pld_item_no").ToString.Trim <> CType(GridView1.Rows(i).FindControl("pld_item_no"), HiddenField).Value OrElse _
                dt.Rows(i).Item("pld_pack_key").ToString.Trim <> CType(GridView1.Rows(i).FindControl("pld_pack_key"), HiddenField).Value OrElse _
                dt.Rows(i).Item("pld_pallet_no").ToString.Trim <> CType(GridView1.Rows(i).FindControl("pld_pallet_no"), HiddenField).Value OrElse _
                dt.Rows(i).Item("pld_batch_no").ToString.Trim <> CType(GridView1.Rows(i).FindControl("pld_batch_no"), HiddenField).Value OrElse _
                dt.Rows(i).Item("pld_loc").ToString.Trim <> CType(GridView1.Rows(i).FindControl("pld_loc"), HiddenField).Value OrElse _
                dt.Rows(i).Item("pld_item_qty").ToString.Trim <> CType(GridView1.Rows(i).FindControl("pld_item_qty"), HiddenField).Value Then

                If Session("gLang") = "E" Then
                    uiFun.displayMsgNew(updtPnlAlert, "", "Pick List has been changed. Please save the WIT first before Post", Session("gLang"))
                Else
                    uiFun.displayMsgNew(updtPnlAlert, "", "取貨單有改動, 請先儲存WIT!", Session("gLang"))
                End If

                Return False
            End If
        Next

        For i = 0 To GridView1.Rows.Count - 1
            If CType(GridView1.Rows(i).FindControl("itm_serial_no_yn"), HiddenField).Value = "Y" AndAlso CType(GridView1.Rows(i).FindControl("itm_type"), HiddenField).Value <> "CABLE" Then
                If CType(GridView1.Rows(i).FindControl("pld_serial_no"), Label).Text <> "" AndAlso CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Text.Trim = "" Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Serial No. cannot be empty!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "序號不能空白!", Session("gLang"))
                    End If
                    CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Focus()
                    Return False
                End If


                If UCase(CType(GridView1.Rows(i).FindControl("pld_serial_no"), Label).Text) <> UCase(CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Text.Trim) Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsgNew(updtPnlAlert, "", "Serial No. not matched!", Session("gLang"))
                    Else
                        uiFun.displayMsgNew(updtPnlAlert, "", "序號不對應!", Session("gLang"))
                    End If
                    CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Focus()

                    Return False
                End If
            End If
            'If CType(GridView1.Rows(i).FindControl("itm_serial_no_yn"), HiddenField).Value <> "Y" Then

            '    If CType(GridView1.Rows(i).FindControl("pld_item_qty"), Label).Text <> "" AndAlso CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Text.Trim = "" Then
            '        If Session("gLang") = "E" Then
            '            uiFun.displayMsgNew(updtPnlAlert, "", "Picked Qty cannot be empty!", Session("gLang"))
            '        Else
            '            uiFun.displayMsgNew(updtPnlAlert, "", "實取貨量不能空白!", Session("gLang"))
            '        End If
            '        CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Focus()
            '        Return False
            '    End If

            '    If CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Text.Trim <> "" AndAlso Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Text.Trim) Then
            '        If Session("gLang") = "E" Then
            '            uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Picked Qty!", Session("gLang"))
            '        Else
            '            uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 實取貨量!", Session("gLang"))
            '        End If
            '        CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Focus()
            '        Return False
            '    End If

            '    If gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("pld_item_qty"), Label).Text, 0) <> gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Text.Trim, 0) Then
            '        If Session("gLang") = "E" Then
            '            uiFun.displayMsgNew(updtPnlAlert, "", "Picked Qty not matched!", Session("gLang"))
            '        Else
            '            uiFun.displayMsgNew(updtPnlAlert, "", "實取貨量不對應!", Session("gLang"))
            '        End If
            '        CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Focus()
            '        Return False
            '    End If

            'Else

            '    If CType(GridView1.Rows(i).FindControl("itm_type"), HiddenField).Value = "CABLE" Then

            '        If CType(GridView1.Rows(i).FindControl("pld_qty2"), Label).Text <> "" AndAlso CType(GridView1.Rows(i).FindControl("pld_qty2_chk"), TextBox).Text.Trim = "" Then
            '            If Session("gLang") = "E" Then
            '                uiFun.displayMsgNew(updtPnlAlert, "", "Qty 2 cannot be empty!", Session("gLang"))
            '            Else
            '                uiFun.displayMsgNew(updtPnlAlert, "", "數量2不能空白!", Session("gLang"))
            '            End If
            '            CType(GridView1.Rows(i).FindControl("pld_qty2_chk"), TextBox).Focus()
            '            Return False
            '        End If

            '        If CType(GridView1.Rows(i).FindControl("pld_qty2_chk"), TextBox).Text.Trim <> "" AndAlso Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("pld_qty2_chk"), TextBox).Text.Trim) Then
            '            If Session("gLang") = "E" Then
            '                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Qty 2!", Session("gLang"))
            '            Else
            '                uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 數量2!", Session("gLang"))
            '            End If
            '            CType(GridView1.Rows(i).FindControl("pld_qty2_chk"), TextBox).Focus()
            '            Return False
            '        End If

            '        If gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("pld_qty2"), Label).Text, 0) <> gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("pld_qty2_chk"), TextBox).Text.Trim, 0) Then
            '            If Session("gLang") = "E" Then
            '                uiFun.displayMsgNew(updtPnlAlert, "", "Qty 2 not matched!", Session("gLang"))
            '            Else
            '                uiFun.displayMsgNew(updtPnlAlert, "", "數量2不對應!", Session("gLang"))
            '            End If
            '            CType(GridView1.Rows(i).FindControl("pld_qty2_chk"), TextBox).Focus()

            '            Return False
            '        End If

            '        If CType(GridView1.Rows(i).FindControl("pld_serial_no"), Label).Text <> "" AndAlso CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Text.Trim = "" Then
            '            If Session("gLang") = "E" Then
            '                uiFun.displayMsgNew(updtPnlAlert, "", "Serial No. cannot be empty!", Session("gLang"))
            '            Else
            '                uiFun.displayMsgNew(updtPnlAlert, "", "序號不能空白!", Session("gLang"))
            '            End If
            '            CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Focus()
            '            Return False
            '        End If


            '        If UCase(CType(GridView1.Rows(i).FindControl("pld_serial_no"), Label).Text) <> UCase(CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Text.Trim) Then
            '            If Session("gLang") = "E" Then
            '                uiFun.displayMsgNew(updtPnlAlert, "", "Serial No. not matched!", Session("gLang"))
            '            Else
            '                uiFun.displayMsgNew(updtPnlAlert, "", "序號不對應!", Session("gLang"))
            '            End If
            '            CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Focus()

            '            Return False
            '        End If

            '    Else

            '        If CType(GridView1.Rows(i).FindControl("pld_item_qty"), Label).Text <> "" AndAlso CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Text.Trim = "" Then
            '            If Session("gLang") = "E" Then
            '                uiFun.displayMsgNew(updtPnlAlert, "", "Picked Qty cannot be empty!", Session("gLang"))
            '            Else
            '                uiFun.displayMsgNew(updtPnlAlert, "", "實取貨量不能空白!", Session("gLang"))
            '            End If
            '            CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Focus()
            '            Return False
            '        End If

            '        If CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Text.Trim <> "" AndAlso Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Text.Trim) Then
            '            If Session("gLang") = "E" Then
            '                uiFun.displayMsgNew(updtPnlAlert, "", "Invalid number, Picked Qty!", Session("gLang"))
            '            Else
            '                uiFun.displayMsgNew(updtPnlAlert, "", "無效的數字, 實取貨量!", Session("gLang"))
            '            End If
            '            CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Focus()
            '            Return False
            '        End If

            '        If gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("pld_item_qty"), Label).Text, 0) <> gU.decodeEmptyCdbl(CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Text.Trim, 0) Then
            '            If Session("gLang") = "E" Then
            '                uiFun.displayMsgNew(updtPnlAlert, "", "Picked Qty not matched!", Session("gLang"))
            '            Else
            '                uiFun.displayMsgNew(updtPnlAlert, "", "實取貨量不對應!", Session("gLang"))
            '            End If
            '            CType(GridView1.Rows(i).FindControl("pld_item_qty_chk"), TextBox).Focus()
            '            Return False
            '        End If

            '        If CType(GridView1.Rows(i).FindControl("pld_serial_no"), Label).Text <> "" AndAlso CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Text.Trim = "" Then
            '            If Session("gLang") = "E" Then
            '                uiFun.displayMsgNew(updtPnlAlert, "", "Serial No. cannot be empty!", Session("gLang"))
            '            Else
            '                uiFun.displayMsgNew(updtPnlAlert, "", "序號不能空白!", Session("gLang"))
            '            End If
            '            CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Focus()
            '            Return False
            '        End If


            '        If UCase(CType(GridView1.Rows(i).FindControl("pld_serial_no"), Label).Text) <> UCase(CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Text.Trim) Then
            '            If Session("gLang") = "E" Then
            '                uiFun.displayMsgNew(updtPnlAlert, "", "Serial No. not matched!", Session("gLang"))
            '            Else
            '                uiFun.displayMsgNew(updtPnlAlert, "", "序號不對應!", Session("gLang"))
            '            End If
            '            CType(GridView1.Rows(i).FindControl("pld_serial_no_chk"), TextBox).Focus()

            '            Return False
            '        End If

            '    End If

            'End If
        Next

        Return True
    End Function

    Protected Sub BindGV()
        Dim sortDt As DataTable
        Dim sort_col As String = "DOD_DISP_SEQ, PLD_SPLIT_FR_INT, ITM_SKU_NO, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_BATCH_NO, PLD_SEQ_INT, PLD_SEQ, PLD_WH, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN, PLD_FOI_QTY, PLD_ITEM_QTY DESC"

        sortDt = dt.Copy

        Dim foundRows As DataRow() = sortDt.Select("", sort_col)

        dt.Rows.Clear()

        For i As Integer = 0 To foundRows.Count - 1
            dt.ImportRow(foundRows(i))
        Next

        dt.AcceptChanges()

        'If Request("moduleAction") = "RELOADPL" Then
        '    Dim objPlt As PickListTable

        '    objPlt = New PickListTable(CO_CODE.Value, DO_CODE.Value, STORER_CODE.Value, Session("IMP_CODE"))

        '    objPlt.updateAvailBal(dt)

        '    objPlt.updateTotalQty(dt)

        '    objPlt.updateHoldBal(dt)
        'End If

        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub

    Protected Sub btnPost_Click(sender As Object, e As System.EventArgs) Handles btnPost.Click
        If validateAll() Then
            ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "POST_GO", "postDO();", True)
        End If
    End Sub

    Protected Sub btnPost2_Click(sender As Object, e As System.EventArgs) Handles btnPost2.Click
        If validateAll() Then
            ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "POST_GO", "postDO();", True)
        End If
    End Sub
End Class
