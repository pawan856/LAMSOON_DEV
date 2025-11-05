Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient


Partial Class PickList
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private dt As New DataTable
    Private tabIndex As Long

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim moduleAction As String
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("OB_DO", Session("usr_id"), Me)
        If ar.sessionExpired = "Y" Then
            Exit Sub
        End If

        moduleAction = Request("moduleAction")

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
            saveBtn1.Text = "OK"
            saveBtn2.Text = "OK"
            newrow.Text = "Add"
            btnReset.OnClientClick = "return confirm(""Confirm to reset all the items from the DO?"");"
            'saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[No. will be auto generated]"
            'End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "執貨清單"
            saveBtn1.Text = "確定"
            saveBtn2.Text = "確定"
            newrow.Text = "新增"
            btnReset.OnClientClick = "return confirm(""確定重置所有的項目?"");"
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

        dt = Session("_M_OB_DO_TMP_pl_dt")

        If Not IsPostBack Or moduleAction = "RELOADPL" Then

            'If moduleAction = "RELOADPL" Then
            '    save_dt()
            'End If

            IMP_CODE.Value = Server.UrlDecode(Request("IMP_CODE"))
            STORER_CODE.Value = Server.UrlDecode(Request("STORER_CODE"))
            CO_CODE.Value = Server.UrlDecode(Request("CO_CODE"))
            DO_CODE.Value = Server.UrlDecode(Request("DO_CODE"))

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                generatePickList()
            End If

            Call BindGV()
        End If

        If Request("DO_STATUS") = "CANCELLED" Then
            ar.sec_write = "N"
        ElseIf Request("DO_STATUS") = "POSTED" Then
            ar.sec_write = "N"
        End If

        ar.hideForm(Me, editMode)

        If Request("moduleAction") = "SAVEOK" Then
            save()
        ElseIf Request("moduleAction") = "SAVEGVDT" Then
            save_dt()
        End If
    End Sub

    Protected Sub Page_LoadComplete(sender As Object, e As System.EventArgs) Handles Me.LoadComplete
        Dim sm As ScriptManager = ScriptManager.GetCurrent(Page)

        For i = 0 To GridView1.Rows.Count - 1
            If CType(GridView1.Rows(i).FindControl("Image_Loc_LookUp"), ImageButton) IsNot Nothing Then
                sm.RegisterAsyncPostBackControl(CType(GridView1.Rows(i).FindControl("Image_Loc_LookUp"), ImageButton))
            End If
        Next

        moduleAction.Value = ""
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                Dim oGridView As GridView = DirectCast(sender, GridView)
                Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                REM **********************
                REM Use for re-create the label to change the Langauge
                REM Modify Here
                Call cU.changeGVLabel(oGridViewRow, e, "Seq No.", "編號", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Picked By", "領料者")
                Call cU.changeGVLabel(oGridViewRow, e, "Item Code", "物件號碼")
                Call cU.changeGVLabel(oGridViewRow, e, "Stock No.", "Stock No.")
                Call cU.changeGVLabel(oGridViewRow, e, "DO Qty", "貨單數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Location FFI Qty", "預取貨量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Picked Qty", "實取貨量", HorizontalAlign.Right)

                Call cU.changeGVLabel(oGridViewRow, e, "CSMS Code", "CSMS Code", HorizontalAlign.Center)

                Call cU.changeGVLabel(oGridViewRow, e, "WH", "倉庫", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Org Loc", "原位", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Location", "位置", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Floor", "樓層", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Area", "地區", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Rack", "架子", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Bin", "箱子", HorizontalAlign.Center)

                Call cU.changeGVLabel(oGridViewRow, e, "Location Available Qty", "架上數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Location Bal Qty", "現存數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Hold Qty<br>/ Total Bal", "留貨數<br>/ 總存數", HorizontalAlign.Right)
                'Call cU.changeGVLabel(oGridViewRow, e, "Tracking No.", "追查編號")
                Call cU.changeGVLabel(oGridViewRow, e, "Expiry Date", "到期日", HorizontalAlign.Center)

                Call cU.changeGVLabel(oGridViewRow, e, "Qty2", "數量2", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Serial No.", "序號")
                Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "批次編號")
                Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內碼", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板編號", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Short Ship Qty", "Short Ship Qty", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Remarks", "備注", HorizontalAlign.Left)
                'Call cU.changeGVLabel(oGridViewRow, e, "Is Loan", "借貨", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)

                'Case DataControlRowType.DataRow

        End Select
    End Sub

    Protected Sub GridView1_DataBinding(sender As Object, e As System.EventArgs) Handles GridView1.DataBinding
        tabIndex = 1
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim nDropDown As DropDownList

        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                REM **********************
                REM Modify Here
                'Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("pld_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_SEQ").ToString.Trim
                CType(e.Row.FindControl("dod_disp_seq"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_DISP_SEQ").ToString.Trim
                CType(e.Row.FindControl("pld_picked_by"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_PICKED_BY").ToString.Trim
                CType(e.Row.FindControl("pld_item_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_NO").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ITM_SKU_NO").ToString.Trim
                CType(e.Row.FindControl("pld_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("pld_pallet_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_PALLET_NO").ToString.Trim
                CType(e.Row.FindControl("pld_foi_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_FOI_QTY").ToString.Trim
                CType(e.Row.FindControl("pld_item_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_QTY").ToString.Trim
                CType(e.Row.FindControl("pld_do_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pld_do_qty").ToString.Trim

                CType(e.Row.FindControl("pld_ss_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pld_ss_qty").ToString.Trim

                CType(e.Row.FindControl("stock_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "stock_qty").ToString.Trim

                If CDbl(gU.decodeNullOrEmpty(DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_QTY").ToString.Trim, "0")) > CDbl(gU.decodeNullOrEmpty(DataBinder.Eval(e.Row.DataItem, "stock_qty").ToString.Trim, "0")) Then
                    CType(e.Row.FindControl("pld_item_qty"), TextBox).BackColor = Drawing.Color.Red
                End If

                CType(e.Row.FindControl("pld_item_qty"), TextBox).TabIndex = tabIndex

                tabIndex = tabIndex + 1

                CType(e.Row.FindControl("iloc_bal_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_BAL_QTY").ToString.Trim

                CType(e.Row.FindControl("pld_qty2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_QTY2").ToString.Trim
                CType(e.Row.FindControl("pld_serial_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_SERIAL_NO").ToString.Trim

                CType(e.Row.FindControl("pld_serial_no"), TextBox).Attributes.Add("readonly", "")

                If CDbl(DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim) > 0 Then
                    CType(e.Row.FindControl("hold_qty_desc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim & " / " & DataBinder.Eval(e.Row.DataItem, "TOTAL_BAL").ToString.Trim
                    CType(e.Row.FindControl("avail_qty"), HiddenField).Value = CDbl(DataBinder.Eval(e.Row.DataItem, "TOTAL_BAL").ToString.Trim) - CDbl(DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim)
                Else
                    CType(e.Row.FindControl("hold_qty_desc"), Label).Text = ""
                    CType(e.Row.FindControl("avail_qty"), HiddenField).Value = 0
                End If
                CType(e.Row.FindControl("hold_qty"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim
                CType(e.Row.FindControl("total_bal"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "TOTAL_BAL").ToString.Trim


                'Dim nDropDown As DropDownList = CType(e.Row.FindControl("pld_wh"), DropDownList)
                'uiFun.load_dropdown(nDropDown, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))
                'nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "PLD_WH").ToString.Trim
                'nDropDown.Attributes.Add("onchange", gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_pld_loc"), Label).ClientID) & ".innerHTML='';" & _
                '                            gU.jsHTMLEncode(CType(e.Row.FindControl("pld_loc"), HiddenField).ClientID) & ".value='';" & _
                '                            gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_pld_floor"), Label).ClientID) & ".innerHTML='';" & _
                '                            gU.jsHTMLEncode(CType(e.Row.FindControl("pld_floor"), HiddenField).ClientID) & ".value='';" & _
                '                            gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_pld_area"), Label).ClientID) & ".innerHTML='';" & _
                '                            gU.jsHTMLEncode(CType(e.Row.FindControl("pld_area"), HiddenField).ClientID) & ".value='';" & _
                '                            gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_pld_rack"), Label).ClientID) & ".innerHTML='';" & _
                '                            gU.jsHTMLEncode(CType(e.Row.FindControl("pld_rack"), HiddenField).ClientID) & ".value='';" & _
                '                            gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_pld_bin"), Label).ClientID) & ".innerHTML='';" & _
                '                            gU.jsHTMLEncode(CType(e.Row.FindControl("pld_bin"), HiddenField).ClientID) & ".value='';")

                'CType(e.Row.FindControl("dod_track_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DOD_TRACK_NO").ToString.Trim
                nDropDown = CType(e.Row.FindControl("pld_wh"), DropDownList)
                uiFun.load_dropdown(nDropDown, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "PLD_WH").ToString.Trim

                CType(e.Row.FindControl("iloc_expiry_date"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_EXPIRY_DATE").ToString.Trim

                uiFun.load_ComboBox(CType(e.Row.FindControl("pld_loc"), AjaxControlToolkit.ComboBox),
                                    "Select distinct Convert(nvarchar(20),a.FL_NUM)+b.AR_CODE+c.RK_CODE+d.BN_CODE as CODE,a.FL_NAME+b.AR_CODE+c.RK_CODE+d.BN_CODE as NAME from WMS_WH_FL a Inner join WMS_WH_AREA b on a.FL_NUM=b.FL_NUM Inner join WMS_WH_RACK c on b.AR_CODE=c.AR_CODE Inner join WMS_WH_BIN d on c.RK_CODE=d.RK_CODE where a.WH_CODE='" & DataBinder.Eval(e.Row.DataItem, "PLD_WH").ToString.Trim & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))

                CType(e.Row.FindControl("pld_loc"), AjaxControlToolkit.ComboBox).SelectedValue = DataBinder.Eval(e.Row.DataItem, "pld_loc").ToString.Trim


                'nDropDown = CType(e.Row.FindControl("pld_loc"), DropDownList)
                'uiFun.load_dropdown(nDropDown, "Select distinct Convert(nvarchar(20),a.FL_NUM)+b.AR_CODE+c.RK_CODE+d.BN_CODE as CODE,a.WH_CODE+a.FL_NAME+b.AR_CODE+c.RK_CODE+d.BN_CODE as NAME from WMS_WH_FL a Inner join WMS_WH_AREA b on a.FL_NUM=b.FL_NUM Inner join WMS_WH_RACK c on b.AR_CODE=c.AR_CODE Inner join WMS_WH_BIN d on c.RK_CODE=d.RK_CODE where a.WH_CODE='" & DataBinder.Eval(e.Row.DataItem, "PLD_WH").ToString.Trim & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))
                'nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "pld_loc").ToString.Trim

                'CType(e.Row.FindControl("dsp_pld_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_LOC").ToString.Trim
                'CType(e.Row.FindControl("pld_loc"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_LOC").ToString.Trim

                CType(e.Row.FindControl("dsp_pld_floor"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_FLOOR").ToString.Trim
                CType(e.Row.FindControl("pld_floor"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_FLOOR").ToString.Trim
                CType(e.Row.FindControl("dsp_pld_area"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_AREA").ToString.Trim
                CType(e.Row.FindControl("pld_area"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_AREA").ToString.Trim
                CType(e.Row.FindControl("dsp_pld_rack"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_RACK").ToString.Trim
                CType(e.Row.FindControl("pld_rack"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_RACK").ToString.Trim
                CType(e.Row.FindControl("dsp_pld_bin"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_BIN").ToString.Trim
                CType(e.Row.FindControl("pld_bin"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_BIN").ToString.Trim

                CType(e.Row.FindControl("bn_csms_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "BN_CSMS_CODE").ToString.Trim

                CType(e.Row.FindControl("pld_org_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_ORG_LOC").ToString.Trim
                CType(e.Row.FindControl("pld_remark"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_REMARK").ToString.Trim

                'uiFun.load_dropdown(CType(e.Row.FindControl("pld_batch_no"), DropDownList), _
                '                    "select distinct dc_date_code " & _
                '                    "from wms_item d, wms_item_loc_bal l, wms_date_code dc " & _
                '                    "where " & _
                '                    "d.imp_code = l.imp_code " & _
                '                    "and d.storer_code = l.storer_code " & _
                '                    "and d.itm_code = l.itm_code " & _
                '                    "and d.pack_key = l.pack_key " & _
                '                    "and l.iloc_batch_no = dc.dc_date_code " & _
                '                    "and l.imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                '                    "and l.storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
                '                    "and l.itm_code = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_NO").ToString.Trim) & "' " & _
                '                    "and l.pack_key = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "PLD_PACK_KEY").ToString.Trim) & "' " & _
                '                    "order by dc_date_code desc ", "dc_date_code", "dc_date_code", , Session("gSelectLabel"))


                'CType(e.Row.FindControl("pld_batch_no"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "pld_batch_no").ToString.Trim
                'CType(e.Row.FindControl("pld_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pld_batch_no").ToString.Trim

                CType(e.Row.FindControl("pld_batch_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "pld_batch_no").ToString.Trim

                'If DataBinder.Eval(e.Row.DataItem, "PLD_IS_LOAN").ToString.Trim = "Y" Then
                '    CType(e.Row.FindControl("pld_is_loan"), CheckBox).Checked = True
                'End If

                Dim nImage As ImageButton = CType(e.Row.FindControl("Image_Loc_LookUp"), ImageButton)
                If nImage IsNot Nothing Then
                    nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                End If


                If DataBinder.Eval(e.Row.DataItem, "ITM_SERIAL_NO_YN").ToString.Trim = "Y" Then
                    'nImage.Attributes.Add("onclick", "saveGVDT();LocSerialLookUp('" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_NO").ToString.Trim) & "', " & _
                    '                  "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "PLD_PACK_KEY").ToString.Trim) & "', " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_pallet_no"), TextBox).ClientID) & "').value, " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_batch_no"), Label).ClientID) & "').innerHTML, " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("itm_sku_no"), Label).ClientID) & "').innerHTML, " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_qty2"), TextBox).ClientID) & "').value, " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_seq"), HiddenField).ClientID) & "').value, " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_org_loc"), Label).ClientID) & "').innerHTML, " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dod_disp_seq"), Label).ClientID) & "').innerHTML)")

                    nImage.OnClientClick = "saveGVDT();LocSerialLookUp('" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_NO").ToString.Trim) & "', " &
                                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "PLD_PACK_KEY").ToString.Trim) & "', " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_pallet_no"), TextBox).ClientID) & "').value, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_batch_no"), TextBox).ClientID) & "').innerHTML, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("itm_sku_no"), Label).ClientID) & "').innerHTML, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_qty2"), TextBox).ClientID) & "').value, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_seq"), HiddenField).ClientID) & "').value, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_org_loc"), Label).ClientID) & "').innerHTML, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dod_disp_seq"), Label).ClientID) & "').innerHTML); return false;"
                Else
                    'nImage.Attributes.Add("onclick", "saveGVDT();LocLookUp('" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_NO").ToString.Trim) & "', " & _
                    '                  "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "PLD_PACK_KEY").ToString.Trim) & "', " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_pallet_no"), TextBox).ClientID) & "').value, " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_batch_no"), Label).ClientID) & "').innerHTML, " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("itm_sku_no"), Label).ClientID) & "').innerHTML, " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_org_loc"), Label).ClientID) & "').innerHTML, " & _
                    '                  "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dod_disp_seq"), Label).ClientID) & "').innerHTML)")
                    If nImage IsNot Nothing Then
                        nImage.OnClientClick = "saveGVDT();LocLookUp('" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_NO").ToString.Trim) & "', " &
                                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "PLD_PACK_KEY").ToString.Trim) & "', " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_pallet_no"), TextBox).ClientID) & "').value, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_batch_no"), TextBox).ClientID) & "').innerHTML, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("itm_sku_no"), Label).ClientID) & "').innerHTML, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_org_loc"), Label).ClientID) & "').innerHTML, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dod_disp_seq"), Label).ClientID) & "').innerHTML); return false;"
                    End If
                    CType(e.Row.FindControl("btnSplit"), Button).Visible = False
                    End If

                    'Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp"), Image)
                    'nImage.Attributes.Add("onmousedown", "MM_swapImage('" & gU.jsHTMLEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                    'nImage.Attributes.Add("onclick", "LocLookUp(document.piform." & gU.jsHTMLEncode(nDropDown.ClientID) & ".value, '" & _
                    '                      gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_pld_loc"), Label).ClientID) & "', " & _
                    '                      "'" & gU.jsHTMLEncode(CType(e.Row.FindControl("pld_loc"), HiddenField).ClientID) & "', " & _
                    '                      "'" & gU.jsHTMLEncode(nDropDown.ClientID) & "', " & _
                    '                      "'" & gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_pld_floor"), Label).ClientID) & "', " & _
                    '                      "'" & gU.jsHTMLEncode(CType(e.Row.FindControl("pld_floor"), HiddenField).ClientID) & "', " & _
                    '                      "'" & gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_pld_area"), Label).ClientID) & "', " & _
                    '                      "'" & gU.jsHTMLEncode(CType(e.Row.FindControl("pld_area"), HiddenField).ClientID) & "', " & _
                    '                      "'" & gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_pld_rack"), Label).ClientID) & "', " & _
                    '                      "'" & gU.jsHTMLEncode(CType(e.Row.FindControl("pld_rack"), HiddenField).ClientID) & "', " & _
                    '                      "'" & gU.jsHTMLEncode(CType(e.Row.FindControl("dsp_pld_bin"), Label).ClientID) & "', " & _
                    '                      "'" & gU.jsHTMLEncode(CType(e.Row.FindControl("pld_bin"), HiddenField).ClientID) & "')")
                    REM **********************

                    'Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                    'If Session("gLang") = "E" Then
                    '    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    '    nButton.Text = "Delete"
                    'ElseIf Session("gLang") = "C" Then
                    '    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    '    nButton.Text = "删除"
                    'End If

                    'If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    '    Call ar.hideGVRow(GridView1, e.Row)
                    'End If

                    'If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                    '    nButton.Enabled = False
                    'End If

                    Dim cableFlag As String = gU.decodeNullOrEmpty(DB.getValueFromSQL("SELECT ITM_TYPE from wms_item where imp_code='" & gU.dbEncode(IMP_CODE.Value.Trim) & "' and storer_code='" & gU.dbEncode(STORER_CODE.Value) & "' and itm_code='" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "pld_item_no").ToString.Trim) & "' and pack_key='" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "pld_pack_key").ToString.Trim) & "'"), "")

                If cableFlag = "CABLE" And DataBinder.Eval(e.Row.DataItem, "pld_serial_no").ToString.Trim <> "" Then
                    Dim drumList As String = hd_drumList.Value
                    Dim drumID As String = gU.decodeNullOrEmpty(DB.getValueFromSQL("Select max(ilbs_drum_id) from wms_item_loc_bal_s " &
                                                                                   "where ilbs_serial_no='" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "pld_serial_no").ToString.Trim) & "'"
                                                                                    ), "")
                    If drumID <> "" Then
                        CType(e.Row.FindControl("from_drum_id"), HiddenField).Value = drumID

                        If Not gU.inList(drumList, drumID) Then
                            CType(e.Row.FindControl("btnREDRUM"), Button).Visible = True
                            drumList = gU.appendToList(drumList, drumID)
                            hd_drumList.Value = drumList
                        End If
                    End If
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    e.Row.Visible = False
                End If
        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        If validateAll() Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True, New String() {"Y", "N"}) Then
                Dim rows_count As Integer = 0
                REM **********************
                REM Modify Here

                Session("_M_OB_DO_TMP_pl_seq") = CStr(CInt(Session("_M_OB_DO_TMP_pl_seq")) + 1)

                dt.Rows.Add()
                rows_count = dt.Rows.Count

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("pld_seq") = Session("_M_OB_DO_TMP_pl_seq").ToString
                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"
                dt.AcceptChanges()

                'Session("pi_dt") = dt
                GridView1.DataSource = dt
                GridView1.DataBind()
            End If
        End If
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Select e.CommandName
            Case "SplitItem"
                If validateAll("SPLIT") Then
                    If cU.gfBuildDataTableforGridView(dt, GridView1, True, New String() {"Y", "N"}) Then
                        'Dim rows_count As Integer = 0
                        Dim rowNum As Integer = -1
                        Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)
                        Dim plSeq As Integer
                        Dim newRow As DataRow

                        If Not IsNothing(gvRow) Then
                            rowNum = gvRow.RowIndex
                        End If

                        plSeq = CInt(Session("_M_OB_DO_TMP_pl_seq"))

                        plSeq = plSeq + 1

                        Session("_M_OB_DO_TMP_pl_seq") = CStr(CInt(Session("_M_OB_DO_TMP_pl_seq")) + 1)

                        'dt.Rows.Add()
                        'rows_count = dt.Rows.Count

                        'dt.Rows(rows_count - 1).Item("pld_seq") = Session("_M_OB_DO_TMP_pl_seq").ToString
                        'dt.Rows(rows_count - 1).Item("pld_item_no") = dt.Rows(rowNum).Item("pld_item_no")
                        'dt.Rows(rows_count - 1).Item("itm_sku_no") = dt.Rows(rowNum).Item("itm_sku_no")
                        'dt.Rows(rows_count - 1).Item("pld_pack_key") = dt.Rows(rowNum).Item("pld_pack_key")
                        'dt.Rows(rows_count - 1).Item("pld_pallet_no") = dt.Rows(rowNum).Item("pld_pallet_no")
                        'dt.Rows(rows_count - 1).Item("pld_foi_qty") = 0
                        'dt.Rows(rows_count - 1).Item("pld_item_qty") = 0
                        'dt.Rows(rows_count - 1).Item("pld_do_qty") = dt.Rows(rowNum).Item("pld_do_qty")
                        'dt.Rows(rows_count - 1).Item("stock_qty") = 0
                        'dt.Rows(rows_count - 1).Item("iloc_bal_qty") = 0

                        'dt.Rows(rows_count - 1).Item("pld_qty2") = 0
                        'dt.Rows(rows_count - 1).Item("pld_serial_no") = ""

                        'dt.Rows(rows_count - 1).Item("PLD_BATCH_NO") = dt.Rows(rowNum).Item("pld_batch_no")

                        '----------------------------------------------------------------------------------------------------------------------------------------------------------
                        newRow = dt.NewRow

                        newRow.Item("PLD_SEQ") = CStr(plSeq)
                        newRow.Item("PLD_SEQ_INT") = plSeq

                        newRow.Item("PLD_PICKED_BY") = Session("usr_id")
                        newRow.Item("PLD_ITEM_NO") = dt.Rows(rowNum).Item("pld_item_no")
                        newRow.Item("ITM_SKU_NO") = dt.Rows(rowNum).Item("itm_sku_no")
                        newRow.Item("PLD_PACK_KEY") = dt.Rows(rowNum).Item("pld_pack_key")
                        newRow.Item("PLD_PALLET_NO") = dt.Rows(rowNum).Item("pld_pallet_no")

                        newRow.Item("PLD_SS_QTY") = dt.Rows(rowNum).Item("pld_ss_qty")

                        newRow.Item("PLD_FOI_QTY") = 0
                        newRow.Item("PLD_ITEM_QTY") = 0

                        newRow.Item("PLD_QTY2") = 0
                        newRow.Item("PLD_SERIAL_NO") = ""

                        newRow.Item("PLD_DO_QTY") = dt.Rows(rowNum).Item("pld_do_qty")

                        newRow.Item("STOCK_QTY") = 0

                        newRow.Item("ILOC_BAL_QTY") = 0

                        newRow.Item("HOLD_QTY") = dt.Rows(rowNum).Item("HOLD_QTY")
                        newRow.Item("TOTAL_BAL") = dt.Rows(rowNum).Item("TOTAL_BAL")

                        'Will be updated by the updateTotalQty function
                        newRow.Item("TOTAL_ITEM_QTY") = dt.Rows(rowNum).Item("TOTAL_ITEM_QTY")
                        newRow.Item("TOTAL_FOI_QTY") = dt.Rows(rowNum).Item("TOTAL_FOI_QTY")

                        newRow.Item("ILOC_EXPIRY_DATE") = dt.Rows(rowNum).Item("ILOC_EXPIRY_DATE")

                        newRow.Item("PLD_EXPIRY_DATE") = dt.Rows(rowNum).Item("PLD_EXPIRY_DATE")
                        newRow.Item("PLD_MANU_DATE") = dt.Rows(rowNum).Item("PLD_MANU_DATE")

                        newRow.Item("PLD_WH") = dt.Rows(rowNum).Item("PLD_WH")

                        newRow.Item("PLD_LOC") = ""
                        newRow.Item("PLD_ORG_LOC") = ""
                        newRow.Item("PLD_REMARK") = ""
                        newRow.Item("PLD_FLOOR") = ""
                        newRow.Item("PLD_AREA") = ""
                        newRow.Item("PLD_RACK") = ""
                        newRow.Item("PLD_BIN") = ""
                        newRow.Item("PLD_IS_LOAN") = "N"

                        newRow.Item("PLD_BATCH_NO") = dt.Rows(rowNum).Item("PLD_BATCH_NO")

                        newRow.Item("DOD_DISP_SEQ") = dt.Rows(rowNum).Item("DOD_DISP_SEQ")
                        newRow.Item("BN_CSMS_CODE") = dt.Rows(rowNum).Item("BN_CSMS_CODE")

                        newRow.Item("ITM_SERIAL_NO_YN") = dt.Rows(rowNum).Item("ITM_SERIAL_NO_YN")

                        newRow.Item("PLD_SPLIT_FR_INT") = dt.Rows(rowNum).Item("PLD_SEQ")


                        newRow.Item("mFlag") = "N"


                        dt.Rows.InsertAt(newRow, rowNum + 1)


                        dt.AcceptChanges()

                        GridView1.DataSource = dt
                        GridView1.DataBind()
                    End If
                End If

            Case "REDRUM"
                Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

                Dim drumID As String = DirectCast(gvRow.FindControl("from_drum_id"), HiddenField).Value
                'Dim toDrumID As String = DirectCast(gvRow.FindControl("ISD_TO_DRUM_ID"), HiddenField).Value
                'Dim toDrumList As String = DirectCast(gvRow.FindControl("ISD_TO_DRUM_CABLE_LIST"), HiddenField).Value

                Dim toDrumID As String = dt.Rows(gvRow.RowIndex).Item("PLD_TO_DRUM_ID").ToString.Trim
                Dim toDrumList As String = dt.Rows(gvRow.RowIndex).Item("PLD_TO_DRUM_CABLE_LIST").ToString.Trim
                Dim toDrumIloc_seq As String = dt.Rows(gvRow.RowIndex).Item("PLD_TO_DRUM_ILOC_SEQ").ToString.Trim
                If drumID <> "" Then
                    ShowRedrum(drumID, toDrumID, toDrumList, toDrumIloc_seq)
                    FromRowIDX.Value = gvRow.RowIndex
                Else
                    uiFun.displayMsg(Me, "", "Invalid Drum ID!", Session("gLang"))
                End If
        End Select
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        If dt.Rows(e.RowIndex).Item("mFlag") = "N" Then
            'If dt.Rows(e.RowIndex).Item("pld_seq") = Session("_M_OB_DO_TMP_pl_seq").ToString Then
            '    Session("_M_OB_DO_TMP_pl_seq") = CStr(CInt(Session("_M_OB_DO_TMP_pl_seq")) - 1)
            'End If
            If cU.gfBuildDataTableforGridView(dt, GridView1, True, New String() {"Y", "N"}) Then
                dt.Rows(e.RowIndex).Delete()

                dt.AcceptChanges()

                GridView1.DataSource = dt
                GridView1.DataBind()
            End If
        Else
            dt.Rows(e.RowIndex).Item("mFlag") = "D"
            GridView1.Rows(e.RowIndex).Visible = False
            dt.AcceptChanges()
        End If

    End Sub

    Private Function validateAll(Optional ByVal flag As String = "") As Boolean
        Dim i As Integer
        Dim javaStr As String

        'Sum total foi qty and picked qty for check stock bal validation
        Dim objPlt As PickListTable
        objPlt = New PickListTable(CO_CODE.Value, DO_CODE.Value, STORER_CODE.Value, Session("IMP_CODE"))
        objPlt.updateTotalQtyForGV(GridView1)
        objPlt = Nothing


        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If GridView1.Rows(i).Visible Then
                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("pld_foi_qty"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid whole number, FFI Qty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的整數, 取貨數量!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("pld_item_qty"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid whole number, Picked Qty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的整數, 實取數量!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If flag = "" Then
                        If CType(GridView1.Rows(i).FindControl("pld_foi_qty"), TextBox).Text.Trim <> "" AndAlso CDbl(CType(GridView1.Rows(i).FindControl("pld_foi_qty"), TextBox).Text) <= 0 Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "FFI Qty must be greater than zero!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "取貨數量必需大過零!", Session("gLang"))
                            End If
                            Return False
                        End If
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("pld_qty2"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid whole number, Qty2!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的整數, 數量2!", Session("gLang"))
                        End If
                        Return False
                    End If

                    'If CType(GridView1.Rows(i).FindControl("pld_org_loc"), Label).Text <> "" AndAlso CType(GridView1.Rows(i).FindControl("pld_org_loc"), Label).Text <> CType(GridView1.Rows(i).FindControl("pld_loc"), HiddenField).Value AndAlso _
                    '    CType(GridView1.Rows(i).FindControl("pld_remark"), TextBox).Text.Trim = "" Then

                    '    If Session("gLang") = "E" Then
                    '        uiFun.displayMsg(Me, "", "Selected location is different from Original Location. Please input remarks!", Session("gLang"))
                    '    Else
                    '        uiFun.displayMsg(Me, "", "位置與原位不同, 請輪入備注!", Session("gLang"))
                    '    End If
                    '    CType(GridView1.Rows(i).FindControl("pld_remark"), TextBox).Focus()
                    '    Return False
                    'End If

                    'If CType(GridView1.Rows(i).FindControl("pld_item_qty"), TextBox).Text.Trim <> "" AndAlso CDbl(CType(GridView1.Rows(i).FindControl("pld_item_qty"), TextBox).Text) <= 0 Then
                    '    If Session("gLang") = "E" Then
                    '        uiFun.displayMsg(Me, "", "Picked Qty must be greater than zero!", Session("gLang"))
                    '    Else
                    '        uiFun.displayMsg(Me, "", "實取數量必需大過零!", Session("gLang"))
                    '    End If
                    '    Return False
                    'End If

                    If CType(GridView1.Rows(i).FindControl("pld_item_qty"), TextBox).Text.Trim <> "" Then
                        If CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("pld_item_qty"), TextBox).Text.Trim, "0")) > CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("pld_foi_qty"), TextBox).Text.Trim, "0")) Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Picked qty of " & CType(GridView1.Rows(i).FindControl("pld_item_no"), Label).Text & " is greater than FFI qty!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "實取貨量 " & CType(GridView1.Rows(i).FindControl("pld_item_no"), Label).Text & " 大於預取貨量!", Session("gLang"))
                            End If
                            Return False
                        End If
                    End If

                    If Request("moduleAction") <> "SAVEOK" Then
                        If CType(GridView1.Rows(i).FindControl("pld_item_qty"), TextBox).Text.Trim <> "" Then
                            If CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("pld_item_qty"), TextBox).Text, "0")) > CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("stock_qty"), Label).Text, "0")) Then
                                If Session("gLang") = "E" Then
                                    javaStr = "if(confirm(""Confirm to proceed?"")) saveok();" 'Picked qty is greater than stock qty, 
                                Else
                                    javaStr = "if(confirm(""確定輸入?"")) saveok();" '實取貨量高於貨存, 
                                End If
                                If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                                    Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
                                End If
                                Return False
                            ElseIf CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("hold_qty"), HiddenField).Value, "0")) > 0 AndAlso _
                                CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("pld_item_qty"), TextBox).Text, "0")) > CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("avail_qty"), HiddenField).Value, "0")) Then

                                If Session("gLang") = "E" Then
                                    javaStr = "if(confirm(""Total picked qty of item " & CType(GridView1.Rows(i).FindControl("pld_item_no"), Label).Text & " is greater than available qty (balance qty - hold qty), confirm to proceed?"")) saveok();"
                                Else
                                    javaStr = "if(confirm(""總實取貨量高於可用貨存(貨存-保留), 確定輸入?"")) saveok();"
                                End If
                                If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                                    Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
                                End If
                                Return False
                            ElseIf CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("pld_foi_qty"), TextBox).Text, "0")) > CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("stock_qty"), Label).Text, "0")) Then
                                If Session("gLang") = "E" Then
                                    javaStr = "if(confirm(""FFI qty is greater than stock qty, confirm to proceed?"")) saveok();"
                                Else
                                    javaStr = "if(confirm(""預取貨量高於貨存, 確定輸入?"")) saveok();"
                                End If
                                If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                                    Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
                                End If
                                Return False
                            ElseIf CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("hold_qty"), HiddenField).Value, "0")) > 0 AndAlso _
                                CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("total_foi_qty"), HiddenField).Value, "0")) > CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("avail_qty"), HiddenField).Value, "0")) Then

                                If Session("gLang") = "E" Then
                                    javaStr = "if(confirm(""Total FFI qty of item " & CType(GridView1.Rows(i).FindControl("pld_item_no"), Label).Text & " is greater than available qty (stock qty - hold qty), confirm to proceed?"")) saveok();"
                                Else
                                    javaStr = "if(confirm(""總預取貨量高於可用貨存(貨存-保留), 確定輸入?"")) saveok();"
                                End If
                                If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                                    Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
                                End If
                                Return False
                            End If
                        End If
                    End If
                End If
            Next
        End If

        Return True
    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        If validateAll() Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True, New String() {"Y", "N"}) Then
                Session("_M_OB_DO_TMP_pl_dt") = dt

                Dim pi_dt As DataTable = Session("_M_OB_DO_TMP_pi_dt")

                For k As Integer = 0 To dt.Rows.Count - 1
                    For t As Integer = 0 To pi_dt.Rows.Count - 1
                        If dt.Rows(k).Item("PLD_ITEM_NO").ToString = pi_dt.Rows(t).Item("pad_itm_code").ToString AndAlso _
                                    dt.Rows(k).Item("PLD_PACK_KEY").ToString = pi_dt.Rows(t).Item("pad_pack_key").ToString Then

                            'pi_dt.Rows(t).Item("pad_qty") = dt.Rows(k).Item("PLD_ITEM_QTY").ToString.Trim
                            pi_dt.Rows(t).Item("pad_qty") = dt.Rows(k).Item("PLD_FOI_QTY").ToString.Trim
                            pi_dt.Rows(t).AcceptChanges()
                            Exit For
                        End If
                    Next
                Next
                Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "close", "<script language=""JavaScript"">window.open('','_self');window.close();</script>")
            End If
        End If
    End Sub

    Private Sub save_dt()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True, New String() {"Y", "N"}) Then
            Session("_M_OB_DO_TMP_pl_dt") = dt
        End If
    End Sub

    'This function should be same as DOMain
    Private Sub delZeroRow()
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

    Protected Sub BindGV()
        Dim sortDt As DataTable
        Dim sort_col As String = "DOD_DISP_SEQ, PLD_SPLIT_FR_INT, ITM_SKU_NO, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, PLD_BATCH_NO, PLD_SEQ_INT, PLD_SEQ, PLD_WH, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN, PLD_FOI_QTY, PLD_ITEM_QTY DESC"

        'Dim query = From q In dt.AsEnumerable() _
        '                     Order By _
        '                     q("DOD_DISP_SEQ"), _
        '                     q("ITM_SKU_NO").ToString, _
        '                     q("PLD_ITEM_NO").ToString, _
        '                     q("PLD_PACK_KEY"), _
        '                     q("PLD_PALLET_NO").ToString, _
        '                     q("PLD_BATCH_NO").ToString



        'IMP_CODE.Value = Server.UrlDecode(Request("IMP_CODE"))
        'STORER_CODE.Value = Server.UrlDecode(Request("STORER_CODE"))
        'CO_CODE.Value = Server.UrlDecode(Request("CO_CODE"))
        'DO_CODE.Value = Server.UrlDecode(Request("DO_CODE"))

        'delZeroRow()

        sortDt = dt.Copy

        'commented for implementation of picking edit options

        'Dim foundRows As DataRow() = sortDt.Select("", sort_col)

        'dt.Rows.Clear()

        'For i As Integer = 0 To foundRows.Count - 1
        '    dt.ImportRow(foundRows(i))
        'Next

        dt.AcceptChanges()

        'Dim query = From q In dt.AsEnumerable() _
        '                     Order By gU.decodeEmptyCInt(q("pld_split_fr_int").ToString, 0), gU.decodeEmptyCInt(q("pld_Seq").ToString, 0), _
        '                     q("PLD_BATCH_NO").ToString, _
        '                     q("PLD_WH").ToString, _
        '                     gU.decodeEmptyCInt(q("PLD_FLOOR").ToString, 0), _
        '                     q("PLD_AREA").ToString, _
        '                     q("PLD_RACK").ToString, _
        '                     q("PLD_BIN").ToString, _
        '                     gU.decodeEmptyCdbl(q("PLD_FOI_QTY").ToString, 0), _
        '                     gU.decodeEmptyCdbl(q("PLD_ITEM_QTY").ToString, 0) Descending

        'dt = query.CopyToDataTable()

        If Request("moduleAction") = "RELOADPL" Then
            Dim objPlt As PickListTable

            objPlt = New PickListTable(CO_CODE.Value, DO_CODE.Value, STORER_CODE.Value, Session("IMP_CODE"))

            objPlt.updateAvailBal(dt)

            objPlt.updateTotalQty(dt)

            objPlt.updateHoldBal(dt)
        End If

        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Private Sub generatePickList()
        Dim dtl_dt As DataTable
        Dim rows_count As Integer
        Dim selectSql, selectSql1, selectSql2, selSqlTmp1, selSqlTmp2 As String
        Dim i, j As Integer
        Dim LOC_dt, pickListDt As DataTable
        Dim objPlt As PickListTable
        Dim plSeq As Integer
        Dim isItemExist As Boolean
        Dim itmDict As Dictionary(Of String, Double)
        Dim reqDict As Dictionary(Of String, Double)
        Dim itmKey As String

        For i = dt.Rows.Count - 1 To 0 Step -1
            If dt.Rows(i).Item("mFlag") = "N" Then
                dt.Rows(i).Delete()
            Else
                Session("_M_OB_DO_TMP_pl_del_list") = gU.appendToList(Session("_M_OB_DO_TMP_pl_del_list"), dt.Rows(i).Item("pld_seq"))
                dt.Rows(i).Delete()
            End If
        Next
        dt.AcceptChanges()

        dtl_dt = Session("dt")

        Session("_M_OB_DO_TMP_pl_seq") = 0

        plSeq = CInt(Session("_M_OB_DO_TMP_pl_seq"))

        If dtl_dt.Rows.Count > 0 Then
            'selectSql = ""
            'selectSql1 = ""
            'selectSql2 = ""
            'itmDict = New Dictionary(Of String, Double)
            'reqDict = New Dictionary(Of String, Double)

            'For i = 0 To dtl_dt.Rows.Count - 1
            '    If DB.decodeDBNull(dtl_dt.Rows(i).Item("DOD_QTY"), 0) > 0 And dtl_dt.Rows(i).Item("mFlag").ToString <> "D" Then
            '        itmKey = dtl_dt.Rows(i).Item("dod_itm_code").ToString & "#_#" & _
            '                dtl_dt.Rows(i).Item("dod_pack_key").ToString & "#_#" & _
            '                gU.decodeNullOrEmpty(dtl_dt.Rows(i).Item("dod_pallet_no").ToString, "000")

            '        '###
            '        If dtl_dt.Rows(i).Item("dod_batch_no").ToString.Trim <> "" Then
            '            itmKey = itmKey & "#_#" & dtl_dt.Rows(i).Item("dod_batch_no").ToString.Trim
            '        End If

            '        If Not reqDict.ContainsKey(itmKey) Then
            '            reqDict.Add(itmKey, 0)
            '        End If

            '        If Not itmDict.ContainsKey(itmKey) Then
            '            itmDict.Add(itmKey, dtl_dt.Rows(i).Item("DOD_QTY"))
            '        Else
            '            itmDict.Item(itmKey) = itmDict.Item(itmKey) + dtl_dt.Rows(i).Item("DOD_QTY")
            '        End If



            '        If dtl_dt.Rows(i).Item("dod_batch_no").ToString.Trim <> "" Then
            '            If selectSql1 <> "" Then
            '                selectSql1 = selectSql1 & ", "
            '            End If
            '            selectSql1 = selectSql1 & "'" & gU.dbEncode(itmKey) & "'"
            '        Else
            '            If selectSql2 <> "" Then
            '                selectSql2 = selectSql2 & ", "
            '            End If
            '            selectSql2 = selectSql2 & "'" & gU.dbEncode(itmKey) & "'"
            '        End If

            '        'If selectSql <> "" Then
            '        '    selectSql = selectSql & ", "
            '        'End If

            '        'selectSql = selectSql & "'" & gU.dbEncode(itmKey) & "'"
            '    End If
            'Next

            ''If selectSql <> "" Then
            ''If selectSql1 <> "" OrElse selectSql2 <> "" Then
            ''### PREVIOUS LOGIC
            ''pickListDt = New DataTable
            ''pickListDt.Columns.Add("PLD_ITEM_NO", Type.GetType("System.String"))
            ''pickListDt.Columns.Add("PLD_PACK_KEY", Type.GetType("System.String"))
            ''pickListDt.Columns.Add("PLD_PALLET_NO", Type.GetType("System.String"))
            ''pickListDt.Columns.Add("PLD_DO_QTY", Type.GetType("System.Double"))
            ''pickListDt.Columns.Add("PLD_ITEM_QTY", Type.GetType("System.Double"))
            ''pickListDt.Columns.Add("ILOC_BAL_QTY", Type.GetType("System.Double"))
            ''pickListDt.Columns.Add("PLD_WH", Type.GetType("System.String"))
            ''pickListDt.Columns.Add("PLD_LOC", Type.GetType("System.String"))
            ''pickListDt.Columns.Add("PLD_FLOOR", Type.GetType("System.String"))
            ''pickListDt.Columns.Add("PLD_AREA", Type.GetType("System.String"))
            ''pickListDt.Columns.Add("PLD_RACK", Type.GetType("System.String"))
            ''pickListDt.Columns.Add("PLD_BIN", Type.GetType("System.String"))
            ''pickListDt.Columns.Add("PLD_BATCH_NO", Type.GetType("System.String"))

            'selSqlTmp1 = "select l.ITM_CODE, l.PACK_KEY, l.ILOC_PALLET_NO, 0.0 as COD_QTY, " & _
            '                "l.ILOC_BAL_QTY, l.ILOC_LOC, " & _
            '                "l.ILOC_WH, l.ILOC_FLOOR, l.ILOC_AREA, l.ILOC_RACK, l.ILOC_BIN, l.ILOC_BATCH_NO, "

            'selSqlTmp2 = "from WMS_ITEM_LOC_BAL l " & _
            '                "left outer join WMS_DATE_CODE DC " & _
            '                "on l.ILOC_BATCH_NO = DC.DC_DATE_CODE " & _
            '                "and l.STORER_CODE = DC.STORER_CODE " & _
            '                "and l.IMP_CODE = DC.IMP_CODE " & _
            '            "where l.STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' " & _
            '            "and l.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
            '            "and ISNULL(l.ILOC_BAL_QTY, 0) > 0 "

            'If selectSql1 <> "" Then
            '    selectSql1 = selSqlTmp1 & " 'Y' as BATCH_FLAG " & selSqlTmp2 & _
            '            "and l.ITM_CODE || '#_#' || l.PACK_KEY || '#_#' || l.ILOC_PALLET_NO || '#_#' || l.ILOC_BATCH_NO in (" & selectSql1 & ") "
            'End If

            'If selectSql2 <> "" Then
            '    selectSql2 = selSqlTmp1 & " 'N' as BATCH_FLAG " & selSqlTmp2 & _
            '            "and l.ITM_CODE || '#_#' || l.PACK_KEY || '#_#' || l.ILOC_PALLET_NO in (" & selectSql2 & ") "
            'End If

            'If selectSql1 <> "" Then
            '    selectSql = selectSql1

            '    If selectSql2 <> "" Then
            '        selectSql = selectSql & " union " & selectSql2
            '    End If
            'Else
            '    selectSql = selectSql2
            'End If

            'selectSql = selectSql & _
            '            "order by ITM_CODE, PACK_KEY, DC_CONV_DATE, ILOC_BATCH_NO, ILOC_BAL_QTY desc "

            ''"order by l.ITM_CODE, DC.DC_CONV_DATE, l.PACK_KEY, l.ILOC_BATCH_NO, l.ILOC_BAL_QTY desc "

            'LOC_dt = gDB.getDataTable(selectSql)

            ''###SAMUEL LOGIC
            'For i = 0 To LOC_dt.Rows.Count - 1
            '    If itmDict.ContainsKey(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString) Then
            '        LOC_dt.Rows(i).Item("COD_QTY") = itmDict.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString)
            '    Else
            '        LOC_dt.Rows(i).Item("COD_QTY") = itmDict.Item(LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString)
            '    End If
            'Next

            objPlt = New PickListTable(CO_CODE.Value, DO_CODE.Value, STORER_CODE.Value, Session("IMP_CODE"))

            pickListDt = objPlt.generatePickList(dtl_dt)
            'pickListDt = objPlt.getPickList()



            '### PREVIOUS LOGIC -----------------------------------------------------------------------------------------------------------------------------
            'For i = 0 To LOC_dt.Rows.Count - 1
            '    If gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY").ToString.Trim, 0) > 0 AndAlso _
            '        LOC_dt.Rows(i).Item("ILOC_LOC").ToString.Trim <> "" Then

            '        Dim itmK As String = LOC_dt.Rows(i).Item("ITM_CODE").ToString & "#_#" & _
            '                                                 LOC_dt.Rows(i).Item("PACK_KEY").ToString & "#_#" & _
            '                                                 LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString
            '        Dim currQty As Double = reqDict.Item(itmK)
            '        Dim reqQty As Double = itmDict.Item(itmK)

            '        If currQty < reqQty Then
            '            Dim balQty As Double = gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY").ToString.Trim, 0)

            '            Dim pNewRow = pickListDt.NewRow
            '            pNewRow.Item("PLD_ITEM_NO") = LOC_dt.Rows(i).Item("ITM_CODE").ToString.Trim
            '            pNewRow.Item("PLD_PACK_KEY") = LOC_dt.Rows(i).Item("PACK_KEY").ToString.Trim
            '            pNewRow.Item("PLD_PALLET_NO") = LOC_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim

            '            pNewRow.Item("PLD_WH") = LOC_dt.Rows(i).Item("ILOC_WH").ToString.Trim
            '            pNewRow.Item("PLD_LOC") = LOC_dt.Rows(i).Item("ILOC_LOC").ToString.Trim
            '            pNewRow.Item("PLD_FLOOR") = LOC_dt.Rows(i).Item("ILOC_FLOOR").ToString.Trim
            '            pNewRow.Item("PLD_AREA") = LOC_dt.Rows(i).Item("ILOC_AREA").ToString.Trim
            '            pNewRow.Item("PLD_RACK") = LOC_dt.Rows(i).Item("ILOC_RACK").ToString.Trim
            '            pNewRow.Item("PLD_BIN") = LOC_dt.Rows(i).Item("ILOC_BIN").ToString.Trim
            '            pNewRow.Item("PLD_BATCH_NO") = LOC_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
            '            pNewRow.Item("ILOC_BAL_QTY") = gU.decodeEmptyCdbl(LOC_dt.Rows(i).Item("ILOC_BAL_QTY").ToString.Trim, 0)
            '            pNewRow.Item("PLD_DO_QTY") = reqQty

            '            If currQty + balQty >= reqQty Then
            '                pNewRow.Item("PLD_ITEM_QTY") = reqQty - currQty
            '                currQty = reqQty
            '            Else

            '                pNewRow.Item("PLD_ITEM_QTY") = balQty

            '                currQty += balQty

            '            End If

            '            reqDict.Item(itmK) = currQty

            '            pickListDt.Rows.Add(pNewRow)
            '            pickListDt.AcceptChanges()
            '        End If
            '    End If
            'Next
            '### END OF PREVIOUS LOGIC -----------------------------------------------------------------------------------------------------------------------------

            For i = 0 To pickListDt.Rows.Count - 1
                dt.Rows.Add()

                plSeq = plSeq + 1

                rows_count = dt.Rows.Count

                dt.Rows(rows_count - 1).Item("PLD_SEQ") = CStr(plSeq)
                dt.Rows(rows_count - 1).Item("PLD_PICKED_BY") = Session("usr_id")
                dt.Rows(rows_count - 1).Item("PLD_ITEM_NO") = pickListDt.Rows(i).Item("PLD_ITEM_NO")
                dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = pickListDt.Rows(i).Item("ITM_SKU_NO")
                dt.Rows(rows_count - 1).Item("PLD_PACK_KEY") = pickListDt.Rows(i).Item("PLD_PACK_KEY")
                dt.Rows(rows_count - 1).Item("PLD_PALLET_NO") = pickListDt.Rows(i).Item("PLD_PALLET_NO")

                dt.Rows(rows_count - 1).Item("PLD_FOI_QTY") = pickListDt.Rows(i).Item("PLD_FOI_QTY")
                dt.Rows(rows_count - 1).Item("PLD_ITEM_QTY") = pickListDt.Rows(i).Item("PLD_FOI_QTY")

                dt.Rows(rows_count - 1).Item("PLD_QTY2") = pickListDt.Rows(i).Item("PLD_QTY2")
                dt.Rows(rows_count - 1).Item("PLD_SERIAL_NO") = pickListDt.Rows(i).Item("PLD_SERIAL_NO")

                dt.Rows(rows_count - 1).Item("PLD_DO_QTY") = pickListDt.Rows(i).Item("PLD_DO_QTY")

                dt.Rows(rows_count - 1).Item("PLD_SS_QTY") = pickListDt.Rows(i).Item("PLD_SS_QTY")

                dt.Rows(rows_count - 1).Item("STOCK_QTY") = pickListDt.Rows(i).Item("STOCK_QTY")

                dt.Rows(rows_count - 1).Item("ILOC_BAL_QTY") = pickListDt.Rows(i).Item("ILOC_BAL_QTY")

                dt.Rows(rows_count - 1).Item("HOLD_QTY") = pickListDt.Rows(i).Item("HOLD_QTY")
                dt.Rows(rows_count - 1).Item("TOTAL_BAL") = pickListDt.Rows(i).Item("TOTAL_BAL")

                'Will be updated by the updateTotalQty function
                dt.Rows(rows_count - 1).Item("TOTAL_ITEM_QTY") = 0
                dt.Rows(rows_count - 1).Item("TOTAL_FOI_QTY") = 0

                dt.Rows(rows_count - 1).Item("ILOC_EXPIRY_DATE") = pickListDt.Rows(i).Item("ILOC_EXPIRY_DATE")

                dt.Rows(rows_count - 1).Item("PLD_EXPIRY_DATE") = pickListDt.Rows(i).Item("ILOC_EXPIRY_DATE")
                dt.Rows(rows_count - 1).Item("PLD_MANU_DATE") = pickListDt.Rows(i).Item("ILOC_MANU_DATE")

                dt.Rows(rows_count - 1).Item("PLD_WH") = pickListDt.Rows(i).Item("PLD_WH")

                dt.Rows(rows_count - 1).Item("PLD_LOC") = pickListDt.Rows(i).Item("PLD_LOC")
                dt.Rows(rows_count - 1).Item("PLD_ORG_LOC") = pickListDt.Rows(i).Item("PLD_LOC")
                dt.Rows(rows_count - 1).Item("PLD_REMARK") = ""
                dt.Rows(rows_count - 1).Item("PLD_FLOOR") = pickListDt.Rows(i).Item("PLD_FLOOR")
                dt.Rows(rows_count - 1).Item("PLD_AREA") = pickListDt.Rows(i).Item("PLD_AREA")
                dt.Rows(rows_count - 1).Item("PLD_RACK") = pickListDt.Rows(i).Item("PLD_RACK")
                dt.Rows(rows_count - 1).Item("PLD_BIN") = pickListDt.Rows(i).Item("PLD_BIN")
                dt.Rows(rows_count - 1).Item("PLD_IS_LOAN") = "N"
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

                dt.Rows(rows_count - 1).Item("PLD_BATCH_NO") = pickListDt.Rows(i).Item("PLD_BATCH_NO")

                dt.Rows(rows_count - 1).Item("DOD_DISP_SEQ") = pickListDt.Rows(i).Item("DOD_DISP_SEQ")
                dt.Rows(rows_count - 1).Item("BN_CSMS_CODE") = pickListDt.Rows(i).Item("BN_CSMS_CODE")

                dt.Rows(rows_count - 1).Item("ITM_SERIAL_NO_YN") = pickListDt.Rows(i).Item("ITM_SERIAL_NO_YN")
            Next
            'End If
            '------------------------------------------------------------------------------------------------------------------

            objPlt.updateTotalQty(dt)

            objPlt = Nothing

            'Check if any item is not existed in the pick list
            For i = 0 To dtl_dt.Rows.Count - 1
                isItemExist = False
                For j = 0 To dt.Rows.Count - 1
                    If dtl_dt.Rows(i).Item("DOD_ITM_CODE").ToString.Trim = dt.Rows(j).Item("PLD_ITEM_NO").ToString.Trim AndAlso _
                        dtl_dt.Rows(i).Item("DOD_PACK_KEY").ToString.Trim = dt.Rows(j).Item("PLD_PACK_KEY").ToString.Trim AndAlso _
                        dtl_dt.Rows(i).Item("DOD_PALLET_NO").ToString.Trim = dt.Rows(j).Item("PLD_PALLET_NO").ToString.Trim AndAlso
                        (dtl_dt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim = "" OrElse dtl_dt.Rows(i).Item("DOD_BATCH_NO").ToString.Trim = dt.Rows(j).Item("PLD_BATCH_NO").ToString.Trim) Then

                        isItemExist = True

                        '###
                        'If dtl_dt.Rows(i).Item("DOD_PALLET_NO").ToString.Trim <> "" Then
                        '    If dtl_dt.Rows(i).Item("DOD_PALLET_NO").ToString.Trim = dt.Rows(j).Item("PLD_BATCH_NO").ToString.Trim Then
                        '        isItemExist = True
                        '    End If
                        'Else
                        '    isItemExist = True
                        'End If

                        Exit For
                    End If
                Next
                If Not isItemExist Then
                    dt.Rows.Add()

                    plSeq = plSeq + 1

                    rows_count = dt.Rows.Count

                    dt.Rows(rows_count - 1).Item("PLD_SEQ") = CStr(plSeq)
                    dt.Rows(rows_count - 1).Item("PLD_PICKED_BY") = Session("usr_id")
                    dt.Rows(rows_count - 1).Item("PLD_ITEM_NO") = dtl_dt.Rows(i).Item("DOD_ITM_CODE")
                    dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = dtl_dt.Rows(i).Item("ITM_SKU_NO")
                    dt.Rows(rows_count - 1).Item("PLD_PACK_KEY") = dtl_dt.Rows(i).Item("DOD_PACK_KEY")
                    dt.Rows(rows_count - 1).Item("PLD_PALLET_NO") = dtl_dt.Rows(i).Item("DOD_PALLET_NO")
                    dt.Rows(rows_count - 1).Item("PLD_FOI_QTY") = 0
                    dt.Rows(rows_count - 1).Item("PLD_ITEM_QTY") = 0

                    dt.Rows(rows_count - 1).Item("PLD_QTY2") = dtl_dt.Rows(i).Item("DOD_QTY2")
                    dt.Rows(rows_count - 1).Item("PLD_SERIAL_NO") = ""
                    dt.Rows(rows_count - 1).Item("PLD_SS_QTY") = 0

                    dt.Rows(rows_count - 1).Item("PLD_DO_QTY") = dtl_dt.Rows(i).Item("DOD_QTY")
                    dt.Rows(rows_count - 1).Item("STOCK_QTY") = 0
                    dt.Rows(rows_count - 1).Item("ILOC_BAL_QTY") = 0

                    dt.Rows(rows_count - 1).Item("HOLD_QTY") = 0
                    dt.Rows(rows_count - 1).Item("TOTAL_BAL") = 0

                    dt.Rows(rows_count - 1).Item("TOTAL_ITEM_QTY") = 0
                    dt.Rows(rows_count - 1).Item("TOTAL_FOI_QTY") = 0

                    dt.Rows(rows_count - 1).Item("ILOC_EXPIRY_DATE") = ""

                    dt.Rows(rows_count - 1).Item("PLD_EXPIRY_DATE") = ""
                    dt.Rows(rows_count - 1).Item("PLD_MANU_DATE") = ""

                    dt.Rows(rows_count - 1).Item("PLD_WH") = ""

                    dt.Rows(rows_count - 1).Item("PLD_LOC") = ""
                    dt.Rows(rows_count - 1).Item("PLD_ORG_LOC") = ""
                    dt.Rows(rows_count - 1).Item("PLD_REMARK") = ""
                    dt.Rows(rows_count - 1).Item("PLD_FLOOR") = ""
                    dt.Rows(rows_count - 1).Item("PLD_AREA") = ""
                    dt.Rows(rows_count - 1).Item("PLD_RACK") = ""
                    dt.Rows(rows_count - 1).Item("PLD_BIN") = ""
                    dt.Rows(rows_count - 1).Item("PLD_IS_LOAN") = "N"
                    dt.Rows(rows_count - 1).Item("mFlag") = "N"

                    dt.Rows(rows_count - 1).Item("PLD_BATCH_NO") = dtl_dt.Rows(i).Item("DOD_BATCH_NO")

                    dt.Rows(rows_count - 1).Item("DOD_DISP_SEQ") = dtl_dt.Rows(i).Item("DOD_DISP_SEQ")
                    dt.Rows(rows_count - 1).Item("BN_CSMS_CODE") = ""


                    dt.Rows(rows_count - 1).Item("PLD_EXPIRY_DATE") = dtl_dt.Rows(i).Item("DOD_EXPIRY_DATE")
                    dt.Rows(rows_count - 1).Item("PLD_MANU_DATE") = dtl_dt.Rows(i).Item("DOD_MANU_DATE")

                    dt.Rows(rows_count - 1).Item("ITM_SERIAL_NO_YN") = dtl_dt.Rows(i).Item("ITM_SERIAL_NO_YN")
                End If
            Next

            Session("_M_OB_DO_TMP_pl_seq") = CStr(plSeq)

            dt.AcceptChanges()

            'Dim query = From q In dt.AsEnumerable() _
            '                 Order By _
            '                 q("DOD_DISP_SEQ"), _
            '                 q("ITM_SKU_NO").ToString, _
            '                 q("PLD_ITEM_NO").ToString, _
            '                 q("PLD_PACK_KEY"), _
            '                 q("PLD_PALLET_NO").ToString, _
            '                 q("PLD_BATCH_NO").ToString

            'dt = query.CopyToDataTable()
        End If

        'For i = 0 To dtl_dt.Rows.Count - 1
        '    dt.Rows.Add()

        '    Session("_M_OB_DO_TMP_pl_seq") = CStr(CInt(Session("_M_OB_DO_TMP_pl_seq")) + 1)

        '    rows_count = dt.Rows.Count
        '    dt.Rows(rows_count - 1).Item("pld_seq") = Session("_M_OB_DO_TMP_pl_seq").ToString
        '    dt.Rows(rows_count - 1).Item("mFlag") = "N"
        '    dt.Rows(rows_count - 1).Item("pld_item_no") = dtl_dt.Rows(i).Item("dod_itm_code").ToString
        '    dt.Rows(rows_count - 1).Item("pld_item_qty") = dtl_dt.Rows(i).Item("dod_qty").ToString
        'Next

        'dt.AcceptChanges()
    End Sub

    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        generatePickList()
        hd_drumList.Value = ""
        BindGV()

        'GridView1.DataSource = dt
        'GridView1.DataBind()
    End Sub

    Protected Sub btnRefresh_Click(sender As Object, e As System.EventArgs) Handles btnRefresh.Click
        Dim objPlt As PickListTable

        If cU.gfBuildDataTableforGridView(dt, GridView1, True, New String() {"Y", "N"}) Then
            objPlt = New PickListTable(CO_CODE.Value, DO_CODE.Value, STORER_CODE.Value, Session("IMP_CODE"))

            objPlt.updateAvailBal(dt)

            objPlt.updateTotalQty(dt)

            objPlt.updateHoldBal(dt)

            GridView1.DataSource = dt
            GridView1.DataBind()

            updtPnl_gvPickList.Update()
        End If
    End Sub

    Private Sub ShowRedrum(ByVal drumID As String, ByVal toDrumID As String, ByVal ToDrumList As String, ByVal toDrumIloc_seq As String)
        Dim sqlString As String = "Select 'C' as dFlag, ILBS_SEQ, ITM_CODE, PACK_KEY, ILOC_SEQ, ILBS_SERIAL_NO, ILBS_DRUM_ID, ILBS_DRUM_LEVEL, ILBS_UOM2, ILBS_QTY2 FROM WMS_ITEM_LOC_BAL_S " & _
                                  "where STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' and IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                  "and ILBS_DRUM_ID='" & gU.dbEncode(drumID) & "' and ILBS_QTY2 > 0 and ILBS_SL='Y' " & _
                                  " order by ILBS_DRUM_LEVEL"
        Dim sqlString2 As String = "Select 'X' as dFlag,ITM_CODE, PACK_KEY, ILBS_SEQ, ILOC_SEQ, ILBS_SERIAL_NO, ILBS_DRUM_ID, ILBS_DRUM_LEVEL, ILBS_UOM2, ILBS_QTY2 FROM WMS_ITEM_LOC_BAL_S " & _
                                   "where 1=3"

        Dim tempDT As DataTable = gDB.getDataTable(sqlString)
        Dim dummyDT As DataTable

        If tempDT.Rows.Count > 0 Then
            FromDrumID_TXT.Text = drumID
            selectedSeqList.Value = ToDrumList
            ToDrumID_TXT.Text = toDrumID.ToUpper.Trim
            ToDrum_HD.Value = toDrumID.ToUpper.Trim
            ToIloc_seq.Value = toDrumIloc_seq

            Dim balCount As Double = 0
            Dim takeCount As Double = 0

            For i = 0 To tempDT.Rows.Count - 1
                For j = 0 To GridView1.Rows.Count - 1
                    'If DirectCast(GridView1.Rows(j).FindControl("pld_item_no"), Label).Text = tempDT.Rows(i).Item("itm_code").ToString And DirectCast(GridView1.Rows(j).FindControl("pld_pack_key"), Label).Text = tempDT.Rows(i).Item("pack_key").ToString And _
                    '    DirectCast(GridView1.Rows(j).FindControl("pld_serial_no"), TextBox).Text.Trim.ToUpper = tempDT.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim.ToUpper Then
                    If DirectCast(GridView1.Rows(j).FindControl("pld_serial_no"), TextBox).Text.Trim.ToUpper = tempDT.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim.ToUpper Then
                        balCount = CDbl(tempDT.Rows(i).Item("ILBS_QTY2").ToString.Trim)

                        If DirectCast(GridView1.Rows(j).FindControl("PLD_QTY2"), TextBox).Text.Trim <> "" AndAlso gU.isDecimal(DirectCast(GridView1.Rows(j).FindControl("pld_qty2"), TextBox).Text.Trim) AndAlso DirectCast(GridView1.Rows(j).FindControl("pld_qty2"), TextBox).Text.Trim > 0 Then
                            takeCount = DirectCast(GridView1.Rows(j).FindControl("pld_qty2"), TextBox).Text.Trim
                        Else
                            takeCount = balCount
                        End If

                        If balCount - takeCount <= 0 Then
                            tempDT.Rows(i).Item("dFlag") = "X"
                        Else
                            tempDT.Rows(i).Item("ILBS_QTY2") = balCount - takeCount
                        End If
                    End If
                Next
            Next
            tempDT.AcceptChanges()

            GVDrumIN.DataSource = tempDT
            GVDrumIN.DataBind()

            If toDrumID <> "" Then

                sqlString2 = "Select 'O' as dFlag,ITM_CODE, PACK_KEY, ILBS_SEQ, ILOC_SEQ, ILBS_SERIAL_NO, ILBS_DRUM_ID, ILBS_DRUM_LEVEL, ILBS_UOM2, ILBS_QTY2 FROM WMS_ITEM_LOC_BAL_S " & _
                             "where STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' and IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                             "and ILBS_DRUM_ID='" & gU.dbEncode(toDrumID) & "' and ILBS_QTY2 > 0 and ILBS_SL='Y' " & _
                             " order by ILBS_DRUM_LEVEL"
                dummyDT = gDB.getDataTable(sqlString2)

                If ToDrumList <> "" Then
                    Dim nRow As DataRow

                    Dim seqList() As String = Split(ToDrumList, ", ")
                    Dim tempArr() As String
                    Dim currSeq As String = ""
                    Dim tempDrumLV As Integer = 0
                    For i = 0 To seqList.Length - 1
                        For j = 0 To tempDT.Rows.Count - 1
                            tempArr = Split(seqList(i), "||")


                            currSeq = tempArr(0)
                            tempDrumLV = tempArr(1)

                            If tempDT.Rows(j).Item("ILBS_SEQ").ToString.Trim = currSeq Then
                                nRow = dummyDT.NewRow

                                nRow.Item("dFlag") = "N"

                                nRow.Item("ITM_CODE") = tempDT.Rows(j).Item("ITM_CODE")
                                nRow.Item("PACK_KEY") = tempDT.Rows(j).Item("PACK_KEY")
                                nRow.Item("ILBS_SEQ") = tempDT.Rows(j).Item("ILBS_SEQ")
                                nRow.Item("ILOC_SEQ") = tempDT.Rows(j).Item("ILOC_SEQ")
                                nRow.Item("ILBS_SERIAL_NO") = tempDT.Rows(j).Item("ILBS_SERIAL_NO")
                                nRow.Item("ILBS_DRUM_ID") = toDrumID
                                nRow.Item("ILBS_DRUM_LEVEL") = gU.decodeEmptyCInt(tempDrumLV, 0)
                                nRow.Item("ILBS_UOM2") = tempDT.Rows(j).Item("ILBS_UOM2")
                                nRow.Item("ILBS_QTY2") = tempDT.Rows(j).Item("ILBS_QTY2")

                                dummyDT.Rows.Add(nRow)

                            End If
                        Next

                    Next

                    dummyDT.AcceptChanges()
                End If

                GVDrumOUT.DataSource = dummyDT
            Else
                dummyDT = gDB.getDataTable(sqlString2)
                GVDrumOUT.DataSource = dummyDT
            End If

            GVDrumOUT.DataBind()
            drumPop.Show()
        Else
            uiFun.displayMsgNew(Me, "", "No available Cable in this Drum!", Session("gLang"))
        End If

    End Sub

    Protected Sub GVDrumIN_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVDrumIN.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("dFlag"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "dFlag").ToString.Trim

                CType(e.Row.FindControl("ILBS_SERIAL_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_SERIAL_NO").ToString.Trim
                CType(e.Row.FindControl("ILOC_SEQ"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ILOC_SEQ").ToString.Trim
                CType(e.Row.FindControl("ILBS_SEQ"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ILBS_SEQ").ToString.Trim
                CType(e.Row.FindControl("ILBS_DRUM_LEVEL"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_DRUM_LEVEL").ToString.Trim
                CType(e.Row.FindControl("ILBS_QTY2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_QTY2").ToString.Trim
                CType(e.Row.FindControl("ILBS_UOM2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_UOM2").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "dFlag").ToString.Trim = "C" Then
                    CType(e.Row.FindControl("selectYN"), CheckBox).Visible = True

                    If InStr(" " & selectedSeqList.Value, DataBinder.Eval(e.Row.DataItem, "ILBS_SEQ").ToString.Trim & "||") Then
                        CType(e.Row.FindControl("selectYN"), CheckBox).Checked = True
                    End If
                Else
                    CType(e.Row.FindControl("selectYN"), CheckBox).Checked = False
                End If

        End Select
    End Sub

    Protected Sub GVDrumOUT_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVDrumOUT.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("dFlag"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "dFlag").ToString.Trim

                CType(e.Row.FindControl("ILBS_SERIAL_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_SERIAL_NO").ToString.Trim
                CType(e.Row.FindControl("ILOC_SEQ"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ILOC_SEQ").ToString.Trim
                CType(e.Row.FindControl("ILBS_SEQ"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "ILBS_SEQ").ToString.Trim
                CType(e.Row.FindControl("ILBS_DRUM_LEVEL"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_DRUM_LEVEL").ToString.Trim
                CType(e.Row.FindControl("ILBS_QTY2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_QTY2").ToString.Trim
                CType(e.Row.FindControl("ILBS_UOM2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILBS_UOM2").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "dFlag").ToString.Trim <> "N" Then CType(e.Row.FindControl("ILBS_DRUM_LEVEL"), TextBox).Enabled = False
        End Select
    End Sub

    Protected Sub btnMove_Click(sender As Object, e As System.EventArgs) Handles btnMove.Click
        Dim SuccessFlag As Boolean = False
        Dim sqlString As String = ""
        Dim tempDT As DataTable

        If GVDrumIN.Rows.Count > 0 Then
            If Not String.IsNullOrWhiteSpace(ToDrumID_TXT.Text) Then
                sqlString = "select count(1) as count from WMS_ITEM_LOC_BAL_S where STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' and IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                          " and ilbs_drum_id=upper('" & gU.dbEncode(ToDrumID_TXT.Text.Trim) & "')"


                Dim dCount As Integer = gU.decodeEmptyCInt(DB.getValueFromSQL(sqlString), 0)

                If dCount > 0 And FromDrumID_TXT.Text.ToUpper.Trim <> ToDrumID_TXT.Text.Trim.ToUpper Then
                    Dim tickCount As Integer = 0
                    For i = 0 To GVDrumIN.Rows.Count - 1
                        If DirectCast(GVDrumIN.Rows(i).FindControl("selectYN"), CheckBox).Checked Then
                            tickCount += 1
                        End If
                    Next

                    If tickCount > 0 Then
                        sqlString = "Select 'O' as dFlag, WMS_ITEM_LOC_BAL.ILOC_SEQ, WMS_ITEM_LOC_BAL_S.ITM_CODE, WMS_ITEM_LOC_BAL_S.PACK_KEY, ILBS_SEQ, ILBS_SERIAL_NO, ILBS_DRUM_ID, ILBS_DRUM_LEVEL, ILBS_UOM2, ILBS_QTY2 FROM WMS_ITEM_LOC_BAL_S " & _
                                    "INNER JOIN WMS_ITEM_LOC_BAL ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                    "where WMS_ITEM_LOC_BAL_S.STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' and WMS_ITEM_LOC_BAL_S.IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                    "and ILBS_DRUM_ID='" & gU.dbEncode(ToDrumID_TXT.Text.ToUpper.Trim) & "' and ILBS_QTY2 > 0 and ILBS_SL='Y' " & _
                                    " order by ILBS_DRUM_LEVEL"

                        tempDT = gDB.getDataTable(sqlString)
                        Dim tIloc_seq As String = ""
                        If tempDT.Rows.Count = 0 Then
                            sqlString = "Select max(WMS_ITEM_LOC_BAL.ILOC_SEQ) as ILOC_SEQ FROM WMS_ITEM_LOC_BAL_S INNER JOIN WMS_ITEM_LOC_BAL ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                        "where WMS_ITEM_LOC_BAL_S.STORER_CODE = '" & gU.dbEncode(STORER_CODE.Value) & "' and WMS_ITEM_LOC_BAL_S.IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                        "and ILBS_DRUM_ID='" & gU.dbEncode(ToDrumID_TXT.Text.ToUpper.Trim) & "'"

                            tIloc_seq = DB.getValueFromSQL(sqlString)
                        Else
                            tIloc_seq = tempDT.Rows(0).Item("ILOC_SEQ").ToString.Trim
                        End If

                        Dim nRow As DataRow
                        Dim maxDrum As Integer
                        For i = GVDrumIN.Rows.Count - 1 To 0 Step -1
                            If DirectCast(GVDrumIN.Rows(i).FindControl("selectYN"), CheckBox).Checked Then

                                If tempDT.Rows.Count = 0 Then
                                    maxDrum = 1
                                Else
                                    maxDrum = gU.decodeEmptyCInt(tempDT.Compute("Max(ILBS_DRUM_LEVEL)", ""), 0) + 1
                                End If


                                nRow = tempDT.NewRow

                                nRow.Item("dFlag") = "N"

                                nRow.Item("ITM_CODE") = ""
                                nRow.Item("PACK_KEY") = ""
                                nRow.Item("ILOC_SEQ") = tIloc_seq
                                nRow.Item("ILBS_SEQ") = DirectCast(GVDrumIN.Rows(i).FindControl("ILBS_SEQ"), HiddenField).Value
                                nRow.Item("ILOC_SEQ") = DirectCast(GVDrumIN.Rows(i).FindControl("ILOC_SEQ"), HiddenField).Value
                                nRow.Item("ILBS_SERIAL_NO") = DirectCast(GVDrumIN.Rows(i).FindControl("ILBS_SERIAL_NO"), Label).Text
                                nRow.Item("ILBS_DRUM_ID") = ToDrumID_TXT.Text.Trim.ToUpper
                                nRow.Item("ILBS_DRUM_LEVEL") = maxDrum
                                nRow.Item("ILBS_UOM2") = DirectCast(GVDrumIN.Rows(i).FindControl("ILBS_UOM2"), Label).Text
                                nRow.Item("ILBS_QTY2") = DirectCast(GVDrumIN.Rows(i).FindControl("ILBS_QTY2"), Label).Text

                                tempDT.Rows.Add(nRow)
                            End If
                        Next

                        tempDT.AcceptChanges()

                        GVDrumOUT.DataSource = tempDT
                        GVDrumOUT.DataBind()
                        ToDrum_HD.Value = ToDrumID_TXT.Text.Trim.ToUpper
                        ToIloc_seq.Value = tIloc_seq

                        SuccessFlag = True
                    Else
                        uiFun.displayMsgNew(Me, "", "Please select Cable to transfer!", Session("gLang"))
                    End If
                Else
                    uiFun.displayMsgNew(Me, "", "Invalid To Drum ID!", Session("gLang"))
                End If
            Else
                uiFun.displayMsgNew(Me, "", "Please enter To Drum ID!", Session("gLang"))
            End If
        End If

        If Not SuccessFlag Then
            sqlString = "Select 'x' as dFlag,ITM_CODE, PACK_KEY, ILBS_SEQ, ILOC_SEQ, ILBS_SERIAL_NO, ILBS_DRUM_ID, ILBS_DRUM_LEVEL, ILBS_UOM2, ILBS_QTY2 FROM WMS_ITEM_LOC_BAL_S " & _
                                   "where 1=3"
            tempDT = gDB.getDataTable(sqlString)

            GVDrumOUT.DataSource = tempDT
            GVDrumOUT.DataBind()
        End If


        drumPop.Show()
    End Sub

    Protected Sub btnConfirmP_Click(sender As Object, e As System.EventArgs) Handles btnConfirmP.Click

        Dim rIndex As Integer = FromRowIDX.Value
        Dim toDrumID As String = ToDrum_HD.Value
        Dim iLocSeq As String = ToIloc_seq.Value
        Dim todrumSeqList As String = ""

        If GVDrumOUT.Rows.Count > 0 Then
            For i = 0 To GVDrumOUT.Rows.Count - 1
                If DirectCast(GVDrumOUT.Rows(i).FindControl("dFlag"), HiddenField).Value = "N" Then
                    todrumSeqList = gU.appendToList(todrumSeqList, DirectCast(GVDrumOUT.Rows(i).FindControl("ILBS_SEQ"), HiddenField).Value & "||" & DirectCast(GVDrumOUT.Rows(i).FindControl("ILBS_DRUM_LEVEL"), TextBox).Text)
                End If
            Next

            If Not String.IsNullOrWhiteSpace(todrumSeqList) Then
                dt.Rows(rIndex).Item("PLD_TO_DRUM_ID") = toDrumID
                dt.Rows(rIndex).Item("PLD_TO_DRUM_CABLE_LIST") = todrumSeqList
                dt.Rows(rIndex).Item("PLD_TO_DRUM_ILOC_SEQ") = iLocSeq
            End If

            dt.AcceptChanges()

        Else
            uiFun.displayMsgNew(Me, "", "Please select Cable for Re-Drum!", Session("gLang"))
            drumPop.Show()
        End If

    End Sub

    Protected Sub pld_wh_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim ddlWH As DropDownList = CType(sender, DropDownList)
        Dim Row = CType(ddlWH.Parent.Parent, GridViewRow)
        Dim idx = Row.RowIndex
        If ddlWH.SelectedValue <> Nothing Then
            uiFun.load_ComboBox(CType(Row.FindControl("pld_loc"), AjaxControlToolkit.ComboBox),
                                    "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & ddlWH.SelectedValue & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))
        End If
    End Sub

End Class
