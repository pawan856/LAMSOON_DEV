Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class PutAWay
    Inherits System.Web.UI.Page
    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private dt As New DataTable
    Private wmsFun As New WMSFunc
    Private dtRejReason As DataTable

    Structure itmStruct
        Dim itmName As String
        Dim qty As Double
        Dim itmSeq As String
        Dim batchNo As String
        Dim skuNo As String
        Dim expDate As String
        Dim manuDate As String
        Dim whCode As String
    End Structure

    Structure itmLocQty
        Dim loc As String
        Dim qty As Double
    End Structure

    Dim DocType As String = ""
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim moduleAction As String
        Dim dtl_dt As DataTable
        Dim paTable As PutAwayTable

        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils("IB_GR", Session("usr_id"), Me)
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

        If Not IsPostBack Then
            uiFun.load_dropdown(app_wh, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))
            app_wh.Attributes.Add("onchange", "dsp_app_loc.innerHTML='';app_loc.value='';")

        End If

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Put Away"
            saveBtn1.Text = "OK"
            saveBtn2.Text = "OK"
            newrow.Text = "Add"
            btnReset.OnClientClick = "return confirm(""Confirm to reset all the items from the GR?"");"
            'saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            'If Session("pagemode") = "N" Then
            '    GR_CODE.Text = "[No. will be auto generated]"
            'End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "上架"
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

        dt = Session("_M_IB_GR_TMP_pa_dt")
        If Not IsPostBack Or moduleAction = "RELOADPL" Then

            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                dtl_dt = Session("dt")

                If dtl_dt.Rows.Count > 0 Then
                    paTable = New PutAwayTable(dtl_dt, dt, Session("_M_IB_GR_TMP_pa_seq"), Session("_M_IB_GR_TMP_pa_del_list"))

                    paTable.genPutAway(True)

                    Session("_M_IB_GR_TMP_pa_del_list") = paTable.putAwayDelList

                    Session("_M_IB_GR_TMP_pa_seq") = paTable.putAwaySeq

                    GridView1.DataSource = paTable.putAwayDataTable
                    GridView1.DataBind()
                End If
            End If

            Call BindGV()
            'Else
            'updtLabelFields()
        End If

        If Request("GR_STATUS") = "CANCELLED" Then
            ar.sec_write = "N"
        ElseIf Request("GR_STATUS") = "POSTED" Then
            ar.sec_write = "N"
        End If

        DocType = GR_DOC_TYPE.Value

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

        If Request("moduleAction") = "SAVEOK" Then
            save()
        End If
    End Sub

    'Protected Sub Page_LoadComplete(sender As Object, e As System.EventArgs) Handles Me.LoadComplete
    '    Dim sm As ScriptManager = ScriptManager.GetCurrent(Page)
    '    Dim i As Integer

    '    sm.RegisterAsyncPostBackControl(btnSelOtherLoc)

    '    For i = 0 To GridView1.Rows.Count - 1
    '        sm.RegisterAsyncPostBackControl(CType(GridView1.Rows(i).FindControl("Image_Loc_LookUp"), ImageButton))
    '    Next

    '    For i = 0 To gvResrvLocList.Rows.Count - 1
    '        sm.RegisterAsyncPostBackControl(CType(gvResrvLocList.Rows(i).FindControl("btnResrvLocSelect"), Button))
    '    Next
    'End Sub

    Private Function customizectrl(ByVal ctl As Control, ByRef ctrlArrayList As ArrayList) As Boolean
        customizectrl = False
    End Function

    Private Function page_customizectrl(ByVal ctl As Control) As Boolean
        page_customizectrl = False

        If ctl.ID = "cSBBtn" Then
            CType(ctl, Button).Enabled = True
            page_customizectrl = True
        End If
    End Function

    Protected Sub GridView1_DataBinding(sender As Object, e As System.EventArgs) Handles GridView1.DataBinding
        Dim selectSql As String

        selectSql = "select colc_code as code, colc_eng_value as name " &
                    "from wms_col_code " &
                    "where colc_tabcol = 'WMS_GOODSRCV_D.GRD_REJ_REASON' " &
                    "order by colc_eng_value "

        dtRejReason = gDB.getDataTable(selectSql)
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
                Call cU.changeGVLabel(oGridViewRow, e, "PA List No.", "儲存物件編號", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Internal Item Code WMS", "貨物編號")
                Call cU.changeGVLabel(oGridViewRow, e, "Item Code", "貨物編號")
                Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物件名稱")
                Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內碼", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板編號", HorizontalAlign.Center)
                'Call cU.changeGVLabel(oGridViewRow, e, "Reference", "參考編號", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "批次號")
                Call cU.changeGVLabel(oGridViewRow, e, "Expiry Date", "失效日期")
                Call cU.changeGVLabel(oGridViewRow, e, "Suggested Qty", "建議數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Put Away Qty", "儲存數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "RCV Qty", "收貨數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Subinventory", "子庫存", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Location", "位置", HorizontalAlign.Center)
                Call cU.changeGVLabel(oGridViewRow, e, "Rej Qty", "拒收數量", HorizontalAlign.Right)
                Call cU.changeGVLabel(oGridViewRow, e, "Rej Reason", "拒收原因")
                Call cU.changeGVLabel(oGridViewRow, e, "Final Location", "Final Location")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                Call cU.changeGVLabel(oGridViewRow, e, "", "")
                REM **********************

                oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
                'Case DataControlRowType.DataRow
        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim dtl_dt As DataTable
        Dim itmQty As Double
        Dim i As Integer

        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                REM **********************
                REM Modify Here
                'Dim xFlag As String = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                CType(e.Row.FindControl("gra_disp_seq"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRA_DISP_SEQ").ToString.Trim
                CType(e.Row.FindControl("gra_itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRA_ITM_CODE").ToString.Trim
                CType(e.Row.FindControl("dsp_itm_name"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "DSP_ITM_NAME").ToString.Trim
                CType(e.Row.FindControl("gra_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRA_PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("gra_pallet_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRA_PALLET_NO").ToString.Trim
                CType(e.Row.FindControl("gra_sug_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRA_SUG_QTY").ToString.Trim
                CType(e.Row.FindControl("gra_pa_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRA_PA_QTY").ToString.Trim
                'CType(e.Row.FindControl("gra_ref_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRA_REF_NO").ToString.Trim

                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ITM_SKU_NO").ToString.Trim
                CType(e.Row.FindControl("GRA_PA_LIST_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRA_PA_LIST_NO").ToString.Trim

                'uiFun.load_dropdown(CType(e.Row.FindControl("GRA_BATCH_NO"), DropDownList), "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.Value) & "' order by 1", "dc_date_code", "dc_date_code", , Session("gSelectLabel"))
                'CType(e.Row.FindControl("GRA_BATCH_NO"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "GRA_BATCH_NO").ToString.Trim
                CType(e.Row.FindControl("GRA_BATCH_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRA_BATCH_NO").ToString.Trim

                CType(e.Row.FindControl("GRA_EXPIRY_DATE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRA_EXPIRY_DATE").ToString.Trim

                dtl_dt = Session("dt")

                If dtl_dt.Rows.Count > 0 Then
                    itmQty = 0

                    For i = 0 To dtl_dt.Rows.Count - 1


                        If DB.decodeDBNull(dtl_dt.Rows(i).Item("GRD_RCV_QTY"), 0) > 0 And dtl_dt.Rows(i).Item("mFlag").ToString <> "D" And
                             DataBinder.Eval(e.Row.DataItem, "GRA_ITM_CODE").ToString.Trim = dtl_dt.Rows(i).Item("grd_itm_code").ToString And
                                DataBinder.Eval(e.Row.DataItem, "GRA_PACK_KEY").ToString.Trim = dtl_dt.Rows(i).Item("grd_pack_key").ToString And
                                    DataBinder.Eval(e.Row.DataItem, "GRA_PALLET_NO").ToString.Trim = dtl_dt.Rows(i).Item("grd_pallet_no").ToString And
                                    DataBinder.Eval(e.Row.DataItem, "GRA_BATCH_NO").ToString.Trim = dtl_dt.Rows(i).Item("GRD_BATCH_NO").ToString Then
                            'DataBinder.Eval(e.Row.DataItem, "GRA_REF_NO").ToString.Trim = dtl_dt.Rows(i).Item("grd_ref_no").ToString Then

                            itmQty = itmQty + gU.decodeEmptyCdbl(dtl_dt.Rows(i).Item("GRD_RCV_QTY"), 0)
                        End If
                    Next
                End If

                CType(e.Row.FindControl("rcv_qty"), Label).Text = Format(itmQty, "0.0000")

                Dim nDropDown As DropDownList = CType(e.Row.FindControl("gra_wh"), DropDownList)
                uiFun.load_dropdown(nDropDown, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "GRA_WH").ToString.Trim
                'nDropDown.Attributes.Add("onchange", HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_gra_loc"), Label).ClientID) & ".innerHTML='';" & _
                '                            HttpUtility.HtmlEncode(CType(e.Row.FindControl("gra_loc"), HiddenField).ClientID) & ".value='';")


                'nDropDown = CType(e.Row.FindControl("gra_loc"), AjaxControlToolkit.ComboBox)
                'uiFun.load_dropdown(nDropDown, "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & DataBinder.Eval(e.Row.DataItem, "gra_wh").ToString.Trim & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))
                'nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "gra_loc").ToString.Trim

                uiFun.load_ComboBox(CType(e.Row.FindControl("gra_loc"), AjaxControlToolkit.ComboBox),
                                    "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & DataBinder.Eval(e.Row.DataItem, "gra_wh").ToString.Trim & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))

                CType(e.Row.FindControl("gra_loc"), AjaxControlToolkit.ComboBox).SelectedValue = DataBinder.Eval(e.Row.DataItem, "gra_loc").ToString.Trim

                'CType(e.Row.FindControl("dsp_gra_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "GRA_LOC").ToString.Trim
                'CType(e.Row.FindControl("gra_loc"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "GRA_LOC").ToString.Trim


                CType(e.Row.FindControl("gra_rej_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRA_REJ_QTY").ToString.Trim

                uiFun.load_dropdown(CType(e.Row.FindControl("gra_rej_reason"), DropDownList), dtRejReason, "CODE", "NAME", , Session("gSelectLabel"), DataBinder.Eval(e.Row.DataItem, "GRA_REJ_REASON").ToString.Trim, , True)

                'CType(e.Row.FindControl("gra_remark"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "GRA_REMARK").ToString.Trim

                uiFun.load_ComboBox(CType(e.Row.FindControl("gra_remark"), AjaxControlToolkit.ComboBox),
                                    "Select (a.WH_CODE + a.LOC_KEY) as CODE,(a.WH_CODE + a.LOC_KEY) as NAME from WMS_WH_BIN a where a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))

                CType(e.Row.FindControl("gra_remark"), AjaxControlToolkit.ComboBox).SelectedValue = DataBinder.Eval(e.Row.DataItem, "gra_remark").ToString.Trim


                'Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp"), Image)
                'nImage.Attributes.Add("onmousedown", "MM_swapImage('" & gU.jsHTMLEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                'nImage.Attributes.Add("onclick", "LocLookUp('" & gU.jsHTMLEncode(DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_NO").ToString.Trim) & "', " & _
                '                      "'" & gU.jsHTMLEncode(DataBinder.Eval(e.Row.DataItem, "PLD_PACK_KEY").ToString.Trim) & "')")

                Dim nImage As ImageButton = CType(e.Row.FindControl("Image_Loc_LookUp"), ImageButton)


                Dim selectSql As String

                selectSql = "select count(*) as value " &
                            "from wms_loc_book_hd h, wms_loc_book_dtl d " &
                            "where h.imp_code = d.imp_code " &
                            "and h.storer_code = d.storer_code " &
                            "and h.lobh_code = d.lobh_code " &
                            "and h.lobh_status = 'NEW' " &
                            "and d.lobd_sku_no = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim) & "' " &
                            "and d.pack_key = '" & gU.dbEncode(DataBinder.Eval(e.Row.DataItem, "gra_pack_key").ToString.Trim) & "' "

                'If DB.getValueFromSQL(selectSql) <> "0" Then
                '    nImage.CommandName = "OPEN_RESERV_LOOKUP"
                'Else
                '    nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                '    nImage.Attributes.Add("onclick", "LocLookUp(document.piform." & HttpUtility.HtmlEncode(nDropDown.ClientID) & ".value, '" & _
                '                          HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_gra_loc"), Label).ClientID) & "', " & _
                '                          "'" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("gra_loc"), HiddenField).ClientID) & "', " & _
                '                          "'" & HttpUtility.HtmlEncode(nDropDown.ClientID) & "');return false;")
                'End If

                REM **********************

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    nButton.Text = "删除"
                End If

                'If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                '    Call ar.hideGVRow(GridView1, e.Row)
                'End If

                'If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                '    nButton.Enabled = False
                'End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    e.Row.Visible = False
                Else
                    dtlSuffixList.Value = gU.appendToList(dtlSuffixList.Value, e.Row.ClientID)
                End If
        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        If validateAll("N") Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True, New String() {"Y", "N"}) Then
                Dim rows_count As Integer = 0
                REM **********************
                REM Modify Here

                Session("_M_IB_GR_TMP_pa_seq") = CStr(CInt(Session("_M_IB_GR_TMP_pa_seq")) + 1)

                dt.Rows.Add()
                rows_count = dt.Rows.Count

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("pld_seq") = Session("_M_IB_GR_TMP_pa_seq").ToString
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
        Dim itmQty As Double
        Dim i As Integer
        Dim selectSql As String
        Dim tmpDt As DataTable

        If e.CommandName = "SplitItem" Then
            If validateAll("L") Then
                If cU.gfBuildDataTableforGridView(dt, GridView1, True, New String() {"Y", "N"}) Then
                    Dim rows_count As Integer = 0
                    Dim rowNum As Integer = -1
                    Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

                    If Not IsNothing(gvRow) Then
                        rowNum = gvRow.RowIndex
                    End If

                    itmQty = 0
                    For i = 0 To dt.Rows.Count - 1
                        If DB.decodeDBNull(dt.Rows(i).Item("gra_pa_qty"), 0) > 0 And dt.Rows(i).Item("mFlag").ToString <> "D" And
                             dt.Rows(i).Item("gra_itm_code").ToString.Trim = dt.Rows(rowNum).Item("gra_itm_code").ToString.Trim And
                                dt.Rows(i).Item("gra_pack_key").ToString.Trim = dt.Rows(rowNum).Item("gra_pack_key").ToString.Trim And
                                    dt.Rows(i).Item("gra_pallet_no").ToString.Trim = dt.Rows(rowNum).Item("gra_pallet_no").ToString.Trim And
                                    dt.Rows(i).Item("GRA_BATCH_NO").ToString.Trim = dt.Rows(rowNum).Item("GRA_BATCH_NO").ToString.Trim Then

                            itmQty = itmQty + dt.Rows(i).Item("gra_pa_qty")
                        End If
                    Next

                    If itmQty >= CDbl(CType(gvRow.FindControl("rcv_qty"), Label).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Put away qty is greater than or equal to received qty, no splitting allowed!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "儲存數量高於或等於領取數量, 不能分拆!", Session("gLang"))
                        End If
                        Exit Sub
                    End If

                    'dt.Rows.InsertAt()

                    Dim newRow As DataRow

                    newRow = dt.NewRow
                    dt.Rows.InsertAt(newRow, rowNum + 1)

                    Session("_M_IB_GR_TMP_pa_seq") = CStr(CInt(Session("_M_IB_GR_TMP_pa_seq")) + 1)

                    'dt.Rows.Add()
                    rows_count = dt.Rows.Count

                    dt.Rows(rowNum + 1).Item("gra_seq") = Session("_M_IB_GR_TMP_pa_seq").ToString
                    'dt.Rows(rowNum + 1).Item("gra_disp_seq") = Session("_M_IB_GR_TMP_pa_seq").ToString
                    dt.Rows(rowNum + 1).Item("gra_disp_seq") = dt.Rows(rowNum).Item("gra_disp_seq")
                    dt.Rows(rowNum + 1).Item("gra_itm_code") = dt.Rows(rowNum).Item("gra_itm_code")
                    dt.Rows(rowNum + 1).Item("dsp_itm_name") = dt.Rows(rowNum).Item("dsp_itm_name")
                    dt.Rows(rowNum + 1).Item("gra_pack_key") = dt.Rows(rowNum).Item("gra_pack_key")
                    dt.Rows(rowNum + 1).Item("gra_pallet_no") = dt.Rows(rowNum).Item("gra_pallet_no")
                    dt.Rows(rowNum + 1).Item("GRA_BATCH_NO") = dt.Rows(rowNum).Item("GRA_BATCH_NO")
                    dt.Rows(rowNum + 1).Item("gra_ref_no") = dt.Rows(rowNum).Item("gra_ref_no")

                    dt.Rows(rowNum + 1).Item("itm_sku_no") = dt.Rows(rowNum).Item("itm_sku_no")

                    'dt.Rows(rowNum + 1).Item("GRA_BATCH_NO") = dt.Rows(rowNum).Item("GRA_BATCH_NO")
                    If CDbl(CType(gvRow.FindControl("rcv_qty"), Label).Text) > itmQty Then
                        dt.Rows(rowNum + 1).Item("gra_pa_qty") = CDbl(CType(gvRow.FindControl("rcv_qty"), Label).Text) - itmQty
                    Else
                        dt.Rows(rowNum + 1).Item("gra_pa_qty") = 0
                    End If
                    dt.Rows(rowNum + 1).Item("gra_sug_qty") = 0
                    dt.Rows(rowNum + 1).Item("gra_split_fr") = dt.Rows(rowNum).Item("gra_split_fr")
                    dt.Rows(rowNum + 1).Item("gra_wh") = ""
                    dt.Rows(rowNum + 1).Item("gra_loc") = ""

                    dt.Rows(rowNum + 1).Item("gra_rej_qty") = 0
                    dt.Rows(rowNum + 1).Item("gra_rej_reason") = ""
                    dt.Rows(rowNum + 1).Item("gra_remark") = ""

                    dt.Rows(rowNum + 1).Item("mFlag") = "N"



                    'dt.Rows.Add()
                    'rows_count = dt.Rows.Count

                    'dt.Rows(rows_count - 1).Item("gra_seq") = Session("_M_IB_GR_TMP_pa_seq").ToString
                    ''dt.Rows(rows_count - 1).Item("gra_disp_seq") = Session("_M_IB_GR_TMP_pa_seq").ToString
                    'dt.Rows(rows_count - 1).Item("gra_disp_seq") = dt.Rows(rowNum).Item("gra_disp_seq")
                    'dt.Rows(rows_count - 1).Item("gra_itm_code") = dt.Rows(rowNum).Item("gra_itm_code")
                    'dt.Rows(rows_count - 1).Item("dsp_itm_name") = dt.Rows(rowNum).Item("dsp_itm_name")
                    'dt.Rows(rows_count - 1).Item("gra_pack_key") = dt.Rows(rowNum).Item("gra_pack_key")
                    'dt.Rows(rows_count - 1).Item("gra_pallet_no") = dt.Rows(rowNum).Item("gra_pallet_no")
                    'If CDbl(CType(gvRow.FindControl("rcv_qty"), Label).Text) > itmQty Then
                    '    dt.Rows(rows_count - 1).Item("gra_pa_qty") = CDbl(CType(gvRow.FindControl("rcv_qty"), Label).Text) - itmQty
                    'Else
                    '    dt.Rows(rows_count - 1).Item("gra_pa_qty") = 0
                    'End If
                    'dt.Rows(rows_count - 1).Item("gra_split_fr") = dt.Rows(rowNum).Item("gra_split_fr")
                    'dt.Rows(rows_count - 1).Item("gra_wh") = ""
                    'dt.Rows(rows_count - 1).Item("gra_loc") = ""

                    'dt.Rows(rows_count - 1).Item("mFlag") = "N"

                    dt.AcceptChanges()

                    GridView1.DataSource = dt
                    GridView1.DataBind()
                End If
            End If
        ElseIf e.CommandName = "OPEN_RESERV_LOOKUP" Then
            Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)

            'DS_ITEM_CODE.Text = CType(gvRow.FindControl("grd_itm_code"), Label).Text
            'DS_PACK_KEY.Text = CType(gvRow.FindControl("grd_pack_key"), Label).Text
            'DS_SKU_NO.Text = CType(gvRow.FindControl("itm_sku_no"), Label).Text
            'DS_ITEM_NAME.Text = CType(gvRow.FindControl("grd_itm_name"), Label).Text

            tarLocID.Value = CType(gvRow.FindControl("gra_loc"), DropDownList).ClientID
            tarWhID.Value = CType(gvRow.FindControl("gra_wh"), DropDownList).ClientID

            'tarDspLocID.Value = CType(gvRow.FindControl("dsp_gra_loc"), Label).ClientID
            'sn_dt = Session("sn_dt")

            selectSql = "select h.LOBH_CODE, d.LOBD_LOC, d.LOBD_SKU_NO, d.PACK_KEY, d.LOBD_ITM_DESC, d.LOBD_CBM, d.LOBD_CBM_PERC, d.LOBD_STATUS " &
                        "from wms_loc_book_hd h, wms_loc_book_dtl d " &
                        "where h.imp_code = d.imp_code " &
                        "and h.storer_code = d.storer_code " &
                        "and h.lobh_code = d.lobh_code " &
                        "and h.lobh_status = 'NEW' " &
                        "and h.lobh_ro_code = '" & gU.dbEncode(GR_DOC_NO.Value) & "' " &
                        "and d.lobd_sku_no = '" & gU.dbEncode(CType(gvRow.FindControl("itm_sku_no"), Label).Text) & "' " &
                        "and d.pack_key = '" & gU.dbEncode(CType(gvRow.FindControl("gra_pack_key"), Label).Text) & "' " &
                        "order by h.LOBH_CODE, d.LOBD_LOC, d.LOBD_STATUS, d.LOBD_CBM, d.LOBD_CBM_PERC "

            tmpDt = gDB.getDataTable(selectSql)

            If tmpDt.Rows.Count <= 0 Then

                selectSql = "select h.LOBH_CODE, d.LOBD_LOC, d.LOBD_SKU_NO, d.PACK_KEY, d.LOBD_ITM_DESC, d.LOBD_CBM, d.LOBD_CBM_PERC, d.LOBD_STATUS " &
                            "from wms_loc_book_hd h, wms_loc_book_dtl d " &
                            "where h.imp_code = d.imp_code " &
                            "and h.storer_code = d.storer_code " &
                            "and h.lobh_code = d.lobh_code " &
                            "and h.lobh_status = 'NEW' " &
                            "and h.lobh_ro_code is null " &
                            "and d.lobd_sku_no = '" & gU.dbEncode(CType(gvRow.FindControl("itm_sku_no"), Label).Text) & "' " &
                            "and d.pack_key = '" & gU.dbEncode(CType(gvRow.FindControl("gra_pack_key"), Label).Text) & "' " &
                            "order by h.LOBH_CODE, d.LOBD_LOC, d.LOBD_STATUS, d.LOBD_CBM, d.LOBD_CBM_PERC "

            End If

            '"and d.lobd_itm_code = '" & gU.dbEncode(CType(gvRow.FindControl("gra_itm_code"), Label).Text) & "' " & _

            tmpDt = gDB.getDataTable(selectSql)

            gvResrvLocList.DataSource = tmpDt
            gvResrvLocList.DataBind()

            updtPnl_ResrvLocList.Update()
            ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "OPEN_RESERV_LOOKUP", "$find('resrvLoc_behavior').show();", True)

        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))

        If dt.Rows(e.RowIndex).Item("mFlag") = "N" Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True, New String() {"Y", "N"}) Then
                If dt.Rows(e.RowIndex).Item("gra_seq") = Session("_M_IB_GR_TMP_pa_seq").ToString Then
                    Session("_M_IB_GR_TMP_pa_seq") = CStr(CInt(Session("_M_IB_GR_TMP_pa_seq")) - 1)
                End If
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

    Protected Sub gvResrvLocList_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvResrvLocList.RowCommand
        If e.CommandName = "SEL_LOC" Then
            Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)
            Dim jsStr As String

            jsStr = "document.getElementById('" & tarLocID.Value & "').value = '" & CType(gvRow.FindControl("LOBD_LOC_HDN"), HiddenField).Value & "'; " &
                    "document.getElementById('" & tarDspLocID.Value & "').innerHTML = '" & CType(gvRow.FindControl("LOBD_LOC_HDN"), HiddenField).Value & "'; " &
                    "$find('resrvLoc_behavior').hide();"


            ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "CLOSE_RESERV_LOOKUP", jsStr, True)

        End If
    End Sub

    Protected Sub gvResrvLocList_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvResrvLocList.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("LOBD_LOC_HDN"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "LOBD_LOC").ToString.Trim

        End Select
    End Sub

    'Private Sub updtLabelFields()
    '    Dim i As Integer

    '    dsp_app_loc.Text = app_loc.Value

    '    If GridView1.Rows.Count > 0 Then
    '        For i = 0 To GridView1.Rows.Count - 1
    '            CType(GridView1.Rows(i).FindControl("dsp_gra_loc"), Label).Text = CType(GridView1.Rows(i).FindControl("GRA_LOC"), DropDownList).Value
    '        Next
    '    End If
    'End Sub

    Private Function validateAll(Optional ByVal flag As String = "") As Boolean
        Dim i, j As Integer
        Dim javaStr As String
        Dim itmQty, itmQty2 As Double
        Dim warnItmList, warnItmList2 As String

        If GridView1.Rows.Count > 0 Then
            warnItmList = ""
            warnItmList2 = ""

            For i = 0 To GridView1.Rows.Count - 1
                If GridView1.Rows(i).Visible Then
                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("gra_pa_qty"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid number, Put Away Qty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的數字, 儲存數量!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If DocType = "IPR" AndAlso CType(GridView1.Rows(i).FindControl("gra_wh"), DropDownList).SelectedItem.Value <> "FGTM" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid warehouse, Please select FGTM warehouse!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "Invalid warehouse, Please select FGTM warehouse!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If DocType = "IPR" AndAlso CType(GridView1.Rows(i).FindControl("gra_remark"), AjaxControlToolkit.ComboBox).SelectedValue.Length > 10 Then
                        Dim IsValidLocation As Boolean = 0
                        'check for valid location
                        Dim selectSqlLoc As String = "select 1 from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' and WH_CODE = '" & gU.dbEncode(CType(GridView1.Rows(i).FindControl("gra_wh"), DropDownList).SelectedItem.Value) & "'"
                        Dim tmpDtLoc As DataTable = gDB.getDataTable(selectSqlLoc)

                        If tmpDtLoc IsNot Nothing AndAlso tmpDtLoc.Rows.Count > 0 Then
                            Dim RemarkStr = CType(GridView1.Rows(i).FindControl("gra_remark"), AjaxControlToolkit.ComboBox).SelectedValue
                            selectSqlLoc = "select 1 from WMS_WH_BIN where WH_CODE='" & gU.dbEncode(RemarkStr.Remove(RemarkStr.Length - 8)) & "' and LOC_KEY='" & gU.dbEncode(Right(RemarkStr, 8)) & "' and IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "'"
                            tmpDtLoc = gDB.getDataTable(selectSqlLoc)

                            If tmpDtLoc IsNot Nothing AndAlso tmpDtLoc.Rows.Count > 0 Then
                                IsValidLocation = 1
                            End If
                        End If

                        If IsValidLocation = 0 Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Invalid Location for Stock Transfer!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "Invalid Location for Stock Transfer!", Session("gLang"))
                            End If
                            Return False
                        End If

                    ElseIf DocType = "IPR" AndAlso CType(GridView1.Rows(i).FindControl("gra_remark"), AjaxControlToolkit.ComboBox).SelectedValue.Length < 10 Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Please select final location for Stock Transfer!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "Please select final location for Stock Transfer!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("gra_rej_qty"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid number, Rejected Qty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的數字, 拒收數量!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If Not gU.isDecimal(CType(GridView1.Rows(i).FindControl("gra_sug_qty"), TextBox).Text) Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Invalid number, Suggested Qty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "無效的數字, 建議數量!", Session("gLang"))
                        End If
                        Return False
                    End If

                    If CType(GridView1.Rows(i).FindControl("gra_rej_qty"), TextBox).Text.Trim <> "" AndAlso CDbl(CType(GridView1.Rows(i).FindControl("gra_rej_qty"), TextBox).Text.Trim) <> 0 Then
                        If CType(GridView1.Rows(i).FindControl("rcv_qty"), Label).Text = "" OrElse CDbl(CType(GridView1.Rows(i).FindControl("gra_rej_qty"), TextBox).Text.Trim) > CDbl(CType(GridView1.Rows(i).FindControl("rcv_qty"), Label).Text) Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Rejected Qty cannot greater than RCV Qty!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "拒收數量不能大於收貨數量!", Session("gLang"))
                            End If
                            Return False
                        End If

                        If CType(GridView1.Rows(i).FindControl("gra_rej_reason"), DropDownList).SelectedValue = "" Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Please select Reject Reason!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "請選擇拒收原因!", Session("gLang"))
                            End If
                            Return False
                        End If
                    End If

                End If

                If flag <> "N" AndAlso flag <> "L" AndAlso Request("moduleAction") <> "SAVEOK" Then
                    If CType(GridView1.Rows(i).FindControl("gra_pa_qty"), TextBox).Text.Trim <> "" OrElse CType(GridView1.Rows(i).FindControl("gra_sug_qty"), TextBox).Text.Trim <> "" Then
                        itmQty = 0
                        itmQty2 = 0
                        For j = 0 To dt.Rows.Count - 1
                            If GridView1.Rows(j).Visible AndAlso
                                CType(GridView1.Rows(i).FindControl("gra_itm_code"), Label).Text = CType(GridView1.Rows(j).FindControl("gra_itm_code"), Label).Text AndAlso
                                CType(GridView1.Rows(i).FindControl("gra_pack_key"), Label).Text = CType(GridView1.Rows(j).FindControl("gra_pack_key"), Label).Text AndAlso
                                CType(GridView1.Rows(i).FindControl("gra_pallet_no"), Label).Text = CType(GridView1.Rows(j).FindControl("gra_pallet_no"), Label).Text AndAlso
                                CType(GridView1.Rows(i).FindControl("GRA_BATCH_NO"), Label).Text = CType(GridView1.Rows(j).FindControl("GRA_BATCH_NO"), Label).Text Then

                                If CType(GridView1.Rows(j).FindControl("gra_pa_qty"), TextBox).Text.Trim <> "" Then
                                    itmQty = itmQty + CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(j).FindControl("gra_pa_qty"), TextBox).Text, "0"))
                                End If

                                If CType(GridView1.Rows(j).FindControl("gra_sug_qty"), TextBox).Text.Trim <> "" Then
                                    itmQty2 = itmQty2 + CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(j).FindControl("gra_sug_qty"), TextBox).Text, "0"))
                                End If

                            End If
                        Next

                        'If CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("gra_pa_qty"), TextBox).Text, "0")) > CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("rcv_qty"), Label).Text, "0")) Then
                        If itmQty > CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("rcv_qty"), Label).Text, "0")) Then
                            warnItmList = gU.appendToList(warnItmList, CType(GridView1.Rows(i).FindControl("gra_itm_code"), Label).Text)
                            'If Session("gLang") = "E" Then
                            '    javaStr = "if(confirm(""Put away qty is greater than received qty, confirm to proceed?"")) saveok();"
                            'Else
                            '    javaStr = "if(confirm(""儲存數量高於領取數量, 確定輸入?"")) saveok();"
                            'End If
                            'If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                            '    Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
                            'End If
                            'Return False
                        End If

                        If itmQty2 > CDbl(gU.decodeNullOrEmpty(CType(GridView1.Rows(i).FindControl("rcv_qty"), Label).Text, "0")) Then
                            warnItmList2 = gU.appendToList(warnItmList2, CType(GridView1.Rows(i).FindControl("gra_itm_code"), Label).Text)
                        End If
                    End If
                End If
            Next

            If warnItmList <> "" Then
                If Session("gLang") = "E" Then
                    javaStr = "if(confirm(""Item: [" & warnItmList & "] Put away qty is greater than received qty, confirm to proceed?"")) saveok();"
                Else
                    javaStr = "if(confirm(""物件: [" & warnItmList & "] 儲存數量高於領取數量, 確定輸入?"")) saveok();"
                End If
                If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                    Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
                End If
                Return False
            End If

            If warnItmList <> "" Then
                If Session("gLang") = "E" Then
                    javaStr = "if(confirm(""Item: [" & warnItmList & "] Suggested qty is greater than received qty, confirm to proceed?"")) saveok();"
                Else
                    javaStr = "if(confirm(""物件: [" & warnItmList & "] 建議數量高於領取數量, 確定輸入?"")) saveok();"
                End If
                If (Not Me.ClientScript.IsStartupScriptRegistered(Me.GetType, "")) Then
                    Me.ClientScript.RegisterStartupScript(Me.GetType, "confirm", javaStr, True)
                End If
                Return False
            End If
        End If

        Return True
    End Function

    Protected Sub save()
        If validateAll() Then
            If cU.gfBuildDataTableforGridView(dt, GridView1, True, New String() {"Y", "N"}) Then
                'uiFun.reOrderDetails(dt, "gra_disp_seq")
                Session("_M_IB_GR_TMP_pa_dt") = dt
                Me.ClientScript.RegisterClientScriptBlock(Me.GetType, "close", "<script language=""JavaScript"">window.open('','_self');window.close();</script>")
            End If
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
                itmKey = dt.Rows(i).Item("gra_itm_code").ToString.Trim & "#_#" &
                        dt.Rows(i).Item("gra_pack_key").ToString.Trim & "#_#" &
                        dt.Rows(i).Item("gra_pallet_no").ToString.Trim & "#_#" &
                        dt.Rows(i).Item("GRA_BATCH_NO").ToString.Trim

                If CDbl(gU.decodeNullOrEmpty(DB.decodeDBNull(dt.Rows(i).Item("gra_pa_qty"), 0), 0)) > 0 Then
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
                If CInt(dt.Rows(delList(i)).Item("PLD_SEQ")) = CInt(Session("_M_IB_GR_TMP_pa_seq")) Then
                    Session("_M_IB_GR_TMP_pa_seq") = CStr(CInt(Session("_M_IB_GR_TMP_pa_seq")) - 1)
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
        Dim sort_col As String = "GRA_DISP_SEQ"

        IMP_CODE.Value = Server.UrlDecode(Request("IMP_CODE"))
        STORER_CODE.Value = Server.UrlDecode(Request("STORER_CODE"))
        GR_DOC_NO.Value = Server.UrlDecode(Request("GR_DOC_NO"))
        GR_DOC_TYPE.Value = Server.UrlDecode(Request("GR_DOC_TYPE"))
        'delZeroRow()

        sortDt = dt.Copy

        Dim foundRows As DataRow() = sortDt.Select("", sort_col)

        dt.Rows.Clear()
        For i As Integer = 0 To foundRows.Count - 1
            dt.ImportRow(foundRows(i))
        Next
        dt.AcceptChanges()
        'Added by Ashwin 20-Nov-2020
        'Dim tmpArr = dt.Select("GRA_PA_QTY > 0")
        'If tmpArr IsNot Nothing AndAlso tmpArr.Length > 0 Then
        '    dt = tmpArr.CopyToDataTable()
        'End If
        'Added by Ashwin 20-Nov-2020
        GridView1.DataSource = dt
        GridView1.DataBind()
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub btnSelOtherLoc_Click(sender As Object, e As System.EventArgs) Handles btnSelOtherLoc.Click
        Dim jsStr As String

        jsStr = "LocLookUp(document.getElementById('" & HttpUtility.HtmlEncode(tarWhID.Value) & "').value, " &
                            "'" & HttpUtility.HtmlEncode(tarDspLocID.Value) & "', " &
                            "'" & HttpUtility.HtmlEncode(tarLocID.Value) & "', " &
                            "'" & HttpUtility.HtmlEncode(tarWhID.Value) & "');" &
                "$find('resrvLoc_behavior').hide();"

        ScriptManager.RegisterStartupScript(updtPnlAlert, updtPnlAlert.GetType, "CLOSE_RESERV_LOOKUP", jsStr, True)

    End Sub

    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Dim dtl_dt As DataTable
        Dim paTable As PutAwayTable

        dtl_dt = Session("dt")

        If dtl_dt.Rows.Count > 0 Then
            paTable = New PutAwayTable(dtl_dt, dt, Session("_M_IB_GR_TMP_pa_seq"), Session("_M_IB_GR_TMP_pa_del_list"))

            paTable.genPutAway(True)

            Session("_M_IB_GR_TMP_pa_del_list") = paTable.putAwayDelList

            Session("_M_IB_GR_TMP_pa_seq") = paTable.putAwaySeq

            GridView1.DataSource = paTable.putAwayDataTable
            GridView1.DataBind()
        End If


        'dtl_dt = Session("dt")

        'If dtl_dt.Rows.Count > 0 Then
        '    genPutAway(Session("dt"), dt, True)

        '    GridView1.DataSource = dt
        '    GridView1.DataBind()
        'End If



        'Dim dtl_dt As DataTable
        'Dim rows_count As Integer
        'Dim i As Integer
        'Dim plSeq As Integer
        'Dim itmDict As Dictionary(Of String, itmStruct)
        'Dim itmKey As String
        'Dim keys As Dictionary(Of String, itmStruct).KeyCollection
        'Dim keyArray As String()
        'Dim tmpStruct As itmStruct

        'For i = dt.Rows.Count - 1 To 0 Step -1
        '    If dt.Rows(i).Item("mFlag") = "N" Then
        '        dt.Rows(i).Delete()
        '    Else
        '        Session("_M_IB_GR_TMP_pa_del_list") = gU.appendToList(Session("_M_IB_GR_TMP_pa_del_list"), dt.Rows(i).Item("gra_seq"))
        '        dt.Rows(i).Delete()
        '    End If
        'Next
        'dt.AcceptChanges()

        'dtl_dt = Session("dt")

        'Session("_M_IB_GR_TMP_pa_seq") = "0"

        'plSeq = CInt(Session("_M_IB_GR_TMP_pa_seq"))

        'If dtl_dt.Rows.Count > 0 Then
        '    itmDict = New Dictionary(Of String, itmStruct)

        '    For i = 0 To dtl_dt.Rows.Count - 1


        '        If DB.decodeDBNull(dtl_dt.Rows(i).Item("GRD_RCV_QTY"), 0) > 0 And dtl_dt.Rows(i).Item("mFlag").ToString <> "D" Then
        '            itmKey = dtl_dt.Rows(i).Item("grd_itm_code").ToString & "#_#" & _
        '                        dtl_dt.Rows(i).Item("grd_pack_key").ToString & "#_#" & _
        '                        dtl_dt.Rows(i).Item("grd_pallet_no").ToString & "#_#" & _
        '                        dtl_dt.Rows(i).Item("GRD_BATCH_NO").ToString

        '            'If Not itmDict.ContainsKey(itmKey) Then
        '            '    itmDict.Add(itmKey, dtl_dt.Rows(i).Item("GRD_RCV_QTY"))
        '            'Else
        '            '    itmDict.Item(itmKey) = itmDict.Item(itmKey) + dtl_dt.Rows(i).Item("GRD_RCV_QTY")
        '            'End If

        '            If Not itmDict.ContainsKey(itmKey) Then
        '                'itmDict.Add(itmKey, dtl_dt.Rows(i).Item("GRD_RCV_QTY"))
        '                tmpStruct = New itmStruct
        '                tmpStruct.itmName = gU.decodeNullOrEmpty(dtl_dt.Rows(i).Item("GRD_ITM_NAME").ToString, "")
        '                tmpStruct.itmSeq = gU.decodeNullOrEmpty(dtl_dt.Rows(i).Item("GRD_SEQ").ToString, "")
        '                tmpStruct.qty = gU.decodeEmptyCdbl(dtl_dt.Rows(i).Item("GRD_RCV_QTY"), 0)
        '                tmpStruct.batchNo = gU.decodeNullOrEmpty(dtl_dt.Rows(i).Item("GRD_BATCH_NO").ToString, "")
        '                tmpStruct.skuNo = gU.decodeNullOrEmpty(dtl_dt.Rows(i).Item("ITM_SKU_NO").ToString, "")

        '                tmpStruct.expDate = dtl_dt.Rows(i).Item("GRD_EXPIRY_DATE").ToString.Trim
        '                tmpStruct.manuDate = dtl_dt.Rows(i).Item("GRD_MANU_DATE").ToString.Trim

        '                itmDict.Add(itmKey, tmpStruct)
        '            Else
        '                'itmDict.Item(itmKey) = itmDict.Item(itmKey) + dtl_dt.Rows(i).Item("GRD_RCV_QTY")
        '                tmpStruct = itmDict.Item(itmKey)
        '                tmpStruct.qty = tmpStruct.qty + gU.decodeEmptyCdbl(dtl_dt.Rows(i).Item("GRD_RCV_QTY"), 0)

        '                itmDict(itmKey) = tmpStruct 
        '            End If
        '        End If
        '    Next

        '    keys = itmDict.Keys

        '    If keys.Count > 0 Then
        '        For i = 0 To keys.Count - 1
        '            dt.Rows.Add()

        '            plSeq = plSeq + 1

        '            rows_count = dt.Rows.Count

        '            keyArray = Split(keys(i), "#_#")

        '            dt.Rows(rows_count - 1).Item("GRA_SEQ") = CStr(plSeq)
        '            dt.Rows(rows_count - 1).Item("GRA_DISP_SEQ") = CStr(plSeq)
        '            dt.Rows(rows_count - 1).Item("GRA_ITM_CODE") = keyArray(0)
        '            dt.Rows(rows_count - 1).Item("DSP_ITM_NAME") = itmDict.Item(keys(i)).itmName
        '            dt.Rows(rows_count - 1).Item("GRA_PACK_KEY") = keyArray(1)
        '            dt.Rows(rows_count - 1).Item("GRA_PALLET_NO") = keyArray(2)
        '            dt.Rows(rows_count - 1).Item("GRA_REF_NO") = ""
        '            dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = itmDict.Item(keys(i)).skuNo

        '            'dt.Rows(rows_count - 1).Item("GRA_PA_QTY") = itmDict.Item(keys(i)).qty
        '            dt.Rows(rows_count - 1).Item("GRA_PA_QTY") = 0
        '            dt.Rows(rows_count - 1).Item("GRA_SPLIT_FR") = itmDict.Item(keys(i)).itmSeq
        '            dt.Rows(rows_count - 1).Item("GRA_WH") = ""
        '            dt.Rows(rows_count - 1).Item("GRA_LOC") = ""
        '            dt.Rows(rows_count - 1).Item("mFlag") = "N"

        '            dt.Rows(rows_count - 1).Item("GRA_BATCH_NO") = keyArray(3)

        '            dt.Rows(rows_count - 1).Item("GRA_EXPIRY_DATE") = itmDict.Item(keys(i)).expDate
        '            dt.Rows(rows_count - 1).Item("GRA_MANU_DATE") = itmDict.Item(keys(i)).manuDate

        '            dt.Rows(rows_count - 1).Item("GRA_REJ_QTY") = 0
        '            dt.Rows(rows_count - 1).Item("GRA_REJ_REASON") = ""
        '        Next

        '        Session("_M_IB_GR_TMP_pa_seq") = CStr(plSeq)

        '        dt.AcceptChanges()
        '    End If

        '    GridView1.DataSource = dt
        '    GridView1.DataBind()
        'End If
    End Sub

    Protected Sub gra_wh_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim ddlWH As DropDownList = CType(sender, DropDownList)
        Dim Row = CType(ddlWH.Parent.Parent, GridViewRow)
        Dim idx = Row.RowIndex
        If ddlWH.SelectedValue <> Nothing Then
            uiFun.load_ComboBox(CType(Row.FindControl("gra_loc"), AjaxControlToolkit.ComboBox),
                                    "Select Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as CODE,Convert(nvarchar(20),a.FL_NUM)+a.AR_CODE+a.RK_CODE+a.BN_CODE as NAME from WMS_WH_BIN a where a.WH_CODE='" & ddlWH.SelectedValue & "' and a.IMP_CODE='" & gU.dbEncode(Session("IMP_CODE")) & "'", "CODE", "NAME", , Session("gSelectLabel"))

        End If
    End Sub


End Class
