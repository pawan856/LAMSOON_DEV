Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OUTBOUND_WAVE_PICK
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private st As New StockTrans
    Private maxPLNO As String = "8"
    Private DDFORMAT As String = "DD/MM/YYYY"

    Private moduleAction As String = ""
    Private dt As New DataTable
    Private exceptionEditList As List(Of String)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DDFORMAT = gU.getConfig("DDFORMATNO")
        REM ****************************
        REM Modify Access Right Here
        ar = New AccessRightUtils(Session("PAGE_SESSION_MENU_CODE"), Session("usr_id"), Me)

        moduleAction = Request("moduleAction")

        If ar.sessionExpired = "Y" Then
            ar.Force_PageEndCtrlClear(Me, False)
            Me.Visible = False
            Exit Sub
        End If

        If Session("PAGE_SESSION_MENU_CODE") Is Nothing Then
            Exit Sub
        End If

        If Not IsPostBack Then
            ViewState("pagemode") = Nothing
            ViewState("pagemode") = Request("mode")
            Session("WPpkListDT") = Nothing
            Session("selectedDO") = ""

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(WP_WH_CODE, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"))

        End If

        If ViewState("pagemode") = "N" Then
            CancelBtn.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If

            If WP_DATE.Text = "" Then
                WP_DATE.Text = Now.Date.ToString("dd/MM/yyyy")
            End If
            btnPrint.Visible = False
        Else
            btnPrint.Visible = True
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Wave Picking Maintenance"
            lbl_ImageHd.Text = "Pick Lists:"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            newrow.Text = "Add"

            selectItemBtn.Text = "Select Item"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"

            If ViewState("pagemode") = "N" Then
                WP_NO.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "Wave Picking Maintenance"
            lbl_ImageHd.Text = "Pick Lists:"


            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "儲存"
            saveBtn2.Text = "儲存"
            CancelBtn.Text = "取消"
            newrow.Text = "新增"

            selectItemBtn.Text = "選擇物料"
            CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"

            If ViewState("pagemode") = "N" Then
                WP_NO.Text = "[號碼會自動產生]"
            End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        WP_WH_CODE.CssClass = "REQUIRED"
        'AD_DATE.CssClass = "REQUIRED"
        REM **********************

        If ViewState("pagemode") = "N" Then
            'WP_NO.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("WP_NO") = ""
            ViewState("WP_NO") = Server.UrlDecode(Request("WP_NO"))
            ViewState("STORER_CODE") = ""
            ViewState("STORER_CODE") = Server.UrlDecode(Request("STORER_CODE"))
            Call BindGV()
        Else
            dt = Session("WPpkListDT")
        End If

        If moduleAction = "SELECTDO" Then
            addItemtoPLIST()
        ElseIf moduleAction = "RELOADPL" Then
            Dim tempDT As DataTable = Session("WPpkListDT")
            SortPLList(tempDT)
            tempDT.AcceptChanges()

            Session("WPpkListDT") = tempDT
            GridView1.DataSource = tempDT
            GridView1.DataBind()
        End If

        selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value, document.myform." & WP_WH_CODE.ClientID & ".value);")

        If WP_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
        ElseIf WP_STATUS.Text = "POSTED" Then
            ar.sec_viewMode = "Y"
        ElseIf WP_STATUS.Text = "ASSIGNED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
        End If

        setPageCtrlAccess()

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)
    End Sub

    Private Sub setPageCtrlAccess()

        exceptionEditList = New List(Of String)
        exceptionEditList.Add("btnPrint")

    End Sub

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

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                'Dim oGridView As GridView = DirectCast(sender, GridView)
                'Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                'REM **********************
                'REM Use for re-create the label to change the Langauge
                'REM Modify Here
                'Call cU.changeGVLabel(oGridViewRow, e, "No.", "編號")
                'Call cU.changeGVLabel(oGridViewRow, e, "Item Code.", "物料編號")
                'Call cU.changeGVLabel(oGridViewRow, e, "SKU No.", "SKU No.")
                'Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內碼")
                'Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物料名稱")
                'Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "Batch No.")
                'Call cU.changeGVLabel(oGridViewRow, e, "Loc.", "位置")
                'Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板")
                'Call cU.changeGVLabel(oGridViewRow, e, "Orig. Qty", "原來數量")
                'Call cU.changeGVLabel(oGridViewRow, e, "Revised Qty", "修訂數量")
                'Call cU.changeGVLabel(oGridViewRow, e, "Variance Qty", "差異數量")
                'Call cU.changeGVLabel(oGridViewRow, e, "Remarks", "備註")
                'Call cU.changeGVLabel(oGridViewRow, e, "Expiry Date", "Expiry Date")
                'Call cU.changeGVLabel(oGridViewRow, e, "ManuFactory Date", "ManuFactory Date")
                'Call cU.changeGVLabel(oGridViewRow, e, "", "")
                'REM **********************

                'oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim nDropDown As DropDownList

        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                CType(e.Row.FindControl("pld_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_SEQ").ToString.Trim
                'CType(e.Row.FindControl("do_co_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "do_co_code").ToString.Trim
                CType(e.Row.FindControl("pld_picked_by"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_PICKED_BY").ToString.Trim
                CType(e.Row.FindControl("pld_item_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_NO").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ITM_SKU_NO").ToString.Trim
                CType(e.Row.FindControl("pld_pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_PACK_KEY").ToString.Trim
                CType(e.Row.FindControl("pld_pallet_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_PALLET_NO").ToString.Trim
                CType(e.Row.FindControl("pld_foi_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_FOI_QTY").ToString.Trim
                CType(e.Row.FindControl("pld_item_qty"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_QTY").ToString.Trim
                CType(e.Row.FindControl("pld_do_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pld_do_qty").ToString.Trim

                CType(e.Row.FindControl("ITM_WEIGHT_TYPE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ITM_WEIGHT_TYPE").ToString.Trim

                CType(e.Row.FindControl("stock_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "stock_qty").ToString.Trim

                CType(e.Row.FindControl("iloc_bal_qty"), Label).Text = DataBinder.Eval(e.Row.DataItem, "ILOC_BAL_QTY").ToString.Trim

                'CType(e.Row.FindControl("DO_EDI_SIR_NO"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DO_EDI_SIR_NO").ToString.Trim

                CType(e.Row.FindControl("DO_CO_CODE"), Label).Text = DataBinder.Eval(e.Row.DataItem, "DO_CO_CODE").ToString.Trim

                CType(e.Row.FindControl("do_code"), HyperLink).Text = DataBinder.Eval(e.Row.DataItem, "do_code").ToString.Trim
                CType(e.Row.FindControl("do_code"), HyperLink).NavigateUrl = "#"
                CType(e.Row.FindControl("do_code"), HyperLink).Attributes.Add("onclick", "OpenDO('" & DataBinder.Eval(e.Row.DataItem, "do_code").ToString.Trim & "');")


                'If CDbl(DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim) > 0 Then
                '    CType(e.Row.FindControl("hold_qty_desc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim & " / " & DataBinder.Eval(e.Row.DataItem, "TOTAL_BAL").ToString.Trim
                '    CType(e.Row.FindControl("avail_qty"), HiddenField).Value = CDbl(DataBinder.Eval(e.Row.DataItem, "TOTAL_BAL").ToString.Trim) - CDbl(DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim)
                'Else
                '    CType(e.Row.FindControl("hold_qty_desc"), Label).Text = ""
                '    CType(e.Row.FindControl("avail_qty"), HiddenField).Value = 0
                'End If
                'CType(e.Row.FindControl("hold_qty"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "HOLD_QTY").ToString.Trim
                'CType(e.Row.FindControl("total_bal"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "TOTAL_BAL").ToString.Trim

                nDropDown = CType(e.Row.FindControl("pld_status"), DropDownList)
                uiFun.load_dropdown(nDropDown, "SELECT WAVE_STATUS_CODE As CODE,WAVE_STATUS_DESC as NAME FROM WMS_STATUS_MASTER Where WAVE_STATUS_TYPE='WAVEPICK'")
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "pld_status").ToString.Trim

                CType(e.Row.FindControl("pld_wh"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_WH").ToString.Trim
                CType(e.Row.FindControl("dsp_pld_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_LOC").ToString.Trim
                CType(e.Row.FindControl("pld_loc"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_LOC").ToString.Trim
                CType(e.Row.FindControl("dsp_pld_floor"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_FLOOR").ToString.Trim
                CType(e.Row.FindControl("pld_floor"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_FLOOR").ToString.Trim
                CType(e.Row.FindControl("dsp_pld_area"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_AREA").ToString.Trim
                CType(e.Row.FindControl("pld_area"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_AREA").ToString.Trim
                CType(e.Row.FindControl("dsp_pld_rack"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_RACK").ToString.Trim
                CType(e.Row.FindControl("pld_rack"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_RACK").ToString.Trim
                CType(e.Row.FindControl("dsp_pld_bin"), Label).Text = DataBinder.Eval(e.Row.DataItem, "PLD_BIN").ToString.Trim
                CType(e.Row.FindControl("pld_bin"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "PLD_BIN").ToString.Trim


                CType(e.Row.FindControl("pld_batch_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pld_batch_no").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "PLD_IS_LOAN").ToString.Trim = "Y" Then
                    CType(e.Row.FindControl("pld_is_loan"), CheckBox).Checked = True
                End If

                Dim nImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp"), Image)
                nImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(nImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                nImage.Attributes.Add("onclick", "LocLookUp('" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "PLD_ITEM_NO").ToString.Trim) & "', " &
                                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "PLD_PACK_KEY").ToString.Trim) & "', " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_pallet_no"), HiddenField).ClientID) & "').value, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("pld_batch_no"), Label).ClientID) & "').innerHTML, " &
                                      "document.getElementById('" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("itm_sku_no"), Label).ClientID) & "').innerHTML," &
                                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "DO_CODE").ToString.Trim) & "'," &
                                      "'" & HttpUtility.HtmlEncode(DataBinder.Eval(e.Row.DataItem, "DO_CO_CODE").ToString.Trim) & "');")



                Dim PL_LIST_NO As String = DataBinder.Eval(e.Row.DataItem, "PLD_PL_LIST_NO").ToString.Trim

                If PL_LIST_NO <> "" Then
                    If Not CType(e.Row.FindControl("PL" & PL_LIST_NO), RadioButton) Is Nothing Then CType(e.Row.FindControl("PL" & PL_LIST_NO), RadioButton).Checked = True

                    Dim PLnoCell As Integer = -1
                    Dim cellNumber As Integer

                    For cellNumber = 0 To e.Row.Cells.Count - 1 Step cellNumber + 1
                        Dim ctrl As Control
                        For Each ctrl In e.Row.Cells(cellNumber).Controls
                            If ctrl.ID = "PL" & PL_LIST_NO Then
                                PLnoCell = cellNumber
                                Exit For
                            End If
                        Next
                    Next

                    If PLnoCell <> -1 Then e.Row.Cells(PLnoCell).Style.Add("background-color", returnStatusColor(DataBinder.Eval(e.Row.DataItem, "PLD_STATUS").ToString.Trim))

                End If



                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" OrElse DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "PD" Then
                    e.Row.Visible = False
                End If

        End Select
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        'GridView1.Rows(e.RowIndex).Visible = False
        Dim plDt As DataTable = Session("WPpkListDT")

        If Not plDt Is Nothing AndAlso plDt.Rows.Count > 0 Then
            Dim dtl_do_code As String = ""

            dtl_do_code = plDt.Rows(e.RowIndex).Item("do_code")

            For i = 0 To plDt.Rows.Count - 1

                If plDt.Rows(i).Item("do_code").ToString.Trim = dtl_do_code Then
                    plDt.Rows(i).Item("mFlag") = "D"
                    GridView1.Rows(i).Visible = False
                End If

            Next
            'plDt.Rows(e.RowIndex).Item("mFlag") = "D"
            plDt.AcceptChanges()
            Session("WPpkListDT") = plDt
        End If

    End Sub

    Private Function validateAll() As Boolean
        Dim selectSql As String = ""
        Dim i As Integer


        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If WP_WH_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_WP_WH.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_WP_WH.Text & "不能空白!", Session("gLang"))
            End If
            Return False
        End If

        If WP_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Date cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "日期不能空白!", Session("gLang"))
            End If
            Return False
        ElseIf Not gU.isValidDate(WP_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid Date!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期!", Session("gLang"))
            End If
            Return False
        End If

        Dim itemCount As Integer = 0
        Dim tempDT As DataTable = Session("WPpkListDT")

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If tempDT.Rows(i).Item("mFlag").ToString.Trim <> "D" AndAlso tempDT.Rows(i).Item("mFlag").ToString.Trim <> "PD" Then
                    If CType(GridView1.Rows(i).FindControl("pld_loc"), HiddenField).Value = "" Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Location Cannot Be Empty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "位置不能空白!", Session("gLang"))
                        End If
                        Return False
                    End If

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

                    If CType(GridView1.Rows(i).FindControl("pld_foi_qty"), TextBox).Text.Trim <> "" AndAlso CDbl(CType(GridView1.Rows(i).FindControl("pld_foi_qty"), TextBox).Text) <= 0 Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "FFI Qty must be greater than zero!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "取貨數量必需大過零!", Session("gLang"))
                        End If
                        Return False
                    End If

                    itemCount += 1
                End If
            Next

            If itemCount = 0 Then
                uiFun.displayMsg(Me, "", "At least one DO must be selected for wave picking!", Session("gLang"))
                Return False
            End If
        Else
            uiFun.displayMsg(Me, "", "At least one DO must be selected for wave picking!", Session("gLang"))
            Return False
        End If

        Return True

    End Function

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim nextNo As String = ""
        Dim gConn As SqlConnection
        Dim dupSQL As String = ""
        Dim dupTbl As New DataTable
        Dim plDT As DataTable
        Dim varQty As Double = 0
        Dim successFlag As Boolean = False
        Dim pk_code As String = ""
        Dim maxSeq As Integer = 0
        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If ViewState("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("WAVEPICK", gConn, transaction)
                    'nextNo = WP_NO.Text
                    REM **********************

                    dupSQL = "select 1 from WMS_WAVE_PICK_H " &
                                "where WP_NO = '" & gU.dbEncode(nextNo) & "' " &
                                "and imp_code='" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                "and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL, gConn, transaction)

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "INSERT INTO WMS_WAVE_PICK_H " &
                                     "(IMP_CODE, STORER_CODE, WP_NO, WP_STATUS, WP_DATE, CUS_CODE, CUS_NAME, WP_WH_CODE, WP_REMARK, SYS_LUB, SYS_LUD, SYS_CB, SYS_CD) " &
                                     "VALUES ('" & gU.dbEncode(Session("IMP_CODE")) & "'," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(nextNo)) & ",'NEW'," &
                                     gU.convdbDate(gU.dbEncode(WP_DATE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_CODE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(CUS_NAME.Text.Trim)) & "," &
                                     gU.convdbNVCData(gU.dbEncode(WP_WH_CODE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(WP_REMARK.Text.Trim)) & "," &
                                     "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                        gDB.amendData(sql_string, gConn, transaction)
                        REM **********************s
                        plDT = Session("WPpkListDT")
                        If cU.gfBuildDataTableforGridView(plDT, GridView1, True) Then
                            updatePLNo(plDT, GridView1)
                            If Not plDT Is Nothing AndAlso plDT.Rows.Count > 0 Then
                                For i = 0 To plDT.Rows.Count - 1
                                    Select Case plDT.Rows(i).Item("mFlag").ToString.Trim
                                        Case "N"
                                            maxSeq = DB.getValueFromSQL("select max(isnull(pld_seq,0)) + 1 as seq from WMS_DO_PICKLIST_D where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                      "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                      "and DO_CODE = '" & gU.dbEncode(plDT.Rows(i).Item("do_code").ToString.Trim) & "'", gConn, transaction)


                                            itemSQL = "insert into WMS_DO_PICKLIST_D " &
                                                         "(IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, " &
                                                         "PLD_ITEM_QTY, PLD_WH, PLD_LOC, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN, PLD_IS_LOAN, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY,PLD_STATUS, " &
                                                         "PLD_EXPIRY_DATE, PLD_MANU_DATE,PLD_WAVE_PICK_NO,PLD_PL_LIST_NO,PLD_PL_SORT_SEQ," &
                                                         "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                         "values " &
                                                         "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("DO_CODE").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(maxSeq, "1")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_picked_by").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_item_no").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_pack_key").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_pallet_no").ToString.Trim, "000")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_item_qty").ToString.Trim, "0")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_wh").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_loc").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_floor").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_area").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_rack").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_bin").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_is_loan").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_do_qty").ToString.Trim, "0")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("PLD_BATCH_NO").ToString.Trim, "")) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_foi_qty").ToString.Trim, "0")) & "', " &
                                                          gU.convdbDate(gU.dbEncode(plDT.Rows(i).Item("pld_status").ToString.Trim)) & ", " &
                                                          gU.convdbDate(gU.dbEncode(plDT.Rows(i).Item("pld_expiry_date").ToString.Trim)) & ", " &
                                                          gU.convdbDate(gU.dbEncode(plDT.Rows(i).Item("pld_manu_date").ToString.Trim)) & ", " &
                                                          gU.dbEncode(nextNo) & "', " &
                                                          "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim, "8")) & "', " &
                                                          gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("PLD_PL_SORT_SEQ").ToString.Trim, "1")) & ", " &
                                                          "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                        Case "U"
                                            itemSQL = "update WMS_DO_PICKLIST_D set " &
                                                          "PLD_PL_LIST_NO='" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim, "8")) & "'," &
                                                          "PLD_WAVE_PICK_NO='" & gU.dbEncode(gU.decodeNull(nextNo, "")) & "'," &
                                                          "PLD_PICKED_BY = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_picked_by").ToString.Trim, "")) & "', " &
                                                          "PLD_ITEM_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_item_qty").ToString.Trim, "0")) & "', " &
                                                          "PLD_WH = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_wh").ToString.Trim, "")) & "', " &
                                                          "PLD_LOC = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_loc").ToString.Trim, "")) & "', " &
                                                          "PLD_STATUS = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_status").ToString.Trim, "")) & "', " &
                                                          "PLD_FLOOR = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_floor").ToString.Trim, "")) & "', " &
                                                          "PLD_AREA = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_area").ToString.Trim, "")) & "', " &
                                                          "PLD_RACK = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_rack").ToString.Trim, "")) & "', " &
                                                          "PLD_BIN = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_bin").ToString.Trim, "")) & "', " &
                                                          "PLD_IS_LOAN = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_is_loan").ToString.Trim, "")) & "', " &
                                                          "PLD_DO_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_do_qty").ToString.Trim, "0")) & "', " &
                                                          "PLD_FOI_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_foi_qty").ToString.Trim, "0")) & "', " &
                                                          "PLD_PL_SORT_SEQ=" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("PLD_PL_SORT_SEQ").ToString.Trim, "1")) & ", " &
                                                          "sys_lub = '" & Session("usr_id") & "', " &
                                                          "sys_lud = Getdate() " &
                                                      "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                      "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                      "and DO_CODE = '" & gU.dbEncode(plDT.Rows(i).Item("do_code").ToString.Trim) & "' " &
                                                      "and PLD_SEQ = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_seq").ToString.Trim, "")) & "'"
                                        Case "PD"
                                            itemSQL = "delete from WMS_DO_PICKLIST_D " &
                                                      "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                      "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                      "and DO_CODE = '" & gU.dbEncode(plDT.Rows(i).Item("do_code").ToString.Trim) & "' " &
                                                      "and PLD_SEQ = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_seq").ToString.Trim, "")) & "'"
                                    End Select
                                    If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                                Next
                            End If
                        End If
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Stock Adjustment!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "貨量調整資料重複!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    If ViewState("WP_NO") <> "" Then
                        pk_code = ViewState("WP_NO")

                    Else
                        pk_code = Server.UrlDecode(Request("WP_NO"))

                    End If
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = " UPDATE WMS_WAVE_PICK_H SET " &
                                    "WP_DATE =convert(datetime,'" & WP_DATE.Text & "'," & DDFORMAT & ")," &
                                    "CUS_CODE=" & gU.convdbNVCData(gU.dbEncode(CUS_CODE.Text.Trim)) & "," &
                                    "CUS_NAME=" & gU.convdbNVCData(gU.dbEncode(CUS_NAME.Text.Trim)) & "," &
                                    "WP_WH_CODE=" & gU.convdbNVCData(gU.dbEncode(WP_WH_CODE.SelectedValue)) & "," &
                                    "WP_REMARK=" & gU.convdbNVCData(gU.dbEncode(WP_REMARK.Text.Trim)) & "," &
                                    "sys_lub = '" & Session("usr_id") & "', " &
                                    "sys_lud = Getdate() " &
                                 "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                 "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                 "and WP_NO = '" & gU.dbEncode(WP_NO.Text) & "' "
                    REM **********************
                    gDB.amendData(sql_string, gConn, transaction)

                    plDT = Session("WPpkListDT")

                    If cU.gfBuildDataTableforGridView(plDT, GridView1, True) Then
                        updatePLNo(plDT, GridView1)
                        If Not plDT Is Nothing AndAlso plDT.Rows.Count > 0 Then
                            For i = 0 To plDT.Rows.Count - 1
                                Select Case plDT.Rows(i).Item("mFlag").ToString.Trim
                                    Case "N"
                                        maxSeq = DB.getValueFromSQL("select max(isnull(pld_seq,0)) + 1 as seq from WMS_DO_PICKLIST_D where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                    "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                    "and DO_CODE = '" & gU.dbEncode(plDT.Rows(i).Item("do_code").ToString.Trim) & "'", gConn, transaction)


                                        itemSQL = "insert into WMS_DO_PICKLIST_D " &
                                                     "(IMP_CODE, STORER_CODE, DO_CODE, PLD_SEQ, PLD_PICKED_BY, PLD_ITEM_NO, PLD_PACK_KEY, PLD_PALLET_NO, " &
                                                     "PLD_ITEM_QTY, PLD_WH, PLD_LOC, PLD_FLOOR, PLD_AREA, PLD_RACK, PLD_BIN, PLD_IS_LOAN, PLD_DO_QTY, PLD_BATCH_NO, PLD_FOI_QTY,PLD_STATUS, " &
                                                     "PLD_EXPIRY_DATE, PLD_MANU_DATE,PLD_WAVE_PICK_NO,PLD_PL_LIST_NO,PLD_PL_SORT_SEQ," &
                                                     "sys_cb, sys_cd, sys_lub, sys_lud) " &
                                                     "values " &
                                                     "('" & Session("imp_code") & "', '" & gU.dbEncode(STORER_CODE.SelectedValue) & "', '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("DO_CODE").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(maxSeq, "1")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_picked_by").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_item_no").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_pack_key").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_pallet_no").ToString.Trim, "000")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_item_qty").ToString.Trim, "0")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_wh").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_loc").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_floor").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_area").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_rack").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_bin").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_is_loan").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_do_qty").ToString.Trim, "0")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("PLD_BATCH_NO").ToString.Trim, "")) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_foi_qty").ToString.Trim, "0")) & "', " &
                                                      gU.convdbDate(gU.dbEncode(plDT.Rows(i).Item("pld_status").ToString.Trim)) & ", " &
                                                      gU.convdbDate(gU.dbEncode(plDT.Rows(i).Item("pld_expiry_date").ToString.Trim)) & ", " &
                                                      gU.convdbDate(gU.dbEncode(plDT.Rows(i).Item("pld_manu_date").ToString.Trim)) & ", " &
                                                      "NULL,'" & gU.dbEncode(WP_NO.Text) & "', " &
                                                      "'" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim, "8")) & "', " &
                                                      gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("PLD_PL_SORT_SEQ").ToString.Trim, "1")) & ", " &
                                                      "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "
                                    Case "U"
                                        itemSQL = "update WMS_DO_PICKLIST_D set " &
                                                      "PLD_PL_LIST_NO='" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim, "")) & "'," &
                                                      "PLD_WAVE_PICK_NO='" & gU.dbEncode(gU.decodeNull(WP_NO.Text, "")) & "'," &
                                                      "PLD_PICKED_BY = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_picked_by").ToString.Trim, "")) & "', " &
                                                      "PLD_ITEM_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_item_qty").ToString.Trim, "0")) & "', " &
                                                      "PLD_WH = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_wh").ToString.Trim, "")) & "', " &
                                                      "PLD_LOC = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_loc").ToString.Trim, "")) & "', " &
                                                      "PLD_FLOOR = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_floor").ToString.Trim, "")) & "', " &
                                                      "PLD_AREA = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_area").ToString.Trim, "")) & "', " &
                                                      "PLD_RACK = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_rack").ToString.Trim, "")) & "', " &
                                                      "PLD_BIN = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_bin").ToString.Trim, "")) & "', " &
                                                      "PLD_IS_LOAN = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_is_loan").ToString.Trim, "")) & "', " &
                                                      "PLD_DO_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_do_qty").ToString.Trim, "0")) & "', " &
                                                      "PLD_FOI_QTY = '" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_foi_qty").ToString.Trim, "0")) & "', " &
                                                       "PLD_STATUS = '" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("pld_status").ToString.Trim, "0")) & "', " &
                                                      "PLD_PL_SORT_SEQ=" & gU.dbEncode(gU.decodeNullOrEmpty(plDT.Rows(i).Item("PLD_PL_SORT_SEQ").ToString.Trim, "0")) & ", " &
                                                      "sys_lub = '" & Session("usr_id") & "', " &
                                                      "sys_lud = Getdate() " &
                                                  "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                  "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                  "and DO_CODE = '" & gU.dbEncode(plDT.Rows(i).Item("do_code").ToString.Trim) & "' " &
                                                  "and PLD_SEQ = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_seq").ToString.Trim, "")) & "'"
                                    Case "D"
                                        itemSQL = "update WMS_DO_PICKLIST_D set " &
                                                     "PLD_PL_LIST_NO=NULL," &
                                                     "PLD_WAVE_PICK_NO=NULL," &
                                                     "PLD_STATUS=NULL," &
                                                     "sys_lub = '" & Session("usr_id") & "', " &
                                                     "sys_lud = Getdate() " &
                                                 "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                 "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                 "and DO_CODE = '" & gU.dbEncode(plDT.Rows(i).Item("do_code").ToString.Trim) & "' " &
                                                 "and PLD_SEQ = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_seq").ToString.Trim, "")) & "'"
                                    Case "PD"
                                        itemSQL = "delete from WMS_DO_PICKLIST_D " &
                                                  "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                                                  "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                                                  "and DO_CODE = '" & gU.dbEncode(plDT.Rows(i).Item("do_code").ToString.Trim) & "' " &
                                                  "and PLD_SEQ = '" & gU.dbEncode(gU.decodeNull(plDT.Rows(i).Item("pld_seq").ToString.Trim, "")) & "'"
                                End Select
                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                            Next
                        End If
                    End If
                End If
                successFlag = True
                'Response.Write(sql_string)
                transaction.Commit()

                If ViewState("pagemode") = "N" Then
                    ViewState("pagemode") = ""
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    WP_NO.Text = nextNo
                    ViewState("WP_NO") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    WP_NO.ForeColor = Drawing.Color.Black
                    WP_NO.Font.Size = 10
                    'WP_NO.CssClass = ""
                    STORER_CODE.CssClass = ""

                    REM **********************
                End If

                If flag <> "Y" Then
                    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                End If


            Catch ex As Exception
                transaction.Rollback()
                Response.Write(sql_string & "<br><br>")
                Response.Write(itemSQL & "<br><br>")
                Response.Write(ex.Message)
                uiFun.displayMsg(Me, "1008", "", Session("gLang"))
            Finally
                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try

            If successFlag Then Call BindGV()
        End If
    End Sub

    Protected Sub BindGV()
        Dim SQLString As String = ""
        Dim dt As New DataTable
        Dim SCString As String = "WHERE"
        Dim WhereStr As String = ""
        Dim pk_code As String = ""
        Dim storerCode As String

        REM **********************
        REM Modify Here
        REM Primary Key Session
        If ViewState("WP_NO") <> "" Then
            pk_code = ViewState("WP_NO")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("WP_NO"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))
        End If
        REM **********************

        WP_STATUS.ForeColor = Drawing.Color.Black

        If ViewState("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            WP_STATUS.Text = "NEW"
            WP_NO.ForeColor = Drawing.Color.Red

            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = " SELECT IMP_CODE, STORER_CODE, WP_NO, WP_STATUS,convert(varchar,WP_DATE," & DDFORMAT & ") as WP_DATE, CUS_CODE, CUS_NAME, WP_WH_CODE, WP_REMARK, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB " &
                        " FROM WMS_WAVE_PICK_H " &
                        "where WMS_WAVE_PICK_H.WP_NO = '" & gU.dbEncode(pk_code) & "' " &
                        "and WMS_WAVE_PICK_H.imp_code = '" & Session("IMP_CODE") & "' " &
                        "and WMS_WAVE_PICK_H.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                WP_NO.Text = dt.Rows(0).Item("WP_NO").ToString
                'STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                WP_STATUS.Text = dt.Rows(0).Item("WP_STATUS").ToString
                WP_DATE.Text = dt.Rows(0).Item("WP_DATE").ToString
                CUS_CODE.Text = dt.Rows(0).Item("CUS_CODE").ToString.Trim
                CUS_NAME.Text = dt.Rows(0).Item("CUS_name").ToString.Trim
                WP_WH_CODE.SelectedValue = dt.Rows(0).Item("WP_WH_CODE").ToString
                WP_REMARK.Text = dt.Rows(0).Item("WP_REMARK").ToString

                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , , True)
                End If

                uiFun.load_dropdown(WP_WH_CODE, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' and wh_code='" & dt.Rows(0).Item("WP_WH_CODE").ToString & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"), , True)

                If WP_NO.Text <> "" Then
                    'WP_NO.ReadOnly = True
                    'WP_NO.BorderWidth = 0
                    WP_NO.BackColor = Drawing.Color.Transparent
                End If

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "SELECT d.PLD_SEQ, d.PLD_PICKED_BY, d.PLD_ITEM_NO, i.ITM_SKU_NO, d.PLD_PACK_KEY, d.PLD_PALLET_NO, d.PLD_ITEM_QTY, l.ILOC_BAL_QTY, " &
                      "d.PLD_WH, d.PLD_LOC, d.PLD_FLOOR, d.PLD_AREA, d.PLD_RACK, d.PLD_BIN, d.PLD_IS_LOAN, d.PLD_DO_QTY, d.PLD_BATCH_NO, d.PLD_FOI_QTY,d.PLD_STATUS,  " &
                      "ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(pick_item.picked_qty, 0) as stock_qty, " &
                      "Convert(varchar, d.pld_expiry_date, " & gU.getConfig("DDFORMATNo") & ") as pld_expiry_date, " &
                      "Convert(varchar,d.pld_manu_date, " & gU.getConfig("DDFORMATNo") & ") as pld_manu_date, " &
                      "0.0 as HOLD_QTY, 0.0 as TOTAL_BAL, 0.0 as AVAIL_QTY, 0.0 as TOTAL_ITEM_QTY, 0.0 as TOTAL_FOI_QTY, 'U' as mFlag,i.ITM_WEIGHT_TYPE,  " &
                      "d.PLD_WAVE_PICK_NO,dm.do_code, dm.do_co_code,dm.DO_EDI_SIR_NO,d.PLD_PL_LIST_NO, d.PLD_PL_SORT_SEQ " &
                  "from WMS_WAVE_PICK_H h inner join WMS_DO_PICKLIST_D d on" &
                      " h.IMP_CODE = d.IMP_CODE AND " &
                      " h.STORER_CODE = d.STORER_CODE AND h.WP_NO = d.PLD_WAVE_PICK_NO " &
                  " INNER JOIN WMS_DELV_ORDER dm ON " &
                      " d.IMP_CODE = dm.IMP_CODE AND d.STORER_CODE = dm.STORER_CODE AND d.DO_CODE = dm.DO_CODE " &
                  "left outer join WMS_ITEM_LOC_BAL l " &
                      "on d.STORER_CODE = l.STORER_CODE " &
                      "and d.PLD_ITEM_NO = l.ITM_CODE " &
                      "and d.PLD_PACK_KEY = l.PACK_KEY " &
                      "and ISNULL(d.PLD_PALLET_NO,'000') = ISNULL(l.ILOC_PALLET_NO,'000') " &
                      "and d.PLD_LOC = l.ILOC_LOC " &
                      "and ISNULL(d.PLD_BATCH_NO, '') = ISNULL(l.ILOC_BATCH_NO, '') " &
                  "left outer join WMS_ITEM i " &
                      "on d.IMP_CODE = i.IMP_CODE " &
                      "and d.IMP_CODE = i.IMP_CODE " &
                      "and d.PLD_ITEM_NO = i.ITM_CODE " &
                      "and d.PLD_PACK_KEY = i.PACK_KEY " &
                  "LEFT OUTER JOIN ( " &
                          "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " &
                          "from wms_do_picklist_d p, wms_delv_order d " &
                          "where d.imp_code = p.imp_code " &
                          "and d.storer_code = p.storer_code " &
                          "and d.do_code = p.do_code " &
                          "and d.do_status = 'PICKED' " &
                          "and d.do_code not in (select do_code from WMS_DO_PICKLIST_D where WMS_DO_PICKLIST_D.PLD_WAVE_PICK_NO='" & gU.dbEncode(pk_code) & "' and IMP_CODE = '" & Session("IMP_CODE") & "' and storer_code='" & gU.dbEncode(storerCode) & "')" &
                          "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " &
                      "ON d.IMP_CODE = PICK_ITEM.IMP_CODE " &
                      "AND d.STORER_CODE = PICK_ITEM.STORER_CODE " &
                      "AND d.PLD_ITEM_NO = PICK_ITEM.PLD_ITEM_NO " &
                      "AND d.PLD_PACK_KEY = PICK_ITEM.PLD_PACK_KEY " &
                      "AND ISNULL(d.PLD_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " &
                      "AND ISNULL(d.PLD_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " &
                      "AND d.PLD_LOC = PICK_ITEM.PLD_LOC " &
                  "where h.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                  "and h.WP_NO = '" & gU.dbEncode(pk_code) & "' " &
                  "and h.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "
        SQLString = SQLString & " order by d.PLD_PL_LIST_NO,d.PLD_PL_SORT_SEQ, dm.do_code, dm.do_co_code "

        'REM **********************
        dt = gDB.getDataTable(SQLString)

        Session("WPpkListDT") = dt

        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If
        'ViewState("dt") = dt
        GridView1.DataBind()

        Dim tempDT As DataTable
        Dim tempDOList As String = ""
        SQLString = "select distinct do_code from WMS_DO_PICKLIST_D where IMP_CODE = '" & Session("IMP_CODE") & "' " &
                    "and PLD_WAVE_PICK_NO = '" & gU.dbEncode(pk_code) & "' " &
                    "and STORER_CODE = '" & gU.dbEncode(storerCode) & "' "
        tempDT = gDB.getDataTable(SQLString)

        DO_NO.Text = tempDT.Rows.Count

        If tempDT.Rows.Count > 0 Then
            For i = 0 To tempDT.Rows.Count - 1
                tempDOList = gU.appendToList(tempDT.Rows(i).Item("do_code").ToString.Trim, tempDOList)
            Next

            selectedDO.Value = tempDOList
            Session("selectedDO") = tempDOList
        Else
            selectedDO.Value = ""
            Session("selectedDO") = tempDOList
        End If

        SQLString = "Select count(distinct PLD_PL_LIST_NO) as count from WMS_DO_PICKLIST_D " &
                    "where IMP_CODE = '" & Session("IMP_CODE") & "' " &
                    "and PLD_WAVE_PICK_NO = '" & gU.dbEncode(pk_code) & "' " &
                    "and STORER_CODE = '" & gU.dbEncode(storerCode) & "' "

        PL_NO.Text = gU.decodeEmptyCInt(DB.getValueFromSQL(SQLString), 0)

        REM **********************
    End Sub

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        Dim pk_code As String = ViewState("WP_NO")
        Dim successFlag As Boolean = False
        Dim cancelSql As String = "update WMS_WAVE_PICK_H " &
                     "set WP_STATUS = 'CANCELLED', " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and WP_NO = '" & gU.dbEncode(ViewState("WP_NO")) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction


        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            cancelSql = " update WMS_DO_PICKLIST_D set " &
                        " PLD_PL_LIST_NO=NULL,PLD_WAVE_PICK_NO=NULL,PLD_STATUS=NULL" &
                        " Where PLD_WAVE_PICK_NO='" & ViewState("WP_NO") & "'" &
                        " and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()
            successFlag = True
            WP_STATUS.Text = "CANCELLED"

            ar.sec_write = "N"
            CancelBtn.Visible = False
            ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

            uiFun.displayMsg(Me, "", "This wave picking has been cancelled.", Session("gLang"))

        Catch ex As Exception
            transaction.Rollback()
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "1008", "", Session("gLang"))
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

        If successFlag Then Call BindGV()
    End Sub

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub addItemtoPLIST()

        Dim returnValue As String
        Dim paraArr As String()
        Dim SQLString As String = ""
        Dim doCode As String = ""
        returnValue = DOList.Value

        Dim tempDT As DataTable
        Dim plDT As DataTable = Session("WPpkListDT")
        If cU.gfBuildDataTableforGridView(plDT, GridView1, True) Then
            If Not String.IsNullOrWhiteSpace(returnValue) Then
                paraArr = gU.listToArray(returnValue)

                For i = 0 To paraArr.Length - 1
                    doCode = paraArr(i)

                    SQLString = "SELECT d.PLD_SEQ, d.PLD_PICKED_BY, d.PLD_ITEM_NO, i.ITM_SKU_NO, d.PLD_PACK_KEY, d.PLD_PALLET_NO, d.PLD_ITEM_QTY, l.ILOC_BAL_QTY, " &
                                "d.PLD_WH, d.PLD_LOC, d.PLD_FLOOR, d.PLD_AREA, d.PLD_RACK, d.PLD_BIN, d.PLD_IS_LOAN, d.PLD_DO_QTY, d.PLD_BATCH_NO, d.PLD_FOI_QTY,d.PLD_STATUS, " &
                                "ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(pick_item.picked_qty, 0) as stock_qty, " &
                                "Convert(varchar, d.pld_expiry_date, " & gU.getConfig("DDFORMATNo") & ") as pld_expiry_date, " &
                                "Convert(varchar,d.pld_manu_date, " & gU.getConfig("DDFORMATNo") & ") as pld_manu_date, " &
                                "Convert(varchar,ISNULL(d.pld_expiry_date, l.ILOC_EXPIRY_DATE), " & gU.getConfig("DDFORMATNo") & ") as ILOC_EXPIRY_DATE, " &
                                "0.0 as HOLD_QTY, 0.0 as TOTAL_BAL, 0.0 as AVAIL_QTY, 0.0 as TOTAL_ITEM_QTY, 0.0 as TOTAL_FOI_QTY, 'U' as mFlag,i.ITM_WEIGHT_TYPE,  " &
                                "d.PLD_WAVE_PICK_NO,dm.do_code, dm.do_co_code,dm.DO_EDI_SIR_NO,'1' as PLD_PL_LIST_NO,Cast(1 as numeric) as PLD_PL_SORT_SEQ " &
                            "from WMS_DO_PICKLIST_D d " &
                            " INNER JOIN WMS_DELV_ORDER dm ON " &
                                " d.IMP_CODE = dm.IMP_CODE AND d.STORER_CODE = dm.STORER_CODE AND d.DO_CODE = dm.DO_CODE " &
                            "left outer join WMS_ITEM_LOC_BAL l " &
                                "on d.STORER_CODE = l.STORER_CODE " &
                                "and d.PLD_ITEM_NO = l.ITM_CODE " &
                                "and d.PLD_PACK_KEY = l.PACK_KEY " &
                                "and ISNULL(d.PLD_PALLET_NO,'000') = ISNULL(l.ILOC_PALLET_NO,'000') " &
                                "and d.PLD_LOC = l.ILOC_LOC " &
                                "and ISNULL(d.PLD_BATCH_NO, '') = ISNULL(l.ILOC_BATCH_NO, '') " &
                            "left outer join WMS_ITEM i " &
                                "on d.IMP_CODE = i.IMP_CODE " &
                                "and d.IMP_CODE = i.IMP_CODE " &
                                "and d.PLD_ITEM_NO = i.ITM_CODE " &
                                "and d.PLD_PACK_KEY = i.PACK_KEY " &
                            "LEFT OUTER JOIN ( " &
                                    "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " &
                                    "from wms_do_picklist_d p, wms_delv_order d " &
                                    "where d.imp_code = p.imp_code " &
                                    "and d.storer_code = p.storer_code " &
                                    "and d.do_code = p.do_code " &
                                    "and d.do_status = 'PICKED' " &
                                    "and d.do_code <> '" & gU.dbEncode(doCode) & "' " &
                                    "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " &
                                "ON d.IMP_CODE = PICK_ITEM.IMP_CODE " &
                                "AND d.STORER_CODE = PICK_ITEM.STORER_CODE " &
                                "AND d.PLD_ITEM_NO = PICK_ITEM.PLD_ITEM_NO " &
                                "AND d.PLD_PACK_KEY = PICK_ITEM.PLD_PACK_KEY " &
                                "AND ISNULL(d.PLD_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " &
                                "AND ISNULL(d.PLD_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " &
                                "AND d.PLD_LOC = PICK_ITEM.PLD_LOC " &
                            "where d.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                            "and d.DO_CODE = '" & gU.dbEncode(doCode) & "' " &
                            "and d.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                            "order by dm.do_code, d.PLD_SEQ "

                    tempDT = gDB.getDataTable(SQLString)

                    If Not plDT Is Nothing Then
                        plDT.Merge(tempDT)
                    Else
                        plDT = tempDT
                    End If

                Next
                Dim selectedDO As String = Session("selectedDO")
                Session("selectedDO") = gU.appendToList(selectedDO, returnValue)

                Dim arDO As String() = gU.listToArray(gU.appendToList(selectedDO, returnValue))
                DO_NO.Text = arDO.Count

                Dim s_code As String = STORER_CODE.SelectedValue
                Dim w_code As String = WP_WH_CODE.SelectedValue
                uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' AND storer_code='" & s_code & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"), s_code, True)
                uiFun.load_dropdown(WP_WH_CODE, "select WH_CODE from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' and wh_code='" & w_code & "' ORDER BY 1", "WH_CODE", "WH_CODE", , Session("gSelectLabel"), w_code, True)

            End If
            SortPLList(plDT, "FL")

            plDT.AcceptChanges()
            Session("WPpkListDT") = plDT
            'ViewState("dt") = dt
            GridView1.DataSource = plDT
            GridView1.DataBind()

            Dim doList As String = ""
            Dim tempDOCOde As String = ""

            'End If
        End If

    End Sub

    Protected Sub SortPLList(ByRef plDT As DataTable, Optional ByVal sort_type As String = "DO")
        Dim copyDT As DataTable
        Dim pl_list_no As Integer = 0
        Dim sortKey As String = ""

        If Not plDT Is Nothing Or plDT.Rows.Count > 0 Then
            copyDT = plDT.Copy
            Dim dv As DataView = copyDT.DefaultView
            Select Case sort_type
                Case "DO"
                    dv.Sort = "do_code,PLD_LOC"
                    sortKey = "do_code"
                    plDT = dv.ToTable
                Case "LOC"
                    dv.Sort = "PLD_LOC"
                    sortKey = "PLD_LOC"
                    plDT = dv.ToTable
                Case "SKU"
                    dv.Sort = "ITM_SKU_NO,PLD_LOC"
                    sortKey = "ITM_SKU_NO"
                    plDT = dv.ToTable
                Case "WT"
                    dv.Sort = "ITM_WEIGHT_TYPE"
                    sortKey = "ITM_WEIGHT_TYPE"
                    plDT = dv.ToTable
                Case "FL"
                    dv.Sort = "PLD_LOC"
                    sortKey = "PLD_LOC"
                    plDT = dv.ToTable
            End Select

            Dim preKey As String = ""
            Dim subSeq As Integer = 0

            For i = 0 To plDT.Rows.Count - 1
                If plDT.Rows(i).Item("mFlag").ToString.Trim <> "D" Then
                    Select Case sort_type
                        Case "LOC"
                            If preKey <> plDT.Rows(i).Item("PLD_WH").ToString.Trim & plDT.Rows(i).Item("PLD_FLOOR").ToString.Trim & plDT.Rows(i).Item("PLD_AREA").ToString.Trim Then
                                preKey = plDT.Rows(i).Item("PLD_WH").ToString.Trim & plDT.Rows(i).Item("PLD_FLOOR").ToString.Trim & plDT.Rows(i).Item("PLD_AREA").ToString.Trim
                                If pl_list_no < 8 Then
                                    pl_list_no += 1
                                    subSeq = 0
                                End If
                            End If

                        Case "FL"
                            If preKey <> plDT.Rows(i).Item("PLD_WH").ToString.Trim & plDT.Rows(i).Item("PLD_FLOOR").ToString.Trim Then
                                preKey = plDT.Rows(i).Item("PLD_WH").ToString.Trim & plDT.Rows(i).Item("PLD_FLOOR").ToString.Trim
                                If pl_list_no < 8 Then
                                    pl_list_no += 1
                                    subSeq = 0
                                End If
                            End If

                        Case Else
                            If preKey <> plDT.Rows(i).Item(sortKey).ToString.Trim Then
                                preKey = plDT.Rows(i).Item(sortKey).ToString.Trim
                                If pl_list_no < 8 Then
                                    pl_list_no += 1
                                    subSeq = 0
                                End If

                            End If

                    End Select

                    plDT.Rows(i).Item("PLD_PL_SORT_SEQ") = subSeq + 1
                    plDT.Rows(i).Item("PLD_PL_LIST_NO") = pl_list_no

                    If pl_list_no = 0 Then plDT.Rows(i).Item("PLD_PL_LIST_NO") = 1
                End If
            Next
            PL_NO.Text = pl_list_no
        End If

    End Sub

    Protected Sub pl_sort_option_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles pl_sort_option.SelectedIndexChanged
        Dim plDt As DataTable
        plDt = Session("WPpkListDT")
        SortPLList(plDt, pl_sort_option.SelectedValue)
        plDt.AcceptChanges()

        GridView1.DataSource = plDt
        GridView1.DataBind()
        Session("WPpkListDT") = plDt

    End Sub

    Protected Sub updatePLNo(ByRef plDT As DataTable, ByRef gv As GridView)
        Dim pl_no As Integer = 1
        If gv.Rows.Count > 0 Then
            For i = 0 To plDT.Rows.Count - 1
                For x = 1 To 8
                    If Not gv.Rows(i).FindControl("PL" & x) Is Nothing AndAlso DirectCast(gv.Rows(i).FindControl("PL" & x), RadioButton).Checked Then
                        plDT.Rows(i).Item("PLD_PL_LIST_NO") = x
                        Exit For
                    End If
                Next
            Next
        End If
        plDT.AcceptChanges()
    End Sub

    Protected Sub btnReset_Click(sender As Object, e As System.EventArgs) Handles btnReset.Click
        Dim DOlist As String = Session("selectedDO")
        Dim DOArr As String()
        Dim tempSQL As String = ""
        Dim SQLString As String = ""
        Dim dt As DataTable

        If Not String.IsNullOrWhiteSpace(DOlist) Then
            DOArr = gU.listToArray(DOlist)
            For i = 0 To DOArr.Length - 1
                tempSQL &= "'" & DOArr(i) & "',"
            Next

            If tempSQL <> "" Then
                tempSQL = Left(tempSQL, Len(tempSQL) - 1)
            End If


            SQLString = "SELECT d.PLD_SEQ, d.PLD_PICKED_BY, d.PLD_ITEM_NO, i.ITM_SKU_NO, d.PLD_PACK_KEY, d.PLD_PALLET_NO, d.PLD_ITEM_QTY, l.ILOC_BAL_QTY, " &
                             "d.PLD_WH, d.PLD_LOC, d.PLD_FLOOR, d.PLD_AREA, d.PLD_RACK, d.PLD_BIN, d.PLD_IS_LOAN, d.PLD_DO_QTY, d.PLD_BATCH_NO, d.PLD_FOI_QTY,d.PLD_STATUS, wms_do_picklist_d.PLD_SERIAL_NO, wms_do_picklist_d.PLD_QTY2, " &
                             "ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(pick_item.picked_qty, 0) as stock_qty, " &
                             "Convert(varchar, d.pld_expiry_date, " & gU.getConfig("DDFORMATNo") & ") as pld_expiry_date, " &
                             "Convert(varchar,d.pld_manu_date, " & gU.getConfig("DDFORMATNo") & ") as pld_manu_date, " &
                             "Convert(varchar,ISNULL(d.pld_expiry_date, l.ILOC_EXPIRY_DATE), " & gU.getConfig("DDFORMATNo") & ") as ILOC_EXPIRY_DATE, " &
                             "0.0 as HOLD_QTY, 0.0 as TOTAL_BAL, 0.0 as AVAIL_QTY, 0.0 as TOTAL_ITEM_QTY, 0.0 as TOTAL_FOI_QTY, 'U' as mFlag,i.ITM_WEIGHT_TYPE," &
                             "d.PLD_WAVE_PICK_NO,dm.do_code, dm.do_co_code,dm.DO_EDI_SIR_NO,'1' as PLD_PL_LIST_NO,cast(1 as numeric) as PLD_PL_SORT_SEQ  " &
                         "from WMS_DO_PICKLIST_D d " &
                         " INNER JOIN WMS_DELV_ORDER dm ON " &
                             " d.IMP_CODE = dm.IMP_CODE AND d.STORER_CODE = dm.STORER_CODE AND d.DO_CODE = dm.DO_CODE " &
                         "left outer join WMS_ITEM_LOC_BAL l " &
                             "on d.STORER_CODE = l.STORER_CODE " &
                             "and d.PLD_ITEM_NO = l.ITM_CODE " &
                             "and d.PLD_PACK_KEY = l.PACK_KEY " &
                             "and ISNULL(d.PLD_PALLET_NO,'000') = ISNULL(l.ILOC_PALLET_NO,'000') " &
                             "and d.PLD_LOC = l.ILOC_LOC " &
                             "and ISNULL(d.PLD_BATCH_NO, '') = ISNULL(l.ILOC_BATCH_NO, '') " &
                         "left outer join WMS_ITEM i " &
                             "on d.IMP_CODE = i.IMP_CODE " &
                             "and d.IMP_CODE = i.IMP_CODE " &
                             "and d.PLD_ITEM_NO = i.ITM_CODE " &
                             "and d.PLD_PACK_KEY = i.PACK_KEY " &
                         "LEFT OUTER JOIN ( " &
                                 "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " &
                                 "from wms_do_picklist_d p, wms_delv_order d " &
                                 "where d.imp_code = p.imp_code " &
                                 "and d.storer_code = p.storer_code " &
                                 "and d.do_code = p.do_code " &
                                 "and d.do_status = 'PICKED' " &
                                 "and d.do_code not in (" & tempSQL & ") " &
                                 "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " &
                             "ON d.IMP_CODE = PICK_ITEM.IMP_CODE " &
                             "AND d.STORER_CODE = PICK_ITEM.STORER_CODE " &
                             "AND d.PLD_ITEM_NO = PICK_ITEM.PLD_ITEM_NO " &
                             "AND d.PLD_PACK_KEY = PICK_ITEM.PLD_PACK_KEY " &
                             "AND ISNULL(d.PLD_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " &
                             "AND ISNULL(d.PLD_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " &
                             "AND d.PLD_LOC = PICK_ITEM.PLD_LOC " &
                         "where d.IMP_CODE = '" & Session("IMP_CODE") & "' " &
                         "and d.DO_CODE in( " & tempSQL & ") " &
                         "and d.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                         "order by dm.do_code, d.PLD_SEQ "

            dt = gDB.getDataTable(SQLString)
            SortPLList(dt, "FL")
            dt.AcceptChanges()

            Session("WPpkListDT") = dt

            If dt.Rows.Count > 0 Then
                GridView1.DataSource = dt
            Else
                GridView1.DataSource = Nothing
            End If

            GridView1.DataBind()

        End If

    End Sub

    Protected Sub btnRelease_Click(sender As Object, e As System.EventArgs) Handles btnRelease.Click
        Dim pk_code As String = ViewState("WP_NO")

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction


        transaction = gConn.BeginTransaction()

        Try
            Dim ReleaseSql As String = "update WMS_WAVE_PICK_H " &
                     "set WP_STATUS = 'ASSIGNED', " &
                     "sys_lub = '" & Session("usr_id") & "', " &
                     "sys_lud = Getdate() " &
                     "where imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' " &
                     "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                     "and WP_NO = '" & gU.dbEncode(ViewState("WP_NO")) & "' "

            gDB.amendData(ReleaseSql, gConn, transaction)


            ReleaseSql = " update WMS_DO_PICKLIST_D set " &
                        " PLD_STATUS='ASSIGNED', " &
                        "sys_lub = '" & Session("usr_id") & "', " &
                        "sys_lud = Getdate() " &
                        " Where PLD_WAVE_PICK_NO='" & ViewState("WP_NO") & "'" &
                        " and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " &
                        " and imp_code = '" & gU.dbEncode(Session("IMP_CODE")) & "' "

            gDB.amendData(ReleaseSql, gConn, transaction)


            ReleaseSql = " update wms_delv_order set do_status='ASSIGNED'," &
                         "sys_lub = '" & Session("usr_id") & "', " &
                         "sys_lud = Getdate() " &
                         " where do_code in (select distinct do_code from WMS_DO_PICKLIST_D " &
                         " where imp_code='" & gU.dbEncode(Session("IMP_CODE")) & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and WMS_DO_PICKLIST_D.PLD_WAVE_PICK_NO='" & ViewState("WP_NO") & "') "

            gDB.amendData(ReleaseSql, gConn, transaction)

            transaction.Commit()

            WP_STATUS.Text = "ASSIGNED"

            ar.sec_write = "N"
            CancelBtn.Visible = False
            ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

            uiFun.displayMsg(Me, "", "This Wave picking has been released.", Session("gLang"))

        Catch ex As Exception
            transaction.Rollback()
            Response.Write(ex.Message)
            uiFun.displayMsg(Me, "1008", "", Session("gLang"))
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try
    End Sub

    Protected Function returnStatusColor(ByVal status As String) As String
        Dim returnColor As String = "Yellow"

        Select Case status
            Case "ASSIGNED"
                returnColor = "Yellow"
            Case "SHORT-PICKED"
                returnColor = "Green"
            Case "PICKING"
                returnColor = "Orange"
            Case "COMPELTE"
                returnColor = "Blue"
            Case "CANCELLED"
                returnColor = "Red"
            Case Else
                returnColor = "Transparent"
        End Select
        Return returnColor
    End Function

    Protected Sub btnSPLF_Click(sender As Object, e As System.EventArgs) Handles btnSPLF.Click
        If GridView1.Rows.Count > 0 Then
            Dim plDt As DataTable
            plDt = Session("WPpkListDT")
            SortPLList(plDt, "FL")
            plDt.AcceptChanges()

            GridView1.DataSource = plDt
            GridView1.DataBind()
            Session("WPpkListDT") = plDt
        End If
    End Sub

    Protected Sub btnSPLW_Click(sender As Object, e As System.EventArgs) Handles btnSPLW.Click
        Dim plDt As DataTable
        plDt = Session("WPpkListDT")

        If Not plDt Is Nothing AndAlso plDt.Rows.Count > 0 Then
            If cU.gfBuildDataTableforGridView(plDt, GridView1, True) Then
                'SortPLList(plDt, "WT")
                updatePLNo(plDt, GridView1)
                SplitbyWgt(plDt, wtPLNO_SEL.SelectedValue)

                plDt.AcceptChanges()

                GridView1.DataSource = plDt
                GridView1.DataBind()
                Session("WPpkListDT") = plDt
            End If
        End If

    End Sub

    Protected Sub sortSubList(ByRef sortDT As DataTable, ByVal sortKey As String, ByVal sortGrp As String)
        If Not sortDT Is Nothing AndAlso sortDT.Rows.Count > 0 Then
            Dim resultDT As DataTable
            Dim tempDT As DataTable
            Dim allSort As String = ""
            tempDT = sortDT.Copy

            Dim dv As DataView = tempDT.DefaultView
            If sortGrp <> "ALL" Then
                dv.RowFilter = "PLD_PL_LIST_NO='" & sortGrp & "'"
            Else
                allSort = "PLD_PL_LIST_NO,"
            End If


            Select Case sortKey
                Case "WT"
                    dv.Sort = allSort & "ITM_WEIGHT_TYPE,PLD_LOC"
                Case "FL"
                    dv.Sort = allSort & "PLD_LOC,ITM_WEIGHT_TYPE"
                Case "DO"
                    dv.Sort = allSort & "DO_CODE,PLD_LOC"
                Case "SKU"
                    dv.Sort = allSort & "ITM_SKU_NO,PLD_FLOOR,PLD_AREA"
            End Select

            resultDT = dv.ToTable
            Dim PrePLNO As String = ""
            Dim seqCount As Integer = 0

            If resultDT.Rows.Count > 0 Then
                For i = 0 To resultDT.Rows.Count - 1
                    If PrePLNO <> resultDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim Then
                        seqCount = 0
                        PrePLNO = resultDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim
                    End If

                    seqCount += 1
                    resultDT.Rows(i).Item("PLD_PL_SORT_SEQ") = seqCount
                Next
                resultDT.AcceptChanges()
            End If

            For i = 0 To sortDT.Rows.Count - 1
                If sortGrp <> "ALL" Then
                    If sortDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim = sortGrp Then
                        For x = 0 To resultDT.Rows.Count - 1
                            If sortDT.Rows(i).Item("DO_CODE").ToString.Trim = resultDT.Rows(x).Item("DO_CODE") AndAlso sortDT.Rows(i).Item("PLD_SEQ").ToString.Trim = resultDT.Rows(x).Item("PLD_SEQ") Then
                                sortDT.Rows(i).Item("PLD_PL_SORT_SEQ") = resultDT.Rows(x).Item("PLD_PL_SORT_SEQ")
                            End If
                        Next
                    End If
                Else
                    For x = 0 To resultDT.Rows.Count - 1
                        If sortDT.Rows(i).Item("DO_CODE").ToString.Trim = resultDT.Rows(x).Item("DO_CODE") AndAlso sortDT.Rows(i).Item("PLD_SEQ").ToString.Trim = resultDT.Rows(x).Item("PLD_SEQ") Then
                            sortDT.Rows(i).Item("PLD_PL_SORT_SEQ") = resultDT.Rows(x).Item("PLD_PL_SORT_SEQ")
                        End If
                    Next
                End If

            Next

            Dim resultDV As DataView
            resultDV = sortDT.DefaultView
            resultDV.Sort = "PLD_PL_LIST_NO,PLD_PL_SORT_SEQ"
            Dim tempcount = resultDV.Count
            sortDT = resultDV.ToTable
        End If
    End Sub

    Protected Sub btnSortPL_Click(sender As Object, e As System.EventArgs) Handles btnSortPL.Click
        Dim tempDT As DataTable = Session("WPpkListDT")
        If Not tempDT Is Nothing AndAlso tempDT.Rows.Count > 0 Then
            If cU.gfBuildDataTableforGridView(tempDT, GridView1, True) Then
                updatePLNo(tempDT, GridView1)

                sortSubList(tempDT, sort_type.SelectedValue, sort_pl_list_no.SelectedValue)
                tempDT.AcceptChanges()
                Session("WPpkListDT") = tempDT
                GridView1.DataSource = tempDT
                GridView1.DataBind()
            End If
        End If
    End Sub

    Protected Sub SplitbyWgt(ByRef sortDT As DataTable, ByVal sortGrp As String)
        If Not sortDT Is Nothing AndAlso sortDT.Rows.Count > 0 Then
            Dim maxPLNOs As String = ""
            Dim resultDT As DataTable
            Dim dv As DataView

            Dim lightSeqCount As Integer = 0
            Dim HeavySeqCount As Integer = 0

            maxPLNOs = sortDT.Compute("max(PLD_PL_LIST_NO)", "")

            If maxPLNOs < maxPLNO Then
                resultDT = sortDT.Copy

                dv = resultDT.DefaultView
                dv.RowFilter = "PLD_PL_LIST_NO='" & sortGrp & "'"
                dv.Sort = "ITM_WEIGHT_TYPE,PLD_LOC"

                resultDT = dv.ToTable

                If resultDT.Rows.Count = 0 Then
                    uiFun.displayMsg(Me, "", "No item in this PL grp.", Session("gLang"))
                ElseIf resultDT.Rows.Count = 1 Then
                    uiFun.displayMsg(Me, "", "Cannot split grp with Only 1 item.", Session("gLang"))
                ElseIf resultDT.Rows.Count > 1 Then

                    Dim hasHeavyChg As Boolean = False
                    Dim hasLightChg As Boolean = False

                    For i = 0 To resultDT.Rows.Count - 1
                        If resultDT.Rows(i).Item("ITM_WEIGHT_TYPE").ToString.Trim = "Heavy" Then
                            resultDT.Rows(i).Item("PLD_PL_LIST_NO") = CInt(resultDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim) + 1
                            HeavySeqCount += 1
                            resultDT.Rows(i).Item("PLD_PL_SORT_SEQ") = HeavySeqCount
                            hasHeavyChg = True
                        Else
                            lightSeqCount += 1
                            resultDT.Rows(i).Item("PLD_PL_SORT_SEQ") = lightSeqCount
                            hasLightChg = True
                        End If

                    Next
                    resultDT.AcceptChanges()

                    If hasHeavyChg AndAlso hasLightChg Then
                        For i = 0 To sortDT.Rows.Count - 1
                            If sortDT.Rows(i).Item("PLD_PL_LIST_NO") = sortGrp Then
                                For x = 0 To resultDT.Rows.Count - 1
                                    If sortDT.Rows(i).Item("DO_CODE").ToString.Trim = resultDT.Rows(x).Item("DO_CODE") AndAlso sortDT.Rows(i).Item("PLD_SEQ").ToString.Trim = resultDT.Rows(x).Item("PLD_SEQ") Then
                                        sortDT.Rows(i).Item("PLD_PL_SORT_SEQ") = resultDT.Rows(x).Item("PLD_PL_SORT_SEQ")
                                        sortDT.Rows(i).Item("PLD_PL_LIST_NO") = resultDT.Rows(x).Item("PLD_PL_LIST_NO")
                                    End If
                                Next

                            ElseIf sortDT.Rows(i).Item("PLD_PL_LIST_NO") > sortGrp Then
                                sortDT.Rows(i).Item("PLD_PL_LIST_NO") = CInt(gU.decodeNullOrEmpty(sortDT.Rows(i).Item("PLD_PL_LIST_NO").ToString.Trim, "7")) + 1
                            End If
                        Next
                    End If

                    Dim resultDV As DataView
                    resultDV = sortDT.DefaultView
                    resultDV.Sort = "PLD_PL_LIST_NO,PLD_PL_SORT_SEQ"
                    Dim tempcount = resultDV.Count
                    sortDT = resultDV.ToTable
                End If

            Else
                uiFun.displayMsg(Me, "", "Max No. of Picklist group has been reached.\nCannot divide into more group.", Session("gLang"))
            End If


        End If
    End Sub

    Protected Sub btPL1_Click(sender As Object, e As System.EventArgs) Handles btPL1.Click
        Dim tempDT As DataTable = Session("WPpkListDT")
        If Not tempDT Is Nothing AndAlso tempDT.Rows.Count > 0 Then
            If cU.gfBuildDataTableforGridView(tempDT, GridView1, True) Then
                For i = 0 To tempDT.Rows.Count - 1
                    tempDT.Rows(i).Item("PLD_PL_LIST_NO") = 1
                Next

                tempDT.AcceptChanges()
                Session("WPpkListDT") = tempDT
                GridView1.DataSource = tempDT
                GridView1.DataBind()
            End If
        End If
    End Sub

    Protected Sub btnReloadPicking_Click(sender As Object, e As System.EventArgs) Handles btnReloadPicking.Click

        Dim plDt As DataTable = Session("WPpkListDT")

        Dim StkBaldt As DataTable
        Dim SQLString As String
        Dim pk_code, storerCode As String

        If ViewState("WP_NO") <> "" Then
            pk_code = ViewState("WP_NO")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("WP_NO"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))
        End If

        SQLString = "SELECT d.PLD_SEQ, d.PLD_PICKED_BY, d.PLD_ITEM_NO, i.ITM_SKU_NO, d.PLD_PACK_KEY, d.PLD_PALLET_NO, d.PLD_ITEM_QTY, l.ILOC_BAL_QTY, " &
              "d.PLD_WH, d.PLD_LOC, d.PLD_FLOOR, d.PLD_AREA, d.PLD_RACK, d.PLD_BIN, d.PLD_IS_LOAN, d.PLD_DO_QTY, d.PLD_BATCH_NO, d.PLD_FOI_QTY,d.PLD_STATUS," &
              "ISNULL(l.ILOC_BAL_QTY, 0) - ISNULL(pick_item.picked_qty, 0) as stock_qty, " &
              "Convert(varchar, d.pld_expiry_date, " & gU.getConfig("DDFORMATNo") & ") as pld_expiry_date, " &
              "Convert(varchar,d.pld_manu_date, " & gU.getConfig("DDFORMATNo") & ") as pld_manu_date, " &
              "0.0 as HOLD_QTY, 0.0 as TOTAL_BAL, 0.0 as AVAIL_QTY, 0.0 as TOTAL_ITEM_QTY, 0.0 as TOTAL_FOI_QTY, 'U' as mFlag,i.ITM_WEIGHT_TYPE,  " &
              "d.PLD_WAVE_PICK_NO,dm.do_code, dm.do_co_code,dm.DO_EDI_SIR_NO,d.PLD_PL_LIST_NO, d.PLD_PL_SORT_SEQ " &
          "from WMS_WAVE_PICK_H h inner join WMS_DO_PICKLIST_D d on" &
              " h.IMP_CODE = d.IMP_CODE AND " &
              " h.STORER_CODE = d.STORER_CODE AND h.WP_NO = d.PLD_WAVE_PICK_NO " &
          " INNER JOIN WMS_DELV_ORDER dm ON " &
              " d.IMP_CODE = dm.IMP_CODE AND d.STORER_CODE = dm.STORER_CODE AND d.DO_CODE = dm.DO_CODE " &
          "left outer join WMS_ITEM_LOC_BAL l " &
              "on d.STORER_CODE = l.STORER_CODE " &
              "and d.PLD_ITEM_NO = l.ITM_CODE " &
              "and d.PLD_PACK_KEY = l.PACK_KEY " &
              "and ISNULL(d.PLD_PALLET_NO,'000') = ISNULL(l.ILOC_PALLET_NO,'000') " &
              "and d.PLD_LOC = l.ILOC_LOC " &
              "and ISNULL(d.PLD_BATCH_NO, '') = ISNULL(l.ILOC_BATCH_NO, '') " &
          "left outer join WMS_ITEM i " &
              "on d.IMP_CODE = i.IMP_CODE " &
              "and d.IMP_CODE = i.IMP_CODE " &
              "and d.PLD_ITEM_NO = i.ITM_CODE " &
              "and d.PLD_PACK_KEY = i.PACK_KEY " &
          "LEFT OUTER JOIN ( " &
                  "select p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000') as pld_pallet_no, ISNULL(p.pld_batch_no, '') as pld_batch_no, p.pld_loc, sum(p.pld_item_qty) as picked_qty " &
                  "from wms_do_picklist_d p, wms_delv_order d " &
                  "where d.imp_code = p.imp_code " &
                  "and d.storer_code = p.storer_code " &
                  "and d.do_code = p.do_code " &
                  "and d.do_status = 'PICKED' " &
                  "and d.do_code not in (select do_code from WMS_DO_PICKLIST_D where WMS_DO_PICKLIST_D.PLD_WAVE_PICK_NO='" & gU.dbEncode(pk_code) & "' and IMP_CODE = '" & Session("IMP_CODE") & "' and storer_code='" & gU.dbEncode(storerCode) & "')" &
                  "group by p.imp_code, p.storer_code, p.pld_item_no, p.pld_pack_key, ISNULL(p.pld_pallet_no, '000'), ISNULL(p.pld_batch_no, ''), p.pld_loc) PICK_ITEM " &
              "ON d.IMP_CODE = PICK_ITEM.IMP_CODE " &
              "AND d.STORER_CODE = PICK_ITEM.STORER_CODE " &
              "AND d.PLD_ITEM_NO = PICK_ITEM.PLD_ITEM_NO " &
              "AND d.PLD_PACK_KEY = PICK_ITEM.PLD_PACK_KEY " &
              "AND ISNULL(d.PLD_PALLET_NO, '000') = PICK_ITEM.PLD_PALLET_NO " &
              "AND ISNULL(d.PLD_BATCH_NO, '') = PICK_ITEM.PLD_BATCH_NO " &
              "AND d.PLD_LOC = PICK_ITEM.PLD_LOC " &
          "where h.IMP_CODE = '" & Session("IMP_CODE") & "' " &
          "and h.WP_NO = '" & gU.dbEncode(pk_code) & "' " &
          "and h.STORER_CODE = '" & gU.dbEncode(storerCode) & "' "
        SQLString = SQLString & " order by d.PLD_PL_LIST_NO,d.PLD_PL_SORT_SEQ, dm.do_code, dm.do_co_code "

        'REM **********************
        dt = gDB.getDataTable(SQLString)

        Session("WPpkListDT") = dt

        If Not plDt Is Nothing AndAlso plDt.Rows.Count > 0 Then
            Dim dtl_do_code As String = ""

            'dtl_do_code = plDt.Rows(e.RowIndex).Item("do_code")

            For i = 0 To plDt.Rows.Count - 1

                If plDt.Rows(i).Item("do_code").ToString.Trim = dtl_do_code Then
                    plDt.Rows(i).Item("mFlag") = "D"
                    GridView1.Rows(i).Visible = False
                End If

            Next
            'plDt.Rows(e.RowIndex).Item("mFlag") = "D"
            plDt.AcceptChanges()
            Session("WPpkListDT") = plDt
        End If


    End Sub
End Class
