Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OPERATION_STQ_STREQMain
    Inherits System.Web.UI.Page

    Private gDB As New GlobalDBFunc
    Private DB As New DBfunc
    Private cU As New CommonUtils
    Private uiFun As New UIfunc
    Private gU As New GeneralUtils
    Private ar As AccessRightUtils
    Private cm As CommonMenu
    Private st As New StockTrans

    Private moduleAction As String = ""
    Private dt As New DataTable
    Private DDFORMAT As String = "DD/MM/YYYY"
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
            Session("IS_Cable") = ""
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            uiFun.load_dropdown(STORER_CODE, "select STORER_CODE, STO_SHORTNAME from WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
            uiFun.load_dropdown(TQ_WH_FR, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"))
            uiFun.load_dropdown(TQ_WH_TO, "select distinct WH_MAIN_WH from WMS_WAREHOUSE WHERE IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' ORDER BY 1", "WH_MAIN_WH", "WH_MAIN_WH", , Session("gSelectLabel"))


            Dim tImage As Image = Image_Loc_LookUp_TQ_TO_LOC
            tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
            tImage.Attributes.Add("onclick", "LocLookUp(2,'','" & HttpUtility.HtmlEncode(dsp_TQ_TO_LOC.ClientID) & "', '" & HttpUtility.HtmlEncode(TQ_TO_LOC.ClientID) & "')")
        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If
            TQ_BY.Text = Session("usr_id")
            If TQ_DATE.Text = "" Then
                TQ_DATE.Text = Now().Date.ToString("dd/MM/yyyy")
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Inter-store Transfer Request Maintenance"

            lbl_ImageHd.Text = "Transfered Items"
            lbl_TQ_CODE.Text = "RFT Code:"
            lbl_TQ_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Storer:"
            lbl_TQ_EDI_RFT_NO.Text = "RFT No."
            lbl_TQ_DATE.Text = "Date:"
            lbl_TQ_BY.Text = "Transfered By:"
            lbl_TQ_BATCH_NO.Text = "Batch No.:"
            lbl_TQ_REF_NO.Text = "Ref. No.:"
            lbl_TQ_WH_FR.Text = "From (Warehouse):"
            lbl_TQ_WH_TO.Text = "To (Warehouse):"
            lbl_TQ_TO_LOC.Text = "To Location"
            lbl_TQ_REM.Text = "Remarks"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"
            newrow.Text = "Add"
            btnPost.Text = "Post"
            selectItemBtn.Text = "Select Item"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            btnPost.OnClientClick = "return confirm(""Are you sure to post this record?\r\n(Please save your work before Posting)"");"
            If Session("pagemode") = "N" Then
                TQ_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "貨品轉移維護"
            lbl_ImageHd.Text = "貨品詳情"
            lbl_TQ_CODE.Text = "轉移號碼:"
            lbl_TQ_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "貨主:"
            lbl_TQ_EDI_RFT_NO.Text = "RFT No."
            lbl_TQ_DATE.Text = "日期:"
            lbl_TQ_BY.Text = "轉移者:"
            lbl_TQ_BATCH_NO.Text = "批號:"
            lbl_TQ_REF_NO.Text = "文件編號:"
            lbl_TQ_WH_FR.Text = "由 (倉庫):"
            lbl_TQ_WH_TO.Text = "到 (倉庫):"
            lbl_TQ_TO_LOC.Text = "To"
            lbl_TQ_REM.Text = "備註:"

            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "儲存"
            saveBtn2.Text = "儲存"
            CancelBtn.Text = "取消"
            newrow.Text = "新增"
            btnPost.Text = "發布"
            selectItemBtn.Text = "選擇物料"
            CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"
            btnPost.OnClientClick = "return confirm(""確定發布資料?"");"
            If Session("pagemode") = "N" Then
                TQ_CODE.Text = "[號碼會自動產生]"
            End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        TQ_WH_FR.CssClass = "REQUIRED"
        TQ_WH_TO.CssClass = "REQUIRED"
        TQ_DATE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            'TQ_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("TQ_CODE") = ""
            ViewState("TQD_loc_fr") = Nothing
            ViewState("TQD_loc_to") = Nothing

            Call BindGV()
        Else
            dt = ViewState("dt")

        End If

        If moduleAction = "SELECTIM" Then
            addItemtoSTF()
            dsp_TQ_TO_LOC.Text = TQ_TO_LOC.Value
            'ElseIf moduleAction = "SELECTTOWH" Then
            'changeToPallet()
        End If

        'selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.getElementById('" & STORER_CODE.ClientID & "').value, document.getElementById('" & TQ_WH_FR.ClientID & "').value);")
        btnAttach.Attributes.Add("onclick", "javascript:goToAttach('OP_STQ','" & Session("imp_code") & "||" & ViewState("STORER_CODE") & "||" & ViewState("TQ_CODE") & "','N');")

        setPageCtrlAccess()

        If TQ_STATUS.Text = "CANCELLED" Or TQ_STATUS.Text = "CLOSED" Then
            ar.sec_write = "N"
            CancelBtn.Visible = False
            btnClose.Visible = False
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)


    End Sub
    Private Sub setPageCtrlAccess()

        exceptionEditList = New List(Of String)

    End Sub
    Private Function customizectrl(ByVal ctl As Control, ByRef ctrlArrayList As ArrayList) As Boolean
        customizectrl = False
    End Function

    Private Function page_customizectrl(ByVal ctl As Control) As Boolean
        page_customizectrl = False

    End Function


    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        'Select Case e.Row.RowType
        '    Case DataControlRowType.Header
        '        Dim oGridView As GridView = DirectCast(sender, GridView)
        '        Dim oGridViewRow As New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

        '        REM **********************
        '        REM Use for re-create the label to change the Langauge
        '        REM Modify Here
        '        Call cU.changeGVLabel(oGridViewRow, e, "No.", "編號")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Item Code.", "物料號碼")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Pack Key", "封裝內碼")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Item Name", "物料名稱")
        '        Call cU.changeGVLabel(oGridViewRow, e, "SKU No.", "SKU No.")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Qty", "數量")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Batch No.", "Batch No.")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Pallet No.", "貨板")
        '        Call cU.changeGVLabel(oGridViewRow, e, "From Loc.", "由位置")
        '        Call cU.changeGVLabel(oGridViewRow, e, "To Loc.", "到位置")
        '        Call cU.changeGVLabel(oGridViewRow, e, "Remarks", "備註")
        '        Call cU.changeGVLabel(oGridViewRow, e, "", "")
        '        REM **********************

        '        oGridView.Controls(0).Controls.AddAt(0, oGridViewRow)
        'End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("itm_code"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim

                CType(e.Row.FindControl("pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim
                CType(e.Row.FindControl("TQD_UOM"), Label).Text = DataBinder.Eval(e.Row.DataItem, "TQD_UOM").ToString.Trim

                Dim pDropDown As DropDownList = CType(e.Row.FindControl("TQD_batch_no_fr"), DropDownList)
                uiFun.load_dropdown(pDropDown, "select dc_date_code from wms_date_code where imp_code = '" & gU.dbEncode(Session("imp_code")) & "' and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' order by 1", "dc_date_code", "dc_date_code", , "N/A", DataBinder.Eval(e.Row.DataItem, "TQD_batch_no_fr").ToString.Trim, False)

                CType(e.Row.FindControl("itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_name").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                CType(e.Row.FindControl("TQD_qty"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "TQD_qty").ToString.Trim)
                CType(e.Row.FindControl("TQD_qty2"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "TQD_qty2").ToString.Trim)

                CType(e.Row.FindControl("TQD_pallet_no_fr"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TQD_pallet_no_fr").ToString.Trim

                CType(e.Row.FindControl("TQD_rem"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TQD_rem").ToString.Trim

                CType(e.Row.FindControl("TQD_SERIAL"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "TQD_SERIAL").ToString.Trim
                CType(e.Row.FindControl("TQD_UOM"), Label).Text = DataBinder.Eval(e.Row.DataItem, "TQD_UOM").ToString.Trim
                CType(e.Row.FindControl("TQD_UOM2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "TQD_UOM2").ToString.Trim

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    nButton.Text = "删除"
                End If


                'CType(e.Row.FindControl("TQD_batch_no_to"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "TQD_batch_no_to").ToString.Trim
                'CType(e.Row.FindControl("TQD_batch_no_fr"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "TQD_batch_no_fr").ToString.Trim

                CType(e.Row.FindControl("mFlag"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    Call ar.hideGVRow(GridView1, e.Row)
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                    'nButton.Enabled = False
                End If
        End Select
    End Sub

    Protected Sub newrow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newrow.Click
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(TQD_SEQ AS int)) + 1 from WMS_STOCK_TRANS_REQ_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and TQ_CODE = '" & gU.dbEncode(TQ_CODE.Text.Trim) & "' "
            REM **********************
            Dim nS_dt As New DataTable
            nS_dt = gDB.getDataTable(seq_string)
            Dim next_seq_no, temp_no As String
            Dim temp_seq_no As Integer = 1

            next_seq_no = ""

            If nS_dt.Rows.Count > 0 Then
                next_seq_no = nS_dt.Rows(0).Item(0).ToString()
            Else
                temp_no = "1"
            End If

            If next_seq_no = "" Then next_seq_no = "1"

            If ViewState("n_cur_seq") = "" Then
                ViewState("n_cur_seq") = next_seq_no
            Else
                temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                ViewState("n_cur_seq") = temp_seq_no.ToString
            End If


            dt.Rows.Add()
            rows_count = dt.Rows.Count

            REM **********************
            REM Modify Here
            dt.Rows(rows_count - 1).Item("TQD_seq") = ViewState("n_cur_seq").ToString
            REM **********************
            dt.Rows(rows_count - 1).Item("mFlag") = "N"
            dt.AcceptChanges()

            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()

        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        CType(GridView1.Rows(e.RowIndex).FindControl("mflag"), HiddenField).Value = "D"
        dt.AcceptChanges()
        reloadHiddenValue(GridView1)
    End Sub

    Private Function validateAll(Optional ByVal flag As String = "") As Boolean
        Dim selectSql As String = ""
        Dim i As Integer


        If STORER_CODE.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_STORER_CODE.Text & "不能空白!", Session("gLang"))
            End If

            If GridView1.Rows.Count > 0 Then reloadHiddenValue(GridView1)
            Return False

        End If


        If TQ_WH_FR.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_TQ_WH_FR.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_TQ_WH_FR.Text & "不能空白!", Session("gLang"))
            End If

            If GridView1.Rows.Count > 0 Then reloadHiddenValue(GridView1)
            Return False
        End If

        If TQ_WH_TO.SelectedValue = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_TQ_WH_TO.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_TQ_WH_TO.Text & "不能空白!", Session("gLang"))
            End If
            If GridView1.Rows.Count > 0 Then reloadHiddenValue(GridView1)
            Return False
        End If

        If TQ_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_TQ_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_TQ_DATE.Text & "不能空白!", Session("gLang"))
            End If

            If GridView1.Rows.Count > 0 Then reloadHiddenValue(GridView1)
            Return False
        ElseIf Not gU.isValidDate(TQ_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_TQ_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_TQ_DATE.Text & "!", Session("gLang"))
            End If

            If GridView1.Rows.Count > 0 Then reloadHiddenValue(GridView1)
            Return False
        End If

        Dim itemCount As Integer = 0


        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If CType(GridView1.Rows(i).FindControl("mflag"), HiddenField).Value <> "D" Then
                    itemCount += 1

                    If CType(GridView1.Rows(i).FindControl("TQD_qty"), TextBox).Text <> "" Then
                        If uiFun.gvValidate(Me, dt, "TQD_qty", "Qty", _
                                         CType(GridView1.Rows(i).FindControl("TQD_qty"), TextBox).Text) = False Then Return False


                        If CInt(CType(GridView1.Rows(i).FindControl("TQD_qty"), TextBox).Text) <= 0 Then
                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Item Qty Cannot Be Zero!", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "物料數量不能零!", Session("gLang"))
                            End If

                            reloadHiddenValue(GridView1)

                            Return False
                        End If
                    Else
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Item Qty Cannot Be Empty!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "物料數量不能空白!", Session("gLang"))
                        End If

                        reloadHiddenValue(GridView1)

                        Return False
                    End If
                End If
            Next
        End If

        If itemCount = 0 Then
            uiFun.displayMsg(Me, "", "Please select item for transferring stock!", Session("gLang"))
            reloadHiddenValue(GridView1)
            Return False
        End If

        Return True

    End Function

    Private Sub reloadHiddenValue(ByRef gv As GridView)
        
    End Sub

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim nextNo As String = ""
        Dim gConn As SqlConnection
        Dim dupSQL As String = ""
        Dim dupTbl As New DataTable
        Dim successFlag As Boolean = False

        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    Dim docType As String = ""
                    
                    docType = "STQ"

                    nextNo = DB.getDocNo(docType, gConn, transaction)
                    'nextNo = TQ_CODE.Text
                    REM **********************

                    dupSQL = "select 1 from WMS_STOCK_TRANS_REQ " & _
                                "where TQ_code = '" & gU.dbEncode(nextNo) & "' " & _
                                "and imp_code= '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and storer_code= '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into WMS_STOCK_TRANS_REQ (" & _
                        "TQ_code, imp_code, storer_code, " & _
                        "TQ_status, TQ_date, TQ_by, " & _
                        "TQ_batch_no, TQ_ref_no, " & _
                        "TQ_wh_fr, TQ_wh_to, TQ_TO_LOC, TQ_rem, " & _
                        "TQ_edi_rft_no, " & _
                         "sys_cb, sys_cd, sys_lub, sys_lud)" & _
                        "values ( " & _
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(TQ_STATUS.Text.Trim)) & "," & gU.convdbDate(gU.dbEncode(TQ_DATE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(TQ_BY.Text.Trim)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(TQ_BATCH_NO.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(TQ_REF_NO.Text.Trim)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(TQ_WH_FR.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(TQ_WH_TO.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(TQ_TO_LOC.Value.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(TQ_REM.Text.Trim)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(TQ_EDI_RFT_NO.Text.Trim)) & "," & _
                         "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                        REM **********************

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                            uiFun.reOrderDetails(dt, "TQD_seq")

                            For Each rows As DataRow In dt.Rows
                                itemSQL = ""

                                REM **********************
                                REM Modify Here
                                Select Case rows.Item("mFlag")
                                    Case "N"
                                        itemSQL = "insert into WMS_STOCK_TRANS_REQ_D (" & _
                                                "TQ_code, imp_code, storer_code, " & _
                                                "TQD_seq, itm_code, pack_key, " & _
                                                "TQD_batch_no_fr, TQD_qty, TQD_org_qty, TQD_loc_fr, " & _
                                                "TQD_loc_to, TQD_rem, " & _
                                                "TQD_pallet_no_fr, TQD_pallet_no_to, TQD_batch_no_to," & _
                                                "TQD_SERIAL, TQD_UOM, TQD_UOM2, TQD_QTY2,TQD_ORG_QTY2,TQD_EXPIRY_DATE, TQD_MANU_DATE, TQD_ITM_NAME, TQD_DRUM_ID_FR, TQD_DRUM_ID_TO, TQD_DRUM_LV_FR, TQD_DRUM_LV_TO," & _
                                                "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                "values (" & _
                                                gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_seq").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("itm_code").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("pack_key").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_qty").ToString.Trim, "0")) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_org_qty").ToString.Trim, "0")) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_loc_fr").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_loc_to").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_rem").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_SERIAL").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_UOM").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_UOM2").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_QTY2").ToString.Trim, "0")) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_ORG_QTY2").ToString.Trim, "0")) & ", " & _
                                                gU.convdbDate(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_EXPIRY_DATE").ToString.Trim, ""))) & "," & _
                                                gU.convdbDate(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_MANU_DATE").ToString.Trim, ""))) & "," & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_ITM_NAME").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_FR").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_TO").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_DRUM_LV_FR").ToString.Trim, "NULL")) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_DRUM_LV_TO").ToString.Trim, "NULL")) & ", " & _
                                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                        '"(N'" & Session("imp_code") & "', N'" & gU.dbEncode(STORER_CODE.Text) & "', N'" & nextNo & "', " & _
                                End Select
                                REM **********************

                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                            Next
                        End If
                    Else
                        If Not transaction Is Nothing Then
                            transaction.Rollback()
                            transaction = Nothing
                        End If

                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Duplicate record has found in Stock Transfer!!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "貨品轉移资料重复!!", Session("gLang"))
                        End If

                        Exit Sub
                    End If
                Else
                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    sql_string = "update WMS_STOCK_TRANS_REQ set " & _
                                    "TQ_status = " & gU.convdbNVCData(gU.dbEncode(TQ_STATUS.Text)) & ", " & _
                                    "TQ_date = " & gU.convdbDate(gU.dbEncode(TQ_DATE.Text.Trim)) & ", " & _
                                    "TQ_by = " & gU.convdbNVCData(gU.dbEncode(TQ_BY.Text.Trim)) & ", " & _
                                    "TQ_batch_no = " & gU.convdbNVCData(gU.dbEncode(TQ_BATCH_NO.Text.Trim)) & ", " & _
                                    "TQ_ref_no = " & gU.convdbNVCData(gU.dbEncode(TQ_REF_NO.Text.Trim)) & ", " & _
                                    "TQ_wh_fr = " & gU.convdbNVCData(gU.dbEncode(TQ_WH_FR.SelectedValue)) & ", " & _
                                    "TQ_wh_to = " & gU.convdbNVCData(gU.dbEncode(TQ_WH_TO.SelectedValue)) & ", " & _
                                    "TQ_TO_LOC = " & gU.convdbNVCData(gU.dbEncode(TQ_TO_LOC.Value.Trim)) & ", " & _
                                    "TQ_rem = " & gU.convdbNVCData(gU.dbEncode(TQ_REM.Text.Trim)) & ", " & _
                                    "TQ_edi_rft_no = " & gU.convdbNVCData(gU.dbEncode(TQ_EDI_RFT_NO.Text.Trim)) & ", " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = Getdate() " & _
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and TQ_code = '" & gU.dbEncode(TQ_CODE.Text) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        uiFun.reOrderDetails(dt, "TQD_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    itemSQL = "insert into WMS_STOCK_TRANS_REQ_D (" & _
                                                "TQ_code, imp_code, storer_code, " & _
                                                "TQD_seq, itm_code, pack_key, " & _
                                                "TQD_batch_no_fr, TQD_qty, TQD_org_qty, TQD_loc_fr, " & _
                                                "TQD_loc_to, TQD_rem, " & _
                                                "TQD_pallet_no_fr, TQD_pallet_no_to, TQD_batch_no_to," & _
                                                "TQD_SERIAL, TQD_UOM, TQD_UOM2, TQD_QTY2,TQD_ORG_QTY2,TQD_EXPIRY_DATE, TQD_MANU_DATE, TQD_ITM_NAME, TQD_DRUM_ID_FR, TQD_DRUM_ID_TO, TQD_DRUM_LV_FR, TQD_DRUM_LV_TO," & _
                                                "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                "values (" & _
                                                gU.convdbNVCData(gU.dbEncode(TQ_CODE.Text)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_seq").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("itm_code").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("pack_key").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_qty").ToString.Trim, "0")) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_org_qty").ToString.Trim, "0")) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_loc_fr").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_loc_to").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_rem").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_SERIAL").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_UOM").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_UOM2").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_QTY2").ToString.Trim, "0")) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_ORG_QTY2").ToString.Trim, "0")) & ", " & _
                                                gU.convdbDate(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_EXPIRY_DATE").ToString.Trim, ""))) & "," & _
                                                gU.convdbDate(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_MANU_DATE").ToString.Trim, ""))) & "," & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_ITM_NAME").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_FR").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_TO").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_DRUM_LV_FR").ToString.Trim, "NULL")) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_DRUM_LV_TO").ToString.Trim, "NULL")) & ", " & _
                                                "'" & Session("usr_id") & "',Getdate(),'" & Session("usr_id") & "',Getdate()) "

                                Case "D"
                                    itemSQL = "delete from WMS_STOCK_TRANS_REQ_D " & _
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and TQ_code = '" & gU.dbEncode(TQ_CODE.Text.Trim) & "' " & _
                                            "and TQD_seq = '" & gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    'itemSQL = "update WMS_STOCK_TRANS_REQ_D set " & _
                                    '            "TQD_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_seq").ToString.Trim, ""))) & ", " & _
                                    '            "itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("itm_code").ToString.Trim, ""))) & ", " & _
                                    '            "pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("pack_key").ToString.Trim, ""))) & ", " & _
                                    '            "TQD_batch_no_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                    '            "TQD_batch_no_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_batch_no_fr").ToString.Trim, ""))) & ", " & _
                                    '            "TQD_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_qty").ToString.Trim, "0")) & ", " & _
                                    '            "TQD_loc_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_loc_fr").ToString.Trim, ""))) & ", " & _
                                    '            "TQD_loc_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_loc_to").ToString.Trim, ""))) & ", " & _
                                    '            "TQD_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_rem").ToString.Trim, ""))) & ", " & _
                                    '            "TQD_pallet_no_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                    '            "TQD_pallet_no_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_fr").ToString.Trim, "000"))) & ", " & _
                                    '            "sys_lub = '" & Session("usr_id") & "', " & _
                                    '            "sys_lud = Getdate() " & _
                                    '        "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                    '        "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                    '        "and TQ_code = '" & gU.dbEncode(TQ_CODE.Text.Trim) & "' " & _
                                    '        "and TQD_seq = '" & gU.dbEncode(gU.decodeNullorEmpty(rows.Item("old_seq").ToString.Trim, "")) & "'"\

                                    itemSQL = "update WMS_STOCK_TRANS_REQ_D set " & _
                                                "TQD_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_qty").ToString.Trim, "0")) & ", " & _
                                                "TQD_org_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_org_qty").ToString.Trim, "0")) & ", " & _
                                                "TQD_loc_fr = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_loc_fr").ToString.Trim, ""))) & ", " & _
                                                "TQD_loc_to = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_loc_to").ToString.Trim, ""))) & ", " & _
                                                "TQD_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_rem").ToString.Trim, ""))) & ", " & _
                                                "TQD_QTY2 = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_QTY2").ToString.Trim, ""))) & ", " & _
                                                "TQD_SERIAL = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_SERIAL").ToString.Trim, ""))) & ", " & _
                                                "TQD_DRUM_ID_TO = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_TO").ToString.Trim, ""))) & ", " & _
                                                "TQD_DRUM_LV_TO=" & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("TQD_DRUM_LV_TO").ToString.Trim, "NULL")) & ", " & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = Getdate() " & _
                                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "and TQ_code = '" & gU.dbEncode(TQ_CODE.Text.Trim) & "' " & _
                                            "and TQD_seq = '" & gU.dbEncode(gU.decodeNullorEmpty(rows.Item("old_seq").ToString.Trim, "")) & "'"

                            End Select
                            REM **********************

                            'Response.Write(itemSQL)
                            If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)

                        Next
                    End If
                End If

                'Response.Write(sql_string)
                If sql_string <> "" Then gDB.amendData(sql_string, gConn, transaction)

                transaction.Commit()
                successFlag = True
                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    TQ_CODE.Text = nextNo
                    ViewState("TQ_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    TQ_CODE.ForeColor = Drawing.Color.Black
                    TQ_CODE.Font.Size = 10
                    'TQ_CODE.CssClass = ""
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

            If successFlag Then
                Call BindGV()
            End If
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
        If ViewState("TQ_CODE") <> "" Then
            pk_code = ViewState("TQ_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("TQ_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))
            ViewState("TQ_CODE") = pk_code
            ViewState("STORER_CODE") = storerCode
        End If
        REM **********************

        TQ_STATUS.ForeColor = Drawing.Color.Black

        If Session("PAGE_SESSION_MENU_CODE") = "OP_SRL" OrElse Session("PAGE_SESSION_MENU_CODE") = "OP_RD" Then
            TQ_WH_TO.Enabled = False
        End If


        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            TQ_STATUS.Text = "NEW"
            TQ_CODE.ForeColor = Drawing.Color.Red

            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = "select COnvert(varchar, WMS_STOCK_TRANS_REQ.TQ_DATE," & DDFORMAT & ") as TQ_DATE, convert(varchar,TQ_CLOSED_DATE," & DDFORMAT & "), WMS_STOCK_TRANS_REQ.* from WMS_STOCK_TRANS_REQ " & _
                        "where WMS_STOCK_TRANS_REQ.TQ_code = '" & gU.dbEncode(pk_code) & "' " & _
                        "and WMS_STOCK_TRANS_REQ.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        "and WMS_STOCK_TRANS_REQ.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                TQ_CODE.Text = dt.Rows(0).Item("TQ_code").ToString
                STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                TQ_STATUS.Text = dt.Rows(0).Item("TQ_STATUS").ToString
                TQ_DATE.Text = dt.Rows(0).Item("TQ_DATE").ToString
                TQ_BY.Text = dt.Rows(0).Item("TQ_BY").ToString
                TQ_BATCH_NO.Text = dt.Rows(0).Item("TQ_BATCH_NO").ToString
                TQ_REF_NO.Text = dt.Rows(0).Item("TQ_REF_NO").ToString
                TQ_WH_FR.SelectedValue = dt.Rows(0).Item("TQ_WH_FR").ToString
                TQ_WH_TO.SelectedValue = dt.Rows(0).Item("TQ_WH_TO").ToString
                TQ_TO_LOC.Value = dt.Rows(0).Item("TQ_TO_LOC").ToString
                dsp_TQ_TO_LOC.Text = dt.Rows(0).Item("TQ_TO_LOC").ToString
                TQ_REM.Text = dt.Rows(0).Item("TQ_REM").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)
                TQ_CLOSED_DATE.text = dt.Rows(0).Item("TQ_CLOSED_DATE").ToString
                TQ_EDI_RFT_NO.Text = dt.Rows(0).Item("TQ_EDI_RFT_NO").ToString

                If dt.Rows(0).Item("TQ_EDI_RFT_NO").ToString = "CSMS" Then
                    TQ_DATA_FR.Text = "Yes"
                Else
                    TQ_DATA_FR.Text = ""
                End If



                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , , True)
                End If

                If TQ_CODE.Text <> "" Then
                    'TQ_CODE.ReadOnly = True
                    'TQ_CODE.BorderWidth = 0
                    TQ_CODE.BackColor = Drawing.Color.Transparent
                End If

                If TQ_BY.Text <> "" Then
                    TQ_BY.ReadOnly = True
                    TQ_BY.BorderWidth = 0
                    TQ_BY.BackColor = Drawing.Color.Transparent
                End If

                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "SELECT  WMS_STOCK_TRANS_REQ_D.IMP_CODE, WMS_STOCK_TRANS_REQ_D.STORER_CODE, WMS_STOCK_TRANS_REQ_D.TQ_CODE, " & _
                    "WMS_STOCK_TRANS_REQ_D.ITM_CODE, WMS_STOCK_TRANS_REQ_D.PACK_KEY, WMS_STOCK_TRANS_REQ_D.TQD_SEQ,  " & _
                    "WMS_STOCK_TRANS_REQ_D.TQD_BATCH_NO, WMS_STOCK_TRANS_REQ_D.TQD_QTY, WMS_STOCK_TRANS_REQ_D.TQD_ORG_QTY, WMS_STOCK_TRANS_REQ_D.TQD_PALLET_NO_FR,  " & _
                    "WMS_STOCK_TRANS_REQ_D.TQD_PALLET_NO_TO, WMS_STOCK_TRANS_REQ_D.TQD_LOC_FR, WMS_STOCK_TRANS_REQ_D.TQD_LOC_TO,  " & _
                    "WMS_STOCK_TRANS_REQ_D.TQD_REM, WMS_STOCK_TRANS_REQ_D.SYS_LUB, WMS_STOCK_TRANS_REQ_D.SYS_LUD, WMS_STOCK_TRANS_REQ_D.SYS_CD,  " & _
                    "WMS_STOCK_TRANS_REQ_D.SYS_CB, WMS_STOCK_TRANS_REQ_D.TQD_BATCH_NO_FR, WMS_STOCK_TRANS_REQ_D.TQD_BATCH_NO_TO,  " & _
                    "WMS_STOCK_TRANS_REQ_D.TQD_SERIAL, WMS_STOCK_TRANS_REQ_D.TQD_UOM, WMS_STOCK_TRANS_REQ_D.TQD_UOM2,  " & _
                    "WMS_STOCK_TRANS_REQ_D.TQD_QTY2, WMS_STOCK_TRANS_REQ_D.TQD_ORG_QTY2, Convert(varchar,WMS_STOCK_TRANS_REQ_D.TQD_EXPIRY_DATE," & DDFORMAT & ") as TQD_EXPIRY_DATE,Convert(varchar,WMS_STOCK_TRANS_REQ_D.TQD_MANU_DATE," & DDFORMAT & ") as TQD_MANU_DATE,  " & _
                    "WMS_STOCK_TRANS_REQ_D.TQD_ITM_NAME, WMS_STOCK_TRANS_REQ_D.TQD_DRUM_ID_FR, WMS_STOCK_TRANS_REQ_D.TQD_DRUM_ID_TO,  " & _
                    "WMS_STOCK_TRANS_REQ_D.TQD_DRUM_LV_FR, WMS_STOCK_TRANS_REQ_D.TQD_DRUM_LV_TO, WMS_STOCK_TRANS_REQ_D.TQD_POST_QTY," & _
                    "wms_item.itm_name,wms_item.itm_sku_no, 'U' as mFlag, WMS_STOCK_TRANS_REQ_D.TQD_seq as old_seq " & _
                    "from WMS_STOCK_TRANS_REQ_D, wms_item " & _
                    "where WMS_STOCK_TRANS_REQ_D.imp_code = wms_item.imp_code " & _
                    "and WMS_STOCK_TRANS_REQ_D.storer_code = wms_item.storer_code " & _
                    "and WMS_STOCK_TRANS_REQ_D.itm_code = wms_item.itm_code " & _
                    "and WMS_STOCK_TRANS_REQ_D.PACK_KEY = wms_item.PACK_KEY " & _
                    "and WMS_STOCK_TRANS_REQ_D.TQ_code = '" & gU.dbEncode(pk_code) & "' " & _
                    "and WMS_STOCK_TRANS_REQ_D.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                    "and WMS_STOCK_TRANS_REQ_D.imp_code = '" & Session("IMP_CODE") & "'"

        SQLString = SQLString & " order by convert(int,WMS_STOCK_TRANS_REQ_D.TQD_seq)"
        REM **********************
        dt = gDB.getDataTable(SQLString)
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If
        ViewState("dt") = dt
        GridView1.DataBind()

        REM **********************
    End Sub

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        Dim cancelSql As String = "update WMS_STOCK_TRANS_REQ " & _
                         "set TQ_status = 'CANCELLED', " & _
                         "sys_lub = '" & Session("usr_id") & "', " & _
                         "sys_lud = Getdate() " & _
                         "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                         "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                         "and TQ_code = '" & gU.dbEncode(TQ_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql)

            transaction.Commit()

            TQ_STATUS.Text = "CANCELLED"

            ar.sec_write = "N"
            CancelBtn.Visible = False
            ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

            uiFun.displayMsg(Me, "1011", "", Session("gLang"))

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

    Protected Sub saveBtn1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn1.Click
        Call save()
    End Sub

    Protected Sub saveBtn2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles saveBtn2.Click
        Call save()
    End Sub

    Protected Sub btnPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPost.Click
        Dim updtSql, qtyString, imString As String
        Dim qtyTbl As New DataTable
        Dim imTbl As New DataTable
        Dim gConn As SqlConnection

        Dim MANU_DATE As String = ""
        Dim EXP_DATE As String = ""

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            If GridView1.Rows.Count > 0 Then
                If validateAll() Then
                    Call save("Y")

                    For Each rows As DataRow In dt.Rows

                        qtyString = "select ISNULL(max(ILOC_BAL_QTY), 0) as ILOC_BAL_QTY,  Convert(varchar,max(ILOC_EXPIRY_DATE)," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, max(ILOC_MANU_DATE)," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & _
                                    "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & _
                                    "' AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & _
                                    "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & _
                                    "' AND ILOC_LOC = '" & gU.dbEncode(rows.Item("TQD_loc_fr").ToString.Trim) & _
                                    "' AND isnull(ILOC_PALLET_NO,'000') = '" & gU.dbEncode(rows.Item("TQD_pallet_no_fr").ToString.Trim) & "' "

                        qtyTbl = gDB.getDataTable(qtyString, gConn, transaction)

                        If qtyTbl.Rows.Count > 0 Then
                            Dim locQty As String = gU.decodeEmptyCInt(qtyTbl.Rows(0).Item("ILOC_BAL_QTY").ToString, "0")

                            EXP_DATE = qtyTbl.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
                            MANU_DATE = qtyTbl.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim

                            'If gU.dbEncode(rows.Item("TQD_pallet_no_fr").ToString.Trim) = "000" Then
                            '    If CInt(gU.decodeNullorEmpty(rows.Item("TQD_qty").ToString.Trim, "0")) > CInt(locQty) Then
                            '        qtyString = "select ISNULL(max(ILOC_BAL_QTY), 0) from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & _
                            '       "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & _
                            '       "' AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & _
                            '       "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & _
                            '       "' AND ILOC_LOC = '" & gU.dbEncode(rows.Item("TQD_loc_fr").ToString.Trim) & _
                            '       "' AND ILOC_PALLET_NO = '' "

                            '        qtyTbl = gDB.getDataTable(qtyString, gConn, transaction)

                            '        locQty = gU.decodeEmptyCInt(qtyTbl.Rows(0).Item(0).ToString, "0")
                            '    End If
                            'End If

                            If CInt(gU.decodeNullorEmpty(rows.Item("TQD_qty").ToString.Trim, "0")) > CInt(locQty) Then
                                If transaction IsNot Nothing Then
                                    transaction.Rollback()
                                End If

                                If Session("gLang") = "E" Then
                                    uiFun.displayMsg(Me, "", "Transfer Qty cannot be Greater than this Location Stock Qty!!\r\nCurrent Location Qty: " & locQty, Session("gLang"))
                                Else
                                    uiFun.displayMsg(Me, "", "轉移數量不能大於這個位置儲存數量!!\r\n現時位置儲存數量: " & locQty, Session("gLang"))
                                End If

                                Exit Sub
                            Else
                                st.STORER_CODE = STORER_CODE.SelectedValue
                                st.ITM_CODE = gU.decodeNullorEmpty(rows.Item("itm_code").ToString.Trim, "")
                                st.PACK_KEY = gU.decodeNullorEmpty(rows.Item("pack_key").ToString.Trim, "")
                                st.IO_CUST_CODE = ""
                                st.IO_AREA = ""
                                st.IO_DOC = "STF"
                                st.IO_DOC_ID = TQ_CODE.Text.Trim
                                st.IO_QTY = gU.decodeEmptyCdbl(gU.decodeNullorEmpty(rows.Item("TQD_qty").ToString.Trim, ""), "0")
                                st.IO_CBM = 0
                                st.IO_KG = 0
                                st.IO_EXPIRY_DATE = EXP_DATE
                                st.IO_MANU_DATE = MANU_DATE

                                st.lO_BATCH_NO = gU.decodeNullorEmpty(rows.Item("TQD_batch_no_fr").ToString.Trim, "")

                                st.IO_WH = TQ_WH_FR.SelectedValue
                                st.IO_LOC = gU.decodeNullorEmpty(rows.Item("TQD_loc_fr").ToString.Trim, "")
                                st.PALLET_NO = gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_fr").ToString.Trim, "")

                                If rows.Item("TQD_SERIAL").ToString.Trim <> "" Then
                                    st.IOS_SERIAL_NO = rows.Item("TQD_SERIAL").ToString.Trim
                                    st.IOS_QTY2 = gU.decodeEmptyCdbl(gU.decodeNullorEmpty(rows.Item("TQD_qty2").ToString.Trim, ""), "0")
                                    'st.IOS_UOM2 = gU.decodeNullorEmpty(rows.Item("TQD_UOM2").ToString.Trim, "")
                                    Call st.setOrgSerialInfo(gU.decodeNullorEmpty(rows.Item("TQD_SERIAL").ToString.Trim, ""), gConn, transaction)
                                End If

                                st.UpdateStockTrans("OUT", gConn, transaction)
                                st.UpdateStockBalTrans("OUT", gConn, transaction)

                                If rows.Item("TQD_SERIAL").ToString.Trim <> "" Then
                                    st.UpdateStockSerialTrans("OUT", gConn, transaction)
                                    st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                                End If

                                If Session("PAGE_SESSION_MENU_CODE") <> "OP_RD" Then
                                    st.IO_LOC = gU.decodeNullorEmpty(rows.Item("TQD_loc_to").ToString.Trim, "")
                                    st.IO_WH = TQ_WH_TO.SelectedValue
                                    'st.PALLET_NO = gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_to").ToString.Trim, "")

                                    'st.lO_BATCH_NO = gU.decodeNullorEmpty(rows.Item("TQD_batch_no_to").ToString.Trim, "")

                                    st.UpdateStockTrans("IN", gConn, transaction)
                                    st.UpdateStockBalTrans("IN", gConn, transaction)


                                    If rows.Item("TQD_SERIAL").ToString.Trim <> "" Then
                                        If gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_TO").ToString.Trim, "") <> "" Then
                                            st.IOS_DRUM_ID = gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_TO").ToString.Trim, "")
                                        Else
                                            st.IOS_DRUM_ID = gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_FR").ToString.Trim, "")
                                        End If

                                        If gU.decodeNullorEmpty(rows.Item("TQD_DRUM_LV_TO").ToString.Trim, "") <> "" Then
                                            st.IOS_DRUM_LEVEL = gU.decodeNullorEmpty(rows.Item("TQD_DRUM_LV_TO").ToString.Trim, "")
                                        Else
                                            st.IOS_DRUM_LEVEL = gU.decodeNullorEmpty(rows.Item("TQD_DRUM_LV_FR").ToString.Trim, "")
                                        End If
                                        st.IOS_SERIAL_NO = rows.Item("TQD_SERIAL").ToString.Trim
                                        st.IOS_QTY2 = gU.decodeEmptyCdbl(gU.decodeNullorEmpty(rows.Item("TQD_qty2").ToString.Trim, ""), "0")
                                        st.IOS_UOM2 = gU.decodeNullorEmpty(rows.Item("TQD_UOM2").ToString.Trim, "")

                                        st.UpdateStockSerialTrans("IN", gConn, transaction)
                                        st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                                    End If
                                End If
                            End If
                        Else
                            'imString = "SELECT M.*, D.* " & _
                            '            "from WMS_ITEM M, WMS_ALT_VEND_ITEM D " & _
                            '            "where M.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            '            "and M.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
                            '            "AND M.ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & "' " & _
                            '            "AND M.PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & "' " & _
                            '            "AND M.IMP_CODE = D.IMP_CODE(+) " & _
                            '            "AND M.STORER_CODE = D.STORER_CODE(+) " & _
                            '            "AND M.PACK_KEY = D.PACK_KEY(+) " & _
                            '            "AND M.ITM_CODE = D.ITM_CODE(+)"

                            'imTbl = gDB.getDataTable(imString)

                            'Dim balQty As String

                            'If imTbl.Rows.Count > 0 Then
                            '    balQty = imTbl.Rows(0).Item("ITM_BALANCE").ToString
                            'Else
                            '    balQty = "0"
                            'End If

                            'If gU.decodeEmptyCInt(gU.decodeNullorEmpty(rows.Item("TQD_qty").ToString.Trim, ""), "0") <= CInt(balQty) Then
                            '    st.STORER_CODE = STORER_CODE.SelectedValue
                            '    st.ITM_CODE = gU.decodeNullorEmpty(rows.Item("itm_code").ToString.Trim, "")
                            '    st.PACK_KEY = gU.decodeNullorEmpty(rows.Item("pack_key").ToString.Trim, "")
                            '    st.IO_CUST_CODE = ""
                            '    st.IO_AREA = ""
                            '    st.IO_DOC = "STF"
                            '    st.IO_DOC_ID = TQ_CODE.Text.Trim
                            '    st.IO_QTY = gU.decodeEmptyCInt(gU.decodeNullorEmpty(rows.Item("TQD_qty").ToString.Trim, ""), "0")
                            '    st.IO_CBM = 0
                            '    st.IO_KG = 0

                            '    st.lO_BATCH_NO = gU.decodeNullorEmpty(rows.Item("TQD_batch_no_fr").ToString.Trim, "")

                            '    st.IO_WH = TQ_WH_FR.SelectedValue
                            '    st.IO_LOC = gU.decodeNullorEmpty(rows.Item("TQD_loc_fr").ToString.Trim, "")
                            '    st.PALLET_NO = gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_fr").ToString.Trim, "")
                            '    st.UpdateStockTrans("OUT", gConn, transaction)
                            '    st.UpdateStockBalTrans("OUT", gConn, transaction)

                            '    st.IO_LOC = gU.decodeNullorEmpty(rows.Item("TQD_loc_to").ToString.Trim, "")
                            '    st.IO_WH = TQ_WH_TO.SelectedValue
                            '    'st.PALLET_NO = gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_to").ToString.Trim, "")

                            '    'st.lO_BATCH_NO = gU.decodeNullorEmpty(rows.Item("TQD_batch_no_to").ToString.Trim, "")

                            '    st.UpdateStockTrans("IN", gConn, transaction)
                            '    st.UpdateStockBalTrans("IN", gConn, transaction)
                            'Else
                            If transaction IsNot Nothing Then
                                transaction.Rollback()
                            End If

                            If Session("gLang") = "E" Then
                                uiFun.displayMsg(Me, "", "Transfer Qty cannot be Greater than Balance Qty!! \r\nCurrent Balance Qty: 0", Session("gLang"))
                            Else
                                uiFun.displayMsg(Me, "", "轉移數量不能大於結餘數量\r\n現時結餘數量: 0", Session("gLang"))
                            End If

                            Exit Sub
                            'End If

                        End If
                    Next

                    If Session("PAGE_SESSION_MENU_CODE") = "OP_RD" Then
                        For Each rows As DataRow In dt.Rows

                            qtyString = "select ISNULL(max(ILOC_BAL_QTY), 0) as ILOC_BAL_QTY,  Convert(varchar,max(ILOC_EXPIRY_DATE)," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, max(ILOC_MANU_DATE)," & DDFORMAT & ") as ILOC_MANU_DATE from WMS_ITEM_LOC_BAL WHERE IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & _
                                        "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & _
                                        "' AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & _
                                        "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & _
                                        "' AND ILOC_LOC = '" & gU.dbEncode(rows.Item("TQD_loc_fr").ToString.Trim) & _
                                        "' AND isnull(ILOC_PALLET_NO,'000') = '" & gU.dbEncode(rows.Item("TQD_pallet_no_fr").ToString.Trim) & "' "

                            qtyTbl = gDB.getDataTable(qtyString, gConn, transaction)

                            If qtyTbl.Rows.Count > 0 Then
                                Dim locQty As String = gU.decodeEmptyCInt(qtyTbl.Rows(0).Item("ILOC_BAL_QTY").ToString, "0")

                                EXP_DATE = qtyTbl.Rows(0).Item("ILOC_EXPIRY_DATE").ToString.Trim
                                MANU_DATE = qtyTbl.Rows(0).Item("ILOC_MANU_DATE").ToString.Trim

                                st.STORER_CODE = STORER_CODE.SelectedValue
                                st.ITM_CODE = gU.decodeNullorEmpty(rows.Item("itm_code").ToString.Trim, "")
                                st.PACK_KEY = gU.decodeNullorEmpty(rows.Item("pack_key").ToString.Trim, "")
                                st.IO_CUST_CODE = ""
                                st.IO_AREA = ""
                                st.IO_DOC = "STF"
                                st.IO_DOC_ID = TQ_CODE.Text.Trim
                                st.IO_QTY = gU.decodeEmptyCdbl(gU.decodeNullorEmpty(rows.Item("TQD_qty").ToString.Trim, ""), "0")
                                st.IO_CBM = 0
                                st.IO_KG = 0
                                st.IO_EXPIRY_DATE = EXP_DATE
                                st.IO_MANU_DATE = MANU_DATE
                                st.lO_BATCH_NO = gU.decodeNullorEmpty(rows.Item("TQD_batch_no_fr").ToString.Trim, "")
                                st.PALLET_NO = gU.decodeNullorEmpty(rows.Item("TQD_pallet_no_fr").ToString.Trim, "")
                                st.IO_LOC = gU.decodeNullorEmpty(rows.Item("TQD_loc_to").ToString.Trim, "")
                                st.IO_WH = TQ_WH_TO.SelectedValue

                                st.UpdateStockTrans("IN", gConn, transaction)
                                st.UpdateStockBalTrans("IN", gConn, transaction)

                                If rows.Item("TQD_SERIAL").ToString.Trim <> "" Then
                                    If gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_TO").ToString.Trim, "") <> "" Then
                                        st.IOS_DRUM_ID = gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_TO").ToString.Trim, "")
                                    Else
                                        st.IOS_DRUM_ID = gU.decodeNullorEmpty(rows.Item("TQD_DRUM_ID_FR").ToString.Trim, "")
                                    End If

                                    If gU.decodeNullorEmpty(rows.Item("TQD_DRUM_LV_TO").ToString.Trim, "") <> "" Then
                                        st.IOS_DRUM_LEVEL = gU.decodeNullorEmpty(rows.Item("TQD_DRUM_LV_TO").ToString.Trim, "")
                                    Else
                                        st.IOS_DRUM_LEVEL = gU.decodeNullorEmpty(rows.Item("TQD_DRUM_LV_FR").ToString.Trim, "")
                                    End If
                                    st.IOS_SERIAL_NO = rows.Item("TQD_SERIAL").ToString.Trim
                                    st.IOS_QTY2 = gU.decodeEmptyCdbl(gU.decodeNullorEmpty(rows.Item("TQD_qty2").ToString.Trim, ""), "0")
                                    st.IOS_UOM2 = gU.decodeNullorEmpty(rows.Item("TQD_UOM2").ToString.Trim, "")

                                    st.UpdateStockSerialTrans("IN", gConn, transaction)
                                    st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                                End If

                            Else
                                If transaction IsNot Nothing Then
                                    transaction.Rollback()
                                End If

                                If Session("gLang") = "E" Then
                                    uiFun.displayMsg(Me, "", "Transfer Qty cannot be Greater than Balance Qty!! \r\nCurrent Balance Qty: 0", Session("gLang"))
                                Else
                                    uiFun.displayMsg(Me, "", "轉移數量不能大於結餘數量\r\n現時結餘數量: 0", Session("gLang"))
                                End If

                                Exit Sub
                                'End If

                            End If
                        Next
                    End If

                    updtSql = "update WMS_STOCK_TRANS_REQ_D " & _
                                "set TQ_status = 'POSTED', " & _
                                "sys_lub = '" & Session("usr_id") & "', " & _
                                "sys_lud = Getdate() " & _
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and TQ_code = '" & gU.dbEncode(TQ_CODE.Text) & "' "

                    gDB.amendData(updtSql, gConn, transaction)

                    transaction.Commit()

                    TQ_STATUS.Text = "POSTED"
                    'ar.sec_write = "N"
                    ar.sec_viewMode = "Y"
                    btnPost.Visible = False
                    ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

                    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                End If
            Else
                If Session("gLang") = "E" Then
                    uiFun.displayMsg(Me, "", "No item can be posted!", Session("gLang"))
                Else
                    uiFun.displayMsg(Me, "", "沒有可供發布的物件!", Session("gLang"))
                End If
            End If

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

    Protected Sub addItemtoSTF()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(TQD_SEQ AS int)) + 1 from WMS_STOCK_TRANS_REQ_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and TQ_CODE = '" & gU.dbEncode(TQ_CODE.Text.Trim) & "' "
            REM **********************
            Dim nS_dt As New DataTable
            nS_dt = gDB.getDataTable(seq_string)
            Dim next_seq_no, temp_no As String
            Dim temp_seq_no As Integer = 1

            next_seq_no = ""

            If nS_dt.Rows.Count > 0 Then
                next_seq_no = nS_dt.Rows(0).Item(0).ToString()
            Else
                temp_no = "1"
            End If

            If next_seq_no = "" Then next_seq_no = "1"

            Dim SQLString As String
            Dim stf_dt As DataTable
            Dim addSQL As String = ""

            Dim itemPackList As String = ""
            Dim itemListarray As String()
            Dim packKeyListarray As String()
            Dim qtyListarray As String()

            itemListarray = Split(itemList.Value, ", ")
            packKeyListarray = Split(packKeyList.Value, ", ")
            qtyListarray = Split(qtyList.Value, ", ")

            Dim qtyitemDict As New Dictionary(Of String, String)

            If qtyList.Value <> "" Then
                For i As Integer = 0 To itemListarray.Count - 1
                    qtyitemDict.Add(itemListarray(i) & "_000_" & packKeyListarray(i), qtyListarray(i))
                Next
            End If

            If itemListarray.Count = 0 Then
                If itemList.Value <> "" Then
                    itemPackList = Server.HtmlDecode(itemList.Value) & "_000_" & Server.HtmlDecode(packKeyList.Value)
                End If
            Else
                For i = 0 To itemListarray.Count - 1
                    itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemListarray(i)) & "_000_" & Server.HtmlDecode(packKeyListarray(i)))
                Next
            End If


            SQLString = ""
            SQLString += "select wms_item.*, wms_alt_vend_item.* "
            SQLString += "from wms_item left outer join wms_alt_vend_item on "
            SQLString += "wms_item.imp_code = wms_alt_vend_item.imp_code "
            SQLString += "and wms_item.storer_code = wms_alt_vend_item.storer_code "
            SQLString += "and wms_item.itm_code = wms_alt_vend_item.itm_code "
            SQLString += "and wms_item.pack_key = wms_alt_vend_item.pack_key "
            SQLString += "where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
            SQLString += "and wms_item.imp_code = '" & Session("IMP_CODE") & "' "
            SQLString += "and wms_item.itm_code + '_000_' + wms_item.pack_key + '_000_' + ISNULL(wms_alt_vend_item.vnd_code,'000') in ('" & Replace(itemPackList, ", ", "', '") & "') "

            SQLString = SQLString & " order by wms_item.itm_code, wms_item.pack_key, wms_alt_vend_item.vnd_code"
            REM **********************
            stf_dt = gDB.getDataTable(SQLString)

            For i As Integer = 0 To stf_dt.Rows.Count - 1
                If ViewState("n_cur_seq") = "" Then
                    ViewState("n_cur_seq") = next_seq_no
                Else
                    temp_seq_no = CInt(ViewState("n_cur_seq")) + 1
                    ViewState("n_cur_seq") = temp_seq_no.ToString
                End If

                dt.Rows.Add()

                rows_count = dt.Rows.Count

                Dim itmQty As Integer = 0
                If qtyitemDict.ContainsKey(stf_dt.Rows(i).Item("ITM_CODE").ToString & "_000_" & stf_dt.Rows(i).Item("PACK_KEY").ToString & "_000_" & gU.decodeNullOrEmpty(stf_dt.Rows(i).Item("vnd_code").ToString, "000")) Then
                    itmQty = gU.decodeEmptyCInt(qtyitemDict(stf_dt.Rows(i).Item("ITM_CODE").ToString & "_000_" & stf_dt.Rows(i).Item("PACK_KEY").ToString & "_000_" & gU.decodeNullOrEmpty(stf_dt.Rows(i).Item("vnd_code").ToString, "000")), 0)
                End If

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("TQD_SEQ") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("ITM_CODE") = stf_dt.Rows(i).Item("ITM_CODE")
                dt.Rows(rows_count - 1).Item("ITM_NAME") = stf_dt.Rows(i).Item("ITM_NAME")
                dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = stf_dt.Rows(i).Item("ITM_SKU_NO")
                dt.Rows(rows_count - 1).Item("PACK_KEY") = stf_dt.Rows(i).Item("PACK_KEY")
                
                dt.Rows(rows_count - 1).Item("TQD_ITM_NAME") = stf_dt.Rows(i).Item("ITM_NAME")
                'dt.Rows(rows_count - 1).Item("TQD_PALLET_NO_FR") = stf_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                'dt.Rows(rows_count - 1).Item("TQD_PALLET_NO_TO") = stf_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                'dt.Rows(rows_count - 1).Item("TQD_BATCH_NO_FR") = stf_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                'dt.Rows(rows_count - 1).Item("TQD_BATCH_NO_TO") = stf_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                'dt.Rows(rows_count - 1).Item("TQD_SERIAL") = stf_dt.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("TQD_UOM") = stf_dt.Rows(i).Item("ITM_UOM")
                dt.Rows(rows_count - 1).Item("TQD_UOM2") = stf_dt.Rows(i).Item("ITM_UOM")
                dt.Rows(rows_count - 1).Item("TQD_QTY2") = stf_dt.Rows(i).Item("ITM_QTY2")
                dt.Rows(rows_count - 1).Item("TQD_ORG_QTY2") = stf_dt.Rows(i).Item("ITM_QTY2")
                'dt.Rows(rows_count - 1).Item("TOD_VND_CODE") = stf_dt.Rows(i).Item("VND_CODE").ToString.Trim
                'dt.Rows(rows_count - 1).Item("TQD_ORG_QTY") = ""

                If stf_dt.Rows(i).Item("ITM_TYPE").ToString.Trim = "CABLE" Then
                    dt.Rows(rows_count - 1).Item("TQD_QTY") = 1
                Else
                    dt.Rows(rows_count - 1).Item("TQD_QTY") = itmQty
                End If


                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

            Next

            STORER_CODE.Enabled = False

            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub

    Protected Sub selectItemBtn_Click(sender As Object, e As System.EventArgs) Handles selectItemBtn.Click
        'selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.getElementById('" & STORER_CODE.ClientID & "').value, document.getElementById('" & TQ_WH_FR.ClientID & "').value);")
        ScriptManager.RegisterStartupScript(LOOKUPUDP, LOOKUPUDP.GetType, "itemLookup", "ItemLookUp(document.getElementById('" & STORER_CODE.ClientID & "').value,document.getElementById('" & TQ_WH_FR.ClientID & "').value);", True)

    End Sub

    Protected Sub btnClose_Click(sender As Object, e As System.EventArgs) Handles btnClose.Click
        Dim cancelSql As String = "update WMS_STOCK_TRANS_REQ " & _
                       "set TQ_status = 'CLOSED', " & _
                       "sys_lub = '" & Session("usr_id") & "', " & _
                       "sys_lud = Getdate() " & _
                       "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                       "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                       "and TQ_code = '" & gU.dbEncode(TQ_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()

            TQ_STATUS.Text = "CLOSED"

            ar.sec_viewMode = "Y"
            btnClose.Visible = False
            setPageCtrlAccess()
            ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl, exceptionEditList)

            uiFun.displayMsg(Me, "", "This Transfer Request has been closed.", Session("gLang"))

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
End Class
