Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class OPERATION_WO_Main
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
    Private r_dt As DataTable
    Private DDFORMAT As String = "DD/MM/YYYY"

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
            Session("pagemode") = Nothing
            Session("pagemode") = Request("mode")

            If Session("usr_pref_storer") <> "" Then
                uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' AND STORER_CODE = '" & Session("usr_pref_storer") & "' ORDER BY 2", "STORER_CODE", "STO_NAME", , , Session("gSelectLabel"), True)
            Else
                uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME AS STO_NAME FROM WMS_STORER WHERE STO_STATUS = 'ACTIVE' ORDER BY 2", "STORER_CODE", "STO_NAME", , Session("gSelectLabel"))
            End If

            uiFun.load_dropdownBy_ColCode(WO_TYPE, "WMS_WORK_ORDER.WO_TYPE", Session("gSelectLabel"))


        End If

        If Session("pagemode") = "N" Then
            CancelBtn.Visible = False
            If STORER_CODE.SelectedValue = "" Then
                STORER_CODE.SelectedValue = Session("usr_pref_storer")
            End If
            If WO_DATE.Text = "" Then
                WO_DATE.Text = Now().Date.ToString("dd/MM/yyyy")
            End If
        End If

        REM ****************************

        REM **********************
        REM Modify Here
        If Session("gLang") = "E" Then
            lheader.Text = "Work Order"
            lbl_ImageHd.Text = "Work Order Items"
            lbl_WO_CODE.Text = "Work Order Code:"
            lbl_WO_STATUS.Text = "Status:"
            lbl_STORER_CODE.Text = "Storer:"
            lbl_WO_DATE.Text = "Date:"
            lbl_WO_TR_CODE.Text = "Stock Relocation No.:"
            lbl_WO_REM.Text = "Remarks"

            lbl_sys_cb.Text = "CB"
            lbl_sys_lub.Text = "LUB"
            lbl_sys_cd.Text = "CD"
            lbl_sys_lud.Text = "LUD"
            saveBtn1.Text = "Save"
            saveBtn2.Text = "Save"
            CancelBtn.Text = "Cancel"

            btnPost.Text = "Post"
            selectItemBtn.Text = "Select Item"
            CancelBtn.OnClientClick = "return confirm(""Are you sure to cancel this record?"");"
            saveBtn1.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            saveBtn2.OnClientClick = "return confirm(""Are you sure to save this record?"");"
            btnPost.OnClientClick = "return confirm(""Are you sure to post this record?\r\n(Please save your work before Posting)"");"
            If Session("pagemode") = "N" Then
                WO_CODE.Text = "[No. will be auto generated]"
            End If

        ElseIf Session("gLang") = "C" Then
            lheader.Text = "貨品轉移維護"
            lbl_ImageHd.Text = "貨品詳情"
            lbl_WO_CODE.Text = "轉移號碼:"
            lbl_WO_STATUS.Text = "狀態:"
            lbl_STORER_CODE.Text = "貨主:"
            lbl_WO_DATE.Text = "日期:"
            lbl_WO_TR_CODE.Text = "Stock Relocation No.:"
            lbl_WO_REM.Text = "備註:"

            lbl_sys_cb.Text = "創建者"
            lbl_sys_lub.Text = "最後更新者"
            lbl_sys_cd.Text = "創建日期"
            lbl_sys_lud.Text = "最後更新日期"
            saveBtn1.Text = "儲存"
            saveBtn2.Text = "儲存"
            CancelBtn.Text = "取消"

            btnPost.Text = "發布"
            selectItemBtn.Text = "選擇物料"
            CancelBtn.OnClientClick = "return confirm(""確定取消資料?"");"
            saveBtn1.OnClientClick = "return confirm(""確定儲存資料?"");"
            saveBtn2.OnClientClick = "return confirm(""確定儲存資料?"");"
            btnPost.OnClientClick = "return confirm(""確定發布資料?"");"
            If Session("pagemode") = "N" Then
                WO_CODE.Text = "[號碼會自動產生]"
            End If
        End If
        REM **********************

        REM **********************
        REM Additional CSS

        WO_DATE.CssClass = "REQUIRED"
        REM **********************

        If Session("pagemode") = "N" Then
            'WO_CODE.CssClass = "REQUIRED"
            STORER_CODE.CssClass = "REQUIRED"
        Else
            'STORER_CODE.Enabled = False
        End If

        If Not IsPostBack Then
            ViewState("dt") = Nothing

            ViewState("n_cur_seq") = ""
            ViewState("WO_CODE") = ""
            ViewState("WOD_LOC") = Nothing

            ViewState("r_dt") = Nothing

            Call BindGV()
        Else
            dt = ViewState("dt")
            r_dt = ViewState("r_dt")
        End If

        If moduleAction = "SELECTIM" Then
            addItemtoWO()
            'ElseIf moduleAction = "SELECTTOWH" Then
            'changeToPallet()
        ElseIf moduleAction = "SELECTIMLOC" Then
            addBalItemtoWO()
        End If

        'cm = New CommonMenu("WO", lheader.Text, WO_CODE.Text)
        'cm.parentDir = "../../"
        'cm.haveCheckList = "N"
        'cm.haveAttachments = "N"
        'cm.haveNotes = "N"
        'cm.haveTasks = "N"
        'cm.haveEmail = "N"
        'cm.haveHistory = "N"

        'cm.genCM(cmBar)

        selectItemBtn.Attributes.Add("onclick", "ItemLookUp(document.myform." & STORER_CODE.ClientID & ".value);return false;")
        selectLocBtn.Attributes.Add("onclick", "ItemLocBalLookUp(document.myform." & STORER_CODE.ClientID & ".value);return false;")

        If WO_TR_CODE.Text <> "" Then
            btnGenSR.Visible = False
        End If

        If WO_STATUS.Text = "CANCELLED" Then
            ar.sec_write = "N"
            ar.sec_viewMode = "Y"
            CancelBtn.Visible = False
        ElseIf WO_STATUS.Text = "POSTED" Then
            ar.sec_viewMode = "Y"
            btnPost.Visible = False
        End If

        ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)
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

    Protected Sub GridView1_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
         Select e.CommandName
            Case "RFID"
                Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)
                If Not IsNothing(gvRow) Then

                    Dim wod_seq As String = DirectCast(gvRow.FindControl("WOD_SEQ"), HiddenField).Value
                    Dim wod_qty As Double = gU.decodeEmptyCdbl(DirectCast(gvRow.FindControl("WOD_qty"), TextBox).Text.Trim, 0)
                    Call BindRFID(wod_seq, wod_qty)
                End If

        End Select
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("itm_code"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_code").ToString.Trim
                CType(e.Row.FindControl("pack_key"), Label).Text = DataBinder.Eval(e.Row.DataItem, "pack_key").ToString.Trim


                CType(e.Row.FindControl("WOD_batch_no"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "WOD_batch_no").ToString.Trim

                CType(e.Row.FindControl("itm_name"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_name").ToString.Trim
                CType(e.Row.FindControl("WOD_qty"), TextBox).Text = cU.FormatIntegerString(DataBinder.Eval(e.Row.DataItem, "WOD_qty").ToString.Trim)

                CType(e.Row.FindControl("wod_type"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "wod_type").ToString.Trim

                'CType(e.Row.FindControl("WOD_loc"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "WOD_loc").ToString.Trim
                CType(e.Row.FindControl("WOD_seq"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "WOD_seq").ToString.Trim
                CType(e.Row.FindControl("itm_sku_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "itm_sku_no").ToString.Trim
                'CType(e.Row.FindControl("dsp_WOD_loc"), Label).Text = DataBinder.Eval(e.Row.DataItem, "WOD_loc").ToString.Trim
                Dim nDropDown As DropDownList = CType(e.Row.FindControl("wod_loc"), DropDownList)
                uiFun.load_dropdown(nDropDown, "Select distinct Convert(nvarchar(20),a.FL_NUM)+b.AR_CODE+c.RK_CODE+d.BN_CODE as CODE,a.WH_CODE+a.FL_NAME+b.AR_CODE+c.RK_CODE+d.BN_CODE as NAME from WMS_WH_FL a Inner join WMS_WH_AREA b on a.FL_NUM=b.FL_NUM Inner join WMS_WH_RACK c on b.AR_CODE=c.AR_CODE Inner join WMS_WH_BIN d on c.RK_CODE=d.RK_CODE")
                nDropDown.SelectedValue = DataBinder.Eval(e.Row.DataItem, "wod_loc").ToString.Trim

                CType(e.Row.FindControl("WOD_pallet_no"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "WOD_pallet_no").ToString.Trim
                CType(e.Row.FindControl("WOD_rem"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "WOD_rem").ToString.Trim

                CType(e.Row.FindControl("WOD_KIT_QTY"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "WOD_KIT_QTY").ToString.Trim
                CType(e.Row.FindControl("WOD_SERIAL_NO"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "WOD_SERIAL_NO").ToString.Trim
                CType(e.Row.FindControl("WOD_DRUM_ID"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "WOD_DRUM_ID").ToString.Trim
                CType(e.Row.FindControl("WOD_DRUM_LV"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "WOD_DRUM_LV").ToString.Trim
                CType(e.Row.FindControl("WOD_qty2"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "WOD_qty2").ToString.Trim

                CType(e.Row.FindControl("itm_uom"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_uom").ToString.Trim
                CType(e.Row.FindControl("itm_uom2"), Label).Text = DataBinder.Eval(e.Row.DataItem, "itm_uom2").ToString.Trim

                CType(e.Row.FindControl("RFID_COUNT"), Label).Text = DataBinder.Eval(e.Row.DataItem, "RFID_COUNT").ToString.Trim

                'Dim tImage As Image = CType(e.Row.FindControl("Image_Loc_LookUp_TO"), Image)
                'tImage.Attributes.Add("onmousedown", "MM_swapImage('" & HttpUtility.HtmlEncode(tImage.ClientID) & "','','../../images/btn_search_over.gif',1)")
                'tImage.Attributes.Add("onclick", "LocLookUp(2,'" & e.Row.RowIndex & "','" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("dsp_wod_loc"), Label).ClientID) & "', '" & HttpUtility.HtmlEncode(CType(e.Row.FindControl("wod_loc"), HiddenField).ClientID) & "')")

                If DataBinder.Eval(e.Row.DataItem, "wod_type").ToString.Trim = "IN" Then
                    DirectCast(e.Row.FindControl("btnRFID"), Button).Visible = True
                    CType(e.Row.FindControl("RFID_COUNT"), Label).Visible = True

                    DirectCast(e.Row.FindControl("wod_qty"), TextBox).Attributes.Add("onchange", "javascript:countKit(this);")
                Else
                    DirectCast(e.Row.FindControl("btnRFID"), Button).Visible = False
                    CType(e.Row.FindControl("RFID_COUNT"), Label).Visible = False
                End If

                Dim nButton As Button = CType(e.Row.FindControl("btnDelete"), Button)

                If Session("gLang") = "E" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('Are you sure you want to delete this record?')")
                    nButton.Text = "Delete"
                ElseIf Session("gLang") = "C" Then
                    nButton.Attributes.Add("onclick", "javascript:return confirm('你是否確定要刪除這個資料?')")
                    nButton.Text = "删除"
                End If


                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "D" Then
                    Call ar.hideGVRow(GridView1, e.Row)
                End If

                If DataBinder.Eval(e.Row.DataItem, "mFlag").ToString.Trim = "N" Then
                    'nButton.Enabled = False
                End If
        End Select
    End Sub


    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Call ar.hideGVRow(GridView1, GridView1.Rows(e.RowIndex))
        dt.Rows(e.RowIndex).Item("mFlag") = "D"
        dt.AcceptChanges()
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

        If WO_DATE.Text.Trim = "" Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", lbl_WO_DATE.Text & " cannot be empty!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", lbl_WO_DATE.Text & "不能空白!", Session("gLang"))
            End If

            Return False
        ElseIf Not gU.isValidDate(WO_DATE.Text.Trim) Then
            If Session("gLang") = "E" Then
                uiFun.displayMsg(Me, "", "Invalid date, " & lbl_WO_DATE.Text & "!", Session("gLang"))
            Else
                uiFun.displayMsg(Me, "", "無效的日期, " & lbl_WO_DATE.Text & "!", Session("gLang"))
            End If
            Return False
        End If

        If GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                Dim frloc As String = CType(GridView1.Rows(i).FindControl("WOD_LOC"), DropDownList).SelectedValue

                If frloc = "" And 1 = 2 Then
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Location Cannot Be Empty!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "位置不能空白!", Session("gLang"))
                    End If

                    'reloadHiddenValue(GridView1)

                    Return False

                End If

                If CType(GridView1.Rows(i).FindControl("WOD_qty"), TextBox).Text <> "" Then
                    If uiFun.gvValidate(Me, dt, "wod_qty", "Qty", _
                                     CType(GridView1.Rows(i).FindControl("wod_qty"), TextBox).Text) = False Then Return False


                    If CInt(CType(GridView1.Rows(i).FindControl("WOD_qty"), TextBox).Text) <= 0 And 1 = 2 Then
                        If Session("gLang") = "E" Then
                            uiFun.displayMsg(Me, "", "Item Qty Cannot Be Zero!", Session("gLang"))
                        Else
                            uiFun.displayMsg(Me, "", "物料數量不能零!", Session("gLang"))
                        End If

                        'reloadHiddenValue(GridView1)

                        Return False
                    End If
                Else
                    If Session("gLang") = "E" Then
                        uiFun.displayMsg(Me, "", "Item Qty Cannot Be Empty!", Session("gLang"))
                    Else
                        uiFun.displayMsg(Me, "", "物料數量不能空白!", Session("gLang"))
                    End If

                    'reloadHiddenValue(GridView1)

                    Return False
                End If
            Next
        End If

        Return True

    End Function

    'Private Sub reloadHiddenValue(ByRef gv As GridView)
    '    For i = 0 To gv.Rows.Count - 1

    '        'CType(gv.Rows(i).FindControl("dsp_WOD_batch_no"), Label).Text = CType(GridView1.Rows(i).FindControl("WOD_batch_no"), HiddenField).Value
    '        CType(gv.Rows(i).FindControl("dsp_WOD_LOC"), Label).Text = CType(GridView1.Rows(i).FindControl("WOD_LOC"), HiddenField).Value
    '        'CType(gv.Rows(i).FindControl("dsp_WOD_pallet_no"), Label).Text = CType(GridView1.Rows(i).FindControl("WOD_pallet_no"), HiddenField).Value
    '    Next
    'End Sub

    Protected Sub save(Optional ByVal flag As String = "")
        Dim alertstr As String = ""
        Dim sql_string As String = ""
        Dim itemSQL As String = ""
        Dim nextNo As String = ""
        Dim gConn As SqlConnection
        Dim dupSQL As String = ""
        Dim dupTbl As New DataTable

        Dim paP As New GlobalDBFunc.DBCmdPara

        If validateAll() Then

            gConn = gDB.getConnection()

            Dim transaction As SqlTransaction
            transaction = gConn.BeginTransaction()
            ' Start a local transaction
            Try
                If Session("pagemode") = "N" Then
                    REM **********************
                    REM Modify Here
                    nextNo = DB.getDocNo("WO", gConn, transaction)
                    'nextNo = WO_CODE.Text
                    REM **********************

                    dupSQL = "select 1 from wms_work_order " & _
                                "where WO_code = '" & gU.dbEncode(nextNo) & "' " & _
                                "and imp_code= '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                "and storer_code= '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

                    dupTbl = gDB.getDataTable(dupSQL)

                    REM **********************
                    REM Modify Here
                    REM char "N" is use for update Unicode
                    If dupTbl.Rows.Count = 0 Then
                        sql_string = "insert into wms_work_order (" & _
                        "WO_code, imp_code, storer_code, " & _
                        "WO_status, WO_date, WO_TR_CODE, WO_type, WO_rem, " & _
                         "sys_cb, sys_cd, sys_lub, sys_lud)" & _
                        "values ( " & _
                        gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(WO_STATUS.Text.Trim)) & "," & gU.convdbDate(gU.dbEncode(WO_DATE.Text.Trim)) & "," & gU.convdbDate(gU.dbEncode(WO_TR_CODE.Text.Trim)) & "," & gU.convdbNVCData(gU.dbEncode(WO_TYPE.SelectedValue)) & "," & _
                        gU.convdbNVCData(gU.dbEncode(WO_REM.Text.Trim)) & "," & _
                         "'" & Session("usr_id") & "',getdate(),'" & Session("usr_id") & "',getdate()) "

                        REM **********************

                        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                            'uiFun.reOrderDetails(dt, "wod_seq")

                            For Each rows As DataRow In dt.Rows
                                itemSQL = ""

                                REM **********************
                                REM Modify Here
                                Select Case rows.Item("mFlag")
                                    Case "N"
                                        itemSQL = "insert into wms_work_order_d (" & _
                                                "WO_code, imp_code, storer_code, " & _
                                                "WOD_seq, itm_code, pack_key, " & _
                                                "WOD_batch_no, WOD_qty, WOD_LOC, WOD_rem, WOD_pallet_no, WOD_TYPE," & _
                                                "WOD_KIT_QTY, WOD_SERIAL_NO, WOD_DRUM_ID, WOD_DRUM_LV, WOD_QTY2," & _
                                                "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                "values (" & _
                                                gU.convdbNVCData(gU.dbEncode(nextNo)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_seq").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("itm_code").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pack_key").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_batch_no").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("WOD_qty").ToString.Trim, "0")) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_LOC").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_rem").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_pallet_no").ToString.Trim, "000"))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_TYPE").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty("1", "NULL")) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_SERIAL_NO").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_DRUM_ID").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_DRUM_LV").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("WOD_QTY2").ToString.Trim, "NULL")) & ", " & _
                                                "'" & Session("usr_id") & "',getdate(),'" & Session("usr_id") & "',getdate()) "
                                End Select
                                REM **********************

                                If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction)
                            Next
                        End If

                        If r_dt.Rows.Count > 0 Then
                            Dim resSeq As Integer = 0

                            If cU.gfBuildDataTableforGridView(r_dt, GridView2, True) Then

                                For Each rows As DataRow In r_dt.Rows
                                    itemSQL = ""
                                    If rows.Item("mFlag").ToString.Trim <> "D" Then
                                        resSeq = gU.decodeEmptyCdbl(DB.getValueFromSQL("Select max(convert(int,wo_res_seq)) + 1 from wms_work_order_res " & _
                                                                                       "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                                                                       "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                                                                       "and WO_CODE = '" & gU.dbEncode(WO_CODE.Text.Trim) & "' ", gConn, transaction), 1)

                                        paP = New GlobalDBFunc.DBCmdPara
                                        itemSQL = "Insert into wms_work_order_res (IMP_CODE,  STORER_CODE,  WO_CODE,  WO_RES_SEQ,  WO_RES_TYPE, " & _
                                                  " WO_RES_DESC,  WO_RES_AMT,  WO_RES_REM,  WO_RES_LIST," & _
                                                  " SYS_CD,  SYS_CB,  SYS_LUD,  SYS_LUB" & _
                                                  ") values (" & _
                                                  paP.AP(Session("imp_code")) & "," & paP.AP(STORER_CODE.SelectedValue) & "," & paP.AP(nextNo) & "," & paP.AP(resSeq) & "," & _
                                                  paP.AP(rows.Item("WO_RES_TYPE").ToString.Trim) & "," & paP.AP(rows.Item("WO_RES_DESC").ToString.Trim) & "," & paP.AP(rows.Item("WO_RES_AMT").ToString.Trim) & "," & _
                                                  paP.AP(rows.Item("WO_RES_REM").ToString.Trim) & "," & paP.AP(rows.Item("WO_RES_LIST").ToString.Trim) & "," & _
                                                  " getdate(), " & paP.AP(Session("usr_id")) & ", getdate()," & paP.AP(Session("usr_id")) & _
                                                  ")"
                                    End If
                                    Dim tempCMD As String = gDB.getCmdSql(itemSQL, paP)
                                    If itemSQL <> "" Then gDB.amendData(itemSQL, gConn, transaction, paP)
                                Next
                            End If

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
                    sql_string = "update wms_work_order set " & _
                                    "WO_status = " & gU.convdbNVCData(gU.dbEncode(WO_STATUS.Text)) & ", " & _
                                    "WO_date = " & gU.convdbDate(gU.dbEncode(WO_DATE.Text.Trim)) & ", " & _
                                    "WO_TR_CODE = " & gU.convdbNVCData(gU.dbEncode(WO_TR_CODE.Text.Trim)) & ", " & _
                                    "WO_TYPE=" & gU.convdbNVCData(gU.dbEncode(WO_TYPE.SelectedValue)) & ", " & _
                                    "WO_rem = " & gU.convdbNVCData(gU.dbEncode(WO_REM.Text.Trim)) & ", " & _
                                    "sys_lub = '" & Session("usr_id") & "', " & _
                                    "sys_lud = getdate() " & _
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and WO_code = '" & gU.dbEncode(WO_CODE.Text) & "' "
                    REM **********************

                    If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then

                        'uiFun.reOrderDetails(dt, "WOD_seq")

                        For Each rows As DataRow In dt.Rows
                            itemSQL = ""
                            REM **********************
                            REM Modify Here
                            Select Case rows.Item("mFlag")
                                Case "N"
                                    itemSQL = "insert into wms_work_order_d (" & _
                                                "WO_code, imp_code, storer_code, " & _
                                                "WOD_seq, itm_code, pack_key, " & _
                                                "WOD_batch_no, WOD_qty, WOD_LOC, WOD_rem, WOD_pallet_no,WOD_TYPE, " & _
                                                "WOD_KIT_QTY, WOD_SERIAL_NO, WOD_DRUM_ID, WOD_DRUM_LV, WOD_QTY2," & _
                                                "sys_cb, sys_cd, sys_lub, sys_lud) " & _
                                                "values (" & _
                                                gU.convdbNVCData(gU.dbEncode(WO_CODE.Text)) & "," & gU.convdbNVCData(gU.dbEncode(Session("IMP_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_seq").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("itm_code").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pack_key").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_batch_no").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("WOD_qty").ToString.Trim, "0")) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_LOC").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_rem").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_pallet_no").ToString.Trim, "000"))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_TYPE").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty("1", "NULL")) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_SERIAL_NO").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_DRUM_ID").ToString.Trim, ""))) & ", " & _
                                                gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_DRUM_LV").ToString.Trim, ""))) & ", " & _
                                                gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("WOD_QTY2").ToString.Trim, "NULL")) & ", " & _
                                                "'" & Session("usr_id") & "',getdate(),'" & Session("usr_id") & "',getdate()) "

                                Case "D"
                                    itemSQL = "delete from wms_work_order_d " & _
                                              "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                              "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                              "and WO_code = '" & gU.dbEncode(WO_CODE.Text.Trim) & "' " & _
                                              "and WOD_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("WOD_seq").ToString.Trim, "")) & "' "
                                Case Else
                                    itemSQL = "update wms_work_order_d set " & _
                                                "WOD_seq = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_seq").ToString.Trim, ""))) & ", " & _
                                                "itm_code = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("itm_code").ToString.Trim, ""))) & ", " & _
                                                "pack_key = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("pack_key").ToString.Trim, ""))) & ", " & _
                                                "WOD_batch_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_batch_no").ToString.Trim, ""))) & ", " & _
                                                "WOD_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("WOD_qty").ToString.Trim, "0")) & ", " & _
                                                "WOD_LOC = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_LOC").ToString.Trim, ""))) & ", " & _
                                                "WOD_rem = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_rem").ToString.Trim, ""))) & ", " & _
                                                "WOD_pallet_no = " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_pallet_no").ToString.Trim, "000"))) & ", " & _
                                                "WOD_TYPE= " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_TYPE").ToString.Trim, ""))) & ", " & _
                                                "WOD_SERIAL_NO= " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_SERIAL_NO").ToString.Trim, ""))) & ", " & _
                                                "WOD_DRUM_ID= " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_DRUM_ID").ToString.Trim, ""))) & ", " & _
                                                "WOD_DRUM_LV= " & gU.convdbNVCData(gU.dbEncode(gU.decodeNull(rows.Item("WOD_DRUM_LV").ToString.Trim, ""))) & ", " & _
                                                "WOD_qty2 = " & gU.dbEncode(gU.decodeNullOrEmpty(rows.Item("WOD_qty2").ToString.Trim, "NULL")) & ", " & _
                                                "sys_lub = '" & Session("usr_id") & "', " & _
                                                "sys_lud = getdate() " & _
                                              "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                              "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                              "and WO_code = '" & gU.dbEncode(WO_CODE.Text.Trim) & "' " & _
                                              "and WOD_seq = '" & gU.dbEncode(gU.decodeNull(rows.Item("old_seq").ToString.Trim, "")) & "'"
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

                If Session("pagemode") = "N" Then
                    Session.Remove("pagemode")
                    'Call BindGV()
                    REM **********************
                    REM Modify Here
                    WO_CODE.Text = nextNo
                    ViewState("WO_CODE") = nextNo
                    ViewState("STORER_CODE") = STORER_CODE.SelectedValue
                    WO_CODE.ForeColor = Drawing.Color.Black
                    WO_CODE.Font.Size = 10
                    'WO_CODE.CssClass = ""
                    STORER_CODE.CssClass = ""

                    REM **********************
                End If

                If flag <> "Y" Then
                    uiFun.displayMsg(Me, "1007", "", Session("gLang"))
                End If

                Call BindGV()
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
        If ViewState("WO_CODE") <> "" Then
            pk_code = ViewState("WO_CODE")
            storerCode = ViewState("STORER_CODE")
        Else
            pk_code = Server.UrlDecode(Request("WO_CODE"))
            storerCode = Server.UrlDecode(Request("STORER_CODE"))

            ViewState("WO_CODE") = pk_code
            ViewState("STORER_CODE") = storerCode
        End If
        REM **********************

        WO_STATUS.ForeColor = Drawing.Color.Black

        If Session("pagemode") = "N" Then
            REM **********************
            REM Modify Here
            WO_STATUS.Text = "NEW"
            WO_CODE.ForeColor = Drawing.Color.Red

            REM **********************
        Else
            REM **********************
            REM Modify Here
            REM Generate Data Table from Header
            SQLString = "select Convert(varchar, wms_work_order.WO_DATE," & DDFORMAT & ") as WO_DATE_1,  WO_TR_CODE, wms_work_order.* from wms_work_order " & _
                        "where wms_work_order.WO_code = '" & gU.dbEncode(pk_code) & "' " & _
                        "and wms_work_order.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        "and wms_work_order.storer_code = '" & gU.dbEncode(storerCode) & "' "

            dt = gDB.getDataTable(SQLString)

            If dt.Rows.Count > 0 Then

                IMP_CODE.Value = dt.Rows(0).Item("IMP_CODE").ToString
                WO_CODE.Text = dt.Rows(0).Item("WO_code").ToString
                STORER_CODE.SelectedValue = dt.Rows(0).Item("STORER_CODE").ToString
                WO_STATUS.Text = dt.Rows(0).Item("WO_STATUS").ToString
                WO_DATE.Text = dt.Rows(0).Item("WO_DATE_1").ToString
                WO_TR_CODE.Text = dt.Rows(0).Item("WO_TR_CODE").ToString

                If dt.Rows(0).Item("WO_TR_CODE").ToString = "" Then gotoSTKREC.Visible = False

                WO_REM.Text = dt.Rows(0).Item("WO_REM").ToString
                WO_TYPE.SelectedValue = dt.Rows(0).Item("WO_TYPE").ToString
                sys_cb.Text = dt.Rows(0).Item("sys_cb").ToString
                sys_lub.Text = dt.Rows(0).Item("sys_lub").ToString
                sys_cd.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_cd").ToString)
                sys_lud.Text = cU.chgToFullDF(dt.Rows(0).Item("sys_lud").ToString)

                If dt.Rows(0).Item("STORER_CODE").ToString = "" Then
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , Session("gSelectLabel"))
                Else
                    uiFun.load_dropdown(STORER_CODE, "SELECT STORER_CODE, STO_SHORTNAME FROM WMS_STORER WHERE STORER_CODE = '" & gU.dbEncode(dt.Rows(0).Item("storer_code").ToString) & "' ORDER BY 2", "STORER_CODE", "STO_SHORTNAME", , , , True)
                End If

                If WO_CODE.Text <> "" Then
                    'WO_CODE.ReadOnly = True
                    'WO_CODE.BorderWidth = 0
                    WO_CODE.BackColor = Drawing.Color.Transparent
                End If

                If WO_TYPE.SelectedValue = "RFID" OrElse WO_TYPE.SelectedValue = "T" Then
                    btnPost.Visible = False
                End If

                If WO_TYPE.SelectedValue = "RFID" Then
                    btnChkBal.Visible = True
                End If
                REM **********************
            End If
            REM **********************
        End If
        REM **********************
        REM Modify Here
        REM Generate Data Table from Detail
        SQLString = "SELECT wms_work_order_d.IMP_CODE, wms_work_order_d.STORER_CODE, wms_work_order_d.WO_CODE, wms_work_order_d.WOD_SEQ, wms_work_order_d.ITM_CODE, wms_work_order_d.PACK_KEY, wms_work_order_d.WOD_BATCH_NO," & _
                    "wms_work_order_d.WOD_TYPE, wms_work_order_d.WOD_QTY, wms_work_order_d.WOD_PALLET_NO, wms_work_order_d.WOD_LOC, wms_work_order_d.WOD_REM, wms_work_order_d.SYS_LUB, wms_work_order_d.SYS_LUD, wms_work_order_d.SYS_CD, wms_work_order_d.SYS_CB," & _
                    "wms_work_order_d.WOD_KIT_QTY, wms_work_order_d.WOD_SERIAL_NO, wms_work_order_d.WOD_DRUM_ID, wms_work_order_d.WOD_DRUM_LV, wms_work_order_d.WOD_QTY2, " & _
                    "wms_item.itm_name, 'U' as mFlag, wms_work_order_d.WOD_seq as old_seq, Count(WORF_SEQ) as RFID_COUNT, wms_item.itm_uom, wms_item.itm_uom2, wms_item.itm_sku_no " & _
                    "from wms_work_order_d left outer join  wms_item on " & _
                    "wms_work_order_d.imp_code = wms_item.imp_code " & _
                    "and wms_work_order_d.storer_code = wms_item.storer_code " & _
                    "and wms_work_order_d.itm_code = wms_item.itm_code " & _
                    "and wms_work_order_d.PACK_KEY = wms_item.PACK_KEY " & _
                    "LEFT OUTER JOIN WMS_WORK_ORDER_RFID ON WMS_WORK_ORDER_D.IMP_CODE = WMS_WORK_ORDER_RFID.IMP_CODE AND WMS_WORK_ORDER_D.STORER_CODE = WMS_WORK_ORDER_RFID.STORER_CODE AND " & _
                    "WMS_WORK_ORDER_D.WO_CODE = WMS_WORK_ORDER_RFID.WO_CODE AND WMS_WORK_ORDER_D.WOD_SEQ = WMS_WORK_ORDER_RFID.WOD_SEQ " & _
                    "where wms_work_order_d.WO_code = '" & gU.dbEncode(pk_code) & "' " & _
                    "and wms_work_order_d.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                    "and wms_work_order_d.imp_code = '" & Session("IMP_CODE") & "'" & _
                    "group by wms_work_order_d.IMP_CODE, wms_work_order_d.STORER_CODE, wms_work_order_d.WO_CODE, wms_work_order_d.WOD_SEQ, wms_work_order_d.ITM_CODE, wms_work_order_d.PACK_KEY, wms_work_order_d.WOD_BATCH_NO," & _
                    "wms_work_order_d.WOD_TYPE, wms_work_order_d.WOD_QTY, wms_work_order_d.WOD_PALLET_NO, wms_work_order_d.WOD_LOC, wms_work_order_d.WOD_REM, wms_work_order_d.SYS_LUB, wms_work_order_d.SYS_LUD, wms_work_order_d.SYS_CD, wms_work_order_d.SYS_CB, wms_item.itm_name, " & _
                    "wms_work_order_d.WOD_KIT_QTY, wms_work_order_d.WOD_SERIAL_NO, wms_work_order_d.WOD_DRUM_ID, wms_work_order_d.WOD_DRUM_LV, wms_work_order_d.WOD_QTY2, wms_item.itm_uom, wms_item.itm_uom2, wms_item.itm_sku_no "


        SQLString = SQLString & " order by convert(int, wms_work_order_d.wod_seq)"
        REM **********************
        dt = gDB.getDataTable(SQLString)
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
        Else
            GridView1.DataSource = Nothing
        End If
        ViewState("dt") = dt
        GridView1.DataBind()

        If WO_TYPE.SelectedValue <> "RFID" AndAlso WO_TYPE.SelectedValue <> "T" Then
            GridView1.Columns(16).Visible = False
            GridView1.Columns(17).Visible = False
        End If

        SQLString = "SELECT 'U' as mFlag, wms_work_order_res.WO_CODE,  wms_work_order_res.WO_RES_SEQ,  wms_work_order_res.WO_RES_TYPE, " & _
                    " wms_work_order_res.WO_RES_DESC,  wms_work_order_res.WO_RES_AMT,  wms_work_order_res.WO_RES_REM,  wms_work_order_res.WO_RES_LIST " & _
                    "FROM wms_work_order_res " & _
                    "Where wms_work_order_res.WO_code = '" & gU.dbEncode(pk_code) & "' " & _
                    "and wms_work_order_res.storer_code = '" & gU.dbEncode(storerCode) & "' " & _
                    "and wms_work_order_res.imp_code = '" & Session("IMP_CODE") & "'" & _
                    " order by convert(int,wo_res_seq)"

        r_dt = gDB.getDataTable(SQLString)

        If r_dt.Rows.Count > 0 Then
            GridView2.DataSource = r_dt
        Else
            GridView2.DataSource = Nothing
        End If
        ViewState("r_dt") = r_dt
        GridView2.DataBind()

        REM **********************
    End Sub

    Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click
        Dim cancelSql As String = "update wms_work_order " & _
                         "set WO_status = 'CANCELLED', " & _
                         "sys_lub = '" & Session("usr_id") & "', " & _
                         "sys_lud = getdate() " & _
                         "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                         "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                         "and WO_code = '" & gU.dbEncode(WO_CODE.Text) & "' "

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            gDB.amendData(cancelSql, gConn, transaction)

            cancelSql = "Update WMS_WORK_ORDER_RFID set WORF_STATUS = 'CANCEL', " & _
                        "sys_lub = '" & Session("usr_id") & "', " & _
                        "sys_lud = getdate() " & _
                        "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                        "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                        "and WO_code = '" & gU.dbEncode(WO_CODE.Text) & "' "
            gDB.amendData(cancelSql, gConn, transaction)

            cancelSql = "Update WMS_RFID_ACTION set WRAT_MASTER_FLAG = 'Y' "
            gDB.amendData(cancelSql, gConn, transaction)

            transaction.Commit()

            WO_STATUS.Text = "CANCELLED"

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
        Dim checkdt As DataTable
        REM DEMO XXXXX
        'WO_STATUS.Text = "POSTED"
        'uiFun.displayMsg(Me, "", "Work Order Posted.", Session("gLang"))
        'Exit Sub
        REM XXXXXX

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Dim iscable As Boolean = False
        Dim sqlString As String = ""
        Dim serial_No As String = ""

        Try
            If GridView1.Rows.Count > 0 Then
                Call save("Y")

                For Each rows As DataRow In dt.Rows
                    sqlString = "select ITM_CODE from WMS_ITEM WHERE ITM_TYPE = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value) & "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.Text) & "' " & _
                                "AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & "' "
                    If gConn Is Nothing Then
                        checkdt = gDB.getDataTable(sqlString)
                    Else
                        If transaction Is Nothing Then
                            checkdt = gDB.getDataTable(sqlString, gConn)
                        Else
                            checkdt = gDB.getDataTable(sqlString, gConn, transaction)
                        End If
                    End If

                    If checkdt.Rows.Count > 0 Then
                        iscable = True
                    End If

                    Select Case rows.Item("wod_type")
                        Case "IN"
                            st.STORER_CODE = STORER_CODE.SelectedValue
                            st.ITM_CODE = gU.decodeNullOrEmpty(rows.Item("itm_code").ToString.Trim, "")
                            st.PACK_KEY = gU.decodeNullOrEmpty(rows.Item("pack_key").ToString.Trim, "")
                            st.IO_CUST_CODE = ""
                            st.IO_AREA = ""
                            st.IO_DOC = "WO"
                            st.IO_DOC_ID = WO_CODE.Text.Trim
                            st.IO_QTY = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("WOD_QTY").ToString.Trim, ""), "0")
                            st.IO_CBM = 0
                            st.IO_KG = 0

                            st.lO_BATCH_NO = gU.decodeNullOrEmpty(rows.Item("WOD_batch_no").ToString.Trim, "")

                            st.IO_WH = ""
                            st.IO_LOC = gU.decodeNullOrEmpty(rows.Item("WOD_LOC").ToString.Trim, "")
                            st.PALLET_NO = gU.decodeNullOrEmpty(rows.Item("WOD_pallet_no").ToString.Trim, "000")



                            If rows.Item("WOD_SERIAL_NO").ToString.Trim <> "" Then
                                st.IOS_DRUM_ID = gU.decodeNull(rows.Item("WOD_DRUM_ID").ToString.Trim, "")
                                st.IOS_DRUM_LEVEL = gU.decodeNull(rows.Item("WOD_DRUM_LV").ToString.Trim, "")
                                serial_No = st.getNewSerialNo(gU.decodeNull(rows.Item("WOD_SERIAL_NO").ToString.Trim, ""), gConn, transaction)
                                st.IOS_SERIAL_NO = serial_No
                                st.IOS_QTY2 = gU.decodeEmptyCdbl(rows.Item("WOD_QTY2").ToString.Trim, 0)
                                st.IOS_UOM2 = gU.decodeNull(rows.Item("ITM_UOM2").ToString.Trim, "")
                                REM check cable
                                sqlString = "select ITM_CODE from WMS_ITEM WHERE ITM_SERIAL_NO_YN = 'Y' AND isnull(ITM_TYPE,'') = 'CABLE' AND IMP_CODE = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                            "AND ITM_CODE = '" & gU.dbEncode(rows.Item("itm_code").ToString.Trim) & "' AND PACK_KEY = '" & gU.dbEncode(rows.Item("pack_key").ToString.Trim) & "' "
                                checkdt = gDB.getDataTable(sqlString)
                                If checkdt.Rows.Count > 0 Then
                                    st.IOS_SL = "Y"
                                Else
                                    st.IOS_SL = "N"
                                End If

                                Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("WOD_SERIAL_NO").ToString.Trim, ""), gConn, transaction)
                            End If
                            st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                            st.UpdateStockTrans("IN", gConn, transaction)
                            st.UpdateStockBalTrans("IN", gConn, transaction)
                            If rows.Item("WOD_SERIAL_NO").ToString.Trim <> "" Then
                                st.UpdateStockSerialTrans("IN", gConn, transaction)
                                st.UpdateStockBalSerialTrans("IN", gConn, transaction)
                            End If

                        Case "OUT"
                            st.STORER_CODE = STORER_CODE.SelectedValue
                            st.ITM_CODE = gU.decodeNullOrEmpty(rows.Item("itm_code").ToString.Trim, "")
                            st.PACK_KEY = gU.decodeNullOrEmpty(rows.Item("pack_key").ToString.Trim, "")
                            st.IO_CUST_CODE = ""
                            st.IO_AREA = ""
                            st.IO_DOC = "WO"
                            st.IO_DOC_ID = WO_CODE.Text.Trim
                            st.IO_QTY = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("WOD_QTY").ToString.Trim, ""), "0")
                            st.IO_CBM = 0
                            st.IO_KG = 0
                            st.lO_BATCH_NO = gU.decodeNullOrEmpty(rows.Item("WOD_batch_no").ToString.Trim, "")
                            st.IO_WH = ""
                            st.IO_LOC = gU.decodeNullOrEmpty(rows.Item("WOD_LOC").ToString.Trim, "")
                            st.PALLET_NO = gU.decodeNullOrEmpty(rows.Item("WOD_pallet_no").ToString.Trim, "000")
                            If rows.Item("WOD_SERIAL_NO").ToString.Trim <> "" Then
                                st.IOS_SERIAL_NO = rows.Item("WOD_SERIAL_NO").ToString.Trim
                                st.IOS_QTY2 = gU.decodeEmptyCdbl(gU.decodeNull(rows.Item("WOD_QTY2").ToString.Trim, ""), "0")
                                st.IOS_UOM2 = gU.decodeNull(rows.Item("ITM_UOM2").ToString.Trim, "")
                                Call st.setOrgSerialInfo(gU.decodeNull(rows.Item("WOD_SERIAL_NO").ToString.Trim, ""), gConn, transaction)
                            End If
                            st.IO_SYS_SEQ = DB.getDocNo("SYS_SEQ", gConn, transaction)
                            st.UpdateStockTrans("OUT", gConn, transaction)
                            st.UpdateStockBalTrans("OUT", gConn, transaction)
                            If rows.Item("WOD_SERIAL_NO").ToString.Trim <> "" Then
                                st.UpdateStockSerialTrans("OUT", gConn, transaction)
                                st.UpdateStockBalSerialTrans("OUT", gConn, transaction)
                            End If
                    End Select

                Next

                updtSql = "update wms_work_order " & _
                            "set WO_status = 'POSTED', " & _
                            "sys_lub = '" & Session("usr_id") & "', " & _
                            "sys_lud = getdate() " & _
                            "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                            "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            "and WO_code = '" & gU.dbEncode(WO_CODE.Text) & "' "

                gDB.amendData(updtSql, gConn, transaction)

                transaction.Commit()

                WO_STATUS.Text = "POSTED"
                'ar.sec_write = "N"
                ar.sec_viewMode = "Y"
                btnPost.Visible = False
                ar.hideForm(Me, editMode, AddressOf page_customizectrl, AddressOf customizectrl)

                uiFun.displayMsg(Me, "1007", "", Session("gLang"))
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

    Protected Sub addItemtoWO()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(WOD_SEQ AS int)) + 1 from wms_work_order_D " & _
                                        "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                        "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and WO_CODE = '" & gU.dbEncode(WO_CODE.Text.Trim) & "' "
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

            Dim itemPackList As String = ""
            Dim itemListarray As String()
            Dim packKeyListarray As String()
            itemListarray = Split(itemList.Value, ", ")
            packKeyListarray = Split(packKeyList.Value, ", ")
            If itemListarray.Count = 0 Then
                If itemList.Value <> "" Then
                    itemPackList = Server.HtmlDecode(itemList.Value) & "_000_" & Server.HtmlDecode(packKeyList.Value)
                End If
            Else
                For i = 0 To itemListarray.Count - 1
                    itemPackList = gU.appendToList(itemPackList, Server.HtmlDecode(itemListarray(i)) & "_000_" & Server.HtmlDecode(packKeyListarray(i)))
                Next
            End If

            'SQLString = "SELECT M.*, D.* " & _
            '"from WMS_ITEM M, WMS_ALT_VEND_ITEM D " & _
            '"where M.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
            '"and M.IMP_CODE = '" & Session("IMP_CODE") & "' " & _
            '"AND M.ITM_CODE + '_000_' + M.PACK_KEY + '_000_' + D.VND_CODE IN ('" & Replace(itemPackList, ", ", "', '") & "') " & _
            '"AND M.IMP_CODE = D.IMP_CODE " & _
            '"AND M.STORER_CODE = D.STORER_CODE " & _
            '"AND M.PACK_KEY = D.PACK_KEY " & _
            '"AND M.ITM_CODE = D.ITM_CODE"


            SQLString = ""
            SQLString += "select wms_item.*, wms_alt_vend_item.* "
            SQLString += "from wms_item left outer join wms_alt_vend_item on "
            SQLString += "wms_item.imp_code = wms_alt_vend_item.imp_code "
            SQLString += "and wms_item.storer_code = wms_alt_vend_item.storer_code "
            SQLString += "and wms_item.itm_code = wms_alt_vend_item.itm_code "
            SQLString += "and wms_item.pack_key = wms_alt_vend_item.pack_key "
            SQLString += "where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "
            SQLString += "and wms_item.imp_code = '" & Session("IMP_CODE") & "' "
            SQLString += "and wms_item.itm_code + '_000_' + wms_item.pack_key + '_000_' + isnull(wms_alt_vend_item.vnd_code,'000') in ('" & Replace(itemPackList, ", ", "', '") & "') "

            SQLString = SQLString & " order by wms_item.ITM_CODE, wms_item.PACK_KEY, VND_CODE"
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

                REM **********************
                REM Modify Here
                dt.Rows(rows_count - 1).Item("WOD_SEQ") = ViewState("n_cur_seq").ToString

                dt.Rows(rows_count - 1).Item("WOD_TYPE") = "OUT"
                dt.Rows(rows_count - 1).Item("ITM_CODE") = stf_dt.Rows(i).Item("ITM_CODE")
                dt.Rows(rows_count - 1).Item("ITM_SKU_NO") = stf_dt.Rows(i).Item("ITM_SKU_NO")
                dt.Rows(rows_count - 1).Item("ITM_NAME") = stf_dt.Rows(i).Item("ITM_NAME")
                dt.Rows(rows_count - 1).Item("PACK_KEY") = stf_dt.Rows(i).Item("PACK_KEY")
                dt.Rows(rows_count - 1).Item("WOD_QTY") = 0

                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

            Next

            STORER_CODE.Enabled = False

            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()

            'For i = 0 To GridView1.Rows.Count - 1
            '    'CType(GridView1.Rows(i).FindControl("dsp_WOD_batch_no"), Label).Text = CType(GridView1.Rows(i).FindControl("WOD_batch_no"), HiddenField).Value
            '    CType(GridView1.Rows(i).FindControl("dsp_WOD_loc"), Label).Text = CType(GridView1.Rows(i).FindControl("WOD_loc"), HiddenField).Value
            '    ' CType(GridView1.Rows(i).FindControl("dsp_WOD_pallet_no"), Label).Text = CType(GridView1.Rows(i).FindControl("WOD_pallet_no"), HiddenField).Value
            'Next

        End If
    End Sub

    Protected Sub btnAddRes_Click(sender As Object, e As System.EventArgs) Handles btnAddRes.Click
        If cU.gfBuildDataTableforGridView(r_dt, GridView2, True) Then
            Dim newRow As DataRow = r_dt.NewRow

            newRow.Item("mFlag") = "N"
            r_dt.Rows.Add(newRow)
            r_dt.AcceptChanges()


            ViewState("r_dt") = r_dt
            GridView2.DataSource = r_dt
            GridView2.DataBind()
        End If
    End Sub


    Protected Sub GridView2_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("WO_RES_TYPE"), DropDownList).SelectedValue = DataBinder.Eval(e.Row.DataItem, "WO_RES_TYPE").ToString.Trim
                CType(e.Row.FindControl("WO_RES_DESC"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "WO_RES_DESC").ToString.Trim
                CType(e.Row.FindControl("WO_RES_AMT"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "WO_RES_AMT").ToString.Trim
                CType(e.Row.FindControl("WO_RES_REM"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "WO_RES_REM").ToString.Trim


        End Select
    End Sub

    Protected Sub GridView2_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView2.RowDeleting
        Call ar.hideGVRow(GridView2, GridView2.Rows(e.RowIndex))
        r_dt.Rows(e.RowIndex).Item("mFlag") = "D"
        r_dt.AcceptChanges()
    End Sub

    Protected Sub BindRFID(ByVal wod_seq As String, ByVal wod_qty As Double)

        If wod_qty <> 0 Then
            Dim updateSql, selectSQL As String
            Dim tempDT As DataTable

            Dim MaxSeq As String = ""
            Dim lRFID_CODE As String = ""

            Dim gConn As SqlConnection

            gConn = gDB.getConnection()
            Dim transaction As SqlTransaction

            transaction = gConn.BeginTransaction()

            Try
                selectSQL = "SELECT WOD_SEQ, WORF_SEQ, WORF_RFID, WORF_STATUS FROM WMS_WORK_ORDER_RFID " & _
                            "where WMS_WORK_ORDER_RFID.WO_code = '" & gU.dbEncode(ViewState("WO_CODE")) & "' " & _
                            "and WMS_WORK_ORDER_RFID.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                            "and WMS_WORK_ORDER_RFID.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            "and WMS_WORK_ORDER_RFID.WOD_SEQ = '" & gU.dbEncode(wod_seq) & "' "
                tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                If tempDT.Rows.Count > 0 Then

                    If wod_qty > tempDT.Rows.Count Then
                        For i = 1 To CInt(wod_qty - tempDT.Rows.Count)
                            selectSQL = "SELECT ISNULL(MAX(convert(numeric,WORF_SEQ)) + 1,1) FROM WMS_WORK_ORDER_RFID " & _
                                        "where WMS_WORK_ORDER_RFID.WO_code = '" & gU.dbEncode(ViewState("WO_CODE")) & "' " & _
                                        "and WMS_WORK_ORDER_RFID.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                                        "and WMS_WORK_ORDER_RFID.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and WMS_WORK_ORDER_RFID.WOD_SEQ = '" & gU.dbEncode(wod_seq) & "' "

                            MaxSeq = gU.decodeNullOrEmpty(DB.getValueFromSQL(selectSQL, gConn, transaction), "1")


                            lRFID_CODE = Now.ToString("yyyyMMdd") & Right(DB.getDocNo("RFID", gConn, transaction), 8)

                            updateSql = "INSERT INTO WMS_WORK_ORDER_RFID " & _
                                        "(IMP_CODE, STORER_CODE, WO_CODE, WOD_SEQ, WORF_SEQ, WORF_RFID, WORF_STATUS, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB) VALUES ( " & _
                                        gU.convdbNVCData(Session("IMP_CODE")) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(ViewState("WO_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(wod_seq)) & "," & _
                                        gU.convdbNVCData(MaxSeq) & "," & gU.convdbNVCData(gU.dbEncode(lRFID_CODE)) & ",'NEW', " & _
                                        "'" & Session("usr_id") & "',Getdate(),Getdate(),'" & Session("usr_id") & "') "

                            gDB.amendData(updateSql, gConn, transaction)

                        Next
                    End If

                    selectSQL = "SELECT WOD_SEQ, WORF_SEQ, WORF_RFID, WORF_STATUS FROM WMS_WORK_ORDER_RFID " & _
                            "where WMS_WORK_ORDER_RFID.WO_code = '" & gU.dbEncode(ViewState("WO_CODE")) & "' " & _
                            "and WMS_WORK_ORDER_RFID.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                            "and WMS_WORK_ORDER_RFID.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            "and WMS_WORK_ORDER_RFID.WOD_SEQ = '" & gU.dbEncode(wod_seq) & "' "
                    tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                    GVRFID.DataSource = tempDT
                    GVRFID.DataBind()
                ElseIf tempDT.Rows.Count = 0 Then

                    For i = 1 To wod_qty
                        selectSQL = "SELECT ISNULL(MAX(convert(numeric,WORF_SEQ)) + 1,1) FROM WMS_WORK_ORDER_RFID " & _
                                        "where WMS_WORK_ORDER_RFID.WO_code = '" & gU.dbEncode(ViewState("WO_CODE")) & "' " & _
                                        "and WMS_WORK_ORDER_RFID.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                                        "and WMS_WORK_ORDER_RFID.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                        "and WMS_WORK_ORDER_RFID.WOD_SEQ = '" & gU.dbEncode(wod_seq) & "' "

                        MaxSeq = gU.decodeNullOrEmpty(DB.getValueFromSQL(selectSQL, gConn, transaction), "1")

                        lRFID_CODE = Now.ToString("yyyyMMdd") & Right(DB.getDocNo("RFID", gConn, transaction), 8)

                        updateSql = "INSERT INTO WMS_WORK_ORDER_RFID " & _
                                    "(IMP_CODE, STORER_CODE, WO_CODE, WOD_SEQ, WORF_SEQ, WORF_RFID, WORF_STATUS, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB) VALUES ( " & _
                                    gU.convdbNVCData(Session("IMP_CODE")) & "," & gU.convdbNVCData(gU.dbEncode(STORER_CODE.SelectedValue)) & "," & gU.convdbNVCData(gU.dbEncode(ViewState("WO_CODE"))) & "," & gU.convdbNVCData(gU.dbEncode(wod_seq)) & "," & _
                                    gU.convdbNVCData(MaxSeq) & "," & gU.convdbNVCData(gU.dbEncode(lRFID_CODE)) & ",'NEW', " & _
                                    "'" & Session("usr_id") & "',Getdate(),Getdate(),'" & Session("usr_id") & "') "

                        gDB.amendData(updateSql, gConn, transaction)

                    Next

                    selectSQL = "SELECT WOD_SEQ, WORF_SEQ, WORF_RFID, WORF_STATUS FROM WMS_WORK_ORDER_RFID " & _
                            "where WMS_WORK_ORDER_RFID.WO_code = '" & gU.dbEncode(ViewState("WO_CODE")) & "' " & _
                            "and WMS_WORK_ORDER_RFID.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                            "and WMS_WORK_ORDER_RFID.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                            "and WMS_WORK_ORDER_RFID.WOD_SEQ = '" & gU.dbEncode(wod_seq) & "' "
                    tempDT = gDB.getDataTable(selectSQL, gConn, transaction)

                    GVRFID.DataSource = tempDT
                    GVRFID.DataBind()
                Else

                    GVRFID.DataSource = tempDT
                    GVRFID.DataBind()

                End If

                transaction.Commit()

            Catch ex As Exception
                transaction.Rollback()
                Response.Write(ex.Message)
                uiFun.displayMsgNew(udp1, "1008", "", Session("gLang"))
            Finally
                If gConn IsNot Nothing Then
                    If gConn.State = ConnectionState.Open Then
                        gConn.Close()
                        gConn.Dispose()
                    End If
                End If
            End Try


            pnlRFID_ModalPopupExtender.Show()

        Else
            uiFun.displayMsgNew(udp1, "", "Please enter item quantity!", Session("gLang"))
        End If



    End Sub

    Protected Sub GVRFID_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GVRFID.RowCommand
         Select e.CommandName
            Case "OUT"
                Dim gvRow As GridViewRow = CType(CType(e.CommandSource, Control).NamingContainer, GridViewRow)
                Dim wod_seq As String = DirectCast(gvRow.FindControl("wod_seq"), HiddenField).Value
                Dim worf_seq As String = DirectCast(gvRow.FindControl("WORF_SEQ"), HiddenField).Value

                ChangeStatus(wod_seq, worf_seq)
        End Select
    End Sub

    Protected Sub GVRFID_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GVRFID.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow

                CType(e.Row.FindControl("WORF_RFID"), TextBox).Text = DataBinder.Eval(e.Row.DataItem, "WORF_RFID").ToString.Trim
                CType(e.Row.FindControl("WORF_SEQ"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "WORF_SEQ").ToString.Trim
                CType(e.Row.FindControl("WOD_SEQ"), HiddenField).Value = DataBinder.Eval(e.Row.DataItem, "WOD_SEQ").ToString.Trim

                Select Case DataBinder.Eval(e.Row.DataItem, "WORF_STATUS").ToString.Trim
                    Case "NEW"
                        DirectCast(e.Row.FindControl("WORF_STATUS"), Panel).CssClass = "BlackCircle"
                        CType(e.Row.FindControl("lbl_status"), Label).ForeColor = Drawing.Color.White
                        CType(e.Row.FindControl("btnOut"), Button).Visible = False
                    Case "IN"
                        DirectCast(e.Row.FindControl("WORF_STATUS"), Panel).CssClass = "GreenCircle"
                        CType(e.Row.FindControl("lbl_status"), Label).ForeColor = Drawing.Color.White
                        CType(e.Row.FindControl("btnOut"), Button).Visible = False
                    Case "OUT"
                        DirectCast(e.Row.FindControl("WORF_STATUS"), Panel).CssClass = "YellowCircle"
                        CType(e.Row.FindControl("lbl_status"), Label).ForeColor = Drawing.Color.Black
                        CType(e.Row.FindControl("btnOut"), Button).Visible = False
                    Case "MISS"
                        DirectCast(e.Row.FindControl("WORF_STATUS"), Panel).CssClass = "RedCircle"
                        CType(e.Row.FindControl("lbl_status"), Label).ForeColor = Drawing.Color.White
                        CType(e.Row.FindControl("btnOut"), Button).Visible = True

                End Select
                CType(e.Row.FindControl("lbl_status"), Label).Text = DataBinder.Eval(e.Row.DataItem, "WORF_STATUS").ToString.Trim

        End Select
    End Sub

    Protected Sub GVRFID_RowDeleting(sender As Object, e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GVRFID.RowDeleting
        Dim wod_seq As String = DirectCast(GVRFID.Rows(e.RowIndex).FindControl("wod_seq"), HiddenField).Value
        Dim worf_seq As String = DirectCast(GVRFID.Rows(e.RowIndex).FindControl("WORF_SEQ"), HiddenField).Value


        Call DeleteRFID(wod_seq, worf_seq)
    End Sub
    Private Sub ChangeStatus(ByVal wod_seq As String, ByVal worf_seq As String)
        Dim updateSql, selectSQL As String
        Dim tempDT As DataTable
        Dim successFlag As Boolean = False

        Dim MaxSeq As String = ""
        Dim lRFID_CODE As String = ""

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            updateSql = "UPDATE WMS_WORK_ORDER_RFID SET " & _
                        "WORF_STATUS = 'OUT', SYS_LUD=getdate(), sys_LUB='" & gU.dbEncode(Session("usr_id")) & "' " & _
                        "where WMS_WORK_ORDER_RFID.WO_code = '" & gU.dbEncode(ViewState("WO_CODE")) & "' " & _
                        "and WMS_WORK_ORDER_RFID.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        "and WMS_WORK_ORDER_RFID.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                        "and WMS_WORK_ORDER_RFID.WOD_SEQ = '" & gU.dbEncode(wod_seq) & "' " & _
                        "and WMS_WORK_ORDER_RFID.WORF_SEQ='" & gU.dbEncode(worf_seq) & "' " & _
                        "and WMS_WORK_ORDER_RFID.WORF_STATUS='MISS'"
            gDB.amendData(updateSql, gConn, transaction)

            transaction.Commit()
            successFlag = True


        Catch ex As Exception
            transaction.Rollback()
            Response.Write(ex.Message)
            uiFun.displayMsgNew(udp1, "1008", "", Session("gLang"))
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

        If successFlag Then
            selectSQL = "SELECT WOD_SEQ, WORF_SEQ, WORF_RFID, WORF_STATUS FROM WMS_WORK_ORDER_RFID " & _
                        "where WMS_WORK_ORDER_RFID.WO_code = '" & gU.dbEncode(ViewState("WO_CODE")) & "' " & _
                        "and WMS_WORK_ORDER_RFID.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        "and WMS_WORK_ORDER_RFID.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                        "and WMS_WORK_ORDER_RFID.WOD_SEQ = '" & gU.dbEncode(wod_seq) & "' "
            tempDT = gDB.getDataTable(selectSQL)

            GVRFID.DataSource = tempDT
            GVRFID.DataBind()
        End If

        pnlRFID_ModalPopupExtender.Show()
    End Sub
    Protected Sub DeleteRFID(ByVal wod_seq As String, ByVal worf_seq As String)

        Dim updateSql, selectSQL As String
        Dim tempDT As DataTable
        Dim successFlag As Boolean = False

        Dim MaxSeq As String = ""
        Dim lRFID_CODE As String = ""

        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            updateSql = "DELETE FROM WMS_WORK_ORDER_RFID " & _
                         "where WMS_WORK_ORDER_RFID.WO_code = '" & gU.dbEncode(ViewState("WO_CODE")) & "' " & _
                         "and WMS_WORK_ORDER_RFID.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                         "and WMS_WORK_ORDER_RFID.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                         "and WMS_WORK_ORDER_RFID.WOD_SEQ = '" & gU.dbEncode(wod_seq) & "' " & _
                         "and WMS_WORK_ORDER_RFID.WORF_SEQ='" & gU.dbEncode(worf_seq) & "'"
            gDB.amendData(updateSql, gConn, transaction)

            transaction.Commit()
            successFlag = True


        Catch ex As Exception
            transaction.Rollback()
            Response.Write(ex.Message)
            uiFun.displayMsgNew(udp1, "1008", "", Session("gLang"))
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try


        If successFlag Then
            selectSQL = "SELECT WOD_SEQ, WORF_SEQ, WORF_RFID, WORF_STATUS FROM WMS_WORK_ORDER_RFID " & _
                        "where WMS_WORK_ORDER_RFID.WO_code = '" & gU.dbEncode(ViewState("WO_CODE")) & "' " & _
                        "and WMS_WORK_ORDER_RFID.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        "and WMS_WORK_ORDER_RFID.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                        "and WMS_WORK_ORDER_RFID.WOD_SEQ = '" & gU.dbEncode(wod_seq) & "' "
            tempDT = gDB.getDataTable(selectSQL)

            GVRFID.DataSource = tempDT
            GVRFID.DataBind()


            For i = 0 To GridView1.Rows.Count - 1
                If DirectCast(GridView1.Rows(i).FindControl("WOD_SEQ"), HiddenField).Value = wod_seq Then
                    'DirectCast(GridView1.Rows(i).FindControl("WOD_QTY"), TextBox).Text = tempDT.Rows.Count
                    DirectCast(GridView1.Rows(i).FindControl("RFID_COUNT"), Label).Text = tempDT.Rows.Count

                    updateSql = "update wms_work_order_d set " & _
                                  "WOD_qty = " & gU.dbEncode(gU.decodeNullOrEmpty(tempDT.Rows.Count, 0)) & ", " & _
                                  "sys_lub = '" & Session("usr_id") & "', " & _
                                  "sys_lud = getdate() " & _
                                "where imp_code = '" & gU.dbEncode(IMP_CODE.Value.Trim) & "' " & _
                                "and storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                "and WO_code = '" & gU.dbEncode(WO_CODE.Text.Trim) & "' " & _
                                "and WOD_seq = '" & gU.dbEncode(wod_seq) & "'"
                    'gDB.amendData(updateSql)

                End If
            Next
        End If

        pnlRFID_ModalPopupExtender.Show()

    End Sub


    Private Sub addBalItemtoWO()
        If cU.gfBuildDataTableforGridView(dt, GridView1, True) Then
            Dim rows_count As Integer = 0
            REM **********************
            REM Modify Here
            Dim seq_string As String = "select MAX(CAST(WOD_SEQ AS int)) + 1 from wms_work_order_D " & _
                                         "where IMP_CODE = '" & gU.dbEncode(Session("IMP_CODE")) & "' " & _
                                         "and STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                                         "and WO_CODE = '" & gU.dbEncode(WO_CODE.Text.Trim) & "' "
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
            Dim sadj_dt As DataTable

            Dim itemPackList As String = ""
            Dim itemListarray As String()
            Dim packKeyListarray As String()
            Dim seqListArray As String()

            Dim seqKeyList As String = ""

            itemListarray = Split(itemList.Value, ", ")
            packKeyListarray = Split(packKeyList.Value, ", ")
            seqListArray = Split(seqList.Value, ", ")

            If seqListArray.Count = 0 Then
                If seqList.Value <> "" Then
                    seqKeyList = Server.HtmlDecode(seqList.Value)
                End If
            Else
                For i = 0 To seqListArray.Count - 1
                    seqKeyList = gU.appendToList(seqKeyList, seqListArray(i))
                Next
            End If

            Dim addSQL As String = ""

            If seqKeyList <> "" Then
                addSQL = "'" & Replace(seqKeyList, ", ", "', '") & "'"
            Else
                addSQL = "NULL"
            End If

            SQLString = " SELECT WMS_ITEM_LOC_BAL.IMP_CODE, WMS_ITEM_LOC_BAL.STORER_CODE, WMS_ITEM_LOC_BAL.ITM_CODE, WMS_ITEM_LOC_BAL.PACK_KEY, WMS_ITEM.ITM_UOM, " & _
                        " WMS_ITEM_LOC_BAL.ILOC_PALLET_NO, WMS_ITEM_LOC_BAL.ILOC_LOC, CASE WHEN WMS_ITEM.ITM_SERIAL_NO_YN = 'Y' THEN 1 ELSE WMS_ITEM_LOC_BAL.ILOC_BAL_QTY END AS ILOC_BAL_QTY, WMS_ITEM_LOC_BAL.ILOC_BATCH_NO, " & _
                        " WMS_ITEM.ITM_NAME, WMS_ITEM.ITM_DESC, WMS_ITEM.ITM_SKU_NO, WMS_ITEM_LOC_BAL.ILOC_SEQ, WMS_ITEM_LOC_BAL_s.ILBS_SEQ, " & _
                        " WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO, WMS_ITEM_LOC_BAL_S.ILBS_UOM2, WMS_ITEM_LOC_BAL_S.ILBS_QTY2, WMS_ITEM_LOC_BAL_S.ILBS_DRUM_ID,WMS_ITEM_LOC_BAL_S.ILBS_DRUM_LEVEL, " & _
                        " Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_EXPIRY_DATE," & DDFORMAT & ") as ILOC_EXPIRY_DATE, Convert(varchar, WMS_ITEM_LOC_BAL.ILOC_MANU_DATE," & DDFORMAT & ") as ILOC_MANU_DATE, WMS_ITEM_LOC_BAL.VND_CODE " & _
                        " FROM WMS_ITEM_LOC_BAL INNER JOIN " & _
                        " WMS_ITEM ON WMS_ITEM.IMP_CODE = WMS_ITEM_LOC_BAL.IMP_CODE AND WMS_ITEM.STORER_CODE = WMS_ITEM_LOC_BAL.STORER_CODE AND " & _
                        " WMS_ITEM.ITM_CODE = WMS_ITEM_LOC_BAL.ITM_CODE AND WMS_ITEM.PACK_KEY = WMS_ITEM_LOC_BAL.PACK_KEY " & _
                        " LEFT OUTER JOIN WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                        " where wms_item.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                        " and wms_item.imp_code = '" & Session("IMP_CODE") & "' " & _
                        " and Cast(WMS_ITEM_LOC_BAL.ILOC_SEQ as varchar) + '#_#' + isnull(cast(WMS_ITEM_LOC_BAL_s.ILBS_SEQ as varchar),'') in (" & addSQL & ") "

            SQLString = SQLString & " order by wms_item.ITM_CODE, wms_item.PACK_KEY, VND_CODE"
            REM **********************
            sadj_dt = gDB.getDataTable(SQLString)

            For i As Integer = 0 To sadj_dt.Rows.Count - 1
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
                dt.Rows(rows_count - 1).Item("wod_seq") = ViewState("n_cur_seq").ToString
                dt.Rows(rows_count - 1).Item("itm_code") = sadj_dt.Rows(i).Item("itm_code")
                dt.Rows(rows_count - 1).Item("itm_sku_no") = sadj_dt.Rows(i).Item("itm_sku_no")
                dt.Rows(rows_count - 1).Item("itm_name") = sadj_dt.Rows(i).Item("itm_name")
                dt.Rows(rows_count - 1).Item("pack_key") = sadj_dt.Rows(i).Item("pack_key")
                dt.Rows(rows_count - 1).Item("itm_uom") = sadj_dt.Rows(i).Item("itm_uom")
                dt.Rows(rows_count - 1).Item("itm_uom2") = sadj_dt.Rows(i).Item("ILBS_UOM2")
                'dt.Rows(rows_count - 1).Item("add_org_qty") = sadj_dt.Rows(i).Item("itm_balance")

                dt.Rows(rows_count - 1).Item("WOD_pallet_no") = sadj_dt.Rows(i).Item("ILOC_PALLET_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("wod_batch_no") = sadj_dt.Rows(i).Item("ILOC_BATCH_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("wod_loc") = sadj_dt.Rows(i).Item("ILOC_LOC").ToString.Trim
                dt.Rows(rows_count - 1).Item("WOD_SERIAL_NO") = sadj_dt.Rows(i).Item("ILBS_SERIAL_NO").ToString.Trim
                dt.Rows(rows_count - 1).Item("WOD_DRUM_ID") = sadj_dt.Rows(i).Item("ILBS_DRUM_ID").ToString.Trim
                dt.Rows(rows_count - 1).Item("WOD_DRUM_LV") = sadj_dt.Rows(i).Item("ILBS_DRUM_LEVEL").ToString.Trim
                dt.Rows(rows_count - 1).Item("WOD_qty2") = sadj_dt.Rows(i).Item("ILBS_QTY2")

                REM **********************
                dt.Rows(rows_count - 1).Item("mFlag") = "N"

            Next
            dt.AcceptChanges()
            ViewState("dt") = dt
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If
    End Sub

    Protected Sub btnGenSR_Click(sender As Object, e As System.EventArgs) Handles btnGenSR.Click
        generateStockRelocation()
    End Sub

    Private Sub generateStockRelocation()
        save("Y")

        Dim insertSQL As String = ""
        Dim successFlag As Boolean = False

        Dim pk_code As String = ViewState("WO_CODE")
        Dim newWOCode As String = ""
        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            newWOCode = DB.getDocNo("STF", gConn, transaction)

            insertSQL = " INSERT INTO WMS_STOCK_TRANSFER " & _
                        " (IMP_CODE, STORER_CODE, TR_CODE, TR_STATUS, TR_DATE, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, TR_BY, TR_WH_FR, TR_WH_TO, TR_TO_LOC) " & _
                        " SELECT IMP_CODE, STORER_CODE, '" & newWOCode & "', 'NEW', convert(date, getdate()), '" & Session("usr_id") & "', getdate(), getdate(), '" & Session("usr_id") & "', '" & Session("usr_id") & "', 'ALC', 'NPT', 'EC000000000000' " & _
                        " FROM WMS_WORK_ORDER AS WMS_WORK_ORDER_1 " & _
                        " where WMS_WORK_ORDER_1.WO_code = '" & gU.dbEncode(pk_code) & "' " & _
                        " and WMS_WORK_ORDER_1.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        " and WMS_WORK_ORDER_1.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

            gDB.amendData(insertSQL, gConn, transaction)


            insertSQL = " INSERT INTO WMS_STOCK_TRANSFER_D " & _
                        " (IMP_CODE, STORER_CODE, TR_CODE, TRD_SEQ, ITM_CODE, PACK_KEY, TRD_BATCH_NO_FR, TRD_BATCH_NO_TO, TRD_QTY, TRD_PALLET_NO_FR, TRD_PALLET_NO_TO, TRD_LOC_FR, TRD_LOC_TO, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, " & _
                        " TRD_UOM, TRD_UOM2, TRD_ITM_NAME, TRD_SERIAL, TRD_DRUM_ID_FR, TRD_DRUM_ID_TO, TRD_DRUM_LV_FR, TRD_DRUM_LV_TO, TRD_QTY2) " & _
                        " SELECT IMP_CODE, STORER_CODE, '" & newWOCode & "', WOD_SEQ, ITM_CODE, PACK_KEY, WOD_BATCH_NO, WOD_BATCH_NO, WOD_QTY, '000', '000', WOD_LOC, 'EC000000000000', '" & Session("usr_id") & "', getdate(), getdate(), '" & Session("usr_id") & "',  " & _
                        " NULL, NULL, NULL, WOD_SERIAL_NO, WOD_DRUM_ID, WOD_DRUM_ID, WOD_DRUM_LV, WOD_DRUM_LV, WOD_QTY2 " & _
                        " FROM WMS_WORK_ORDER_D AS WMS_WORK_ORDER_D_1 " & _
                        " where WMS_WORK_ORDER_D_1.WO_code = '" & gU.dbEncode(pk_code) & "' " & _
                        " and WMS_WORK_ORDER_D_1.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        " and WMS_WORK_ORDER_D_1.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' " & _
                        " and WMS_WORK_ORDER_D_1.WOD_TYPE = 'OUT' "

            gDB.amendData(insertSQL, gConn, transaction)

            insertSQL = " UPDATE WMS_WORK_ORDER SET WO_TR_CODE = '" & newWOCode & "' " & _
                        " where WMS_WORK_ORDER.WO_code = '" & gU.dbEncode(pk_code) & "' " & _
                        " and WMS_WORK_ORDER.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        " and WMS_WORK_ORDER.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

            gDB.amendData(insertSQL, gConn, transaction)

            transaction.Commit()
            successFlag = True


        Catch ex As Exception
            transaction.Rollback()
            Response.Write(ex.Message)
            uiFun.displayMsgNew(udp1, "1008", "", Session("gLang"))
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

        If successFlag Then
            Dim rmtPost As New RemotePost
            rmtPost.Url = "../STF/STFMain.aspx"
            rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
            rmtPost.Add("TR_CODE", newWOCode)
            Session("PAGE_SESSION_MENU_CODE") = "OP_SRL"
            rmtPost.alertMsg = "New Stock Relocation Record has been generated!"
            rmtPost.Post()
        End If

    End Sub
    Protected Sub btnCopy_Click(sender As Object, e As System.EventArgs) Handles btnCopy.Click
        CloneWO()
    End Sub

    Private Sub CloneWO()
        save("Y")

        Dim insertSQL As String = ""
        Dim successFlag As Boolean = False

        Dim pk_code As String = ViewState("WO_CODE")
        Dim newWOCode As String = ""
        Dim gConn As SqlConnection

        gConn = gDB.getConnection()
        Dim transaction As SqlTransaction

        transaction = gConn.BeginTransaction()

        Try
            newWOCode = DB.getDocNo("WO", gConn, transaction)

            insertSQL = " INSERT INTO WMS_WORK_ORDER " & _
                        " (IMP_CODE, STORER_CODE, WO_CODE, WO_STATUS, WO_DATE, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, WO_TYPE, WO_REM, WO_TR_CODE, WO_PARENT_CODE) " & _
                        " SELECT IMP_CODE, STORER_CODE, '" & newWOCode & "', WO_STATUS, convert(date, getdate()), SYS_LUB, getdate(), getdate(), SYS_LUB, 'RFID', WO_REM, WO_TR_CODE, '" & gU.dbEncode(pk_code) & "' " & _
                        " FROM WMS_WORK_ORDER AS WMS_WORK_ORDER_1 " & _
                        " where WMS_WORK_ORDER_1.WO_code = '" & gU.dbEncode(pk_code) & "' " & _
                        " and WMS_WORK_ORDER_1.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        " and WMS_WORK_ORDER_1.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

            gDB.amendData(insertSQL, gConn, transaction)


            insertSQL = " INSERT INTO WMS_WORK_ORDER_D " & _
                        " (IMP_CODE, STORER_CODE, WO_CODE, WOD_SEQ, ITM_CODE, PACK_KEY, WOD_BATCH_NO, WOD_TYPE, WOD_QTY, WOD_PALLET_NO, WOD_LOC, WOD_REM, SYS_LUB, SYS_LUD, SYS_CD, SYS_CB, " & _
                        " WOD_KIT_QTY, WOD_SERIAL_NO, WOD_DRUM_ID, WOD_DRUM_LV, WOD_QTY2) " & _
                        " SELECT IMP_CODE, STORER_CODE, '" & newWOCode & "', WOD_SEQ, ITM_CODE, PACK_KEY, WOD_BATCH_NO, WOD_TYPE, WOD_QTY, WOD_PALLET_NO, WOD_LOC, WOD_REM, SYS_LUB, getdate(), getdate(), SYS_LUB,  " & _
                        " WOD_QTY, WOD_SERIAL_NO, WOD_DRUM_ID, WOD_DRUM_LV, WOD_QTY2 " & _
                        " FROM WMS_WORK_ORDER_D AS WMS_WORK_ORDER_D_1 " & _
                        " where WMS_WORK_ORDER_D_1.WO_code = '" & gU.dbEncode(pk_code) & "' " & _
                        " and WMS_WORK_ORDER_D_1.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        " and WMS_WORK_ORDER_D_1.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

            gDB.amendData(insertSQL, gConn, transaction)

            insertSQL = " INSERT INTO WMS_WORK_ORDER_RES " & _
                        " (IMP_CODE, STORER_CODE, WO_CODE, WO_RES_SEQ, WO_RES_TYPE, WO_RES_DESC, WO_RES_AMT, WO_RES_REM, WO_RES_LIST, SYS_CD, SYS_CB, SYS_LUD, SYS_LUB) " & _
                        " SELECT IMP_CODE, STORER_CODE, '" & newWOCode & "', WO_RES_SEQ, WO_RES_TYPE, WO_RES_DESC, WO_RES_AMT, WO_RES_REM, WO_RES_LIST, getdate(), SYS_LUB, getdate(), SYS_LUB " & _
                        " FROM WMS_WORK_ORDER_RES AS WMS_WORK_ORDER_RES_1 " & _
                        " where WMS_WORK_ORDER_RES_1.WO_code = '" & gU.dbEncode(pk_code) & "' " & _
                        " and WMS_WORK_ORDER_RES_1.imp_CODE = '" & Session("IMP_CODE") & "' " & _
                        " and WMS_WORK_ORDER_RES_1.storer_code = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' "

            gDB.amendData(insertSQL, gConn, transaction)

            transaction.Commit()
            successFlag = True


        Catch ex As Exception
            transaction.Rollback()
            Response.Write(ex.Message)
            uiFun.displayMsgNew(udp1, "1008", "", Session("gLang"))
        Finally
            If gConn IsNot Nothing Then
                If gConn.State = ConnectionState.Open Then
                    gConn.Close()
                    gConn.Dispose()
                End If
            End If
        End Try

        If successFlag Then
            Dim rmtPost As New RemotePost
            rmtPost.Url = "WOMain.aspx"
            rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
            rmtPost.Add("WO_CODE", newWOCode)
            rmtPost.alertMsg = "New WO has been copied!"
            rmtPost.Post()
        End If

    End Sub

    Protected Sub gotoSTKREC_Click(sender As Object, e As System.EventArgs) Handles gotoSTKREC.Click
        Dim rmtPost As New RemotePost
        rmtPost.Url = "..\STF\STFMain.aspx"
        rmtPost.Add("STORER_CODE", STORER_CODE.SelectedValue)
        rmtPost.Add("TR_CODE", WO_TR_CODE.Text)
        rmtPost.Add("frSTF", "Y")
        Session("PAGE_SESSION_MENU_CODE") = "OP_SRL"
        rmtPost.Post()
    End Sub

    Protected Sub btnChkBal_Click(sender As Object, e As System.EventArgs) Handles btnChkBal.Click
        CheckStockBalance()
    End Sub

    Private Function CheckStockBalance() As Boolean
        Dim flag As Boolean = True

        Dim itm_code, pack_key, batch_no, serial_no, pallet_no, loc_code, SERIAL_YN, skuNo As String
        Dim stock_bal As Double = 0
        Dim inputBal As Double = 0
        Dim selectSQL As String = ""

        skuNo = ""
        If GridView1 IsNot Nothing AndAlso GridView1.Rows.Count > 0 Then
            For i = 0 To GridView1.Rows.Count - 1
                If DirectCast(GridView1.Rows(i).FindControl("WOD_TYPE"), DropDownList).SelectedValue = "OUT" Then
                    itm_code = DirectCast(GridView1.Rows(i).FindControl("itm_code"), Label).Text
                    pack_key = DirectCast(GridView1.Rows(i).FindControl("pack_key"), Label).Text
                    batch_no = DirectCast(GridView1.Rows(i).FindControl("wod_batch_no"), TextBox).Text.Trim
                    serial_no = DirectCast(GridView1.Rows(i).FindControl("WOD_SERIAL_NO"), TextBox).Text.Trim
                    pallet_no = DirectCast(GridView1.Rows(i).FindControl("WOD_pallet_no"), HiddenField).Value
                    loc_code = DirectCast(GridView1.Rows(i).FindControl("wod_loc"), DropDownList).SelectedValue

                    SERIAL_YN = DB.getValueFromSQL("Select ITM_SERIAL_NO_YN from wms_item where imp_code='" & Session("imp_code") & "' and storer_code='" & gU.dbEncode(STORER_CODE.SelectedValue) & "' and itm_code='" & gU.dbEncode(itm_code) & "' and pack_key='" & gU.dbEncode(pack_key) & "'")

                    If SERIAL_YN = "Y" Then
                        If DirectCast(GridView1.Rows(i).FindControl("WOD_qty2"), TextBox).Text.Trim <> "" AndAlso gU.isDecimal(DirectCast(GridView1.Rows(i).FindControl("WOD_qty2"), TextBox).Text.Trim) Then

                            selectSQL = " SELECT ILBS_QTY2 " & _
                                        " FROM WMS_ITEM_LOC_BAL INNER JOIN " & _
                                        " WMS_ITEM_LOC_BAL_S ON WMS_ITEM_LOC_BAL.ILOC_SEQ = WMS_ITEM_LOC_BAL_S.ILOC_SEQ " & _
                                        " WHERE WMS_ITEM_LOC_BAL.IMP_CODE = '" & Session("imp_code") & "' AND WMS_ITEM_LOC_BAL.STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND WMS_ITEM_LOC_BAL.ITM_CODE = '" & gU.dbEncode(itm_code) & "' AND WMS_ITEM_LOC_BAL.PACK_KEY = '" & gU.dbEncode(pack_key) & "' " & _
                                        " AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_PALLET_NO,'000') = ISNULL('" & gU.dbEncode(pallet_no) & "','0000') AND WMS_ITEM_LOC_BAL.ILOC_LOC = '" & gU.dbEncode(loc_code) & "' AND ISNULL(WMS_ITEM_LOC_BAL.ILOC_BATCH_NO,'') = isnull('" & gU.dbEncode(batch_no) & "','') " & _
                                        " AND ISNULL(WMS_ITEM_LOC_BAL_S.ILBS_SERIAL_NO,0) = ISNULL('" & gU.dbEncode(serial_no) & "','')"
                            stock_bal = gU.decodeEmptyCdbl(DB.getValueFromSQL(selectSQL), 0)

                            inputBal = DirectCast(GridView1.Rows(i).FindControl("WOD_qty2"), TextBox).Text.Trim

                            If inputBal > stock_bal Then
                                skuNo = gU.appendToList(skuNo, DirectCast(GridView1.Rows(i).FindControl("itm_sku_no"), HiddenField).Value)
                                DirectCast(GridView1.Rows(i).FindControl("WOD_qty2"), TextBox).BackColor = Drawing.Color.LightPink
                                flag = False
                            Else
                                DirectCast(GridView1.Rows(i).FindControl("WOD_qty2"), TextBox).BackColor = Drawing.Color.White
                            End If

                        Else
                            uiFun.displayMsgNew(chkBalUDP, "", "Please enter the WO Qty2.", Session("gLang"))

                            DirectCast(GridView1.Rows(i).FindControl("WOD_qty2"), TextBox).BackColor = Drawing.Color.LightPink
                            Return False
                        End If


                    Else
                        If DirectCast(GridView1.Rows(i).FindControl("WOD_qty"), TextBox).Text.Trim <> "" AndAlso gU.isDecimal(DirectCast(GridView1.Rows(i).FindControl("WOD_qty"), TextBox).Text.Trim) Then

                            selectSQL = " SELECT ILOC_BAL_QTY " & _
                                        " FROM WMS_ITEM_LOC_BAL " & _
                                        " WHERE IMP_CODE = '" & Session("imp_code") & "' AND STORER_CODE = '" & gU.dbEncode(STORER_CODE.SelectedValue) & "' AND ITM_CODE = '" & gU.dbEncode(itm_code) & "' AND PACK_KEY = '" & gU.dbEncode(pack_key) & "' " & _
                                        " AND ISNULL(ILOC_PALLET_NO,'000') = ISNULL('" & gU.dbEncode(pallet_no) & "','0000') AND ILOC_LOC = '" & gU.dbEncode(loc_code) & "' AND ISNULL(ILOC_BATCH_NO,'') = isnull('" & gU.dbEncode(batch_no) & "','') "
                            stock_bal = gU.decodeEmptyCdbl(DB.getValueFromSQL(selectSQL), 0)

                            inputBal = DirectCast(GridView1.Rows(i).FindControl("WOD_qty"), TextBox).Text.Trim

                            If inputBal > stock_bal Then
                                skuNo = gU.appendToList(skuNo, DirectCast(GridView1.Rows(i).FindControl("itm_sku_no"), HiddenField).Value)
                                DirectCast(GridView1.Rows(i).FindControl("WOD_qty"), TextBox).BackColor = Drawing.Color.LightPink
                                flag = False
                            Else
                                DirectCast(GridView1.Rows(i).FindControl("WOD_qty"), TextBox).BackColor = Drawing.Color.White
                            End If

                        Else
                            uiFun.displayMsgNew(chkBalUDP, "", "Please enter the WO Qty.", Session("gLang"))

                            DirectCast(GridView1.Rows(i).FindControl("WOD_qty"), TextBox).BackColor = Drawing.ColorTranslator.FromHtml("#d4d0c8")
                            Return False
                        End If
                    End If
                End If
            Next
        End If

        If Not flag Then
            uiFun.displayMsgNew(chkBalUDP, "", "Stock No.: " & skuNo & " Qty larger than stock Balance. " & vbCrLf & "Please check your input Qty.", Session("gLang"))
            Return flag
        End If

        uiFun.displayMsgNew(chkBalUDP, "", "Balance Checking finished successfully.", Session("gLang"))
        Return flag
    End Function
End Class
